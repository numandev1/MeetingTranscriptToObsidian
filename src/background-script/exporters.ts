import type { ErrorObject, Meeting, ResultLocal, ResultSync, WebhookBody } from "@lib/types"
import { getTranscriptString, getChatMessagesString } from "./utils"
import { getFolderPrefix, DEFAULT_DOWNLOAD_SUBFOLDER } from "@lib/paths"
import { buildMeetingHtml } from "@lib/html-export"
import { TIMEFORMAT } from "./config"

/**
 * Filenames this extension has asked for, in request order, each awaiting its
 * onDeterminingFilename event.
 */
const pendingDownloadFilenames: string[] = []

/**
 * Starts a download and records the intended filename so the guard below can re-assert it.
 *
 * Passing `filename` to downloads.download() is only a *suggestion*: any extension
 * holding the downloads permission can rewrite it from onDeterminingFilename. A download
 * manager doing exactly that is why transcripts were landing in the Downloads root as
 * "download (N).txt" instead of the configured folder.
 */
function startDownload(filename: string, url: string): Promise<number> {
    pendingDownloadFilenames.push(filename)
    return chrome.downloads.download({ url, filename, conflictAction: "uniquify" })
}

/**
 * Re-asserts our own filename for downloads this extension started.
 *
 * Chrome awards a contested filename to the most recently installed extension that calls
 * suggest(), so registering here is what lets the configured folder survive another
 * extension's rewrite. Downloads started by anything else are left alone.
 */
export function registerDownloadFilenameGuard(): void {
    if (!chrome.downloads.onDeterminingFilename) {
        return
    }
    chrome.downloads.onDeterminingFilename.addListener((item, suggest) => {
        if (item.byExtensionId !== chrome.runtime.id) {
            return
        }
        const intendedFilename = pendingDownloadFilenames.shift()
        if (!intendedFilename) {
            return
        }
        suggest({ filename: intendedFilename, conflictAction: "uniquify" })
    })
}

/**
 * Downloads a stored meeting (by index) or a meeting passed in directly — the one still
 * in progress. A live meeting's file says "so far", so it is never mistaken for the
 * complete transcript that is saved when the meeting ends.
 *
 * @throws error codes: 009, 010
 */
export function downloadTranscript(index: number | Meeting, isWebhookEnabled?: boolean, format: "txt" | "html" = "txt"): Promise<string> {
    return new Promise((resolve, reject) => {
        chrome.storage.local.get(["meetings"], function (resultLocalUntyped) {
            const resultLocal = resultLocalUntyped as Partial<ResultLocal>

            chrome.storage.sync.get(["downloadSubfolder"], function (resultSyncUntyped) {
                const resultSync = resultSyncUntyped as Partial<ResultSync>
                // "" when the user cleared the field, meaning save straight into Downloads
                const folderPrefix = getFolderPrefix(resultSync.downloadSubfolder, DEFAULT_DOWNLOAD_SUBFOLDER)

                const isLive = typeof index !== "number"
                const meeting = isLive ? index : resultLocal.meetings?.[index]
                if (meeting) {

                    // Sanitise meeting title to prevent invalid file name errors
                    // https://stackoverflow.com/a/78675894
                    const invalidFilenameRegex = /[:?"*<>|~/\\\u{1}-\u{1f}\u{7f}\u{80}-\u{9f}\p{Cf}\p{Cn}]|^[.\u{0}\p{Zl}\p{Zp}\p{Zs}]|[.\u{0}\p{Zl}\p{Zp}\p{Zs}]$|^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?=\.|$)/gui
                    let sanitisedMeetingTitle = "Meeting"
                    if (meeting.meetingTitle) {
                        sanitisedMeetingTitle = meeting.meetingTitle.replaceAll(invalidFilenameRegex, "_")
                    }
                    else if (meeting.title) {
                        sanitisedMeetingTitle = meeting.title.replaceAll(invalidFilenameRegex, "_")
                    }

                    // Format timestamp for human-readable filename and sanitise to prevent invalid filenames
                    const timestamp = new Date(meeting.meetingStartTimestamp)
                    const formattedTimestamp = timestamp.toLocaleString("default", TIMEFORMAT).replace(/[\/:]/g, "-")

                    const kind = isLive ? "transcript so far" : "transcript"
                    const prefix = meeting.meetingSoftware ? `${meeting.meetingSoftware} ${kind}` : (isLive ? "Transcript so far" : "Transcript")

                    const fileName = `${folderPrefix}${prefix}-${sanitisedMeetingTitle} at ${formattedTimestamp} on.${format}`


                    // Format transcript and chatMessages content
                    let content = getTranscriptString(meeting.transcript)
                    content += `\n\n---------------\nCHAT MESSAGES\n---------------\n\n`
                    content += getChatMessagesString(meeting.chatMessages)

                    const blob = new Blob([format === "html" ? buildMeetingHtml(meeting, "transcript", isLive) : content], { type: format === "html" ? "text/html;charset=utf-8" : "text/plain" })

                    // Read the blob as a data URL
                    const reader = new FileReader()

                    // Read the blob
                    reader.onerror = () => reject({ errorCode: "009", errorMessage: "Failed to read blob" } satisfies ErrorObject)
                    reader.readAsDataURL(blob)

                    // Download as text file, once blob is read
                    reader.onload = function (event) {
                        if (event.target?.result) {
                            const dataUrl = event.target.result as string

                            // Create a download with Chrome Download API
                            startDownload(fileName, dataUrl).then(() => {
                                console.log("Transcript downloaded")
                                resolve("Transcript downloaded successfully")
                            }).catch((err) => {
                                console.error(err)
                                startDownload(`${folderPrefix}Transcript.${format}`, dataUrl).then(() => {
                                    console.log("Invalid file name. Transcript downloaded to the configured directory with simple file name.")
                                    resolve("Transcript downloaded successfully with default file name")
                                }).catch(reject)

                            })
                        }
                        else {
                            const error: ErrorObject = { errorCode: "009", errorMessage: "Failed to read blob" }
                            reject(error)
                        }
                    }
                }
                else {
                    const error: ErrorObject = { errorCode: "010", errorMessage: "Meeting at specified index not found" }
                    reject(error)
                }
            })
        })
    })
}

/** @throws error codes: 010, 011, 012 */
export function postTranscriptToWebhook(index: number): Promise<string> {
    return new Promise((resolve, reject) => {
        // Get webhook URL and meetings
        chrome.storage.local.get(["meetings"], function (resultLocalUntyped) {
            const resultLocal = resultLocalUntyped as Partial<ResultLocal>
            chrome.storage.sync.get(["webhookUrl", "webhookBodyType"], function (resultSyncUntyped) {
                const resultSync = resultSyncUntyped as Partial<ResultSync>

                if (resultSync.webhookUrl) {
                    const meetings = resultLocal.meetings
                    if (meetings && meetings[index]) {
                        const meeting = meetings[index]

                        let webhookData: WebhookBody
                        if (resultSync.webhookBodyType === "advanced") {
                            webhookData = {
                                webhookBodyType: "advanced",
                                meetingSoftware: meeting.meetingSoftware ? meeting.meetingSoftware : "",
                                meetingTitle: meeting.meetingTitle || meeting.title || "",
                                meetingStartTimestamp: new Date(meeting.meetingStartTimestamp).toISOString(),
                                meetingEndTimestamp: new Date(meeting.meetingEndTimestamp).toISOString(),
                                transcript: meeting.transcript,
                                chatMessages: meeting.chatMessages
                            }
                        }
                        else {
                            webhookData = {
                                webhookBodyType: "simple",
                                meetingSoftware: meeting.meetingSoftware ? meeting.meetingSoftware : "",
                                meetingTitle: meeting.meetingTitle || meeting.title || "",
                                meetingStartTimestamp: new Date(meeting.meetingStartTimestamp).toLocaleString("default", TIMEFORMAT).toUpperCase(),
                                meetingEndTimestamp: new Date(meeting.meetingEndTimestamp).toLocaleString("default", TIMEFORMAT).toUpperCase(),
                                transcript: getTranscriptString(meeting.transcript),
                                chatMessages: getChatMessagesString(meeting.chatMessages)
                            }
                        }

                        // Post to webhook
                        fetch(resultSync.webhookUrl, {
                            method: "POST",
                            headers: {
                                "Content-Type": "application/json"
                            },
                            body: JSON.stringify(webhookData)
                        }).then(response => {
                            if (!response.ok) {
                                throw new Error(`Webhook request failed with HTTP status code ${response.status} ${response.statusText}`)
                            }
                        }).then(() => {
                            // Update success status.
                            meetings[index].webhookPostStatus = "successful"
                            chrome.storage.local.set({ meetings: meetings }, function () {
                                resolve("Webhook posted successfully")
                            })
                        }).catch(error => {
                            console.error(error)
                            // Update failure status.
                            meetings[index].webhookPostStatus = "failed"
                            chrome.storage.local.set({ meetings: meetings }, function () {
                                // Create notification and open webhooks page
                                chrome.notifications.create({
                                    type: "basic",
                                    iconUrl: "icon.png",
                                    title: "Could not post webhook!",
                                    message: "Click to view status and retry. Check console for more details."
                                }, function (notificationId) {
                                    // Handle notification click
                                    chrome.notifications.onClicked.addListener(function (clickedNotificationId) {
                                        if (clickedNotificationId === notificationId) {
                                            chrome.tabs.create({ url: "meetings.html" })
                                        }
                                    })
                                })
                                const parsedError: ErrorObject = { errorCode: "011", errorMessage: error instanceof Error ? error.message : String(error) }
                                reject(parsedError)
                            })
                        })
                    }
                    else {
                        const error: ErrorObject = { errorCode: "010", errorMessage: "Meeting at specified index not found" }
                        reject(error)
                    }
                }
                else {
                    const error: ErrorObject = { errorCode: "012", errorMessage: "No webhook URL configured" }
                    reject(error)
                }
            })
        })
    })
}

/**
 * Saves the summary as its own text file, alongside the transcript file and in the same
 * configured folder.
 *
 * @throws error codes: 009, 010, 022
 */
export function downloadSummary(index: number | Meeting, format: "txt" | "html" = "txt"): Promise<string> {
    return new Promise((resolve, reject) => {
        chrome.storage.local.get(["meetings"], function (resultLocalUntyped) {
            const resultLocal = resultLocalUntyped as Partial<ResultLocal>

            chrome.storage.sync.get(["downloadSubfolder"], function (resultSyncUntyped) {
                const resultSync = resultSyncUntyped as Partial<ResultSync>
                const folderPrefix = getFolderPrefix(resultSync.downloadSubfolder, DEFAULT_DOWNLOAD_SUBFOLDER)

                const isLive = typeof index !== "number"
                const meeting = isLive ? index : resultLocal.meetings?.[index]
                if (!meeting) {
                    reject({ errorCode: "010", errorMessage: "Meeting at specified index not found" } satisfies ErrorObject)
                    return
                }
                if (!meeting.summary) {
                    reject({ errorCode: "022", errorMessage: "Nothing to summarise for this meeting" } satisfies ErrorObject)
                    return
                }

                // Same sanitising as the transcript file
                // https://stackoverflow.com/a/78675894
                const invalidFilenameRegex = /[:?"*<>|~/\\\u{1}-\u{1f}\u{7f}\u{80}-\u{9f}\p{Cf}\p{Cn}]|^[.\u{0}\p{Zl}\p{Zp}\p{Zs}]|[.\u{0}\p{Zl}\p{Zp}\p{Zs}]$|^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?=\.|$)/gui
                let sanitisedMeetingTitle = "Meeting"
                if (meeting.meetingTitle) {
                    sanitisedMeetingTitle = meeting.meetingTitle.replaceAll(invalidFilenameRegex, "_")
                }
                else if (meeting.title) {
                    sanitisedMeetingTitle = meeting.title.replaceAll(invalidFilenameRegex, "_")
                }

                const formattedTimestamp = new Date(meeting.meetingStartTimestamp)
                    .toLocaleString("default", TIMEFORMAT).replace(/[\/:]/g, "-")

                const kind = isLive ? "summary so far" : "summary"
                const prefix = meeting.meetingSoftware ? `${meeting.meetingSoftware} ${kind}` : (isLive ? "Summary so far" : "Summary")
                const fileName = `${folderPrefix}${prefix}-${sanitisedMeetingTitle} at ${formattedTimestamp}.${format}`

                let content = `${meeting.meetingTitle || meeting.title || "Meeting"}\n`
                content += `${new Date(meeting.meetingStartTimestamp).toLocaleString("default", TIMEFORMAT).toUpperCase()}\n\n`
                content += `${meeting.summary.trim()}\n`

                const reader = new FileReader()
                reader.onerror = () => reject({ errorCode: "009", errorMessage: "Failed to read blob" } satisfies ErrorObject)
                reader.readAsDataURL(new Blob([format === "html" ? buildMeetingHtml(meeting, "summary", isLive) : content], { type: format === "html" ? "text/html;charset=utf-8" : "text/plain" }))
                reader.onload = function (event) {
                    if (!event.target?.result) {
                        reject({ errorCode: "009", errorMessage: "Failed to read blob" } satisfies ErrorObject)
                        return
                    }
                    const dataUrl = event.target.result as string

                    startDownload(fileName, dataUrl).then(() => {
                        console.log("Summary downloaded")
                        resolve("Summary downloaded successfully")
                    }).catch((err) => {
                        console.error(err)
                        // Fall back to a plain name, exactly as the transcript download does
                        startDownload(`${folderPrefix}Summary.${format}`, dataUrl)
                            .then(() => resolve("Summary downloaded successfully with default file name"))
                            .catch(reject)
                    })
                }
            })
        })
    })
}
