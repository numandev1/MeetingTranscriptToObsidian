import { createCaptureStore } from "./capture-sessions"
import type { ErrorObject, ExtensionMessage, ExtensionResponse, Platform, ResultLocal, ResultSync } from "@lib/types"
import { processLastMeeting, readLiveMeeting, recoverLastMeeting, repairInvalidMerges } from "./meetings"
import { normaliseMeetingUrl } from "@lib/meeting-merge"
import { downloadTranscript, downloadSummary, postTranscriptToWebhook, registerDownloadFilenameGuard } from "./exporters"
import { postTranscriptToObsidian, postSummaryToObsidian, syncAllMeetingsToObsidian } from "./obsidian"
import { summarizeAndFileMeeting, summarizeLiveMeeting } from "./summarizer"
import {
    getPermissionStatus,
    requestPlatformPermission,
    getContentScriptStatus,
    registerContentScript,
    deregisterContentScript,
    reRegisterContentScripts,
    registerZoomRedirect,
    deregisterZoomRedirect,
} from "./platforms"
import {
    openExtensionPopup,
    openSidePanel,
    captureAccountEmail
} from "./utils"
import { DEFAULT_DOWNLOAD_SUBFOLDER, DEFAULT_MEETINGS_TO_KEEP, DEFAULT_OBSIDIAN_FOLDER, DEFAULT_OBSIDIAN_SUMMARY_FOLDER } from "@lib/paths"
import { DEFAULT_OBSIDIAN_BASE_URL } from "./obsidian"

// Registered at the top level so it is in place on every service worker wake
registerDownloadFilenameGuard()

const captureStore = createCaptureStore(
    () => chrome.storage.local.get(["captureSessions", "meetingTabId"]),
    patch => chrome.storage.local.set(patch),
)

let processingCaptures = 0
async function finishCapture(tabId: number, captureId?: string, recover = false): Promise<void> {
    processingCaptures++
    try {
        const snapshot = await captureStore.finish(tabId, captureId)
            ?? (recover ? await captureStore.pending(tabId) : undefined)
        if (snapshot) {
            try {
                await processLastMeeting(snapshot)
                await captureStore.complete(tabId, snapshot.captureId)
            } catch (error) {
                // An empty lobby has nothing to recover; export failures retain the snapshot.
                if ((error as ErrorObject).errorCode === "014") await captureStore.complete(tabId, snapshot.captureId)
                throw error
            }
        }
    } finally {
        processingCaptures--
        const stored = await chrome.storage.local.get(["meetingTabId", "isDeferredUpdatedAvailable"])
        if (!processingCaptures && !stored.meetingTabId && stored.isDeferredUpdatedAvailable) chrome.runtime.reload()
    }
}

chrome.runtime.onMessage.addListener(function (messageUnTyped, sender, sendResponse) {
    const message = messageUnTyped as ExtensionMessage
    console.log(message.type)

    // Chat now lives in a cross-origin chat.google.com iframe inside the Meet tab. Its
    // content script cannot reach the Meet frame's script directly, so hand the batch to
    // the tab's top frame, where the meeting state lives.
    if (message.type === "meet_embedded_chat_messages") {
        const tabId = sender.tab?.id
        if (typeof tabId !== "number") {
            sendResponse({ success: false, message: "No sender tab" } satisfies ExtensionResponse)
            return false
        }
        chrome.tabs.sendMessage(tabId, message, { frameId: 0 })
            .then(() => sendResponse({ success: true } satisfies ExtensionResponse))
            .catch((err) => sendResponse({ success: false, message: String(err) } satisfies ExtensionResponse))
        return true
    }

    if (message.type === "capture_update" || message.type === "meeting_ended" || message.type === "new_meeting_started") {
        const tabId = sender.tab?.id ?? (message.type === "capture_update" ? message.captureTabId : undefined)
        if (typeof tabId !== "number" || !message.captureId) {
            sendResponse({ success: false, message: "Missing meeting identity" } satisfies ExtensionResponse)
            return false
        }
        const captureId = message.captureId
        const patch = message.type === "new_meeting_started"
            ? { meetingUrl: normaliseMeetingUrl(sender.tab?.url) }
            : message.capturePatch ?? {}
        captureStore.update(tabId, captureId, patch, !!message.capturePatch?.meetingSoftware)
            .then(async () => {
                if (message.type === "new_meeting_started") {
                    captureAccountEmail(tabId, accountEmail => {
                        void captureStore.update(tabId, captureId, { accountEmail }, false)
                    })
                }
                if (message.type === "meeting_ended") await finishCapture(tabId, captureId)
                sendResponse({ success: true } satisfies ExtensionResponse)
            })
            .catch(error => sendResponse({ success: false, message: error } satisfies ExtensionResponse))
        return true
    }

    if (message.type === "download_transcript_at_index") {
        if ((typeof message.index === "number") && (message.index >= 0)) {
            // Download the requested item
            downloadTranscript(message.index, false, message.format)
                .then(() => {
                    const response: ExtensionResponse = { success: true }
                    sendResponse(response)
                })
                .catch((error) => {
                    // Fails with error codes: 009, 010
                    const parsedError = error as ErrorObject

                    const response: ExtensionResponse = { success: false, message: parsedError }
                    sendResponse(response)
                })
        }
        else {
            const response: ExtensionResponse = { success: false, message: { errorCode: "015", errorMessage: "Invalid index" } }
            sendResponse(response)
        }
    }

    if (message.type === "post_webhook_at_index") {
        if ((typeof message.index === "number") && (message.index >= 0)) {
            // Handle webhook retry
            postTranscriptToWebhook(message.index)
                .then(() => {
                    const response: ExtensionResponse = { success: true }
                    sendResponse(response)
                })
                .catch(error => {
                    // Fails with error codes: 009, 010, 011, 012
                    const parsedError = error as ErrorObject

                    console.error("Webhook retry failed:", parsedError)
                    const response: ExtensionResponse = { success: false, message: parsedError }
                    sendResponse(response)
                })
        }
        else {
            const response: ExtensionResponse = { success: false, message: { errorCode: "015", errorMessage: "Invalid index" } }
            sendResponse(response)
        }
    }

    if (message.type === "download_summary_at_index") {
        if ((typeof message.index === "number") && (message.index >= 0)) {
            downloadSummary(message.index, message.format)
                .then(() => sendResponse({ success: true } satisfies ExtensionResponse))
                .catch((error) => {
                    // Fails with error codes: 009, 010, 022
                    const parsedError = error as ErrorObject
                    sendResponse({ success: false, message: parsedError } satisfies ExtensionResponse)
                })
        }
        else {
            const response: ExtensionResponse = { success: false, message: { errorCode: "015", errorMessage: "Invalid index" } }
            sendResponse(response)
        }
    }

    if (message.type === "post_summary_obsidian_at_index") {
        if ((typeof message.index === "number") && (message.index >= 0)) {
            postSummaryToObsidian(message.index)
                .then(() => sendResponse({ success: true } satisfies ExtensionResponse))
                .catch((error) => {
                    // Fails with error codes: 010, 017, 018, 019
                    const parsedError = error as ErrorObject
                    console.error("Obsidian summary save failed:", parsedError)
                    sendResponse({ success: false, message: parsedError } satisfies ExtensionResponse)
                })
        }
        else {
            const response: ExtensionResponse = { success: false, message: { errorCode: "015", errorMessage: "Invalid index" } }
            sendResponse(response)
        }
    }

    if (message.type === "sync_all_to_obsidian") {
        syncAllMeetingsToObsidian()
            .then((result) => {
                sendResponse({ success: true, message: JSON.stringify(result) } satisfies ExtensionResponse)
            })
            .catch((error) => {
                // Fails with error codes: 013, 018
                const parsedError = error as ErrorObject
                sendResponse({ success: false, message: parsedError } satisfies ExtensionResponse)
            })
    }

    if (message.type === "post_obsidian_at_index") {
        if ((typeof message.index === "number") && (message.index >= 0)) {
            postTranscriptToObsidian(message.index)
                .then(() => {
                    const response: ExtensionResponse = { success: true }
                    sendResponse(response)
                })
                .catch(error => {
                    // Fails with error codes: 010, 017, 018, 019
                    const parsedError = error as ErrorObject

                    console.error("Obsidian save failed:", parsedError)
                    const response: ExtensionResponse = { success: false, message: parsedError }
                    sendResponse(response)
                })
        }
        else {
            const response: ExtensionResponse = { success: false, message: { errorCode: "015", errorMessage: "Invalid index" } }
            sendResponse(response)
        }
    }

    if (message.type === "summarize_at_index") {
        if ((typeof message.index === "number") && (message.index >= 0)) {
            const index = message.index
            summarizeAndFileMeeting(index)
                .then((summary) => {
                    sendResponse({ success: true, message: summary } satisfies ExtensionResponse)
                })
                .catch((error) => {
                    // Fails with error codes: 010, 020, 021, 022
                    const parsedError = error as ErrorObject

                    console.error("Summarise failed:", parsedError)
                    sendResponse({ success: false, message: parsedError } satisfies ExtensionResponse)
                })
        }
        else {
            const response: ExtensionResponse = { success: false, message: { errorCode: "015", errorMessage: "Invalid index" } }
            sendResponse(response)
        }
    }

    if (message.type === "recover_last_meeting") {
        // A newly opened tab must never recover (and file) another tab's live call.
        const tabId = sender.tab?.id
        const recover = typeof tabId === "number"
            ? finishCapture(tabId, undefined, true)
            : recoverLastMeeting()
        recover.then(() => sendResponse({ success: true } satisfies ExtensionResponse))
            .catch(error => sendResponse({ success: false, message: error } satisfies ExtensionResponse))
    }

    if (message.type === "get_platform_enablement_status") {
        const platform: Platform | Platform[] = message.platform || "google_meet"

        getContentScriptStatus(platform)
            .then((status) => {
                const response: ExtensionResponse = {
                    success: true,
                    message: status
                }
                sendResponse(response)
            })
            .catch((error) => {
                const parsedError = error as ErrorObject
                const response: ExtensionResponse = { success: false, message: parsedError }
                sendResponse(response)
            })
    }

    if (message.type === "get_platform_permission_status") {
        const platform: Platform | Platform[] = message.platform || "google_meet"

        getPermissionStatus(platform)
            .then((status) => {
                const response: ExtensionResponse = {
                    success: true,
                    message: status
                }
                sendResponse(response)
            })
            .catch((error) => {
                const parsedError = error as ErrorObject
                const response: ExtensionResponse = { success: false, message: parsedError }
                sendResponse(response)
            })
    }

    if ((message.type === "enable_platform")) {
        const platform: Platform | Platform[] = message.platform || "google_meet"

        requestPlatformPermission(platform).then(() => {
            // After permissions are granted, register both the scripts and the redirect rule for zoom
            const promises: Promise<string | string[]>[] = [registerContentScript(platform)]

            const platformsArray = Array.isArray(platform) ? platform : [platform]
            const hasZoom = platformsArray.includes("zoom")

            if (hasZoom) {
                promises.push(registerZoomRedirect())
            }

            Promise.all(promises).then((results) => {
                const response: ExtensionResponse = { success: true, message: results[0] }
                sendResponse(response)
            }).catch((error) => {
                // Fails with error codes: not defined
                const parsedError = error as ErrorObject

                const response: ExtensionResponse = { success: false, message: parsedError }
                sendResponse(response)
            })
        })
            .catch((error) => {
                // Fails with error codes: not defined
                const parsedError = error as ErrorObject

                const response: ExtensionResponse = { success: false, message: parsedError }
                sendResponse(response)
            })
    }

    if (message.type === "disable_platform") {
        const platform: Platform | Platform[] = message.platform || "google_meet"

        // To disable, we simply unregister the content scripts and any redirect rules
        const promises: Promise<string | string[]>[] = [deregisterContentScript(platform)]

        const platformsArray = Array.isArray(platform) ? platform : [platform]
        const hasZoom = platformsArray.includes("zoom")

        if (hasZoom) {
            promises.push(deregisterZoomRedirect())
        }

        Promise.all(promises).then((results) => {
            const response: ExtensionResponse = { success: true, message: results[0] }
            sendResponse(response)
        }).catch((error) => {
            const parsedError = error as ErrorObject

            const response: ExtensionResponse = { success: false, message: parsedError }
            sendResponse(response)
        })
    }

    if (message.type === "open_popup") {
        openExtensionPopup().then((resolvedMessage) => {
            const response: ExtensionResponse = { success: true, message: resolvedMessage }
            sendResponse(response)
        }).catch((error) => {
            const parsedError = error as ErrorObject

            const response: ExtensionResponse = { success: false, message: parsedError }
            sendResponse(response)
        })
    }

    // --- The meeting still in progress -------------------------------------------
    // Summarise or download the transcript so far without ending the meeting. Nothing
    // here files anything: the meeting is filed as usual when it ends.

    if (message.type === "summarize_live_meeting") {
        summarizeLiveMeeting()
            .then(liveSummary => sendResponse({ success: true, message: liveSummary.summary } satisfies ExtensionResponse))
            .catch((error) => {
                // Fails with error codes: 020, 021, 022, 023
                console.error("Live summary failed:", error)
                sendResponse({ success: false, message: error as ErrorObject } satisfies ExtensionResponse)
            })
    }

    if (message.type === "download_live_transcript") {
        readLiveMeeting()
            .then(meeting => downloadTranscript(meeting, false, message.format))
            .then(() => sendResponse({ success: true } satisfies ExtensionResponse))
            .catch((error) => {
                // Fails with error codes: 009, 022, 023
                sendResponse({ success: false, message: error as ErrorObject } satisfies ExtensionResponse)
            })
    }

    if (message.type === "download_live_summary") {
        chrome.storage.local.get(["liveSummary", "meetingTitle", "meetingSoftware", "meetingStartTimestamp"], (resultUntyped) => {
            const result = resultUntyped as Partial<ResultLocal>
            const live = result.liveSummary
            if (!live || live.meetingStartTimestamp !== result.meetingStartTimestamp) {
                sendResponse({ success: false, message: { errorCode: "022", errorMessage: "Nothing to summarise for this meeting" } } satisfies ExtensionResponse)
                return
            }
            // Works after the meeting ends too: the summary carries its own start time
            downloadSummary({
                meetingSoftware: result.meetingSoftware ?? "",
                meetingTitle: result.meetingTitle,
                meetingStartTimestamp: live.meetingStartTimestamp,
                meetingEndTimestamp: live.createdAt,
                transcript: [],
                chatMessages: [],
                webhookPostStatus: "new",
                summary: live.summary,
            }, message.format)
                .then(() => sendResponse({ success: true } satisfies ExtensionResponse))
                .catch(error => sendResponse({ success: false, message: error as ErrorObject } satisfies ExtensionResponse))
        })
    }

    if (message.type === "open_side_panel") {
        const select = typeof sender.tab?.id === "number" ? captureStore.select(sender.tab.id) : Promise.resolve()
        // Open synchronously to preserve Chrome's user-gesture requirement.
        const opened = openSidePanel()
        Promise.all([select, opened]).then(([, resolvedMessage]) => {
            const response: ExtensionResponse = { success: true, message: resolvedMessage }
            sendResponse(response)
        }).catch((error) => {
            const parsedError = error as ErrorObject

            const response: ExtensionResponse = { success: false, message: parsedError }
            sendResponse(response)
        })
    }


    return true
})

// Download last meeting if meeting tab is closed
chrome.tabs.onRemoved.addListener(function (tabId) {
    void finishCapture(tabId).catch(error => console.error("Meeting finalization failed", error))
})

// Listen for extension updates
chrome.runtime.onUpdateAvailable.addListener(() => {
    // Check if there is an active meeting
    chrome.storage.local.get(["meetingTabId"], function (resultUntyped) {
        const result = resultUntyped as Partial<ResultLocal>

        if (result.meetingTabId || processingCaptures > 0) {
            // There is an active meeting(values: tabId or processing), defer the update
            chrome.storage.local.set({ isDeferredUpdatedAvailable: true }, function () {
                console.log("Deferred update flag set")
            })
        } else {
            // No active meeting, apply the update immediately. Meeting tab id is nullified only post meeting operations are done, so no race conditions.
            console.log("No active meeting, applying update immediately")
            chrome.runtime.reload()
        }
    })
})

// Register content scripts and Zoom redirect whenever runtime permission change—mostly serves as a backup for changes made outside the UI.
chrome.permissions.onAdded.addListener(() => {
    // Prevent competing with explicit content script registrations
    setTimeout(() => {
        const platforms: Platform[] = ["google_meet", "teams", "zoom"]
        const deRegisterPromises: Promise<unknown>[] = [deregisterContentScript(platforms), deregisterZoomRedirect()]

        Promise.all(deRegisterPromises)
            .then(() => {
                console.log("De-registered all content scripts. Starting re-registration...")

            })

        reRegisterContentScripts()
    }, 2000)
})


chrome.runtime.onInstalled.addListener(() => {
    // Undo merges of separate meetings at one address made before rejoins were bounded
    repairInvalidMerges().then(result => console.log(result))

    // Set defaults values
    chrome.storage.sync.get(null, function (resultSyncUntyped) {
        const resultSync = resultSyncUntyped as Partial<ResultSync>

        console.log(resultSync.wantTeams)

        chrome.storage.sync.set({
            autoPostWebhookAfterMeeting: resultSync.autoPostWebhookAfterMeeting === false ? false : true,
            autoDownloadFileAfterMeeting: resultSync.autoDownloadFileAfterMeeting === false ? false : true,
            // Opt-in: it sends the transcript to ChatGPT and takes the screen for a
            // few seconds, so it may never switch itself on.
            autoSummarizeAfterMeeting: resultSync.autoSummarizeAfterMeeting === true ? true : false,
            // Only seed the default when unset — an empty string is a deliberate
            // "save straight into Downloads" choice and must survive updates.
            downloadSubfolder: typeof resultSync.downloadSubfolder === "string" ? resultSync.downloadSubfolder : DEFAULT_DOWNLOAD_SUBFOLDER,
            obsidianEnabled: resultSync.obsidianEnabled === true ? true : false,
            obsidianFolder: typeof resultSync.obsidianFolder === "string" ? resultSync.obsidianFolder : DEFAULT_OBSIDIAN_FOLDER,
            obsidianSummaryFolder: typeof resultSync.obsidianSummaryFolder === "string" ? resultSync.obsidianSummaryFolder : DEFAULT_OBSIDIAN_SUMMARY_FOLDER,
            meetingsToKeep: typeof resultSync.meetingsToKeep === "number" ? resultSync.meetingsToKeep : DEFAULT_MEETINGS_TO_KEEP,
            obsidianBaseUrl: typeof resultSync.obsidianBaseUrl === "string" && resultSync.obsidianBaseUrl ? resultSync.obsidianBaseUrl : DEFAULT_OBSIDIAN_BASE_URL,
            operationMode: resultSync.operationMode === "manual" ? "manual" : "auto",
            hideCaptions: resultSync.hideCaptions === true ? true : false,
            participantName: typeof resultSync.participantName === "string" ? resultSync.participantName : "",
            webhookBodyType: resultSync.webhookBodyType === "advanced" ? "advanced" : "simple",
            wantGoogleMeet: resultSync.wantGoogleMeet === false ? false : true,
            wantTeams: resultSync.wantTeams === true ? true : false,
            wantZoom: resultSync.wantZoom === true ? true : false
        }, function () {
            // Re-register content scripts whenever extension is installed or updated, provided permissions are available. Suppress notification for silent background operation.
            reRegisterContentScripts()
        })

    })
})

