import type { ErrorObject, Meeting, ResultLocal, ResultSync, TranscriptBlock } from "@lib/types"
import { DEFAULT_MEETINGS_TO_KEEP } from "@lib/paths"
import { trimMeetings } from "@lib/retention"
import { appendOrMergeMeeting, coversStartTimestamp, hasEnoughToSummarize, splitInvalidMerges } from "@lib/meeting-merge"
import { downloadTranscript, postTranscriptToWebhook } from "./exporters"
import { postTranscriptToObsidian } from "./obsidian"
import { summarizeAndFileMeeting } from "./summarizer"

/**
 * Download transcripts, post webhook if URL is enabled and available.
 * Fails if transcript is empty or webhook request fails or if no meetings in storage.
 * @throws error codes: 009, 010, 011, 012, 013, 014
 */
let processingQueue: Promise<unknown> = Promise.resolve()

export function processLastMeeting(snapshot?: Partial<ResultLocal> & { captureId?: string; meetingEndTimestamp?: string }): Promise<string> {
    const run = processingQueue.then(() => processMeeting(snapshot))
    processingQueue = run.catch(() => {})
    return run
}

function processMeeting(snapshot?: Partial<ResultLocal> & { captureId?: string; meetingEndTimestamp?: string }): Promise<string> {
    return new Promise((resolve, reject) => {
        pickupLastMeetingFromStorage(snapshot)
            .then(() => {
                chrome.storage.local.get(["meetings"], function (resultLocalUntyped) {
                    const resultLocal = resultLocalUntyped as Partial<ResultLocal>
                    chrome.storage.sync.get(["webhookUrl", "autoPostWebhookAfterMeeting", "autoDownloadFileAfterMeeting", "obsidianEnabled", "autoSummarizeAfterMeeting"], function (resultSyncUntyped) {
                        const resultSync = resultSyncUntyped as Partial<ResultSync>

                        // pickupLastMeetingFromStorage resolved, so at least one meeting exists
                        const meetings = resultLocal.meetings as Meeting[]

                        // Create an array of promises to execute in parallel
                        const promises: Promise<unknown>[] = []

                        // Meeting index to download and post webhook
                        const lastIndex = snapshot
                            ? meetings.findIndex(meeting => (snapshot.captureId && meeting.captureId === snapshot.captureId)
                                || (coversStartTimestamp(meeting, snapshot.meetingStartTimestamp ?? "")
                                    && meeting.meetingUrl === snapshot.meetingUrl
                                    && meeting.meetingSoftware === snapshot.meetingSoftware))
                            : meetings.length - 1
                        if (lastIndex < 0) {
                            reject({ errorCode: "010", errorMessage: "Meeting at specified index not found" } satisfies ErrorObject)
                            return
                        }

                        // Promise to download transcript
                        if (resultSync.autoDownloadFileAfterMeeting) {
                            promises.push(
                                downloadTranscript(lastIndex)
                            )
                        }

                        // Promise to post webhook if enabled
                        if (resultSync.autoPostWebhookAfterMeeting && resultSync.webhookUrl) {
                            promises.push(postTranscriptToWebhook(lastIndex))
                        }

                        // Promise to save an Obsidian note if enabled. Obsidian runs
                        // alongside the text file rather than replacing it: the app may
                        // well be closed when a meeting ends, and the file is the backup.
                        if (resultSync.obsidianEnabled) {
                            promises.push(postTranscriptToObsidian(lastIndex))
                        }

                        // Execute all promises in parallel
                        Promise.all(promises)
                            .then(() => {
                                // Summarising is sequenced after the fan-out, not part
                                // of it: it drives a real browser tab for up to three
                                // minutes, and the transcript must already be safely
                                // filed before anything that slow and that fragile runs.
                                if (!resultSync.autoSummarizeAfterMeeting) {
                                    resolve("Meeting processing and download/webhook posting complete")
                                    return
                                }

                                // A few seconds before a reload or rejoin are filed as a
                                // meeting of their own, and summarising those put ChatGPT
                                // in front mid-call to produce a one-line summary. The
                                // Summarise button still works on them.
                                if (!hasEnoughToSummarize(meetings[lastIndex])) {
                                    console.log("Too little said to summarise automatically")
                                    resolve("Meeting processing complete, too short to summarise")
                                    return
                                }

                                // A failed summary is not a failed meeting. The
                                // transcript is downloaded, filed and posted by now, so
                                // ChatGPT being signed out or unreachable is reported to
                                // the console and the recent-errors list, never by
                                // rejecting the meeting-end path.
                                summarizeAndFileMeeting(lastIndex)
                                    .then(() => resolve("Meeting processing, exports and summary complete"))
                                    .catch((error) => {
                                        const parsedError = error as ErrorObject
                                        console.error("Automatic summary failed:", parsedError.errorMessage ?? error)
                                        recordSummaryFailure(parsedError)
                                        resolve("Meeting processing complete, automatic summary failed")
                                    })
                            })
                            .catch(error => {
                                // Fails with error codes: 009, 010, 011, 012
                                const parsedError = error as ErrorObject
                                console.error("Operation failed:", parsedError.errorMessage)
                                reject({ errorCode: parsedError.errorCode, errorMessage: parsedError.errorMessage } satisfies ErrorObject)
                            })
                    })
                })
            })
            .catch((error) => {
                // Fails with error codes: 013, 014
                const parsedError = error as ErrorObject
                reject({ errorCode: parsedError.errorCode, errorMessage: parsedError.errorMessage } satisfies ErrorObject)
            })
    })
}

/**
 * Notes a failed automatic summary in the same list the capture errors go to.
 *
 * The automatic run has no dialog to show a message in — nobody clicked anything — so
 * without this a signed-out ChatGPT would fail completely silently, which is the exact
 * failure mode the recent-errors list exists to prevent.
 */
function recordSummaryFailure(error: ErrorObject): void {
    chrome.storage.local.get(["recentErrors"], function (resultLocalUntyped) {
        const resultLocal = resultLocalUntyped as Partial<ResultLocal>
        const recentErrors = resultLocal.recentErrors ?? []
        recentErrors.push({
            at: new Date().toISOString(),
            code: error?.errorCode ?? "021",
            meetingSoftware: "",
            message: error?.errorMessage ?? "Automatic summary failed"
        })
        chrome.storage.local.set({ recentErrors: recentErrors.slice(-20) })
    })
}

/**
 * Appends the utterance that was still being spoken when the meeting ended.
 *
 * Closing the meeting tab kills the content script before it can flush its buffer, so
 * the buffer is persisted separately as it is spoken. It is skipped when the content
 * script did manage to flush — the same text would already be the last block.
 */
function mergeLiveTranscriptBlock(transcript: TranscriptBlock[], liveTranscriptBlock: TranscriptBlock | null | undefined): TranscriptBlock[] {
    if (!liveTranscriptBlock || !liveTranscriptBlock.transcriptText) {
        return transcript
    }

    const lastBlock = transcript[transcript.length - 1]
    if (lastBlock
        && (lastBlock.personName === liveTranscriptBlock.personName)
        // The persisted copy lags the final text, so the flushed block starts with it
        && (lastBlock.transcriptText.startsWith(liveTranscriptBlock.transcriptText))) {
        return transcript
    }

    console.log("Recovered the utterance in progress when the meeting ended")
    return [...transcript, liveTranscriptBlock]
}

/**
 * The meeting still in progress, as it stands right now, built the same way a finished
 * meeting is filed — including the utterance still being spoken — but left in place.
 * Nothing is written: the meeting carries on and is filed normally when it ends.
 *
 * @throws error codes: 022, 023
 */
export function readLiveMeeting(): Promise<Meeting> {
    return new Promise((resolve, reject) => {
        chrome.storage.local.get([
            "meetingTabId",
            "activeCaptureId",
            "meetingSoftware",
            "meetingTitle",
            "meetingStartTimestamp",
            "transcript",
            "chatMessages",
            "liveTranscriptBlock",
            "accountEmail",
        ], function (resultUntyped) {
            const result = resultUntyped as Partial<ResultLocal>

            // A tab id means a meeting is live; "processing" means it has just ended and
            // is being filed, which the finished-meeting path owns.
            if (typeof result.meetingTabId !== "number" || !result.meetingStartTimestamp) {
                reject({ errorCode: "023", errorMessage: "No meeting is in progress" } satisfies ErrorObject)
                return
            }

            const transcript = mergeLiveTranscriptBlock(result.transcript ?? [], result.liveTranscriptBlock)
            if (transcript.length === 0) {
                reject({ errorCode: "022", errorMessage: "Nothing to summarise for this meeting" } satisfies ErrorObject)
                return
            }

            resolve({
                captureId: result.activeCaptureId,
                meetingSoftware: result.meetingSoftware ? result.meetingSoftware : "",
                accountEmail: result.accountEmail,
                meetingTitle: result.meetingTitle,
                meetingStartTimestamp: result.meetingStartTimestamp,
                meetingEndTimestamp: new Date().toISOString(),
                transcript: transcript,
                chatMessages: result.chatMessages ?? [],
                webhookPostStatus: "new",
                obsidianPostStatus: "new",
            })
        })
    })
}

/**
 * Process transcript and chat messages of the meeting that just ended from storage,
 * format them into strings, and save as a new entry in meetings (keeping last 10).
 * @throws error codes: 013, 014
 */
export function pickupLastMeetingFromStorage(snapshot?: Partial<ResultLocal> & { captureId?: string; meetingEndTimestamp?: string }): Promise<string> {
    return new Promise((resolve, reject) => {
        chrome.storage.local.get([
            "meetingSoftware",
            "meetingTitle",
            "meetingStartTimestamp",
            "transcript",
            "chatMessages",
            "liveTranscriptBlock",
            "accountEmail",
            "captionsSeen",
            "meetingUrl",
        ], function (resultUntyped) {
            const result = snapshot ?? resultUntyped as Partial<ResultLocal>

            if (result.meetingStartTimestamp) {
                const transcript = mergeLiveTranscriptBlock(result.transcript ?? [], result.liveTranscriptBlock)
                const chatMessages = result.chatMessages ?? []

                if ((transcript.length > 0) || (chatMessages.length > 0)) {
                    // Create new transcript entry
                    const newMeetingEntry: Meeting = {
                        captureId: snapshot?.captureId,
                        meetingSoftware: result.meetingSoftware ? result.meetingSoftware : "",
                        accountEmail: result.accountEmail,
                        captionsSeen: result.captionsSeen === true,
                        meetingTitle: result.meetingTitle,
                        meetingStartTimestamp: result.meetingStartTimestamp,
                        meetingEndTimestamp: snapshot?.meetingEndTimestamp ?? new Date().toISOString(),
                        transcript: transcript,
                        chatMessages: chatMessages,
                        webhookPostStatus: "new",
                        obsidianPostStatus: "new",
                        meetingUrl: result.meetingUrl
                    }

                    // Get existing recent meetings and add the new meeting
                    chrome.storage.local.get(["meetings"], function (resultLocalUntyped) {
                        const resultLocal = resultLocalUntyped as Partial<ResultLocal>
                        chrome.storage.sync.get(["meetingsToKeep"], function (resultSyncUntyped) {
                        const resultSync = resultSyncUntyped as Partial<ResultSync>
                        let meetings = appendOrMergeMeeting(resultLocal.meetings ?? [], newMeetingEntry)

                        // Trim to the configured limit. trimMeetings has no access to the
                        // vault, downloads or the webhook, so this can only ever drop the
                        // extension's own copy.
                        const keep = typeof resultSync.meetingsToKeep === "number"
                            ? resultSync.meetingsToKeep
                            : DEFAULT_MEETINGS_TO_KEEP
                        meetings = trimMeetings(meetings, keep)

                        // Save updated recent transcripts, and drop the in-flight buffer
                        // now that it is part of a saved meeting
                        chrome.storage.local.set({ meetings: meetings, ...(snapshot ? {} : { liveTranscriptBlock: null }) }, function () {
                            console.log("Last meeting picked up")
                            resolve("Last meeting picked up")
                        })
                        })
                    })
                }
                else {
                    reject({ errorCode: "014", errorMessage: "Empty transcript and empty chatMessages" } satisfies ErrorObject)
                }
            }
            else {
                reject({ errorCode: "013", errorMessage: "No meetings found. May be attend one?" } satisfies ErrorObject)
            }
        })
    })
}

/**
 * Splits saved meetings that hold several separate meetings — what an unbounded rejoin
 * rule produced for a recurring meeting that keeps one address — and files the
 * resulting meetings to Obsidian when it is set up.
 *
 * Filing is sequential and stops at the first failure, as the bulk sync does: an
 * unreachable vault fails for every note. Anything not filed stays marked failed or new
 * in the meetings list, where its Obsidian button retries it.
 */
export function repairInvalidMerges(): Promise<string> {
    return new Promise((resolve) => {
        chrome.storage.local.get(["meetings"], function (resultLocalUntyped) {
            const resultLocal = resultLocalUntyped as Partial<ResultLocal>
            const meetings = resultLocal.meetings ?? []
            const repaired = splitInvalidMerges(meetings)
            if (repaired.length === meetings.length) {
                resolve("No merged meetings needed splitting")
                return
            }
            const kept = new Set(meetings)
            const toFile = repaired
                .map((meeting, index) => kept.has(meeting) ? -1 : index)
                .filter(index => index >= 0)

            chrome.storage.local.set({ meetings: repaired }, function () {
                console.log(`Split ${repaired.length - meetings.length} meeting(s) out of merged entries`)
                chrome.storage.sync.get(["obsidianEnabled"], function (resultSyncUntyped) {
                    const resultSync = resultSyncUntyped as Partial<ResultSync>
                    if (!resultSync.obsidianEnabled) {
                        resolve("Split merged meetings")
                        return
                    }
                    toFile.reduce<Promise<void>>(
                        (previous, index) => previous.then(() => postTranscriptToObsidian(index).then(() => undefined)),
                        Promise.resolve()
                    ).then(() => {
                        resolve("Split merged meetings and filed them to Obsidian")
                    }).catch((error) => {
                        console.error("Filing split meetings to Obsidian stopped:", error)
                        resolve("Split merged meetings; some were not filed to Obsidian")
                    })
                })
            })
        })
    })
}

/** @throws error codes: 009, 010, 011, 012, 013, 014 */
export function recoverLastMeeting(): Promise<string> {
    return new Promise((resolve, reject) => {
        chrome.storage.local.get(["meetings", "meetingStartTimestamp"], function (resultLocalUntyped) {
            const resultLocal = resultLocalUntyped as Partial<ResultLocal>
            // Check if user ever attended a meeting
            if (resultLocal.meetingStartTimestamp) {
                // A resumed meeting can have been moved within the saved list.
                const alreadySaved = (resultLocal.meetings ?? []).some(meeting =>
                    coversStartTimestamp(meeting, resultLocal.meetingStartTimestamp!))
                if (!alreadySaved) {
                    processLastMeeting().then(() => {
                        resolve("Recovered last meeting to the best possible extent")
                    }).catch((error) => {
                        // Fails with error codes: 009, 010, 011, 013, 014
                        const parsedError = error as ErrorObject
                        reject({ errorCode: parsedError.errorCode, errorMessage: parsedError.errorMessage } satisfies ErrorObject)
                    })
                }
                else {
                    resolve("No recovery needed")
                }
            }
            else {
                reject({ errorCode: "013", errorMessage: "No meetings found. May be attend one?" } satisfies ErrorObject)
            }
        })
    })
}
