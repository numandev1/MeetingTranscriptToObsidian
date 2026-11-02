import type { ContentScriptState, ErrorObject, ExtensionMessage } from "@lib/types"
import { extensionStatusJSON_bug, reportErrorMessage } from "../common-config"
import {
    applyParticipantNameSetting,
    broadcastLiveBuffer,
    checkExtensionStatus,
    clearLiveBuffer,
    createContentScriptState,
    logError,
    overWriteChromeStorage,
    pushBufferToTranscript,
    recoverLastMeeting,
    renderFab,
    startTranscriptMonitor,
    unmountFab,
    updateFabState,
    updateMeetingTitle,
    waitForElement,
} from "../common-utils"
import { SELECTORS_ZOOM } from "./config"
import { findNewPart, getPersonName, hasIframeLoaded, showNotificationZoom } from "./utils"

let isZoomInjected = false

export function watchForZoomMeeting(): void {
    setInterval(() => {
        // Meeting page
        const zoomUrlPattern = /^https:\/\/app\.zoom\.us\/wc\/\d+\/.+$/
        const isZoomUrlMatching = zoomUrlPattern.test(location.href)

        // On the meeting page and main zoom function is not running, inject it
        // This won't cause multiple main zoom injections into the current meeting because when the previous meeting ends, all UI elements are gone, destroying the corresponding event listeners
        if (isZoomUrlMatching && !isZoomInjected) {
            initZoom()
            isZoomInjected = true
        }
        // Set flag to false when meetings ends and the tab navigates to a non matching URL, or simply the current URL is a non meeting URL
        if (!isZoomUrlMatching) {
            isZoomInjected = false
        }
    }, 2000)
}

function initZoom(): void {
    // Attempt to recover last meeting, if any. Abort if it takes more than 2 seconds to prevent current meeting getting messed up.
    Promise.race([
        recoverLastMeeting(),
        new Promise((_, reject) =>
            setTimeout(() => reject({ errorCode: "016", errorMessage: "Recovery timed out" } satisfies ErrorObject), 2000)
        )
    ]).
        catch((error) => {
            const parsedError = error as ErrorObject
            if ((parsedError.errorCode !== "013") && (parsedError.errorCode !== "014")) {
                console.error(parsedError.errorMessage)
            }
        }).
        finally(() => {
            // Initialise new state for current meeting
            const state = createContentScriptState("Zoom", "zoom")
            applyParticipantNameSetting(state)
            // Push fresh state to chrome storage
            overWriteChromeStorage(state, ["meetingSoftware", "meetingStartTimestamp", "meetingTitle", "transcript", "chatMessages"], false)
            // Drop any buffer left over from a meeting that was never picked up
            clearLiveBuffer(state)

            checkExtensionStatus(state).finally(() => {
                console.log("Extension status " + state.extensionStatusJSON.status)

                // Enable extension functions only if status is 200
                if (state.extensionStatusJSON.status === 200) {

                    zoomMeetingRoutines(state)
                }
                else {
                    // Show downtime message as extension status is 400
                    showNotificationZoom(state.extensionStatusJSON)
                }
            })
        })
}

function zoomMeetingRoutines(state: ContentScriptState): void {
    renderFab(state)

    waitForElement(SELECTORS_ZOOM.IFRAME).then(() => {
        console.log(`Found iframe`)
        const iframe = document.querySelector(SELECTORS_ZOOM.IFRAME) as HTMLIFrameElement | null

        if (iframe) {
            hasIframeLoaded(iframe).then(() => {
                console.log("Iframe loaded")
                const iframeDOM = iframe.contentDocument

                // CRITICAL DOM DEPENDENCY. Wait until the meeting end icon appears, used to detect meeting start
                if (iframeDOM) {
                    waitForElement(SELECTORS_ZOOM.AUDIO_OPTION_MENU, undefined, iframe).then(() => {
                        console.log("Meeting started")
                        const message: ExtensionMessage = {
                            type: "new_meeting_started",
            captureId: state.captureId
                        }
                        chrome.runtime.sendMessage(message, function () { })
                        state.hasMeetingStarted = true
                        // Update meeting startTimestamp
                        state.meetingStartTimestamp = new Date().toISOString()
                        overWriteChromeStorage(state, ["meetingStartTimestamp"], false)

                        //*********** MEETING START ROUTINES **********//
                        updateMeetingTitle(state)

                        // Ask user to switch on captions
                        showNotificationZoom(state.extensionStatusJSON)

                        // **** REGISTER TRANSCRIPT LISTENER **** //
                        // Wait for transcript node to be visible
                        waitForElement(SELECTORS_ZOOM.TRANSCRIPT_CONTAINER, undefined, iframe).
                            then((element) => {
                                console.log("Found captions container")
                                // CRITICAL DOM DEPENDENCY. Grab the transcript element.
                                state.transcriptTargetNode = element

                                if (state.transcriptTargetNode) {
                                    console.log(`Registering mutation observer on ${SELECTORS_ZOOM.TRANSCRIPT_CONTAINER}`)

                                    // Create transcript observer instance linked to the callback function. Registered irrespective of operation mode, so that any visible transcript can be picked up during the meeting, independent of the operation mode.
                                    // Initial attach and monitor every 2s
                                    startTranscriptMonitor(state, transcriptMutationCallbackZoom)
                                }
                                else {
                                    throw new Error("Transcript element not found in DOM")
                                }
                            })
                            .catch((err) => {
                                console.error(err)
                                state.isTranscriptDomErrorCaptured = true
                                showNotificationZoom(extensionStatusJSON_bug)

                                logError(state, "001", err)
                            })


                        //*********** MEETING END ROUTINES **********//
                        try {
                            // CRITICAL DOM DEPENDENCY. Event listener to capture meeting end button click by user
                            const endCallElement = iframeDOM.querySelector(SELECTORS_ZOOM.LEAVE_BUTTON_CONTAINER)
                            endCallElement?.firstChild?.addEventListener("click", function meetingEndRoutines() {
                                endCallElement.removeEventListener("click", meetingEndRoutines)
                                console.log("Meeting ended")
                                // To suppress further errors
                                state.hasMeetingEnded = true
                                if (state.transcriptObserver) {
                                    state.transcriptObserver.disconnect()
                                }

                                // Push any data in the buffer variables to the transcript array. Needed to handle one or more speaking when meeting ends.

                                pushBufferToTranscript(state)
                                // Save to chrome storage and send message to download transcript from background script
                                overWriteChromeStorage(state, ["transcript", "chatMessages"], true)

                                unmountFab()
                            })
                        } catch (err) {
                            console.error(err)
                            showNotificationZoom(extensionStatusJSON_bug)

                            logError(state, "004", err)
                        }
                    })
                }
            })
        }
    })
}

/** Callback function to execute when transcription mutations are observed. */
function transcriptMutationCallbackZoom(state: ContentScriptState, mutationsList: MutationRecord[]): void {
    mutationsList.forEach((mutation) => {
        try {
            if (mutation.type === "characterData") {
                const mutationTargetElement = mutation.target.parentElement
                const currentTranscriptText = mutationTargetElement?.textContent

                if (currentTranscriptText) {
                    // Find person name using various strategies
                    const currentPersonName = getPersonName(mutationTargetElement) || "Person"

                    if (currentPersonName && currentTranscriptText) {
                        // Starting fresh in a meeting or resume from no active transcript
                        if (state.stateTranscriptBlock.transcriptTextBuffer === "") {
                            state.stateTranscriptBlock.personName = currentPersonName
                            state.stateTranscriptBlock.timestamp = new Date().toISOString()
                            state.stateTranscriptBlock.transcriptTextBuffer = currentTranscriptText
                        }
                        // Some prior transcript buffer exists
                        else {
                            // New person started speaking
                            if (state.stateTranscriptBlock.personName !== currentPersonName) {
                                // Push previous person's transcript as a block
                                pushBufferToTranscript(state)

                                // Update buffers for next mutation and store transcript block timestamp
                                state.stateTranscriptBlock.personName = currentPersonName
                                state.stateTranscriptBlock.timestamp = new Date().toISOString()
                                state.stateTranscriptBlock.transcriptTextBuffer = currentTranscriptText
                            }
                            // Same person speaking more
                            else {
                                // Update buffers for next mutation
                                // Append only the new part of the transcript
                                state.stateTranscriptBlock.transcriptTextBuffer = state.stateTranscriptBlock.transcriptTextBuffer + findNewPart(state.stateTranscriptBlock.transcriptTextBuffer, currentTranscriptText)
                            }
                        }
                    }
                }

                // Rendered by the side panel
                broadcastLiveBuffer(state)
                updateFabState(state)
            }
        }
        catch (err) {
            console.error(err)
            if (!state.isTranscriptDomErrorCaptured && !state.hasMeetingEnded) {
                console.log(reportErrorMessage)
                showNotificationZoom(extensionStatusJSON_bug)

                logError(state, "005", err)
            }
            state.isTranscriptDomErrorCaptured = true
        }
    })
}
