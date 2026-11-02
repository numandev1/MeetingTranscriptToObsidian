import type { ContentScriptState, ErrorObject, ExtensionMessage, ResultSync } from "@lib/types"
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
import { SELECTORS_TEAMS } from "./config"
import { dispatchLiveCaptionsShortcut, showNotificationTeams } from "./utils"

let isTeamsInjected = false

export function watchForTeamsMeeting(): void {
  setInterval(() => {
    // Meeting lobby
    const isJoinButtonFound = document.querySelector(SELECTORS_TEAMS.PREJOIN_JOIN_BUTTON)

    // On the meeting lobby and main teams function is not running, inject it
    // This won't cause multiple main teams injections into the current meeting because when the previous meeting ends, all UI elements are gone, destroying the corresponding event listeners
    if (isJoinButtonFound && !isTeamsInjected) {
      initTeams()
      isTeamsInjected = true
    }
    // Reset flag for next meeting lobby visit
    if (!isJoinButtonFound) {
      isTeamsInjected = false
    }
  }, 2000)
}

function initTeams(): void {
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
      const state = createContentScriptState("Teams", "teams")
      applyParticipantNameSetting(state)
      // Push fresh state to chrome storage
      overWriteChromeStorage(state, ["meetingSoftware", "meetingStartTimestamp", "meetingTitle", "transcript", "chatMessages"], false)
      // Drop any buffer left over from a meeting that was never picked up
      clearLiveBuffer(state)

      checkExtensionStatus(state).finally(() => {
        console.log("Extension status " + state.extensionStatusJSON.status)

        // Enable extension functions only if status is 200
        if (state.extensionStatusJSON.status === 200) {

          teamsMeetingRoutines(state)
        }
        else {
          // Show downtime message as extension status is 400
          showNotificationTeams(state.extensionStatusJSON)
        }
      })
    })
}

function teamsMeetingRoutines(state: ContentScriptState): void {
  renderFab(state)

  // CRITICAL DOM DEPENDENCY. Wait until the meeting end icon appears, used to detect meeting start
  waitForElement(SELECTORS_TEAMS.HANGUP_BUTTON).then(() => {
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

    // Fire captions shortcut based on operation mode. Async operation.
    chrome.storage.sync.get(["operationMode"], function (resultSyncUntyped) {
      const resultSync = resultSyncUntyped as Partial<ResultSync>
      if (resultSync.operationMode === "manual") {
        console.log("Manual mode selected, leaving transcript off")
        showNotificationTeams({ status: 400, message: "<strong>MeetingTranscriptToObsidian is not running</strong> <br /> Turn on captions, if needed (More > Language > Captions)" })
      }
      else {
        // Allow keyboard event listener to be ready
        setTimeout(() => {
          dispatchLiveCaptionsShortcut()
          // Show message to enable because keyboard shortcut does not work in guest meetings
          showNotificationTeams(state.extensionStatusJSON)
        }, 2000)
      }
    })

    // **** REGISTER TRANSCRIPT LISTENER **** //
    // Wait for transcript node to be visible
    waitForElement(SELECTORS_TEAMS.CAPTIONS_REGION).then((element) => {
      console.log("Found captions container")
      // CRITICAL DOM DEPENDENCY. Grab the transcript element.
      state.transcriptTargetNode = element

      if (state.transcriptTargetNode) {
        // Create transcript observer instance linked to the callback function. Registered irrespective of operation mode, so that any visible transcript can be picked up during the meeting, independent of the operation mode.
        // Initial attach and monitor every 2s
        startTranscriptMonitor(state, transcriptMutationCallbackTeams)
      }
      else {
        throw new Error("Transcript element not found in DOM")
      }
    })
      .catch((err) => {
        console.error(err)
        state.isTranscriptDomErrorCaptured = true
        showNotificationTeams(extensionStatusJSON_bug)

        logError(state, "001", err)
      })


    //*********** MEETING END ROUTINES **********//
    waitForElement(SELECTORS_TEAMS.HANGUP_BUTTON).then((element) => {
      // For some reason, capturing the reference to #hangup-button immediately is not working. Need to wait for a moment.
      setTimeout(() => {
        // CRITICAL DOM DEPENDENCY. Event listener to capture meeting end button click by user
        let endCallElement = element
        if (endCallElement?.nextElementSibling?.tagName === "BUTTON") {
          endCallElement = endCallElement?.parentElement as Element
        }
        endCallElement?.addEventListener("click", meetingEndRoutines)

        function meetingEndRoutines() {
          endCallElement?.removeEventListener("click", meetingEndRoutines)
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
        }
      }, 1000)
    })
  })
}

/** Callback function to execute when transcription mutations are observed. */
function transcriptMutationCallbackTeams(state: ContentScriptState, mutationsList: MutationRecord[]): void {
  mutationsList.forEach((mutation) => {
    try {
      if (mutation.type === "characterData") {
        const mutationTargetElement = mutation.target.parentElement

        const currentPersonName = mutationTargetElement?.parentElement?.previousSibling?.textContent
        const currentTranscriptText = mutationTargetElement?.textContent

        if (currentPersonName && currentTranscriptText) {
          // Starting fresh in a meeting
          if (!state.stateTranscriptBlock.mutationTargetElement) {
            state.stateTranscriptBlock.mutationTargetElement = mutation.target.parentElement
            state.stateTranscriptBlock.personName = currentPersonName
            state.stateTranscriptBlock.timestamp = new Date().toISOString()
            state.stateTranscriptBlock.transcriptTextBuffer = currentTranscriptText
          }
          // Some prior transcript buffer exists
          else {
            // New transcript UI block
            if (state.stateTranscriptBlock.mutationTargetElement !== mutation.target.parentElement) {
              // Push previous transcript block
              pushBufferToTranscript(state)

              // Update stateTranscriptBlock for next mutation and store transcript block timestamp
              state.stateTranscriptBlock.mutationTargetElement = mutation.target.parentElement
              state.stateTranscriptBlock.personName = currentPersonName
              state.stateTranscriptBlock.timestamp = new Date().toISOString()
              state.stateTranscriptBlock.transcriptTextBuffer = currentTranscriptText
            }
            // Same transcript UI block being appended
            else {
              // Update stateTranscriptBlock for next mutation
              state.stateTranscriptBlock.transcriptTextBuffer = currentTranscriptText
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
        showNotificationTeams(extensionStatusJSON_bug)

        logError(state, "005", err)
      }
      state.isTranscriptDomErrorCaptured = true
    }
  })
}
