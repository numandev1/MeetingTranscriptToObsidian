import type { ChatMessage, ContentScriptState, ErrorObject, ExtensionMessage, ExtensionResponse, ResultSync } from "@lib/types"
import { extensionStatusJSON_bug, mutationConfig, reportErrorMessage } from "../common-config"
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
import { SELECTORS_GOOGLE_MEET } from "./config"
import { captureUserName, pushChatBlock, selectControlButton, selectElements, showNotificationGoogleMeet } from "./utils"

export function initGoogleMeet(): void {
    // Attempt to recover last meeting, if any. Abort if it takes more than 15 seconds to prevent current meeting getting messed up.
    Promise.race([
        recoverLastMeeting(),
        new Promise((_, reject) =>
            setTimeout(() => reject({ errorCode: "016", errorMessage: "Recovery timed out" } satisfies ErrorObject), 15000)
        )
    ]).catch((error) => {
        const parsedError = error as ErrorObject
        if ((parsedError.errorCode !== "013") && (parsedError.errorCode !== "014")) {
            console.error(parsedError.errorMessage)
        }
    }).finally(() => {
        // Initialise new state for current meeting
        const state = createContentScriptState("Google Meet", "google_meet")
        applyParticipantNameSetting(state)
        // Push fresh state to chrome storage
        overWriteChromeStorage(state, ["meetingSoftware", "meetingStartTimestamp", "meetingTitle", "transcript", "chatMessages"], false)
        listenForEmbeddedChatMessages(state)
        // Drop any buffer left over from a meeting that was never picked up
        clearLiveBuffer(state)

        checkExtensionStatus(state).finally(() => {
            console.log("Extension status " + state.extensionStatusJSON.status)

            // Enable extension functions only if status is 200
            if (state.extensionStatusJSON.status === 200) {
                googleMeetRoutines(state)
            }
            else {
                // Show downtime message as extension status is 400
                showNotificationGoogleMeet(state.extensionStatusJSON)
            }
        })
    })
}

function googleMeetRoutines(state: ContentScriptState): void {
    renderFab(state)

    // NON CRITICAL DOM DEPENDENCY
    captureUserName(state)

    // CRITICAL DOM DEPENDENCY. Wait until the meeting end icon appears, used to detect meeting start
    waitForElement(SELECTORS_GOOGLE_MEET.GOOGLE_SYMBOLS, SELECTORS_GOOGLE_MEET.TEXT_CALL_END).then(() => {
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

        // **** REGISTER TRANSCRIPT AND CHAT MESSAGES LISTENERS **** //

        // REGISTER TRANSCRIPT LISTENER
        // Wait for captions icon to be visible. When user is waiting in meeting lobbing for someone to let them in, the call end icon is visible, but the captions icon is still not visible.
        waitForElement(SELECTORS_GOOGLE_MEET.GOOGLE_SYMBOLS,
            [SELECTORS_GOOGLE_MEET.TEXT_CAPTIONS_OFF, SELECTORS_GOOGLE_MEET.TEXT_CAPTIONS_ON])
            .then((captionsIcon) => {
                // Which ligature matched tells us the current state. Clicking when
                // captions are already on would switch them off, losing the transcript.
                const captionsAlreadyOn =
                    (captionsIcon.textContent ?? "").trim() === SELECTORS_GOOGLE_MEET.TEXT_CAPTIONS_ON

                chrome.storage.sync.get(["operationMode"], function (resultSyncUntyped) {
                    const resultSync = resultSyncUntyped as Partial<ResultSync>
                    if (resultSync.operationMode === "manual") {
                        console.log("Manual mode selected, leaving captions as they are")
                    }
                    else if (captionsAlreadyOn) {
                        console.log("Captions already on, nothing to switch")
                    }
                    else {
                        // CRITICAL DOM DEPENDENCY
                        selectControlButton(SELECTORS_GOOGLE_MEET.GOOGLE_SYMBOLS,
                            SELECTORS_GOOGLE_MEET.TEXT_CAPTIONS_OFF)?.click()
                    }
                })

                // Allow DOM to be updated. Once updated, next "then" block will be executed.
                return waitForElement(SELECTORS_GOOGLE_MEET.TRANSCRIPT_REGION)
                    .then(targetNode => (targetNode))
            })
            .then((targetNode) => {
                if (targetNode) {
                    // CRITICAL DOM DEPENDENCY. Grab the transcript element. This element is present, irrespective of captions ON/OFF, so this executes independent of operation mode.
                    state.transcriptTargetNode = targetNode
                    // Initial attach and monitor every 2s
                    startTranscriptMonitor(state, transcriptMutationCallbackGoogleMeet)

                    // Show confirmation message from extensionStatusJSON, once observation has started, based on operation mode
                    chrome.storage.sync.get(["operationMode"], function (resultSyncUntyped) {
                        const resultSync = resultSyncUntyped as Partial<ResultSync>
                        if (resultSync.operationMode === "manual") {
                            showNotificationGoogleMeet({ status: 400, message: "<strong>MeetingTranscriptToObsidian is not running</strong> <br /> Turn on captions using the CC icon, if needed" })
                        }
                        else {
                            showNotificationGoogleMeet(state.extensionStatusJSON)
                        }
                    })
                }
                else {
                    throw new Error("Transcript element not found in DOM")
                }
            })
            .catch((err) => {
                console.error(err)
                state.isTranscriptDomErrorCaptured = true
                showNotificationGoogleMeet(extensionStatusJSON_bug)

                logError(state, "001", err)
            })


        // REGISTER CHAT MESSAGES LISTENER
        // Wait for chat icon to be visible. When user is waiting in meeting lobbing for someone to let them in, the call end icon is visible, but the chat icon is still not visible.
        waitForElement(SELECTORS_GOOGLE_MEET.GOOGLE_SYMBOLS, SELECTORS_GOOGLE_MEET.TEXT_CHAT)
            .then(() => {
                const chatMessagesButton = selectControlButton(SELECTORS_GOOGLE_MEET.GOOGLE_SYMBOLS, SELECTORS_GOOGLE_MEET.TEXT_CHAT)
                // Force open chat messages to make the required DOM to appear. Otherwise, the required chatMessages DOM element is not available.
                chatMessagesButton?.click()

                // Allow DOM to be updated. Once updated, next "then" block will be executed.
                // Whichever chat implementation this meeting uses appears first.
                return waitForElement(`${SELECTORS_GOOGLE_MEET.CHAT_LIVE_REGION}, ${SELECTORS_GOOGLE_MEET.EMBEDDED_CHAT_IFRAME}`)
                    .then(targetNode => ({ targetNode, chatMessagesButton }))
            })
            .then(({ targetNode, chatMessagesButton }) => {
                // Click again to close the chat messages
                chatMessagesButton?.click()

                // Embedded Google Chat: the iframe's own content script reads the messages
                // and they arrive through listenForEmbeddedChatMessages(), so there is no
                // Meet-side DOM to observe.
                if (targetNode.matches(SELECTORS_GOOGLE_MEET.EMBEDDED_CHAT_IFRAME)) {
                    console.log("Chat is an embedded Google Chat iframe; relying on relayed messages")
                    return
                }
                // CRITICAL DOM DEPENDENCY. Grab the chat messages element. This element is present, irrespective of chat ON/OFF, once it appears for this first time.
                state.chatMessagesTargetNode = targetNode

                // Create chat messages observer instance linked to the callback function. Registered irrespective of operation mode.
                if (state.chatMessagesTargetNode) {
                    state.chatMessagesObserver = new MutationObserver((mutations) =>
                        chatMessagesMutationCallback(state, mutations)
                    )
                    state.chatMessagesObserver.observe(state.chatMessagesTargetNode, mutationConfig)
                    // Messages sent before the observer attached are already in the DOM
                    captureChatMessages(state)
                }
                else {
                    throw new Error("Chat messages element not found in DOM")
                }
            })
            .catch((err) => {
                console.error(err)
                state.isChatMessagesDomErrorCaptured = true
                showNotificationGoogleMeet(extensionStatusJSON_bug)

                logError(state, "003", err)
            })

        //*********** MEETING END ROUTINES **********//
        try {
            // CRITICAL DOM DEPENDENCY. Event listener to capture meeting end button click by user
            const callEndButton = selectElements(SELECTORS_GOOGLE_MEET.GOOGLE_SYMBOLS, SELECTORS_GOOGLE_MEET.TEXT_CALL_END)[0]
            const callEndClickTarget = callEndButton.parentElement!.parentElement!
            callEndClickTarget.addEventListener("click", () => {
                // To suppress further errors
                state.hasMeetingEnded = true

                if (state.transcriptObserver) {
                    state.transcriptObserver.disconnect()
                }
                if (state.chatMessagesObserver) {
                    state.chatMessagesObserver.disconnect()
                }

                // Push any data in the buffer variables to the transcript array. Needed to handle one or more speaking when meeting ends.
                pushBufferToTranscript(state)
                // Final sweep, so anything the observer missed while the chat panel was
                // closed is still picked up before the meeting is filed
                captureChatMessages(state)
                // Save to chrome storage and send message to download transcript from background script
                overWriteChromeStorage(state, ["transcript", "chatMessages"], true)

                unmountFab()
            })
        } catch (err) {
            console.error(err)
            showNotificationGoogleMeet(extensionStatusJSON_bug)

            logError(state, "004", err)
        }
    })
}

/** Callback function to execute when transcription mutations are observed. */
function transcriptMutationCallbackGoogleMeet(state: ContentScriptState, mutationsList: MutationRecord[]): void {
    mutationsList.forEach((mutation) => {
        try {
            if (mutation.type === "characterData") {
                const mutationTargetElement = mutation.target.parentElement
                const transcriptUIBlocks = [...mutationTargetElement?.parentElement?.parentElement?.children || []]
                const isLastButSecondElement = transcriptUIBlocks[transcriptUIBlocks.length - 3] === mutationTargetElement?.parentElement ? true : false

                // Pick up only last second element (the last and last but one are non transcript elements), since Meet mutates previous blocks to make minor corrections. Picking them up leads to repetitive transcript blocks in the result.
                if (isLastButSecondElement) {
                    const currentPersonName = mutationTargetElement?.previousSibling?.textContent
                    const currentTranscriptText = mutationTargetElement?.textContent

                    if (currentPersonName && currentTranscriptText) {
                        // Starting fresh in a meeting or resume from no active transcript
                        if (!state.stateTranscriptBlock.mutationTargetElement) {
                            state.stateTranscriptBlock.mutationTargetElement = mutation.target.parentElement
                            state.stateTranscriptBlock.personName = currentPersonName
                            state.stateTranscriptBlock.timestamp = new Date().toISOString()
                            state.stateTranscriptBlock.transcriptTextBuffer = currentTranscriptText
                        }
                        // Some prior transcript buffer exists
                        else {
                            // New person started speaking
                            if (state.stateTranscriptBlock.mutationTargetElement !== mutation.target.parentElement) {
                                // Push previous person's transcript as a block
                                pushBufferToTranscript(state)

                                // Update stateTranscriptBlock for next mutation and store transcript block timestamp
                                state.stateTranscriptBlock.mutationTargetElement = mutation.target.parentElement
                                state.stateTranscriptBlock.personName = currentPersonName
                                state.stateTranscriptBlock.timestamp = new Date().toISOString()
                                state.stateTranscriptBlock.transcriptTextBuffer = currentTranscriptText
                            }
                            // Same person speaking more
                            else {
                                // When the same person speaks for more than 30 min (approx), Meet drops very long transcript for current person and starts over, which is detected by current transcript string being significantly smaller than the previous one
                                // TO VERIFY IF NEEDED
                                if ((currentTranscriptText.length - state.stateTranscriptBlock.transcriptTextBuffer.length) < -250) {
                                    // Push the long transcript
                                    pushBufferToTranscript(state)

                                    // Store transcript block timestamp for next transcript block of same person
                                    state.stateTranscriptBlock.mutationTargetElement = mutation.target.parentElement
                                    state.stateTranscriptBlock.timestamp = new Date().toISOString()
                                }

                                // Update stateTranscriptBlock for next mutation
                                state.stateTranscriptBlock.transcriptTextBuffer = currentTranscriptText
                            }
                        }
                    }
                    // No people found in transcript DOM
                    else {
                        // No transcript yet or the last person stopped speaking(and no one has started speaking next)
                        console.log("No active transcript")
                        // Push data in the buffer variables to the transcript array
                        pushBufferToTranscript(state)
                        // Update stateTranscriptBlock for the next person in the next mutation
                        state.stateTranscriptBlock.mutationTargetElement = null
                        state.stateTranscriptBlock.personName = ""
                        state.stateTranscriptBlock.transcriptTextBuffer = ""
                    }
                }
            }

            // Rendered by the side panel
            broadcastLiveBuffer(state)
            updateFabState(state)
        } catch (err) {
            console.error(err)
            if (!state.isTranscriptDomErrorCaptured && !state.hasMeetingEnded) {
                console.log(reportErrorMessage)
                showNotificationGoogleMeet(extensionStatusJSON_bug)

                logError(state, "005", err)
            }
            state.isTranscriptDomErrorCaptured = true
        }
    })
}

/**
 * Reads every chat message currently in the DOM and stores any not seen before.
 *
 * The previous implementation read only the single newest message on each mutation,
 * which silently dropped whole batches: Google Meet renders a run of messages at once
 * when the chat panel is opened, and two messages arriving between mutation callbacks
 * would leave only the last one. Sweeping the whole region and de-duplicating on
 * data-message-id captures all of them, and makes the sweep safe to repeat.
 */
export function captureChatMessages(state: ContentScriptState): void {
    try {
        // CRITICAL DOM DEPENDENCY
        const chatMessagesElement = document.querySelector(SELECTORS_GOOGLE_MEET.CHAT_LIVE_REGION)
        if (!chatMessagesElement) {
            return
        }

        // Each run of consecutive messages from one person is a group with its own header
        for (const group of Array.from(chatMessagesElement.children)) {
            const header = group.querySelector(SELECTORS_GOOGLE_MEET.CHAT_GROUP_HEADER)
            // Own messages show only a timestamp; everyone else gets a name first
            const personName = (header && header.childNodes.length > 1)
                ? (header.firstChild?.textContent ?? state.userName)
                : state.userName

            const seenInGroup = new Set<string>()
            for (const messageElement of Array.from(group.querySelectorAll(SELECTORS_GOOGLE_MEET.CHAT_MESSAGE))) {
                const messageId = messageElement.getAttribute("data-message-id")
                // The id is repeated on nested elements; only the outermost is the message
                if (!messageId || seenInGroup.has(messageId)) {
                    continue
                }
                seenInGroup.add(messageId)

                if (state.capturedChatMessageIds.includes(messageId)) {
                    continue
                }

                const chatMessageText = getChatMessageText(messageElement)
                if (!chatMessageText) {
                    continue
                }

                state.capturedChatMessageIds.push(messageId)
                pushChatBlock(state, {
                    personName: personName,
                    timestamp: new Date().toISOString(),
                    chatMessageText: chatMessageText
                })
            }
        }
    }
    catch (err) {
        console.error(err)
        if (!state.isChatMessagesDomErrorCaptured && !state.hasMeetingEnded) {
            console.log(reportErrorMessage)
            showNotificationGoogleMeet(extensionStatusJSON_bug)

            logError(state, "006", err)
        }
        state.isChatMessagesDomErrorCaptured = true
    }
}

/**
 * Pulls the message body out of a chat message element.
 *
 * The element also contains hover controls ("keep", "Pin message"), so its innerText
 * cannot be used directly. The body is the first element branch; the controls sit in a
 * sibling that always contains a button, which is what the fallback strips.
 */
function getChatMessageText(messageElement: Element): string {
    const bodyElement = messageElement.firstElementChild?.firstElementChild?.firstElementChild
    const body = (bodyElement as HTMLElement | null | undefined)?.innerText?.trim()
    if (body) {
        return body
    }

    // Structure changed: drop anything holding a control and take what is left
    const clone = messageElement.cloneNode(true) as HTMLElement
    Array.from(clone.querySelectorAll("button")).forEach(button => {
        const container = button.parentElement?.parentElement ?? button
        container.remove()
    })
    return (clone.innerText || "").trim()
}

/**
 * Receives messages read by the content script inside Meet's embedded Google Chat
 * iframe, relayed to this frame by the background script, and stores new ones.
 */
function listenForEmbeddedChatMessages(state: ContentScriptState): void {
    chrome.runtime.onMessage.addListener((messageUnTyped, _sender, sendResponse) => {
        const message = messageUnTyped as ExtensionMessage
        if (message.type !== "meet_embedded_chat_messages") {
            return
        }
        // Acknowledged either way, so the iframe stops resending a batch the meeting no longer wants
        sendResponse({ success: true } satisfies ExtensionResponse)
        if (state.hasMeetingEnded) {
            return
        }
        for (const chatMessage of message.embeddedChatMessages ?? []) {
            // Prefixed so these ids can never collide with legacy data-message-id values
            const messageId = `embedded:${chatMessage.messageId}`
            if (state.capturedChatMessageIds.includes(messageId)) {
                continue
            }
            state.capturedChatMessageIds.push(messageId)
            pushChatBlock(state, {
                personName: chatMessage.personName,
                timestamp: chatMessage.timestamp,
                chatMessageText: chatMessage.chatMessageText
            })
        }
    })
}

/** Callback function to execute when chat messages mutations are observed. */
function chatMessagesMutationCallback(state: ContentScriptState, mutationsList: MutationRecord[]): void {
    if (mutationsList.length > 0) {
        captureChatMessages(state)
    }
}
