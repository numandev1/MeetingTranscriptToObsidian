import type {
    ChatMessage,
    ContentScriptState,
    ErrorObject,
    ExtensionMessage,
    ExtensionResponse,
    MeetingSoftware,
    Platform,
    ResultLocal,
    RecentError,
    ResultSync,
    TranscriptBlock,
} from "@lib/types"
import { LOGO_URL, NOTIFICATION_PLATFORM_CONFIGS, mutationConfig } from "./common-config"
import { SELECTORS_GOOGLE_MEET } from "./google-meet/config"
import { SELECTORS_TEAMS } from "./teams/config"
import { SELECTORS_ZOOM } from "./zoom/config"

/**
 * How long any single DOM wait may run before it is treated as a breakage.
 * Generous, because a user can sit in a meeting lobby for minutes before the controls
 * render, but finite so a vendor change surfaces as an error rather than silence.
 */
export const DEFAULT_ELEMENT_TIMEOUT_MS = 5 * 60 * 1000

/** Keys of the content script state that can be flushed to chrome.storage.local. */
export type StorableStateKey = "meetingSoftware" | "meetingTitle" | "meetingStartTimestamp" | "transcript" | "chatMessages"

/**
 * Platform specific handler invoked for every batch of captions mutations.
 * Passed in by each platform's entry point so that this module does not have to
 * import the platform modules that import it.
 */
export type TranscriptMutationCallback = (state: ContentScriptState, mutations: MutationRecord[]) => void

/** State Factory: Returns a pristine, isolated meeting state block. */
export function createContentScriptState(meetingSoftware: MeetingSoftware, platform: Platform): ContentScriptState {
    return {
        captureId: crypto.randomUUID(),
        meetingSoftware: meetingSoftware,
        platform: platform,
        userName: "You",
        transcript: [],
        chatMessages: [],
        stateTranscriptBlock: {
            timestamp: "",
            mutationTargetElement: null,
            personName: "",
            transcriptTextBuffer: "",

        },
        meetingStartTimestamp: new Date().toISOString(),
        meetingTitle: document.title,
        transcriptTargetNode: null,
        transcriptObserver: null,
        chatMessagesTargetNode: null,
        chatMessagesObserver: null,
        isUserNameLocked: false,
        capturedChatMessageIds: [],
        isTranscriptDomErrorCaptured: false,
        isChatMessagesDomErrorCaptured: false,
        hasMeetingStarted: false,
        hasMeetingEnded: false,
        extensionStatusJSON: {
            status: 200,
            message: "<strong>MeetingTranscriptToObsidian is running</strong> <br /> Do not turn off captions"
        }
    }
}

/**
 * Applies the user's configured display name, so captions attributed to "You" are
 * written out under their real name.
 *
 * Auto-detection from the meeting UI only exists for Google Meet, so without this
 * Teams and Zoom transcripts would always say "You".
 */
export function applyParticipantNameSetting(state: ContentScriptState): void {
    chrome.storage.sync.get(["participantName"], function (resultSyncUntyped) {
        const resultSync = resultSyncUntyped as Partial<ResultSync>
        const participantName = resultSync.participantName?.trim()

        if (participantName) {
            state.userName = participantName
            // Beats anything scraped from the meeting UI
            state.isUserNameLocked = true
        }
    })
}

/**
 * Prepares the in-meeting banner text.
 *
 * This used to fetch a status document from the upstream project's site, which acted as
 * a remote kill switch: a non-200 status there, or a minVersion above the installed
 * build, disabled capture on every user's machine. A fork must not depend on another
 * project's infrastructure to keep working, and a privacy-focused extension should not
 * announce every meeting to a third party, so the check is gone and capture always runs.
 */
export function checkExtensionStatus(state: ContentScriptState): Promise<string> {
    state.extensionStatusJSON = {
        status: 200,
        message: state.meetingSoftware ? NOTIFICATION_PLATFORM_CONFIGS[state.platform].notificationText : ""
    }
    return Promise.resolve("Extension status set")
}

/** Overwrite state to chrome storage */
export function overWriteChromeStorage(state: ContentScriptState, keys: StorableStateKey[], sendDownloadMessage: boolean): void {
    const objectToSave: Partial<ResultLocal> = {}
    if (keys.includes("meetingSoftware")) objectToSave.meetingSoftware = state.meetingSoftware
    if (keys.includes("meetingTitle")) objectToSave.meetingTitle = state.meetingTitle
    if (keys.includes("meetingStartTimestamp")) objectToSave.meetingStartTimestamp = state.meetingStartTimestamp
    if (keys.includes("transcript")) objectToSave.transcript = state.transcript
    if (keys.includes("chatMessages")) objectToSave.chatMessages = state.chatMessages

    chrome.runtime.sendMessage({
        type: sendDownloadMessage ? "meeting_ended" : "capture_update",
        captureId: state.captureId,
        capturePatch: objectToSave,
    } satisfies ExtensionMessage, () => {})
}

/** Persist only to the capture owned by this content script. */
function saveCapturePatch(state: ContentScriptState, patch: Partial<ResultLocal>): void {
    chrome.runtime.sendMessage({ type: "capture_update", captureId: state.captureId, capturePatch: patch } satisfies ExtensionMessage, () => {})
}

/** Attempts to recover last meeting to the best possible extent. */
export function recoverLastMeeting(): Promise<string> {
    return new Promise((resolve, reject) => {
        const message: ExtensionMessage = {
            type: "recover_last_meeting",
        }
        chrome.runtime.sendMessage(message, function (responseUntyped) {
            const response = responseUntyped as ExtensionResponse
            if (response.success) {
                resolve("Last meeting recovered successfully or recovery not needed")
            }
            else {
                reject(response.message)
            }
        })
    })
}

/**
 * Waits for an element, giving up after a deadline.
 *
 * The timeout is the point: without one, a single renamed class or ligature leaves the
 * caller awaiting forever, its .catch() never runs, and the extension reports nothing at
 * all. Every error code this file can raise depends on this rejecting.
 *
 * @param text when given, matches an element whose exact trimmed text equals it, or any
 *             of several alternatives — platforms rename these between states
 * @throws when the element does not appear before timeoutMs
 */
export async function waitForElement(
    selector: string,
    text?: string | string[],
    iframe: HTMLIFrameElement | null = null,
    timeoutMs: number = DEFAULT_ELEMENT_TIMEOUT_MS
): Promise<Element> {
    const wanted = text === undefined ? undefined : (Array.isArray(text) ? text : [text])
    const startedAt = Date.now()

    while (true) {
        // Resolved every pass: an iframe's document is replaced on navigation
        const targetDoc = iframe ? iframe.contentDocument : document

        if (targetDoc) {
            if (wanted) {
                const match = Array.from(targetDoc.querySelectorAll(selector))
                    .find(element => wanted.includes((element.textContent ?? "").trim()))
                if (match) return match
            }
            else {
                const match = targetDoc.querySelector(selector)
                if (match) return match
            }
        }

        if (Date.now() - startedAt > timeoutMs) {
            throw new Error(`Timed out after ${Math.round(timeoutMs / 1000)}s waiting for ${selector}${wanted ? ` ("${wanted.join('" or "')}")` : ""}`)
        }

        await new Promise((resolve) => requestAnimationFrame(resolve))
    }
}

/**
 * Waits until an element matching the selector has the specified computed CSS property value.
 * @param selector The selector to query (e.g., 'div[role="region"]')
 * @param cssProp The camelCase or kebab-case CSS property (e.g., 'containerName')
 * @param cssPropValue The expected value of the CSS property (e.g., 'captions-history')
 */
export async function waitForElementByStyle(selector: string, cssProp: string, cssPropValue: string): Promise<Element> {
    while (true) {
        const elements = Array.from(document.querySelectorAll(selector))
        const matchedElement = elements.find(element => {
            const computedStyle = window.getComputedStyle(element)
            // Cast the string to a valid key type of CSSStyleDeclaration to satisfy the compiler
            return computedStyle[cssProp as keyof CSSStyleDeclaration] === cssPropValue
        })

        if (matchedElement) {
            return matchedElement
        }

        await new Promise((resolve) => requestAnimationFrame(resolve))
    }
}

/**
 * Single, flat polling monitor that handles initial attachment and all re-attachments.
 * @param transcriptMutationCallback platform specific mutation handler
 */
export function startTranscriptMonitor(state: ContentScriptState, transcriptMutationCallback: TranscriptMutationCallback): void {
    state.transcriptTargetNode = null
    let monitorInterval: number | undefined = undefined

    // Call immediately
    transcriptMonitor()
    // Start monitoring
    monitorInterval = setInterval(transcriptMonitor, 2000) as unknown as number

    function transcriptMonitor() {
        if (state.hasMeetingEnded) {
            clearInterval(monitorInterval)
            return
        }

        let activeNode: Element | null | undefined

        switch (state.platform) {
            case "google_meet":
                activeNode = document.querySelector(SELECTORS_GOOGLE_MEET.TRANSCRIPT_REGION)
                break
            case "teams":
                activeNode = document.querySelector(SELECTORS_TEAMS.CAPTIONS_REGION)
                break
            case "zoom": {
                const iframeDOM = (document.querySelector(SELECTORS_ZOOM.IFRAME) as HTMLIFrameElement | null)?.contentDocument
                activeNode = iframeDOM?.querySelector(SELECTORS_ZOOM.TRANSCRIPT_CONTAINER)
                break
            }
            default:
                break
        }

        if (!activeNode) {
            return
        }

        // If the active node is new, replaced, or disconnected, re-attach the observer
        if (!state.transcriptTargetNode || activeNode !== state.transcriptTargetNode || !state.transcriptTargetNode.isConnected) {
            if (!state.transcriptTargetNode) {
                console.log("Captions region detected. Attaching observer...")
            }
            else if (activeNode !== state.transcriptTargetNode) {
                console.log("Captions region replaced. Re-attaching observer...")
            }

            // Flush any in-flight buffer to prevent losing text on transitions
            pushBufferToTranscript(state)
            state.stateTranscriptBlock.personName = ""
            state.stateTranscriptBlock.transcriptTextBuffer = ""
            state.stateTranscriptBlock.timestamp = ""

            if (state.transcriptObserver) {
                state.transcriptObserver.disconnect()
            }

            state.transcriptTargetNode = activeNode
            // Proof that capture actually started, so a short transcript can be told
            // apart from a transcript that never began
            saveCapturePatch(state, { captionsSeen: true })
            state.transcriptObserver = new MutationObserver((mutations) => {
                transcriptMutationCallback(state, mutations)
            })
            state.transcriptObserver.observe(activeNode, mutationConfig)

            // If specified, hide the whole transcript node
            chrome.storage.sync.get(["hideCaptions"], function (resultSyncUntyped) {
                const resultSync = resultSyncUntyped as Partial<ResultSync>
                if ((resultSync.hideCaptions === true) && (state.transcriptTargetNode)) {
                    if (state.platform === "teams") {
                        waitForElement(SELECTORS_TEAMS.CAPTIONS_REGION_WRAPPER).then((element) => {
                            element?.setAttribute("style", `height:40px`)
                        })
                        waitForElement(SELECTORS_TEAMS.CAPTIONS_REGION).then((element) => {
                            element?.setAttribute("style", `opacity:0`)
                        })
                    }
                    else {
                        state.transcriptTargetNode.setAttribute("style", `opacity:0; height:0px`)
                    }
                }
            })
        }
    }
}

/**
 * Last time the in-flight buffer was written to storage. Throttled because this runs on
 * every captions mutation, which can be several times a second.
 */
let lastLiveBufferPersistedAt = 0
const LIVE_BUFFER_PERSIST_INTERVAL = 2000

/**
 * Persists the utterance currently being spoken.
 *
 * A closed tab kills the content script without running the meeting-end routine, so an
 * unflushed buffer would otherwise be lost — and for a single speaker talking without
 * interruption nothing is ever flushed, losing the whole meeting. The background script
 * merges this back in when it picks the meeting up.
 */
function persistLiveBuffer(state: ContentScriptState): void {
    const now = Date.now()
    if (now - lastLiveBufferPersistedAt < LIVE_BUFFER_PERSIST_INTERVAL) {
        return
    }
    lastLiveBufferPersistedAt = now

    const { personName, timestamp, transcriptTextBuffer } = state.stateTranscriptBlock
    if ((personName === "") || (transcriptTextBuffer === "")) {
        return
    }

    const liveTranscriptBlock: TranscriptBlock = {
        personName: personName === "You" ? state.userName : personName,
        timestamp: timestamp,
        transcriptText: transcriptTextBuffer
    }
    saveCapturePatch(state, { liveTranscriptBlock })
}

/** Drops the persisted in-flight buffer, once it is safely part of the transcript. */
export function clearLiveBuffer(state: ContentScriptState): void {
    lastLiveBufferPersistedAt = 0
    saveCapturePatch(state, { liveTranscriptBlock: null })
}

export function broadcastLiveBuffer(state: ContentScriptState): void {
    persistLiveBuffer(state)

    const message: ExtensionMessage = {
        type: "broadcast_live_buffer",
        captureId: state.captureId,
        stateTranscriptBlock: {
            mutationTargetElement: null,
            personName: state.stateTranscriptBlock.personName,
            timestamp: state.stateTranscriptBlock.timestamp,
            transcriptTextBuffer: state.stateTranscriptBlock.transcriptTextBuffer
        }
    }
    chrome.runtime.sendMessage(message, () => { })
}

export function pushBufferToTranscript(state: ContentScriptState): void {
    if ((state.stateTranscriptBlock.personName !== "") && (state.stateTranscriptBlock.transcriptTextBuffer !== "")) {
        state.transcript.push({
            "personName": state.stateTranscriptBlock.personName === "You" ? state.userName : state.stateTranscriptBlock.personName,
            "timestamp": state.stateTranscriptBlock.timestamp,
            "transcriptText": state.stateTranscriptBlock.transcriptTextBuffer
        })
        overWriteChromeStorage(state, ["transcript"], false)
        clearLiveBuffer(state)
    }
}


/** Waits and grabs meeting title from document title */
export function updateMeetingTitle(state: ContentScriptState): void {
    setTimeout(() => {
        // NON CRITICAL DOM DEPENDENCY
        state.meetingTitle = document.title
        overWriteChromeStorage(state, ["meetingTitle"], false)
    }, 5000)
}

export function renderFab(state: ContentScriptState): void {
    const fabCss = `
        position: fixed;
        top: 50%;
        bottom: 50%;
        right: 8px;
        height: 36px;
        width: 36px;
        border-radius: 36px;
        z-index: 100;
        display: flex;
        align-items: center;
        justify-content: center;
        background-color: #4F46E5;
        box-shadow: 0px 0px 4px 0px #2A9ACA;
        cursor: pointer;
        border: none;
        padding: 0;
        overflow: visible;
    `

    const html = document.querySelector("html")
    const fab = document.createElement("button")
    fab.id = "meetingtranscripttoobsidian-fab"
    fab.ariaLabel = "MeetingTranscriptToObsidian"
    fab.title = "MeetingTranscriptToObsidian"
    fab.style.cssText = fabCss

    fab.innerHTML = `
        <div id="fab-main-content" style="display: flex; align-items: center; justify-content: center; width: 100%; height: 100%;">
            <img id="fab-default-logo" src="${LOGO_URL}" alt="MeetingTranscriptToObsidian" draggable="false" style="width: 20px; height: 20px; object-fit: contain;" />
            <span id="fab-letter-mark" style="display: none; color: #ffffff; font-weight: bold; font-size: 16px; text-transform: uppercase; font-family: sans-serif;"></span>
        </div>

        <img id="fab-mini-badge" src="${LOGO_URL}" alt="Active Badge" draggable="false" style="
            display: none;
            position: absolute;
            bottom: -2px;
            right: -2px;
            width: 14px;
            height: 14px;
            border-radius: 50%;
            background-color: #4F46E5;
            box-shadow: 0 0 2px rgba(0,0,0,0.5);
            object-fit: contain;
            pointer-events: none;
        " />
    `

    html?.appendChild(fab)
    makeVerticallyDraggable(fab)

    fab.addEventListener("click", () => {
        const message: ExtensionMessage = { type: "open_side_panel" }
        chrome.runtime.sendMessage(message, () => { })
    })

    updateFabState(state)

}

function makeVerticallyDraggable(fab: HTMLButtonElement): void {
    let isDragging = false
    let startY = 0
    let initialTop = 0
    let hasMoved = false

    const onPointerDown = (e: MouseEvent | TouchEvent) => {
        isDragging = true
        hasMoved = false

        const clientY = "touches" in e ? e.touches[0].clientY : e.clientY
        startY = clientY
        initialTop = fab.getBoundingClientRect().top

        // Attach movement listeners to document so fast drags aren't lost
        document.addEventListener("mousemove", onPointerMove)
        document.addEventListener("mouseup", onPointerUp)
        document.addEventListener("touchmove", onPointerMove, { passive: false })
        document.addEventListener("touchend", onPointerUp)
    }

    const onPointerMove = (e: MouseEvent | TouchEvent) => {
        if (!isDragging) return

        const clientY = "touches" in e ? e.touches[0].clientY : e.clientY
        const deltaY = clientY - startY

        // Threshold (3px) to differentiate click from drag
        if (Math.abs(deltaY) > 3) {
            hasMoved = true
            if (e.cancelable) e.preventDefault() // Prevent scrolling on touch
        }

        let newTop = initialTop + deltaY

        // Bound vertical position inside the visible viewport
        const maxTop = window.innerHeight - fab.offsetHeight
        newTop = Math.max(0, Math.min(newTop, maxTop))

        fab.style.top = `${newTop}px`
    }

    const onPointerUp = () => {
        isDragging = false
        document.removeEventListener("mousemove", onPointerMove)
        document.removeEventListener("mouseup", onPointerUp)
        document.removeEventListener("touchmove", onPointerMove)
        document.removeEventListener("touchend", onPointerUp)
    }

    fab.addEventListener("mousedown", onPointerDown)
    fab.addEventListener("touchstart", onPointerDown, { passive: true })

    // Block the 'click' event if the user dragged the button
    fab.addEventListener("click", (e) => {
        if (hasMoved) {
            e.stopImmediatePropagation()
            e.preventDefault()
            hasMoved = false
        }
    }, true) // Capture phase ensures it runs before the side-panel click handler
}

/** Updates the FAB visual state based on the current transcript data. */
export function updateFabState(state: ContentScriptState): void {
    const fab = document.querySelector("#meetingtranscripttoobsidian-fab")

    const defaultLogo = fab?.querySelector<HTMLElement>("#fab-default-logo")
    const letterMark = fab?.querySelector<HTMLElement>("#fab-letter-mark")
    const miniBadge = fab?.querySelector<HTMLElement>("#fab-mini-badge")

    if (state.transcript && state.transcript.length > 0) {
        const currentSpeaker = state.stateTranscriptBlock?.personName

        if (currentSpeaker && currentSpeaker.trim() !== "") {
            const initial = currentSpeaker.trim().charAt(0)

            // Active Speaker State: Show letter mark + corner badge, hide central logo
            if (defaultLogo) defaultLogo.style.display = "none"
            if (letterMark) {
                letterMark.textContent = initial
                letterMark.style.display = "inline"
            }
            if (miniBadge) miniBadge.style.display = "block"
            return
        }
    }

    // Default State: Fallback to central logo, hide mark + badge
    if (defaultLogo) defaultLogo.style.display = "block"
    if (letterMark) letterMark.style.display = "none"
    if (miniBadge) miniBadge.style.display = "none"
}

export function unmountFab(): void {
    const fab = document.querySelector("#meetingtranscripttoobsidian-fab")
    if (fab) {
        fab.remove()
    }
}


/** How many recent errors to keep. Enough to see a pattern, small enough to stay cheap. */
const MAX_RECENT_ERRORS = 20

/**
 * Records a capture failure locally and surfaces it on the meetings page.
 *
 * Replaces the upstream project's remote error reporting, which sent this fork's
 * failures to a third party. Kept on the device so a user whose transcript broke has
 * something concrete to look at, without anything leaving the machine.
 */
export function logError(state: ContentScriptState, code: string, err: unknown): void {
    const message = err instanceof Error ? err.message : String(err)
    console.error(`[MeetingTranscriptToObsidian] error ${code} on ${state.meetingSoftware || "unknown platform"}:`, err)

    chrome.storage.local.get(["recentErrors"], function (resultUntyped) {
        const result = resultUntyped as Partial<ResultLocal>
        const entry: RecentError = {
            at: new Date().toISOString(),
            code: code,
            meetingSoftware: state.meetingSoftware,
            message: message.slice(0, 300)
        }
        const recentErrors = [...(result.recentErrors ?? []), entry].slice(-MAX_RECENT_ERRORS)
        chrome.storage.local.set({ recentErrors })
    })
}

/** Checks if the installed extension version meets the minimum required version. */
export function meetsMinVersion(oldVer: string, newVer: string): boolean {
    const oldParts = oldVer.split('.')
    const newParts = newVer.split('.')
    for (let i = 0; i < newParts.length; i++) {
        const a = ~~newParts[i] // parse int
        const b = ~~oldParts[i] // parse int
        if (a > b) return false
        if (a < b) return true
    }
    return true
}

