import type { ChatMessage, ContentScriptState, ExtensionStatusJSON } from "@lib/types"
import { LOGO_URL, commonCSS } from "../common-config"
import { overWriteChromeStorage, waitForElement } from "../common-utils"
import { SELECTORS_GOOGLE_MEET } from "./config"

/**
 * Returns all elements of the specified selector type whose text is exactly `text`.
 *
 * Matching is exact rather than a substring test: callers pass Google Symbols ligature
 * names, and an unanchored match makes "chat" also select "chat_bubble", so the first
 * result could be the wrong control entirely. waitForElement already compares exactly,
 * so this keeps the two in agreement.
 */
export function selectElements(selector: string, text: string): HTMLElement[] {
    const elements = document.querySelectorAll<HTMLElement>(selector)
    return Array.prototype.filter.call(elements, function (element: HTMLElement) {
        return (element.textContent ?? "").trim() === text
    })
}

/**
 * Returns the button that actually toggles a control, given its Google Symbols ligature.
 *
 * Meet renders more than one element per ligature — "chat" matches both an offscreen
 * copy with no button ancestor and the real toolbar control. Taking the first match
 * therefore clicked a hidden, non-clickable element and the chat panel never opened,
 * so no chat messages were ever rendered to be captured.
 */
export function selectControlButton(selector: string, text: string): HTMLElement | undefined {
    const matches = selectElements(selector, text)
    const clickable = matches
        .map(element => element.closest("button") as HTMLElement | null)
        .filter((button): button is HTMLElement => button !== null)

    return clickable.find(button => button.getBoundingClientRect().width > 0)
        ?? clickable[0]
        ?? matches[0]
}

/** How long after joining the extension keeps trying to get captions on */
const CAPTIONS_ENFORCE_WINDOW_MS = 2 * 60 * 1000
const CAPTIONS_CHECK_INTERVAL_MS = 1500
/** Attempts of a stuck button click before also trying Meet's "c" shortcut */
const CAPTIONS_CLICKS_BEFORE_SHORTCUT = 2

/** The visible toolbar button whose icon currently shows `ligature`, if any */
function visibleButtonWithIcon(ligature: string): HTMLElement | undefined {
    return selectElements(SELECTORS_GOOGLE_MEET.GOOGLE_SYMBOLS, ligature)
        .map(element => element.closest("button") as HTMLElement | null)
        .filter((button): button is HTMLElement => button !== null)
        .find(button => button.getBoundingClientRect().width > 0)
}

const CAPTIONS_ICONS = [
    SELECTORS_GOOGLE_MEET.TEXT_CAPTIONS_OFF,
    SELECTORS_GOOGLE_MEET.TEXT_CAPTIONS_ON,
    SELECTORS_GOOGLE_MEET.TEXT_CAPTIONS_TRANSLATE
]

/**
 * Current captions state, and the button that toggles them.
 *
 * Meet has two toolbars. The classic one swaps the CC icon between
 * "closed_caption_off" and "closed_caption". In meetings with translated speech the
 * control is a "Captions and translation" button whose icon is always "translate_cc",
 * so its state can only be read from aria-pressed.
 */
export function getCaptionsState(): { on: boolean, button: HTMLElement } | undefined {
    const translateButton = visibleButtonWithIcon(SELECTORS_GOOGLE_MEET.TEXT_CAPTIONS_TRANSLATE)
    if (translateButton) {
        return { on: translateButton.getAttribute("aria-pressed") === "true", button: translateButton }
    }
    const onButton = visibleButtonWithIcon(SELECTORS_GOOGLE_MEET.TEXT_CAPTIONS_ON)
    if (onButton) return { on: true, button: onButton }
    const offButton = visibleButtonWithIcon(SELECTORS_GOOGLE_MEET.TEXT_CAPTIONS_OFF)
    if (offButton) return { on: false, button: offButton }
    return undefined
}

/**
 * Turns captions on and keeps checking until Meet actually shows them on.
 *
 * Meetings with translated speech use the "translate_cc" toggle (see getCaptionsState);
 * before it was recognised, the extension never found a captions button in those
 * meetings and captions stayed off. A single click is also not verified by Meet in any
 * way — the button can be disabled or re-rendered at the moment it lands — so this
 * re-checks the state on an interval and clicks again whenever captions are off,
 * until they have stayed on through two consecutive checks or the window ends. It
 * stops for good as soon as the user clicks the captions button themselves
 * (a trusted click), so turning captions off by hand is respected.
 */
export function ensureCaptionsOn(state: ContentScriptState): void {
    const startedAt = Date.now()
    let userTookOver = false
    let attempts = 0
    let confirmedChecks = 0

    // A real user click on the CC control means they are managing captions themselves
    function onUserClick(event: MouseEvent) {
        if (!event.isTrusted) return
        const target = event.target as Element | null
        const button = target?.closest("button")
        if (!button) return
        const icon = (button.querySelector(SELECTORS_GOOGLE_MEET.GOOGLE_SYMBOLS)?.textContent ?? "").trim()
        if (CAPTIONS_ICONS.includes(icon)) {
            userTookOver = true
        }
    }
    document.addEventListener("click", onUserClick, true)

    function stop(reason: string) {
        clearInterval(interval)
        document.removeEventListener("click", onUserClick, true)
        console.log(`Captions enforcer stopped: ${reason}`)
    }

    function check() {
        if (state.hasMeetingEnded) return stop("meeting ended")
        if (userTookOver) return stop("user toggled captions")
        if (Date.now() - startedAt > CAPTIONS_ENFORCE_WINDOW_MS) {
            return stop(confirmedChecks > 0 ? "captions on" : "gave up, captions still off")
        }

        const captions = getCaptionsState()
        if (!captions) return  // control is being re-rendered; look again next tick

        if (captions.on) {
            // Require captions to survive two consecutive checks: translated speech
            // setup can switch them off again right after they first come on
            confirmedChecks++
            if (confirmedChecks >= 2) stop("captions on")
            return
        }
        confirmedChecks = 0

        const offButton = captions.button
        if ((offButton as HTMLButtonElement).disabled || offButton.getAttribute("aria-disabled") === "true") {
            return  // not ready yet, a click now would be swallowed
        }

        attempts++
        console.log(`Turning captions on (attempt ${attempts})`)
        offButton.click()

        // If clicks keep failing to stick, also try Meet's own keyboard shortcut
        if (attempts > CAPTIONS_CLICKS_BEFORE_SHORTCUT) {
            for (const type of ["keydown", "keyup"]) {
                document.body.dispatchEvent(new KeyboardEvent(type, {
                    key: "c", code: "KeyC", keyCode: 67, which: 67, bubbles: true, cancelable: true
                }))
            }
        }
    }

    const interval = setInterval(check, CAPTIONS_CHECK_INTERVAL_MS)
    check()
}

/** Shows a responsive notification of specified type and message */
export function showNotificationGoogleMeet(extensionStatusJSON: ExtensionStatusJSON): void {
    // Banner CSS
    const html = document.querySelector("html")
    const obj = document.createElement("div")
    const logo = document.createElement("img")
    const text = document.createElement("p")

    logo.setAttribute(
        "src",
        LOGO_URL
    )
    logo.setAttribute("height", "32px")
    logo.setAttribute("width", "32px")
    logo.style.cssText = "border-radius: 4px"

    // Remove banner after 5s
    setTimeout(() => {
        obj.style.display = "none"
    }, 5000)

    if (extensionStatusJSON.status === 200) {
        obj.style.cssText = `color: #2A9ACA; top: 5%; ${commonCSS}`
        text.innerHTML = extensionStatusJSON.message
    }
    else {
        obj.style.cssText = `color: orange; top: 5%; ${commonCSS}`
        text.innerHTML = extensionStatusJSON.message
    }

    obj.prepend(text)
    obj.prepend(logo)
    if (html)
        html.append(obj)
}


/**
 * Attempt to get username before meeting starts. Abort interval if valid username is
 * found or if meeting starts and default to "You".
 */
export function captureUserName(state: ContentScriptState): void {
    waitForElement(SELECTORS_GOOGLE_MEET.USER_NAME).then(() => {
        // Poll the element until the textContent loads from network or until meeting starts
        const captureUserNameInterval = setInterval(() => {
            if (state.isUserNameLocked) {
                clearInterval(captureUserNameInterval)
                return
            }
            if (!state.hasMeetingStarted) {
                const capturedUserName = document.querySelector(SELECTORS_GOOGLE_MEET.USER_NAME)?.textContent
                if (capturedUserName) {
                    state.userName = capturedUserName
                    clearInterval(captureUserNameInterval)
                }
            }
            else {
                clearInterval(captureUserNameInterval)
            }
        }, 100)
    })
}

/**
 * Stores a chat message.
 *
 * De-duplication happens on data-message-id in the caller, so two identical messages
 * from the same person are both kept — collapsing them lost real content.
 */
export function pushChatBlock(state: ContentScriptState, chatBlock: ChatMessage): void {
    state.chatMessages.push(chatBlock)
    overWriteChromeStorage(state, ["chatMessages"], false)
}
