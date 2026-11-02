import type { ErrorObject, LiveSummary, Meeting, ResultLocal, ResultSync } from "@lib/types"
import { readLiveMeeting } from "./meetings"
import { getTranscriptString } from "./utils"
import { downloadSummary } from "./exporters"
import { postSummaryToObsidian } from "./obsidian"
import { END_MARKER, isAnswerComplete, isStableAnswer, pickAnswerText, stripEndMarker } from "@lib/summary-text"

const CHATGPT_URL = "https://chatgpt.com/"
const CHATGPT_ORIGIN = "https://chatgpt.com/*"

/** ChatGPT refuses very long prompts, and a full day of captions can be enormous. */
const MAX_TRANSCRIPT_CHARS = 100000

/** How long to wait for the whole answer to finish streaming. */
const RESPONSE_TIMEOUT_MS = 180000

/**
 * Consecutive unchanged polls before a marker-less answer is considered complete.
 *
 * Deliberately long. This is the fallback for when the model omits the end marker, and
 * it is the only thing standing between a mid-stream pause and a summary that gets
 * stored half-written: ChatGPT streams in bursts, so a few seconds of unchanged text
 * proves nothing. Eight seconds of no new text *and* no stop button does.
 */
const STABLE_POLLS_REQUIRED = 8

/**
 * Consecutive failed reads before the run is abandoned.
 *
 * A tab that has stopped answering will not start again, and continuing to poll it for
 * the full timeout only delays the error while risking a stale half-answer being
 * mistaken for the finished one.
 */
const MAX_FAILED_READS = 10

const POLL_INTERVAL_MS = 1000

/** How long to wait for the ChatGPT tab to finish loading before trying anyway. */
const LOAD_TIMEOUT_MS = 20000

/**
 * Every ChatGPT DOM hook this module depends on, in one place.
 *
 * Each entry lists the current markup and the one before it, so a staged rollout —
 * where some accounts still get the old UI — works either way. The September 2026
 * redesign dropped `#prompt-textarea`, the send/stop `data-testid`s and
 * `data-message-author-role` all at once, and every summary failed at the composer.
 *
 * The injected functions run in the page and cannot see module scope, so this object
 * is passed to them as an argument.
 */
const CHATGPT_SELECTORS = {
    composer: '#prompt-textarea, div.ProseMirror[contenteditable="true"]',
    sendButton: 'button[data-testid="send-button"], button[aria-label="Send"], button[type="submit"]',
    stopButton: 'button[data-testid="stop-button"], button[aria-label="Stop"], button[aria-label^="Stop"]',
    userMessage: '[data-message-author-role="user"], [data-chatgpt-search-unit-key$=":user"]',
    assistantMessage: '[data-message-author-role="assistant"], [data-chatgpt-search-unit-key$=":assistant"]',
    codeBlock: 'pre code, [data-markdown-copy="code-block"] [dir="ltr"]',
    login: 'a[href*="/auth/login"], button[data-testid="login-button"]'
}
type ChatGptSelectors = typeof CHATGPT_SELECTORS

/** The instructions used when the user has not written their own. */
export const DEFAULT_SUMMARY_INSTRUCTIONS = [
    "Summarise the following meeting transcript. The transcript comes from automatic captions, so expect some transcription errors and ignore them.",
    "",
    "Reply with a single fenced code block containing markdown, and nothing outside it.",
    `The very last line inside the block must be exactly ${END_MARKER}`,
    "Inside the block use exactly these sections:",
    "",
    "## Summary",
    "Two to four sentences.",
    "",
    "## Key points",
    "Short bullets.",
    "",
    "## Decisions",
    'Short bullets, or "None recorded".',
    "",
    "## Action items",
    'Bullets shaped "- [ ] Owner — task", or "None recorded".',
].join("\n")

export function buildSummaryPrompt(meeting: Meeting, instructions?: string): string {
    const title = meeting.meetingTitle || meeting.title || "Meeting"
    const participants = [...new Set(meeting.transcript.map(block => block.personName))].filter(name => name.trim())

    let transcript = getTranscriptString(meeting.transcript)
    if (transcript.length > MAX_TRANSCRIPT_CHARS) {
        transcript = transcript.slice(0, MAX_TRANSCRIPT_CHARS) + "\n\n[transcript truncated]"
    }

    // The end marker is how completion is detected, so it is appended even to custom
    // instructions rather than trusting the user to include it.
    const chosen = (instructions ?? "").trim() || DEFAULT_SUMMARY_INSTRUCTIONS
    const withMarker = chosen.includes(END_MARKER)
        ? chosen
        : `${chosen}\n\nEnd your reply with a final line containing exactly ${END_MARKER}`

    return [
        withMarker,
        "",
        `Meeting: ${title}`,
        `Date: ${new Date(meeting.meetingStartTimestamp).toLocaleString()}`,
        participants.length > 0 ? `Participants: ${participants.join(", ")}` : "",
        "",
        "Transcript:",
        "",
        transcript
    ].filter(line => line !== "").join("\n")
}

/**
 * How long the ChatGPT tab stays in front if the answer never starts streaming. Bounds
 * the interruption when something is wrong (signed out, rate limited) rather than
 * leaving a foreign tab in the user's face.
 */
const FOREGROUND_GRACE_MS = 20000

/** The ChatGPT tab, plus what was in front before it, so focus can be handed back. */
interface ChatGptTab {
    tabId: number
    previousTabId: number | undefined
    previousWindowId: number | undefined
}

function noAccessError(): ErrorObject {
    return { errorCode: "020", errorMessage: "Could not open ChatGPT. Grant access to chatgpt.com and make sure you are signed in." }
}

/** The tab opened, but the prompt could not be sent; the reason is read off the page. */
function notSubmittedError(reason: string): ErrorObject {
    return { errorCode: "020", errorMessage: `Could not send the prompt to ChatGPT \u2014 ${reason}` }
}

/**
 * Opens a dedicated ChatGPT tab for this summary and waits until it has finished loading.
 *
 * A new tab is used rather than the user's existing ChatGPT tab on purpose: navigating
 * that tab to a fresh conversation would throw away whatever they had open in it.
 *
 * The tab is opened **in front**, which a background tab cannot substitute for: the
 * composer is a ProseMirror editor and `execCommand("insertText")` needs the document
 * itself focused, which only the active tab of a focused window has. Opened hidden, the
 * insert silently does nothing and the run only completes if the user switches to the
 * tab by hand. Whatever was in front is recorded here so `restoreFocus` can put it back
 * as soon as the answer starts streaming.
 */
function openChatGptTab(): Promise<ChatGptTab> {
    return new Promise((resolve, reject) => {
        chrome.permissions.contains({ origins: [CHATGPT_ORIGIN] }).then(async (hasPermission) => {
            if (!hasPermission) {
                reject(noAccessError() satisfies ErrorObject)
                return
            }

            const [previous] = await chrome.tabs.query({ active: true, lastFocusedWindow: true }).catch(() => [])

            chrome.tabs.create({ url: CHATGPT_URL, active: true }, async function (tab) {
                const tabId = tab.id
                if (typeof tabId !== "number") {
                    reject(noAccessError() satisfies ErrorObject)
                    return
                }

                const opened: ChatGptTab = {
                    tabId: tabId,
                    previousTabId: previous?.id,
                    previousWindowId: previous?.windowId
                }

                // Chrome discards hidden tabs under memory pressure, and this one spends
                // most of its life hidden behind however many tabs the user has open.
                // A discarded tab stops answering executeScript, which looks exactly like
                // an answer that stopped growing.
                chrome.tabs.update(tabId, { autoDiscardable: false }).catch(() => { })

                // tabs.create returns as soon as the tab exists, long before the page is
                // usable, so injecting straight away would race the navigation.
                //
                // The wait polls rather than listening for tabs.onUpdated, because the
                // automatic run has no open message port holding the service worker up:
                // a chrome.* call each second is what keeps the worker past its idle
                // timeout, where a listener plus one long setTimeout could let it be
                // torn down mid-summary. The listener would also be lost if that
                // happened, so the run would hang rather than fail.
                for (let waited = 0; waited < LOAD_TIMEOUT_MS; waited += POLL_INTERVAL_MS) {
                    const current = await chrome.tabs.get(tabId).catch(() => undefined)
                    if (!current) {
                        // Tab closed under us — nothing left to drive
                        reject(noAccessError() satisfies ErrorObject)
                        return
                    }
                    if (current.status === "complete") {
                        resolve(opened)
                        return
                    }
                    await new Promise(r => setTimeout(r, POLL_INTERVAL_MS))
                }

                // Out of time, but the composer check in submitPrompt is the real gate
                resolve(opened)
            })
        })
    })
}

/**
 * Puts back whatever the user was looking at before the ChatGPT tab appeared.
 *
 * Safe to call more than once and safe when the previous tab has since been closed —
 * the meeting tab is a likely previous tab and is often gone by now, and failing to
 * restore focus must never fail the summary.
 */
async function restoreFocus(tab: ChatGptTab): Promise<void> {
    if (typeof tab.previousTabId !== "number") {
        return
    }
    await chrome.tabs.update(tab.previousTabId, { active: true }).catch(() => { })
    if (typeof tab.previousWindowId === "number") {
        await chrome.windows.update(tab.previousWindowId, { focused: true }).catch(() => { })
    }
}

/**
 * Runs fn inside the ChatGPT tab, in the page's own JavaScript world.
 *
 * MAIN world is required, not a preference: the composer is a ProseMirror editor, and
 * from an isolated world execCommand mutates the DOM without ProseMirror's model picking
 * it up. The editor then still reports empty, and because the send button is a
 * type="submit" inside a form, clicking it performs a native GET navigation to
 * "chatgpt.com/?prompt-textarea=" that destroys the page instead of sending anything.
 */
async function runInTab<T>(tabId: number, func: (...args: never[]) => T, args: unknown[] = []): Promise<T | undefined> {
    const results = await chrome.scripting.executeScript({
        target: { tabId },
        world: "MAIN",
        func: func as (...args: unknown[]) => T,
        args: args
    }).catch((error) => {
        console.error("ChatGPT injection failed:", error)
        return null
    })
    return results?.[0]?.result as T | undefined
}

/**
 * Types the prompt into the composer and sends it.
 *
 * The composer is a ProseMirror contenteditable, so the text has to be inserted with
 * execCommand: assigning textContent leaves ProseMirror's internal model empty and the
 * send button never becomes available.
 */
type SubmitOutcome = "sent" | "no-composer" | "not-inserted" | "no-send-button"

async function submitPrompt(tabId: number, prompt: string): Promise<SubmitOutcome> {
    // Wait for the composer, which appears well after load on a cold tab
    let ready = false
    for (let attempt = 0; attempt < 40; attempt++) {
        ready = !!await runInTab(tabId, (selectors: ChatGptSelectors) => !!document.querySelector(selectors.composer), [CHATGPT_SELECTORS])
        if (ready) break
        await new Promise(resolve => setTimeout(resolve, 500))
    }
    if (!ready) return "no-composer"

    const inserted = await runInTab(tabId, (text: string, selectors: ChatGptSelectors) => {
        const composer = document.querySelector<HTMLElement>(selectors.composer)
        if (!composer) return false
        composer.focus()
        return document.execCommand("insertText", false, text)
    }, [prompt, CHATGPT_SELECTORS])

    if (!inserted) return "not-inserted"

    // The send button is rendered only once the composer has content, one frame later.
    // The composer is re-checked on every attempt because clicking submit while it is
    // empty triggers a native form GET that navigates the page away.
    for (let attempt = 0; attempt < 20; attempt++) {
        await new Promise(resolve => setTimeout(resolve, 250))
        const clicked = await runInTab(tabId, (selectors: ChatGptSelectors) => {
            const composer = document.querySelector<HTMLElement>(selectors.composer)
            if (!composer || !(composer.innerText || "").trim()) return false

            // Scoped to the composer's own form: the bare submit fallback must never
            // match a button belonging to some other form on the page.
            const scope = composer.closest("form") ?? document
            const button = scope.querySelector<HTMLButtonElement>(selectors.sendButton)
            if (!button || button.disabled) return false
            button.click()
            return true
        }, [CHATGPT_SELECTORS])
        if (clicked) return "sent"
    }
    return "no-send-button"
}

/**
 * Reads why a run produced nothing, off the page, while the page is still open.
 *
 * This is what makes closing the tab on failure safe. The tab used to be left behind
 * so the cause could be seen, but the automatic run has nobody watching it: every
 * failure left another ChatGPT tab open, and the reason still had to be guessed. The
 * signals are read here instead and travel with the error.
 */
async function describeFailure(tabId: number): Promise<string> {
    const found = await runInTab(tabId, (selectors: ChatGptSelectors) => {
        const text = document.body.textContent || ""
        const signedOut = !!document.querySelector(selectors.login)
        const alerts = [...document.querySelectorAll<HTMLElement>('[role="alert"], [class*="text-token-text-error"]')]
            .map(node => (node.textContent || "").trim())
            .filter(message => message.length > 0)
        return {
            signedOut: signedOut,
            sentPrompt: document.querySelectorAll(selectors.userMessage).length > 0,
            gotAnswer: document.querySelectorAll(selectors.assistantMessage).length > 0,
            alert: alerts[0]?.slice(0, 160) ?? "",
            rateLimited: /reached (our|the) limit|too many requests|usage limit/i.test(text)
        }
    }, [CHATGPT_SELECTORS])

    if (!found) return "the ChatGPT tab could not be read"
    if (found.signedOut) return "you are signed out of ChatGPT"
    if (found.rateLimited) return "ChatGPT reported a usage limit"
    if (found.alert) return `ChatGPT said: ${found.alert}`
    if (!found.sentPrompt) return "the prompt never reached the composer"
    if (!found.gotAnswer) return "ChatGPT never started an answer"
    return "the answer never finished"
}

/**
 * Polls until the answer stops changing.
 *
 * `onStreaming` fires once, as soon as there is answer text in the page. That is the
 * point the tab no longer needs to be in front: the request is in flight and the rest
 * of the work is reading the DOM from the service worker, which does not care whether
 * the tab is visible.
 */
interface AnswerReading {
    text: string
    /**
     * Whether the answer is known to have finished. An unfinished answer is never a
     * summary: storing one produced a note that stopped mid-sentence and still reported
     * success, which is worse than reporting nothing.
     */
    complete: boolean
    /** Why it did not finish, when it did not. */
    reason?: string
}

async function readResponse(tabId: number, onStreaming: () => void): Promise<AnswerReading> {
    const startedAt = Date.now()
    let lastText = ""
    let stablePolls = 0
    let notifiedStreaming = false
    let failedReads = 0
    let sawGenerating = false

    function noteStreaming() {
        if (!notifiedStreaming) {
            notifiedStreaming = true
            onStreaming()
        }
    }

    while (Date.now() - startedAt < RESPONSE_TIMEOUT_MS) {
        await new Promise(resolve => setTimeout(resolve, POLL_INTERVAL_MS))

        // Nothing has appeared and the grace period is up — hand focus back anyway
        // rather than sitting on the user's screen while a failure plays out.
        if (Date.now() - startedAt > FOREGROUND_GRACE_MS) {
            noteStreaming()
        }

        // The page hands back both readings unchosen; picking between them happens in
        // pickAnswerText, where it can be tested. textContent rather than innerText
        // because it needs no layout, and this tab is deliberately not in front.
        const raw = await runInTab(tabId, (selectors: ChatGptSelectors) => {
            const messages = document.querySelectorAll<HTMLElement>(selectors.assistantMessage)
            const last = messages[messages.length - 1]
            const codeBlock = last?.querySelector<HTMLElement>(selectors.codeBlock)
            return {
                codeBlockText: codeBlock?.textContent ?? "",
                messageText: last?.textContent ?? "",
                busy: !!document.querySelector(selectors.stopButton)
            }
        }, [CHATGPT_SELECTORS])

        if (!raw) {
            failedReads++
            if (failedReads >= MAX_FAILED_READS) {
                return { text: lastText, complete: false, reason: "the ChatGPT tab stopped responding" }
            }
            continue
        }
        failedReads = 0
        if (raw.busy) {
            sawGenerating = true
        }

        const reading = {
            text: pickAnswerText(raw.codeBlockText, raw.messageText),
            busy: raw.busy
        }

        if (reading.text) {
            noteStreaming()
        }

        // Definitive: the model has emitted its final line.
        if (isAnswerComplete(reading.text)) {
            return { text: stripEndMarker(reading.text), complete: true }
        }

        // Fallback for when the model omits the marker. The stop button disappears a beat
        // before the answer text is in the DOM, so an empty reading is never treated as a
        // finished answer, and the text must hold still for several polls.
        //
        // `sawGenerating` is required too: without it a pause before generation has even
        // started looks exactly like a finished answer, and the first burst of an answer
        // gets stored as the whole of it.
        if (isStableAnswer({ text: reading.text, previousText: lastText, busy: reading.busy, sawGenerating: sawGenerating })) {
            stablePolls++
            if (stablePolls >= STABLE_POLLS_REQUIRED) {
                return { text: reading.text, complete: true }
            }
        }
        else {
            stablePolls = 0
            lastText = reading.text
        }
    }

    // Out of time. Whatever arrived is half an answer, and half an answer is not a
    // summary — it is a note that stops mid-sentence. The caller reports it as failed.
    return { text: lastText, complete: false, reason: "the answer never finished" }
}

/**
 * Summarises a meeting by handing the transcript to the ChatGPT web app and reading the
 * answer back out of the page.
 *
 * This drives ChatGPT's UI rather than an API, so it depends on that page's DOM and will
 * need the selectors revisited whenever OpenAI restructures the composer or message list.
 *
 * @throws error codes: 010, 020, 021, 022
 */
export function summarizeMeeting(index: number): Promise<string> {
    return new Promise((resolve, reject) => {
        chrome.storage.local.get(["meetings"], async function (resultLocalUntyped) {
            const resultLocal = resultLocalUntyped as Partial<ResultLocal>
            const meetings = resultLocal.meetings

            if (!meetings || !meetings[index]) {
                reject({ errorCode: "010", errorMessage: "Meeting at specified index not found" } satisfies ErrorObject)
                return
            }

            try {
                const summary = await runChatGptSummary(meetings[index])

                // Re-read: the run takes minutes, and the list may have changed meanwhile
                // (a meeting ended and was filed). Write back to the same meeting by its
                // start time rather than trusting a stale array and index.
                await updateStoredMeeting(meetings[index], meeting => {
                    meeting.summary = summary
                    meeting.summaryStatus = "successful"
                })
                resolve(summary)
            }
            catch (error) {
                await updateStoredMeeting(meetings[index], meeting => {
                    meeting.summaryStatus = "failed"
                })
                reject(error)
            }
        })
    })
}

/** Applies a change to the stored meeting with this start time, if it is still there. */
function updateStoredMeeting(identity: Pick<Meeting, "captureId" | "meetingStartTimestamp">, change: (meeting: Meeting) => void): Promise<void> {
    return new Promise(resolve => {
        chrome.storage.local.get(["meetings"], (resultUntyped) => {
            const meetings = (resultUntyped as Partial<ResultLocal>).meetings ?? []
            const meeting = meetings.find(m => identity.captureId ? m.captureId === identity.captureId : m.meetingStartTimestamp === identity.meetingStartTimestamp)
            if (!meeting) {
                resolve()
                return
            }
            change(meeting)
            chrome.storage.local.set({ meetings }, () => resolve())
        })
    })
}

/**
 * Summarises the meeting still in progress from its transcript so far, and keeps the
 * result as `liveSummary` so the side panel shows it — and still shows it if the panel
 * is closed and reopened. The meeting itself is untouched and is filed as usual when it
 * ends.
 *
 * @throws error codes: 020, 021, 022, 023
 */
export async function summarizeLiveMeeting(): Promise<LiveSummary> {
    const meeting = await readLiveMeeting()
    const summary = await runChatGptSummary(meeting)
    const liveSummary: LiveSummary = {
        meetingStartTimestamp: meeting.meetingStartTimestamp,
        summary,
        createdAt: new Date().toISOString(),
        blockCount: meeting.transcript.length,
    }
    await new Promise<void>(resolve => chrome.storage.local.set({ liveSummary }, () => resolve()))
    return liveSummary
}

/**
 * ChatGPT runs one at a time. Each one brings a tab to the front, so two at once —
 * a summary asked for mid-meeting still running as the meeting ends and the automatic
 * one starts — would fight over focus and could each read the other's answer.
 */
let chatGptQueue: Promise<unknown> = Promise.resolve()

/**
 * Sends a meeting's transcript to ChatGPT and returns the finished summary. Storage is
 * the caller's business; this only drives the tab.
 *
 * @throws error codes: 020, 021, 022
 */
function runChatGptSummary(meeting: Meeting): Promise<string> {
    const run = chatGptQueue.then(() => driveChatGpt(meeting))
    chatGptQueue = run.catch(() => { })
    return run
}

async function driveChatGpt(meeting: Meeting): Promise<string> {
    if (meeting.transcript.length === 0) {
        throw { errorCode: "022", errorMessage: "Nothing to summarise for this meeting" } satisfies ErrorObject
    }

    let tab: ChatGptTab | undefined
    try {
        const instructions = await new Promise<string | undefined>(resolve =>
            chrome.storage.sync.get(["summaryPrompt"], (r) =>
                resolve((r as Partial<ResultSync>).summaryPrompt)))

        tab = await openChatGptTab()
        const submitted = await submitPrompt(tab.tabId, buildSummaryPrompt(meeting, instructions))

        if (submitted === "no-composer") {
            // Signed out and a changed page both look like "no composer"; the
            // page says which, and a changed page is what to fix in CHATGPT_SELECTORS.
            const reason = await describeFailure(tab.tabId)
            throw notSubmittedError(reason === "you are signed out of ChatGPT"
                ? reason
                : "the message box was not found (ChatGPT's page may have changed)") satisfies ErrorObject
        }
        if (submitted === "not-inserted") {
            throw notSubmittedError("the transcript could not be typed into the message box") satisfies ErrorObject
        }
        if (submitted === "no-send-button") {
            throw notSubmittedError("the send button was not found or stayed disabled") satisfies ErrorObject
        }

        const openedTab = tab
        const answer = await readResponse(openedTab.tabId, () => { restoreFocus(openedTab) })

        if (!answer.complete) {
            // A partial answer is reported, never stored. Storing one is how a
            // summary that stopped mid-word came to be filed as successful.
            const reason = answer.reason ?? await describeFailure(openedTab.tabId)
            const got = answer.text
                ? ` (${answer.text.length} characters arrived before it stopped)`
                : ` (${await describeFailure(openedTab.tabId)})`
            throw { errorCode: "021", errorMessage: `ChatGPT did not finish the summary — ${reason}${got}` } satisfies ErrorObject
        }

        if (!answer.text) {
            const reason = await describeFailure(openedTab.tabId)
            throw { errorCode: "021", errorMessage: `ChatGPT did not return a summary — ${reason}` } satisfies ErrorObject
        }
        return answer.text
    }
    catch (error) {
        const parsedError = error as ErrorObject
        throw parsedError.errorCode
            ? parsedError
            : { errorCode: "021", errorMessage: String(error) } satisfies ErrorObject
    }
    finally {
        // Always, on every path: the user gets their place back and the tab
        // goes away. Anything else leaves a ChatGPT tab behind and the screen
        // parked on someone else's page.
        if (tab) {
            await restoreFocus(tab)
            chrome.tabs.remove(tab.tabId).catch(() => { })
        }
    }
}

/**
 * Summarises a meeting and puts the result everywhere a summary belongs: the summary
 * text file and the Obsidian summary folder, each subject to its own setting.
 *
 * Saving the file and filing the note are **bonuses**, not part of the result. By the
 * time either runs the summary is already in storage, so neither failing may report the
 * summarise itself as failed — a closed Obsidian is the ordinary case, not an error
 * worth losing a summary over.
 *
 * Both the Summarise button and the automatic run after a meeting go through here, so
 * a summary is filed the same way however it was triggered.
 *
 * @throws error codes: 010, 020, 021, 022
 */
export async function summarizeAndFileMeeting(index: number): Promise<string> {
    const summary = await summarizeMeeting(index)

    const resultSync = await new Promise<Partial<ResultSync>>(resolve =>
        chrome.storage.sync.get(["obsidianEnabled", "autoDownloadFileAfterMeeting"], (r) =>
            resolve(r as Partial<ResultSync>)))

    const followUps: Promise<unknown>[] = []
    if (resultSync.autoDownloadFileAfterMeeting !== false) {
        followUps.push(downloadSummary(index)
            .catch((error) => console.error("Could not save the summary file:", error)))
    }
    if (resultSync.obsidianEnabled) {
        followUps.push(postSummaryToObsidian(index)
            .catch((error) => console.error("Could not save summary to Obsidian:", error)))
    }
    await Promise.all(followUps)

    return summary
}
