import type { ErrorObject, ExtensionMessage, ExtensionResponse, LiveSummary, ResultLocal, StateTranscriptBlock, TranscriptBlock } from "@lib/types"
import { buildSummaryDocument } from "@lib/summary-render"

const meetingTitle = document.querySelector<HTMLElement>("#meeting-title")
const container = document.querySelector<HTMLElement>("#transcript-container")
const SCROLL_THRESHOLD = 50
const LIVE_BLOCK_ID = 'live-transcript-placeholder'

// Make the element editable
if (meetingTitle) {
    meetingTitle.setAttribute("contenteditable", "true")

    // Helper function to save current title
    const saveTitle = () => {
        const titleText = meetingTitle.innerText.trim()
        chrome.storage.local.set({ meetingTitle: titleText })
    }

    // Save only when focus leaves the element
    meetingTitle.addEventListener("blur", saveTitle)

    // Prevent line breaks when pressing 'Enter'
    meetingTitle.addEventListener("keydown", (e) => {
        if (e.key === "Enter") {
            e.preventDefault()
            meetingTitle.blur()
        }
    })
}

chrome.storage.local.get(["meetingTitle", "transcript"], (resultUntyped) => {
    const result = resultUntyped as Partial<ResultLocal>
    if (result.meetingTitle) {
        const titleText = result.meetingTitle
        if (titleText && meetingTitle) {
            console.log(result.meetingTitle)
            meetingTitle.innerHTML = titleText
        }
    }

    if (result.transcript && result.transcript.length > 0) {
        updateTranscriptDOM(result.transcript)
        return
    }

    // Nothing live: show the most recent meeting instead of an empty panel, so this is
    // still useful for re-reading the last call without leaving the meeting.
    chrome.storage.local.get(["meetings"], (storedUntyped) => {
        const stored = storedUntyped as Partial<ResultLocal>
        const last = stored.meetings?.[stored.meetings.length - 1]
        if (!last || last.transcript.length === 0) return

        if (meetingTitle) {
            meetingTitle.innerText = last.meetingTitle || last.title || "Last meeting"
        }
        const indicator = document.querySelector<HTMLElement>(".live-indicator span:last-child")
        if (indicator) indicator.textContent = "Showing your last meeting"
        const dot = document.querySelector<HTMLElement>(".live-dot")
        if (dot) dot.style.background = "var(--ink-subtle)"

        updateTranscriptDOM(last.transcript)
    })
})

// Update on change
chrome.storage.onChanged.addListener((changes, areaName) => {
    if (areaName === 'local' && changes.meetingTitle) {
        const titleText = changes.meetingTitle.newValue
        if (titleText && meetingTitle) {
            meetingTitle.innerHTML = titleText
        }
    }

    if (areaName === 'local' && changes.transcript) {
        const newTranscript = (changes.transcript.newValue || []) as TranscriptBlock[]
        updateTranscriptDOM(newTranscript)
    }
})


// Live Message Listener
chrome.runtime.onMessage.addListener((messageUnTyped) => {
    const message = messageUnTyped as ExtensionMessage
    if (message.type === "broadcast_live_buffer") {
        if (message.stateTranscriptBlock) {
            renderLiveBuffer(message.stateTranscriptBlock)
        }
    }
})

/** Renders or updates the temporary live block at the absolute bottom */
function renderLiveBuffer(liveBlock: StateTranscriptBlock): void {
    // If the buffer is empty, don't show an empty block
    if (!liveBlock.transcriptTextBuffer.trim()) return
    if (!container) return

    const isUserAtBottom =
        (container.scrollHeight - container.scrollTop - container.clientHeight) <= SCROLL_THRESHOLD

    let liveEl = document.getElementById(LIVE_BLOCK_ID)

    if (!liveEl) {
        // Create the live element if it doesn't exist yet
        liveEl = createTranscriptBlockElement({
            personName: liveBlock.personName,
            timestamp: liveBlock.timestamp,
            transcriptText: liveBlock.transcriptTextBuffer,
        })
        liveEl.id = LIVE_BLOCK_ID
        container.appendChild(liveEl)
    } else {
        // Update the existing live element text content safely
        const speakerName = liveEl.querySelector<HTMLElement>('.speaker-name')
        const blockText = liveEl.querySelector<HTMLElement>('.block-text')
        if (speakerName) speakerName.textContent = liveBlock.personName
        if (blockText) blockText.textContent = liveBlock.transcriptTextBuffer
    }

    if (isUserAtBottom) {
        container.scrollTop = container.scrollHeight
    }
}

/** Appends new blocks dynamically and cleans up the live placeholder */
function updateTranscriptDOM(currentTranscriptArray: TranscriptBlock[]): void {
    if (!container) return

    const liveEl = document.getElementById(LIVE_BLOCK_ID)

    // Exclude the live placeholder from our rendered count calculation
    const currentRenderedCount = liveEl ? container.children.length - 1 : container.children.length

    if (currentTranscriptArray.length <= currentRenderedCount) {
        if (currentTranscriptArray.length === 0) container.innerHTML = ''
        return
    }

    const isUserAtBottom =
        (container.scrollHeight - container.scrollTop - container.clientHeight) <= SCROLL_THRESHOLD

    // REMOVE the live placeholder before appending new solid blocks so the new official block takes its correct chronological place.
    if (liveEl) {
        liveEl.remove()
    }

    const fragment = document.createDocumentFragment()
    const newBlocks = currentTranscriptArray.slice(currentRenderedCount)

    newBlocks.forEach(block => {
        const blockEl = createTranscriptBlockElement(block)
        fragment.appendChild(blockEl)
    })

    container.appendChild(fragment)

    if (isUserAtBottom) {
        container.scrollTop = container.scrollHeight
    }
}

function createTranscriptBlockElement(block: TranscriptBlock): HTMLDivElement {
    const div = document.createElement('div')
    div.className = 'transcript-block'

    div.innerHTML = `
    <div class="block-header">
      <span class="speaker-name"></span>
      <span class="block-time"></span>
    </div>
    <p class="block-text"></p>
  `

    const speakerName = div.querySelector<HTMLElement>('.speaker-name')
    const blockTime = div.querySelector<HTMLElement>('.block-time')
    const blockText = div.querySelector<HTMLElement>('.block-text')

    if (speakerName) speakerName.textContent = block.personName

    const timeString = new Date(block.timestamp).toLocaleTimeString([], {
        hour: '2-digit',
        minute: '2-digit'
    })
    if (blockTime) blockTime.textContent = timeString
    if (blockText) blockText.textContent = block.transcriptText

    return div
}

// ---------------------------------------------------------------------------
// The meeting in progress: summarise or download it without ending it
// ---------------------------------------------------------------------------

const panelActions = document.querySelector<HTMLElement>("#panel-actions")
const summarizeButton = document.querySelector<HTMLButtonElement>("#summarize-live")
const downloadButton = document.querySelector<HTMLButtonElement>("#download-live")
const actionStatus = document.querySelector<HTMLElement>("#action-status")
const liveSummarySection = document.querySelector<HTMLDetailsElement>("#live-summary")
const liveSummaryBody = document.querySelector<HTMLElement>("#live-summary-body")
const liveSummaryMeta = document.querySelector<HTMLElement>("#live-summary-meta")
const downloadSummaryButton = document.querySelector<HTMLButtonElement>("#download-live-summary")
const copySummaryButton = document.querySelector<HTMLButtonElement>("#copy-live-summary")

let currentMeetingStart: string | undefined
let currentLiveSummary: LiveSummary | null | undefined

function setStatus(text: string, isError = false): void {
    if (!actionStatus) return
    actionStatus.textContent = text
    actionStatus.classList.toggle("is-error", isError)
}

function describeError(response: ExtensionResponse | undefined): string {
    const error = response?.message as ErrorObject | undefined
    if (error?.errorCode === "023") return "The meeting has ended. Its transcript is in your meetings list."
    if (error?.errorCode === "022") return "Nothing has been said yet, so there is nothing to summarise."
    return error?.errorMessage ?? "Something went wrong. Try again in a moment."
}

/** Actions only make sense while a meeting is live; the last-meeting view has its own page. */
function renderLiveState(meetingTabId: ResultLocal["meetingTabId"] | undefined): void {
    if (panelActions) panelActions.hidden = typeof meetingTabId !== "number"
}

/** Shows the summary so far, but only the one that belongs to the meeting on screen. */
function renderLiveSummary(): void {
    if (!liveSummarySection || !liveSummaryBody) return
    const summary = currentLiveSummary
    if (!summary || !currentMeetingStart || summary.meetingStartTimestamp !== currentMeetingStart) {
        liveSummarySection.hidden = true
        if (summarizeButton) summarizeButton.textContent = "Summarise so far"
        return
    }

    liveSummaryBody.replaceChildren(buildSummaryDocument(summary.summary))
    if (liveSummaryMeta) {
        const time = new Date(summary.createdAt).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })
        liveSummaryMeta.textContent = `as of ${time} · ${summary.blockCount} ${summary.blockCount === 1 ? "line" : "lines"}`
    }
    liveSummarySection.hidden = false
    if (summarizeButton) summarizeButton.textContent = "Update summary"
}

chrome.storage.local.get(["meetingTabId", "meetingStartTimestamp", "liveSummary"], (resultUntyped) => {
    const result = resultUntyped as Partial<ResultLocal>
    currentMeetingStart = result.meetingStartTimestamp
    currentLiveSummary = result.liveSummary
    renderLiveState(result.meetingTabId)
    renderLiveSummary()
})

chrome.storage.onChanged.addListener((changes, areaName) => {
    if (areaName !== "local") return
    if (changes.meetingTabId) renderLiveState(changes.meetingTabId.newValue as ResultLocal["meetingTabId"])
    if (changes.meetingStartTimestamp) currentMeetingStart = changes.meetingStartTimestamp.newValue as string | undefined
    if (changes.liveSummary) currentLiveSummary = changes.liveSummary.newValue as LiveSummary | null | undefined
    if (changes.meetingStartTimestamp || changes.liveSummary) renderLiveSummary()
})

summarizeButton?.addEventListener("click", () => {
    const label = summarizeButton.textContent
    summarizeButton.disabled = true
    summarizeButton.textContent = "Summarising…"
    setStatus("ChatGPT opens for a few seconds to take the transcript, then you are brought back to the meeting.")

    const message: ExtensionMessage = { type: "summarize_live_meeting" }
    chrome.runtime.sendMessage(message, (responseUntyped) => {
        const response = responseUntyped as ExtensionResponse | undefined
        summarizeButton.disabled = false
        summarizeButton.textContent = label
        if (response?.success) {
            // The summary itself arrives through storage, which also covers a panel
            // that was closed and reopened while ChatGPT was answering
            setStatus("")
            if (liveSummarySection) liveSummarySection.open = true
        }
        else {
            setStatus(describeError(response), true)
            renderLiveSummary()
        }
    })
})

downloadButton?.addEventListener("click", () => {
    downloadButton.disabled = true
    const message: ExtensionMessage = {
        type: "download_live_transcript",
        format: document.querySelector<HTMLSelectElement>("#live-download-format")?.value === "html" ? "html" : "txt"
    }
    chrome.runtime.sendMessage(message, (responseUntyped) => {
        const response = responseUntyped as ExtensionResponse | undefined
        downloadButton.disabled = false
        setStatus(response?.success ? "Transcript so far saved to your downloads." : describeError(response), !response?.success)
    })
})

downloadSummaryButton?.addEventListener("click", () => {
    downloadSummaryButton.disabled = true
    const message: ExtensionMessage = {
        type: "download_live_summary",
        format: document.querySelector<HTMLSelectElement>("#live-summary-download-format")?.value === "html" ? "html" : "txt"
    }
    chrome.runtime.sendMessage(message, (responseUntyped) => {
        const response = responseUntyped as ExtensionResponse | undefined
        downloadSummaryButton.disabled = false
        setStatus(response?.success ? "Summary so far saved to your downloads." : describeError(response), !response?.success)
    })
})

copySummaryButton?.addEventListener("click", () => {
    if (!currentLiveSummary) return
    navigator.clipboard.writeText(currentLiveSummary.summary).then(() => {
        copySummaryButton.textContent = "Copied"
        setTimeout(() => { copySummaryButton.textContent = "Copy" }, 1500)
    }).catch(() => setStatus("Could not copy. Select the text instead.", true))
})
