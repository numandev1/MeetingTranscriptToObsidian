import type { Meeting } from "./types"
import { parseSummaryMarkdown, type InlineRun } from "./summary-markdown"

function escapeHtml(text: string): string {
    return text.replace(/[&<>"']/g, character => ({
        "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
    })[character]!)
}

function renderRuns(runs: InlineRun[]): string {
    return runs.map(run => run.bold ? `<strong>${escapeHtml(run.text)}</strong>` : escapeHtml(run.text)).join("")
}

function renderSummary(summary: string): string {
    return parseSummaryMarkdown(summary).map(block => {
        if (block.kind === "list") {
            return `<ul>${block.items.map(item => `<li>${item.checked === undefined ? "" : `<input type="checkbox" disabled${item.checked ? " checked" : ""}> `}${renderRuns(item.runs)}</li>`).join("")}</ul>`
        }
        const tag = block.kind === "heading" ? "h3" : "p"
        return `<${tag}>${renderRuns(block.runs)}</${tag}>`
    }).join("\n")
}

function renderEntries(heading: string, entries: { personName: string, timestamp: string, text: string }[], empty: string): string {
    return `<section><h2>${heading}</h2>${entries.length ? entries.map(entry => `<article><header><strong>${escapeHtml(entry.personName)}</strong> <time datetime="${escapeHtml(entry.timestamp)}">${escapeHtml(new Date(entry.timestamp).toLocaleString())}</time></header><p>${escapeHtml(entry.text)}</p></article>`).join("\n") : `<p>${empty}</p>`}</section>`
}

/** A self-contained document; all meeting and summary text is escaped before rendering. */
export function buildMeetingHtml(meeting: Meeting, kind: "transcript" | "summary", isLive = false): string {
    const title = escapeHtml(meeting.meetingTitle || meeting.title || "Meeting")
    const label = `${kind === "summary" ? "Summary" : "Transcript"}${isLive ? " so far" : ""}`
    const metadata = [meeting.meetingSoftware, new Date(meeting.meetingStartTimestamp).toLocaleString()].filter(Boolean).map(value => escapeHtml(value!)).join(" · ")
    const content = kind === "summary"
        ? `<section><h2>${label}</h2>${renderSummary(meeting.summary || "")}</section>`
        : renderEntries(label, meeting.transcript.map(block => ({ ...block, text: block.transcriptText })), "No transcript was captured for this meeting.")
            + renderEntries("Chat messages", meeting.chatMessages.map(message => ({ ...message, text: message.chatMessageText })), "No chat messages were captured for this meeting.")
    return `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${title} — ${label}</title>
<style>
body { font: 16px/1.6 system-ui, sans-serif; color: #222; background: #fff; margin: 0; }
main { max-width: 800px; margin: auto; padding: 32px 24px; overflow-wrap: anywhere; }
h1 { line-height: 1.2; } h2 { margin-top: 2rem; } h3 { margin-top: 1.5rem; }
.meta, time { color: #555; } time { font-size: .875rem; margin-left: .5rem; }
article { padding: 1rem 0; border-bottom: 1px solid #ddd; break-inside: avoid; }
p { white-space: pre-wrap; } article p { margin: .5rem 0 0; }
li { margin: .4rem 0; } input { margin-right: .4rem; }
@media print { main { max-width: none; padding: 0; } }
</style>
</head>
<body><main><h1>${title}</h1><p class="meta">${metadata}</p>${content}</main></body>
</html>`
}
