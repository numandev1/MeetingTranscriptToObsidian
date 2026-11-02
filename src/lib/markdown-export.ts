import type { Meeting } from "./types"

/** Caption and chat text is content, even when it contains HTML or Markdown syntax. */
function escapeMarkdown(text: string): string {
    return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
        .replace(/[\\`*_\[\]#]/g, "\\$&")
}

/** Full transcript and chat history, in capture order, suitable for a Markdown note. */
export function buildTranscriptMarkdown(meeting: Meeting, isLive = false): string {
    const section = (heading: string, blocks: { personName: string; timestamp: string; text: string }[]) => {
        const body = blocks.map(block =>
            `### ${escapeMarkdown(block.personName)} — ${escapeMarkdown(block.timestamp)}\n\n${escapeMarkdown(block.text).replace(/\r?\n/g, "  \n")}`
        ).join("\n\n")
        return `## ${heading}\n\n${body || "Nothing captured."}`
    }
    return [
        `# ${escapeMarkdown(meeting.meetingTitle || meeting.title || "Meeting")}`,
        `${isLive ? "Transcript so far" : "Full transcript"} · ${meeting.meetingStartTimestamp}`,
        section("Transcript", meeting.transcript.map(block => ({ ...block, text: block.transcriptText }))),
        section("Chat messages", meeting.chatMessages.map(message => ({ ...message, text: message.chatMessageText }))),
    ].join("\n\n") + "\n"
}
