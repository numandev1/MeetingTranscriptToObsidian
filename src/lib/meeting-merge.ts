import type { ChatMessage, Meeting, TranscriptBlock } from "./types"

/**
 * A piece at the same Meet address continues the saved meeting when it starts on the
 * same local day as that meeting's last piece ended, or within this long of it — the
 * latter only matters for a call that runs past midnight.
 *
 * The bound is what tells "left and rejoined" apart from "the next occurrence": a
 * daily standup keeps one Meet code for months, and with no bound every day's standup
 * was appended to the first one — three days filed as a single 49-hour meeting under
 * the first day's note.
 */
export const REJOIN_WINDOW_MS = 2 * 60 * 60 * 1000

/**
 * The earlier piece's end is stamped when it is filed, which can be a moment after the
 * next piece started.
 */
const REJOIN_CLOCK_SLACK_MS = 60 * 1000

/**
 * Below this many words a meeting is not summarised automatically. A reload or a
 * rejoin files the few seconds before it as a meeting of its own, and summarising
 * that opened ChatGPT in the middle of the call to produce a one-line summary.
 * The Summarise button still works on any meeting.
 */
export const MIN_WORDS_TO_AUTO_SUMMARIZE = 50

/** Titles Meet shows before the real meeting name has loaded. */
const GENERIC_TITLES = ["", "google meet", "meet"]

/**
 * Meet ignores account-selection queries on rejoin. Other platforms retain their
 * complete address because query strings and fragments can identify the meeting.
 */
export function normaliseMeetingUrl(url: string | undefined): string | undefined {
    if (!url) {
        return undefined
    }
    try {
        const parsed = new URL(url)
        return parsed.hostname === "meet.google.com"
            ? `${parsed.origin}${parsed.pathname}`.replace(/\/+$/, "")
            : parsed.href
    }
    catch {
        return undefined
    }
}

/**
 * Whether `next` is the same meeting as `previous`, resumed after the page reloaded or
 * the user left and rejoined.
 *
 * Google Meet only, because only Meet has an address that identifies one meeting: a
 * Teams or Zoom web client address can be shared by consecutive, unrelated meetings,
 * and merging those would be far worse than a meeting split in two.
 *
 * The meeting code (the address without its query string) identifies the meeting, but
 * not on its own: a recurring meeting, or someone's personal room, reuses one code for
 * unrelated calls. So the next piece must start on the same local day as the previous
 * one ended, or within REJOIN_WINDOW_MS of it, and two real titles that differ mean two
 * meetings.
 */
export function isRejoinOf(previous: Meeting | undefined, next: Meeting): boolean {
    if (!previous) {
        return false
    }
    if (previous.meetingSoftware !== "Google Meet" || next.meetingSoftware !== "Google Meet") {
        return false
    }
    const previousUrl = normaliseMeetingUrl(previous.meetingUrl)
    if (!previousUrl || previousUrl !== normaliseMeetingUrl(next.meetingUrl)) {
        return false
    }
    if (previous.accountEmail && next.accountEmail && previous.accountEmail !== next.accountEmail) {
        return false
    }
    if (!isGenericTitle(previous.meetingTitle) && !isGenericTitle(next.meetingTitle)
        && previous.meetingTitle!.trim() !== next.meetingTitle!.trim()) {
        return false
    }

    const previousEnd = new Date(previous.meetingEndTimestamp)
    const nextStart = new Date(next.meetingStartTimestamp)
    if (previous.captureId && next.captureId && previous.captureId !== next.captureId
        && nextStart.getTime() < previousEnd.getTime()) return false
    const gap = nextStart.getTime() - previousEnd.getTime()
    if (!Number.isFinite(gap) || gap < -REJOIN_CLOCK_SLACK_MS) {
        return false
    }
    return gap <= REJOIN_WINDOW_MS || previousEnd.toDateString() === nextStart.toDateString()
}

function normaliseText(text: string): string {
    return text.toLowerCase().replace(/[^\p{L}\p{N}]+/gu, "")
}

/**
 * Joins the two transcripts, dropping the earlier piece's last block when the later
 * piece opens with the same speaker saying the same thing. Meet shows the caption in
 * progress again after a rejoin, so without this the sentence appears twice, once cut
 * off.
 */
function joinTranscripts(first: TranscriptBlock[], second: TranscriptBlock[]): TranscriptBlock[] {
    const last = first[first.length - 1]
    const next = second[0]
    const gap = last && next ? Date.parse(next.timestamp) - Date.parse(last.timestamp) : NaN
    if (last && next
        && Number.isFinite(gap) && gap >= 0 && gap <= 60_000
        && normaliseText(last.transcriptText).length > 0
        && last.personName === next.personName
        && normaliseText(next.transcriptText).startsWith(normaliseText(last.transcriptText))) {
        return [...first.slice(0, -1), ...second]
    }
    return [...first, ...second]
}

function joinChatMessages(first: ChatMessage[], second: ChatMessage[]): ChatMessage[] {
    const key = (message: ChatMessage) => `${message.personName}\u0000${message.chatMessageText}`
    const seen = new Set(first.map(key))
    return [...first, ...second.filter(message => !seen.has(key(message)))]
}

/** "Meet - abc-defg-hij": the tab title until, or unless, Meet shows the event name. */
const MEET_CODE_TITLE = /^meet\s*-\s*[a-z]{3}-[a-z]{4}-[a-z]{3}$/

function isGenericTitle(title: string | undefined): boolean {
    const normalised = (title ?? "").trim().toLowerCase()
    return GENERIC_TITLES.includes(normalised) || MEET_CODE_TITLE.test(normalised)
}

/**
 * One meeting out of two pieces of it.
 *
 * It keeps the earlier piece's start time, so the Obsidian note and the file names
 * stay the same and the note is overwritten rather than duplicated. Every delivery is
 * reset to "new" and the earlier summary is dropped: it described only the first
 * piece, and left in place it would read as the summary of the whole meeting.
 */
export function mergeRejoin(previous: Meeting, next: Meeting): Meeting {
    return {
        ...previous,
        meetingTitle: isGenericTitle(previous.meetingTitle) ? next.meetingTitle : previous.meetingTitle,
        meetingEndTimestamp: next.meetingEndTimestamp,
        accountEmail: previous.accountEmail ?? next.accountEmail,
        captionsSeen: previous.captionsSeen === true || next.captionsSeen === true,
        transcript: joinTranscripts(previous.transcript, next.transcript),
        chatMessages: joinChatMessages(previous.chatMessages, next.chatMessages),
        mergedStartTimestamps: [...(previous.mergedStartTimestamps ?? []), next.meetingStartTimestamp],
        webhookPostStatus: "new",
        obsidianPostStatus: "new",
        summary: undefined,
        summaryStatus: undefined,
    }
}

/** Keep the continued meeting last so exports and retention use the updated entry. */
export function appendOrMergeMeeting(meetings: Meeting[], next: Meeting): Meeting[] {
    if (meetings.some(meeting => next.captureId && meeting.captureId
        ? meeting.captureId === next.captureId
        : coversStartTimestamp(meeting, next.meetingStartTimestamp)
            && normaliseMeetingUrl(meeting.meetingUrl) === normaliseMeetingUrl(next.meetingUrl)
            && meeting.meetingSoftware === next.meetingSoftware)) {
        return meetings
    }
    for (let index = meetings.length - 1; index >= 0; index--) {
        if (isRejoinOf(meetings[index], next)) {
            return [...meetings.slice(0, index), ...meetings.slice(index + 1), mergeRejoin(meetings[index], next)]
        }
    }
    return [...meetings, next]
}

/**
 * Whether this saved meeting already holds the piece that started at `startTimestamp`.
 * Recovery relies on it: a merged meeting keeps its first piece's start time, so the
 * later piece would otherwise look like one that was never filed, and be filed again.
 */
export function coversStartTimestamp(meeting: Meeting | undefined, startTimestamp: string): boolean {
    if (!meeting) {
        return false
    }
    return meeting.meetingStartTimestamp === startTimestamp
        || (meeting.mergedStartTimestamps ?? []).includes(startTimestamp)
}

/**
 * Takes apart meetings that were merged although their pieces are not rejoins of each
 * other by today's rules, and returns the list with each such meeting replaced by the
 * meetings it really was. Meetings that hold up are returned as the same objects.
 *
 * This repairs what an unbounded rejoin rule left in storage: pieces of one address
 * from different days, joined into one entry. Each piece is rebuilt from its start
 * time in mergedStartTimestamps, taking the transcript blocks and chat messages
 * stamped from then until the next piece started, and the pieces are then merged
 * again with isRejoinOf, so a genuine rejoin inside a split meeting stays joined.
 *
 * A meeting that is split loses its summary, which described all of it, and every
 * resulting meeting is marked unposted: the first one keeps its start time and so its
 * note and file names, whose contents now cover too much.
 */
export function splitInvalidMerges(meetings: Meeting[]): Meeting[] {
    return meetings.flatMap(meeting => {
        const pieces = splitIntoPieces(meeting)
        if (pieces.length < 2) {
            return [meeting]
        }
        const regrouped: Meeting[] = []
        for (const piece of pieces) {
            const last = regrouped[regrouped.length - 1]
            if (last && isRejoinOf(last, piece)) {
                regrouped[regrouped.length - 1] = mergeRejoin(last, piece)
            }
            else {
                regrouped.push(piece)
            }
        }
        return regrouped.length < 2 ? [meeting] : regrouped
    })
}

function splitIntoPieces(meeting: Meeting): Meeting[] {
    const starts = [meeting.meetingStartTimestamp, ...(meeting.mergedStartTimestamps ?? [])]
        .filter(start => Number.isFinite(Date.parse(start)))
        .sort((a, b) => Date.parse(a) - Date.parse(b))
    if (starts.length < 2 || starts[0] !== meeting.meetingStartTimestamp) {
        return []
    }
    const pieceOf = (timestamp: string) => {
        const time = Date.parse(timestamp)
        let index = 0
        while (index + 1 < starts.length && Date.parse(starts[index + 1]) <= time) {
            index++
        }
        return index
    }
    return starts.map((start, index) => {
        const transcript = meeting.transcript.filter(block => pieceOf(block.timestamp) === index)
        const chatMessages = meeting.chatMessages.filter(message => pieceOf(message.timestamp) === index)
        const isLast = index === starts.length - 1
        const lastSaid = [...transcript, ...chatMessages]
            .map(item => item.timestamp)
            .sort((a, b) => Date.parse(a) - Date.parse(b))
            .pop()
        return {
            ...meeting,
            meetingStartTimestamp: start,
            // Only the last piece's real end is known; the others ended around the
            // last thing said in them.
            meetingEndTimestamp: isLast ? meeting.meetingEndTimestamp : lastSaid ?? start,
            transcript,
            chatMessages,
            mergedStartTimestamps: undefined,
            webhookPostStatus: "new",
            obsidianPostStatus: "new",
            summary: undefined,
            summaryStatus: undefined,
        }
    })
}

/** Whether enough was said in a meeting for it to be worth summarising automatically. */
export function hasEnoughToSummarize(meeting: Meeting): boolean {
    const words = meeting.transcript
        .map(block => block.transcriptText)
        .concat(meeting.chatMessages.map(message => message.chatMessageText))
        .join(" ")
        .split(/\s+/)
        .filter(word => word.length > 0)
    return words.length >= MIN_WORDS_TO_AUTO_SUMMARIZE
}
