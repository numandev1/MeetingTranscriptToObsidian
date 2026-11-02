import type { ErrorObject, Meeting, ResultLocal, ResultSync } from "@lib/types"
import { DEFAULT_OBSIDIAN_FOLDER, DEFAULT_OBSIDIAN_SUMMARY_FOLDER, getFolderPrefix, withAccountSubfolder } from "@lib/paths"
import { TIMEFORMAT } from "./config"

export const DEFAULT_OBSIDIAN_BASE_URL = "http://127.0.0.1:27123"

/**
 * Characters Obsidian refuses in note names, plus the path separators that would
 * otherwise silently nest the note in an unintended folder.
 */
const INVALID_NOTE_NAME_CHARS = /[\\/:*?"<>|#^[\]]/g

/** Local date (not UTC) in YYYY-MM-DD, so a late evening meeting files under that day. */
function getLocalDateStamp(isoTimestamp: string): string {
    const date = new Date(isoTimestamp)
    const month = `${date.getMonth() + 1}`.padStart(2, "0")
    const day = `${date.getDate()}`.padStart(2, "0")
    return `${date.getFullYear()}-${month}-${day}`
}

/** Local 24-hour time as HHMM, to order several meetings on the same day. */
function getLocalTimeStamp(isoTimestamp: string): string {
    const date = new Date(isoTimestamp)
    return `${`${date.getHours()}`.padStart(2, "0")}${`${date.getMinutes()}`.padStart(2, "0")}`
}

/**
 * Note name for a meeting, e.g. "2026-09-14 0930 Standup".
 *
 * The timestamp leads because Obsidian's file explorer sorts alphabetically by file
 * name. With the title first, a new meeting lands wherever its title happens to fall
 * alphabetically; with the date first, the folder is always in chronological order.
 * The time is included so several meetings on one day still order correctly, and
 * because it keeps recurring meetings of the same name from colliding.
 */
export function getObsidianNoteName(meeting: Meeting): string {
    const rawTitle = meeting.meetingTitle || meeting.title || "Meeting"
    const title = rawTitle
        .replace(INVALID_NOTE_NAME_CHARS, "")
        .trim()
        // A note name cannot start with a dot, and Obsidian trims trailing dots
        .replace(/^\.+/, "")
        .replace(/\.+$/, "")
        .trim()

    const startedAt = meeting.meetingStartTimestamp
    return `${getLocalDateStamp(startedAt)} ${getLocalTimeStamp(startedAt)} ${title || "Meeting"}${meeting.captureId ? ` ${meeting.captureId}` : ""}`
}

/** Escapes a value for a single-quoted YAML scalar. */
function toYamlString(value: string): string {
    return `'${value.replace(/'/g, "''")}'`
}

function getDurationString(meeting: Meeting): string {
    const duration = new Date(meeting.meetingEndTimestamp).getTime() - new Date(meeting.meetingStartTimestamp).getTime()
    const totalMinutes = Math.round(duration / (1000 * 60))
    const hours = Math.floor(totalMinutes / 60)
    const minutes = totalMinutes % 60
    return hours > 0 ? `${hours}h ${minutes}m` : `${totalMinutes}m`
}

/** Everyone who spoke or chatted, in first-appearance order. */
function getParticipants(meeting: Meeting): string[] {
    const seen = new Set<string>()
    meeting.transcript.forEach(block => seen.add(block.personName))
    meeting.chatMessages.forEach(message => seen.add(message.personName))
    return [...seen].filter(name => name.trim().length > 0)
}

/** Builds the markdown note, with frontmatter matching the vault's Clippings convention. */
export function getObsidianNoteMarkdown(meeting: Meeting): string {
    const title = meeting.meetingTitle || meeting.title || "Meeting"
    const participants = getParticipants(meeting)

    let note = "---\n"
    note += `title: ${toYamlString(title)}\n`
    note += `meetingSoftware: ${toYamlString(meeting.meetingSoftware || "")}\n`
    note += `created: ${toYamlString(meeting.meetingStartTimestamp)}\n`
    note += `ended: ${toYamlString(meeting.meetingEndTimestamp)}\n`
    note += `duration: ${toYamlString(getDurationString(meeting))}\n`
    if (participants.length > 0) {
        note += "participants:\n"
        participants.forEach(name => {
            note += `  - ${toYamlString(name)}\n`
        })
    }
    note += "tags:\n  - meeting\n  - transcript\n"
    note += "---\n\n"

    note += `# ${title}\n\n`

    note += "## Transcript\n\n"
    if (meeting.transcript.length > 0) {
        meeting.transcript.forEach(block => {
            note += `**${block.personName}** — ${new Date(block.timestamp).toLocaleString("default", TIMEFORMAT).toUpperCase()}\n\n`
            note += `${block.transcriptText}\n\n`
        })
    }
    else {
        note += "_No transcript captured._\n\n"
    }

    note += "## Chat messages\n\n"
    if (meeting.chatMessages.length > 0) {
        meeting.chatMessages.forEach(message => {
            note += `**${message.personName}** — ${new Date(message.timestamp).toLocaleString("default", TIMEFORMAT).toUpperCase()}\n\n`
            note += `${message.chatMessageText}\n\n`
        })
    }
    else {
        note += "_No chat messages captured._\n\n"
    }

    return note
}

/** Builds the summary note, linking back to the full transcript note. */
export function getObsidianSummaryMarkdown(meeting: Meeting, summary: string): string {
    const title = meeting.meetingTitle || meeting.title || "Meeting"

    let note = "---\n"
    note += `title: ${toYamlString(title + " (summary)")}\n`
    note += `meetingSoftware: ${toYamlString(meeting.meetingSoftware || "")}\n`
    note += `created: ${toYamlString(meeting.meetingStartTimestamp)}\n`
    note += `duration: ${toYamlString(getDurationString(meeting))}\n`
    note += "tags:\n  - meeting\n  - summary\n"
    note += "---\n\n"

    note += `# ${title}\n\n`
    // Wiki link to the transcript note. A bare name resolves regardless of which
    // account subfolder each note ended up in.
    note += `Full transcript: [[${getObsidianNoteName(meeting)}]]\n\n`
    note += `${summary.trim()}\n`

    return note
}

/**
 * Writes a meeting summary into the configured summary folder.
 * @throws error codes: 010, 017, 018, 019
 */
export function postSummaryToObsidian(index: number): Promise<string> {
    return new Promise((resolve, reject) => {
        chrome.storage.local.get(["meetings", "obsidianApiKey"], function (resultLocalUntyped) {
            const resultLocal = resultLocalUntyped as Partial<ResultLocal>
            chrome.storage.sync.get(["obsidianSummaryFolder", "obsidianBaseUrl"], function (resultSyncUntyped) {
                const resultSync = resultSyncUntyped as Partial<ResultSync>

                if (!resultLocal.obsidianApiKey) {
                    reject({ errorCode: "018", errorMessage: "Obsidian is not configured" } satisfies ErrorObject)
                    return
                }

                const meetings = resultLocal.meetings
                if (!meetings || !meetings[index] || !meetings[index].summary) {
                    reject({ errorCode: "010", errorMessage: "Meeting at specified index not found" } satisfies ErrorObject)
                    return
                }

                const meeting = meetings[index]
                const folderPrefix = withAccountSubfolder(
                    getFolderPrefix(resultSync.obsidianSummaryFolder, DEFAULT_OBSIDIAN_SUMMARY_FOLDER),
                    meeting.accountEmail
                )

                putVaultNote(
                    resultSync.obsidianBaseUrl,
                    resultLocal.obsidianApiKey,
                    `${folderPrefix}${getObsidianNoteName(meeting)}.md`,
                    getObsidianSummaryMarkdown(meeting, meeting.summary as string)
                ).then(() => {
                    console.log("Obsidian summary saved")
                    resolve("Obsidian summary saved successfully")
                }).catch(reject)
            })
        })
    })
}

/** True when a note already exists at this vault path. */
async function vaultNoteExists(baseUrl: string | undefined, apiKey: string, vaultPath: string): Promise<boolean> {
    const url = (baseUrl || DEFAULT_OBSIDIAN_BASE_URL).replace(/\/+$/, "")
    const encodedPath = vaultPath.split("/").map(encodeURIComponent).join("/")
    const response = await fetch(`${url}/vault/${encodedPath}`, {
        headers: { "Authorization": `Bearer ${apiKey}` }
    }).catch(() => null)
    return response?.ok === true
}

/** PUTs a note into the vault, creating parent folders. */
function putVaultNote(baseUrl: string | undefined, apiKey: string, vaultPath: string, body: string): Promise<void> {
    const url = (baseUrl || DEFAULT_OBSIDIAN_BASE_URL).replace(/\/+$/, "")
    // Each path segment is encoded separately so the folder separators survive
    const encodedPath = vaultPath.split("/").map(encodeURIComponent).join("/")

    return fetch(`${url}/vault/${encodedPath}`, {
        method: "PUT",
        headers: {
            "Authorization": `Bearer ${apiKey}`,
            "Content-Type": "text/markdown"
        },
        body: body
    }).then(response => {
        if (!response.ok) {
            throw new Error(`Obsidian request failed with HTTP status code ${response.status} ${response.statusText}`)
        }
    }).catch(error => {
        // A TypeError from fetch means the connection never landed — by far the most
        // common cause is Obsidian simply not being open.
        const isUnreachable = error instanceof TypeError
        throw {
            errorCode: isUnreachable ? "019" : "017",
            errorMessage: isUnreachable
                ? "Could not reach Obsidian. Is the app running with the Local REST API plugin enabled?"
                : (error instanceof Error ? error.message : String(error))
        } satisfies ErrorObject
    })
}

/**
 * Writes the meeting as a note in the user's vault through the Obsidian Local REST API
 * plugin. Parent folders are created by the plugin.
 *
 * Uses the plugin's plain-HTTP port by default: its HTTPS port serves a self-signed
 * certificate, which fetch() inside a service worker rejects with no way to opt out.
 *
 * @throws error codes: 010, 017, 018, 019
 */
export function postTranscriptToObsidian(index: number): Promise<string> {
    return new Promise((resolve, reject) => {
        chrome.storage.local.get(["meetings", "obsidianApiKey"], function (resultLocalUntyped) {
            const resultLocal = resultLocalUntyped as Partial<ResultLocal>
            chrome.storage.sync.get(["obsidianFolder", "obsidianBaseUrl"], function (resultSyncUntyped) {
                const resultSync = resultSyncUntyped as Partial<ResultSync>

                if (!resultLocal.obsidianApiKey) {
                    reject({ errorCode: "018", errorMessage: "Obsidian is not configured" } satisfies ErrorObject)
                    return
                }

                const meetings = resultLocal.meetings
                if (!meetings || !meetings[index]) {
                    reject({ errorCode: "010", errorMessage: "Meeting at specified index not found" } satisfies ErrorObject)
                    return
                }

                const meeting = meetings[index]
                const baseUrl = (resultSync.obsidianBaseUrl || DEFAULT_OBSIDIAN_BASE_URL).replace(/\/+$/, "")
                const folderPrefix = withAccountSubfolder(
                    getFolderPrefix(resultSync.obsidianFolder, DEFAULT_OBSIDIAN_FOLDER),
                    meeting.accountEmail
                )
                // Each path segment is encoded separately so the folder separators survive
                const vaultPath = `${folderPrefix}${getObsidianNoteName(meeting)}.md`
                    .split("/")
                    .map(encodeURIComponent)
                    .join("/")

                fetch(`${baseUrl}/vault/${vaultPath}`, {
                    method: "PUT",
                    headers: {
                        "Authorization": `Bearer ${resultLocal.obsidianApiKey}`,
                        "Content-Type": "text/markdown"
                    },
                    body: getObsidianNoteMarkdown(meeting)
                }).then(response => {
                    if (!response.ok) {
                        throw new Error(`Obsidian request failed with HTTP status code ${response.status} ${response.statusText}`)
                    }
                    meetings[index].obsidianPostStatus = "successful"
                    chrome.storage.local.set({ meetings: meetings }, function () {
                        console.log("Obsidian note saved")
                        resolve("Obsidian note saved successfully")
                    })
                }).catch(error => {
                    console.error(error)
                    meetings[index].obsidianPostStatus = "failed"
                    chrome.storage.local.set({ meetings: meetings }, function () {
                        // A TypeError from fetch means the connection never landed — by far
                        // the most common cause is Obsidian simply not being open.
                        const isUnreachable = error instanceof TypeError
                        reject({
                            errorCode: isUnreachable ? "019" : "017",
                            errorMessage: isUnreachable
                                ? "Could not reach Obsidian. Is the app running with the Local REST API plugin enabled?"
                                : (error instanceof Error ? error.message : String(error))
                        } satisfies ErrorObject)
                    })
                })
            })
        })
    })
}

export interface VaultSyncResult {
    transcripts: number
    summaries: number
    /** Notes already in the vault, left untouched so user edits survive. */
    skipped: number
    failed: number
    firstError?: string
}

/**
 * Writes every stored meeting into the vault, transcripts and any summaries.
 *
 * Runs sequentially rather than in parallel: the target is a local plugin inside the
 * user's editor, and firing ten concurrent writes at it gains nothing while making a
 * partial failure much harder to report accurately.
 *
 * Notes are addressed by meeting title and date, so re-running replaces the existing
 * note instead of creating duplicates, which makes this safe to use as a repair action.
 *
 * @throws error codes: 013, 018
 */
export function syncAllMeetingsToObsidian(): Promise<VaultSyncResult> {
    return new Promise((resolve, reject) => {
        chrome.storage.local.get(["meetings", "obsidianApiKey"], async function (resultLocalUntyped) {
            const resultLocal = resultLocalUntyped as Partial<ResultLocal>

            if (!resultLocal.obsidianApiKey) {
                reject({ errorCode: "018", errorMessage: "Obsidian is not configured" } satisfies ErrorObject)
                return
            }

            const meetings = resultLocal.meetings ?? []
            if (meetings.length === 0) {
                reject({ errorCode: "013", errorMessage: "No meetings found. May be attend one?" } satisfies ErrorObject)
                return
            }

            const result: VaultSyncResult = { transcripts: 0, summaries: 0, skipped: 0, failed: 0 }

            const settings = await new Promise<Partial<ResultSync>>(res =>
                chrome.storage.sync.get(["obsidianFolder", "obsidianSummaryFolder", "obsidianBaseUrl"],
                    r => res(r as Partial<ResultSync>)))

            for (let index = 0; index < meetings.length; index++) {
                const meeting = meetings[index]
                const noteName = `${getObsidianNoteName(meeting)}.md`
                const transcriptPath = withAccountSubfolder(
                    getFolderPrefix(settings.obsidianFolder, DEFAULT_OBSIDIAN_FOLDER), meeting.accountEmail) + noteName
                const summaryPath = withAccountSubfolder(
                    getFolderPrefix(settings.obsidianSummaryFolder, DEFAULT_OBSIDIAN_SUMMARY_FOLDER), meeting.accountEmail) + noteName

                try {
                    // Existing notes are left alone: this action backfills a vault, and a
                    // note already there may carry the user's own edits — ticked action
                    // items, added context — which a blind overwrite would destroy. The
                    // per-meeting save button still overwrites, because that is explicit.
                    if (await vaultNoteExists(settings.obsidianBaseUrl, resultLocal.obsidianApiKey, transcriptPath)) {
                        result.skipped++
                    }
                    else {
                        await postTranscriptToObsidian(index)
                        result.transcripts++
                    }
                }
                catch (error) {
                    result.failed++
                    const parsedError = error as ErrorObject
                    result.firstError = result.firstError ?? parsedError.errorMessage

                    // An unreachable vault will fail for every remaining meeting too,
                    // so stop rather than grinding through nine more timeouts.
                    if (parsedError.errorCode === "019") {
                        break
                    }
                    continue
                }

                if (meeting.summary) {
                    try {
                        if (await vaultNoteExists(settings.obsidianBaseUrl, resultLocal.obsidianApiKey, summaryPath)) {
                            result.skipped++
                        }
                        else {
                            await postSummaryToObsidian(index)
                            result.summaries++
                        }
                    }
                    catch (error) {
                        const parsedError = error as ErrorObject
                        result.firstError = result.firstError ?? parsedError.errorMessage
                    }
                }
            }

            resolve(result)
        })
    })
}

