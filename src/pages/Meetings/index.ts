import type { ErrorObject, ExtensionMessage, ExtensionResponse, Meeting, Platform, ResultLocal, ResultSync } from "@lib/types"
import { DEFAULT_DOWNLOAD_SUBFOLDER, DEFAULT_MEETINGS_TO_KEEP, DEFAULT_OBSIDIAN_FOLDER, DEFAULT_OBSIDIAN_SUMMARY_FOLDER, getFolderPrefix } from "@lib/paths"
import { DEFAULT_OBSIDIAN_BASE_URL } from "../../background-script/obsidian"
import { parseTranscriptFile } from "@lib/transcript-import"
import { meetingsDroppedBy, trimMeetings } from "@lib/retention"
import { DEFAULT_SUMMARY_INSTRUCTIONS } from "../../background-script/summarizer"
import { buildSummaryDocument } from "@lib/summary-render"
import {
    buildSettingsFile,
    getWantedPlatforms,
    parseSettingsFile,
    TRANSFERABLE_SYNC_KEYS,
    withObsidianEnabled,
} from "@lib/settings-transfer"

document.addEventListener("DOMContentLoaded", function () {
    const webhookUrlForm = document.querySelector("#webhook-url-form")
    const webhookUrlInput = document.querySelector("#webhook-url")
    const saveButton = document.querySelector("#save-webhook")
    const autoPostCheckbox = document.querySelector("#auto-post-webhook")
    const autoDownloadCheckbox = document.querySelector("#auto-download-file")
    const simpleWebhookBodyRadio = document.querySelector("#simple-webhook-body")
    const advancedWebhookBodyRadio = document.querySelector("#advanced-webhook-body")
    const recoverLastMeetingButton = document.querySelector("#recover-last-meeting")
    const downloadFolderForm = document.querySelector("#download-folder-form")
    const downloadFolderInput = document.querySelector("#download-folder")

    // Initial load of transcripts
    loadMeetings()

    // Reload transcripts when page becomes visible
    document.addEventListener("visibilitychange", function () {
        if (document.visibilityState === "visible") {
            loadMeetings()
        }
    })

    chrome.storage.onChanged.addListener(() => {
        loadMeetings()
    })

    if (recoverLastMeetingButton instanceof HTMLButtonElement) {
        recoverLastMeetingButton.addEventListener("click", function () {
            const message: ExtensionMessage = {
                type: "recover_last_meeting",
            }
            chrome.runtime.sendMessage(message, function (responseUntyped) {
                const response = responseUntyped as ExtensionResponse
                loadMeetings()
                scrollTo({ top: 0, behavior: "smooth" })
                if (response.success) {
                    if (response.message === "No recovery needed") {
                        alert("Nothing to recover—you're on top of the world!")
                    }
                    else {
                        alert("Last meeting recovered successfully!")
                    }
                }
                else {
                    const parsedError = response.message as ErrorObject
                    if (parsedError.errorCode === "013") {
                        alert(parsedError.errorMessage)
                    }
                    else if (parsedError.errorCode === "014") {
                        alert("Nothing to recover—you're on top of the world!")
                    }
                    else {
                        alert("Could not recover last meeting!")
                        console.error(parsedError.errorMessage)
                    }
                }
            })
        })
    }

    if (saveButton instanceof HTMLButtonElement && webhookUrlForm instanceof HTMLFormElement && webhookUrlInput instanceof HTMLInputElement && autoPostCheckbox instanceof HTMLInputElement && simpleWebhookBodyRadio instanceof HTMLInputElement && advancedWebhookBodyRadio instanceof HTMLInputElement) {
        // Initially disable the save button
        saveButton.disabled = true

        // Load saved webhook URL, auto-post setting, and webhook body type
        chrome.storage.sync.get(["webhookUrl", "autoPostWebhookAfterMeeting", "autoDownloadFileAfterMeeting", "webhookBodyType"], function (resultSyncUntyped) {
            const resultSync = resultSyncUntyped as Partial<ResultSync>

            if (resultSync.webhookUrl) {
                webhookUrlInput.value = resultSync.webhookUrl
                saveButton.disabled = !webhookUrlInput.checkValidity()
            }

            // Set checkbox state
            autoPostCheckbox.checked = resultSync.autoPostWebhookAfterMeeting === true
            if (autoDownloadCheckbox instanceof HTMLInputElement) {
                autoDownloadCheckbox.checked = resultSync.autoDownloadFileAfterMeeting !== false
            }
            updateAutoDownloadCheckBox()

            // Set radio button state
            if (resultSync.webhookBodyType === "advanced") {
                advancedWebhookBodyRadio.checked = true
            } else {
                simpleWebhookBodyRadio.checked = true
            }
        })

        // Handle URL input changes
        webhookUrlInput.addEventListener("input", function () {
            saveButton.disabled = !webhookUrlInput.checkValidity()
        })

        // Save webhook URL, auto-post setting, and webhook body type
        webhookUrlForm.addEventListener("submit", function (e) {
            e.preventDefault()
            const webhookUrl = webhookUrlInput.value
            if (webhookUrl === "") {
                // Save webhook URL and settings
                chrome.storage.sync.set({
                    webhookUrl: webhookUrl
                }, function () {
                    alert("Webhook URL saved!")
                })
            }
            else if (webhookUrl && webhookUrlInput.checkValidity()) {
                // Request runtime permission for the webhook URL
                requestWebhookAndNotificationPermission(webhookUrl).then(() => {
                    // Save webhook URL and settings
                    chrome.storage.sync.set({
                        webhookUrl: webhookUrl
                    }, function () {
                        alert("Webhook URL saved!")
                    })
                }).catch((error) => {
                    alert("Fine! No webhooks for you!")
                    console.error("Webhook permission error:", error)
                })
            }
        })

        // Auto save auto-post setting
        autoPostCheckbox.addEventListener("change", function () {
            // Save webhook URL and settings
            chrome.storage.sync.set({
                autoPostWebhookAfterMeeting: autoPostCheckbox.checked,
            }, function () {
                updateAutoDownloadCheckBox()
            })
        })

        if (autoDownloadCheckbox instanceof HTMLInputElement) {
            autoDownloadCheckbox.addEventListener("change", function () {
                if (!autoDownloadCheckbox.checked) {
                    if (!confirm("Text file serves as a harmless backup, you sure you don't need it?")) {
                        autoDownloadCheckbox.checked = true
                        return
                    }
                }
                chrome.storage.sync.set({
                    autoDownloadFileAfterMeeting: autoDownloadCheckbox.checked,
                }, function () { })
            })
        }

        function updateAutoDownloadCheckBox() {
            if (autoDownloadCheckbox?.parentElement instanceof HTMLDivElement && autoPostCheckbox instanceof HTMLInputElement) {
                autoDownloadCheckbox.parentElement.style.display = autoPostCheckbox.checked ? "flex" : "none"
                if (!autoPostCheckbox.checked && autoDownloadCheckbox instanceof HTMLInputElement) {
                    autoDownloadCheckbox.checked = true
                    chrome.storage.sync.set({
                        autoDownloadFileAfterMeeting: true,
                    }, function () { })
                }
            }
        }

        // Auto save webhook body type
        simpleWebhookBodyRadio.addEventListener("change", function () {
            // Save webhook URL and settings
            chrome.storage.sync.set({ webhookBodyType: "simple" }, function () { })
        })

        // Auto save webhook body type
        advancedWebhookBodyRadio.addEventListener("change", function () {
            // Save webhook URL and settings
            chrome.storage.sync.set({ webhookBodyType: advancedWebhookBodyRadio.checked ? "advanced" : "simple" }, function () { })
        })
    }

    if (downloadFolderForm instanceof HTMLFormElement && downloadFolderInput instanceof HTMLInputElement) {
        chrome.storage.sync.get(["downloadSubfolder"], function (resultSyncUntyped) {
            const resultSync = resultSyncUntyped as Partial<ResultSync>
            // An empty string is a real setting ("save into Downloads"), so only fall
            // back to the default when nothing has ever been stored.
            downloadFolderInput.value = typeof resultSync.downloadSubfolder === "string"
                ? resultSync.downloadSubfolder
                : DEFAULT_DOWNLOAD_SUBFOLDER
        })

        downloadFolderForm.addEventListener("submit", function (e) {
            e.preventDefault()
            // Store the cleaned value so the input always shows what will actually be used
            const downloadSubfolder = getFolderPrefix(downloadFolderInput.value, DEFAULT_DOWNLOAD_SUBFOLDER).replace(/\/$/, "")
            chrome.storage.sync.set({ downloadSubfolder: downloadSubfolder }, function () {
                downloadFolderInput.value = downloadSubfolder
                alert(downloadSubfolder
                    ? `Transcripts will be saved to Downloads/${downloadSubfolder}/`
                    : "Transcripts will be saved straight into Downloads")
            })
        })
    }

    setupParticipantNameSetting()
    setupSettingsTransfer()
    setupTranscriptImport()
    setupRetentionSetting()
    setupMeetingsFilter()
    setupSummaryPrompt()
    showRecentErrors()
    setupObsidianSettings()

    setupViews()

    // handlePermissions()
})


/** Request runtime permission for webhook URL */
function requestWebhookAndNotificationPermission(url: string): Promise<string> {
    return new Promise((resolve, reject) => {
        try {
            const urlObj = new URL(url)
            const originPattern = `${urlObj.protocol}//${urlObj.hostname}/*`

            // Request both host and notifications permissions
            chrome.permissions.request({
                origins: [originPattern],
                permissions: ["notifications"]
            }).then((granted) => {
                if (granted) {
                    resolve("Permission granted")
                } else {
                    reject(new Error("Permission denied"))
                }
            }).catch((error) => {
                reject(error)
            })
        } catch (error) {
            reject(error)
        }
    })
}

/**
 * Shows a meeting's transcript and chat in a dialog, so it can be read without
 * downloading the text file first.
 *
 * Everything from the meeting is written with textContent rather than innerHTML:
 * transcript text is captured verbatim from the meeting UI and must never be parsed
 * as markup here.
 */
function openTranscriptDialog(meeting: Meeting, index: number, focus: "transcript" | "summary" = "transcript"): void {
    const dialog = document.querySelector<HTMLDialogElement>("#transcript-dialog")
    const titleElement = document.querySelector<HTMLElement>("#transcript-dialog-title")
    const metaElement = document.querySelector<HTMLElement>("#transcript-dialog-meta")
    const bodyElement = document.querySelector<HTMLElement>("#transcript-dialog-body")
    const closeButton = document.querySelector<HTMLButtonElement>("#transcript-dialog-close")
    const copyButton = document.querySelector<HTMLButtonElement>("#transcript-dialog-copy")
    const downloadButton = document.querySelector<HTMLButtonElement>("#transcript-dialog-download")
    const obsidianButton = document.querySelector<HTMLButtonElement>("#transcript-dialog-obsidian")

    if (!dialog || !titleElement || !metaElement || !bodyElement) {
        return
    }

    const meetingName = meeting.meetingTitle || meeting.title || "Meeting"
    titleElement.textContent = focus === "summary" ? `${meetingName} — summary` : meetingName
    metaElement.textContent = [
        meeting.meetingSoftware,
        formatWhen(meeting.meetingStartTimestamp),
        getDuration(meeting.meetingStartTimestamp, meeting.meetingEndTimestamp)
    ].filter(part => part).join("  ·  ")

    bodyElement.innerHTML = ""
    if (meeting.summary) {
        const heading = document.createElement("h3")
        heading.textContent = "Summary"
        bodyElement.appendChild(heading)

        // A failed attempt keeps the previous summary rather than destroying one that may
        // have been good, so the reader has to be told which one they are looking at. This
        // is the screen where a truncated summary was read as the real one.
        if (meeting.summaryStatus === "failed") {
            const warning = document.createElement("p")
            warning.className = "status status-failed dialog-warning"
            warning.textContent = "The last attempt to summarise failed. This is from an earlier attempt and may be cut short — summarise again to replace it."
            bodyElement.appendChild(warning)
        }

        bodyElement.appendChild(buildSummaryDocument(meeting.summary))
    }
    else if (focus === "summary") {
        const empty = document.createElement("p")
        empty.className = "sub-text"
        empty.textContent = "This meeting has not been summarised yet."
        bodyElement.appendChild(empty)
    }
    bodyElement.appendChild(buildSection("Transcript", meeting.transcript.map(block => ({
        personName: block.personName,
        timestamp: block.timestamp,
        text: block.transcriptText
    })), "No transcript was captured for this meeting."))
    bodyElement.appendChild(buildSection("Chat messages", meeting.chatMessages.map(message => ({
        personName: message.personName,
        timestamp: message.timestamp,
        text: message.chatMessageText
    })), "No chat messages were captured for this meeting."))
    bodyElement.scrollTop = 0

    // Listeners are replaced per open, so the buttons always act on the meeting on screen
    if (closeButton) {
        closeButton.onclick = () => dialog.close()
    }

    // Click outside to dismiss. A click on the backdrop reports the dialog itself as the
    // target, and the dialog has no padding of its own, so nothing inside can match.
    // The closedby="any" attribute does this natively but only on very recent browsers.
    dialog.onclick = (event) => {
        if (event.target === dialog) {
            dialog.close()
        }
    }
    if (copyButton) {
        copyButton.textContent = "Copy"
        copyButton.onclick = () => {
            const toCopy = (focus === "summary" && meeting.summary) ? meeting.summary : getPlainText(meeting)
            navigator.clipboard.writeText(toCopy).then(() => {
                copyButton.textContent = "Copied!"
            }).catch((error) => {
                copyButton.textContent = "Copy failed"
                console.error("Could not copy transcript:", error)
            })
        }
    }
    if (obsidianButton) {
        // Follows whichever note is on screen, like the download action beside it
        const savesSummary = (focus === "summary") && !!meeting.summary
        obsidianButton.textContent = savesSummary ? "Save summary to Obsidian" : "Save to Obsidian"
        obsidianButton.disabled = false
        obsidianButton.onclick = () => {
            obsidianButton.disabled = true
            obsidianButton.textContent = "Saving..."

            const message: ExtensionMessage = {
                type: savesSummary ? "post_summary_obsidian_at_index" : "post_obsidian_at_index",
                index: index
            }
            chrome.runtime.sendMessage(message, (responseUntyped) => {
                const response = responseUntyped as ExtensionResponse
                obsidianButton.disabled = false

                if (response.success) {
                    obsidianButton.textContent = "Saved to Obsidian"
                    loadMeetings()
                    return
                }

                obsidianButton.textContent = savesSummary ? "Save summary to Obsidian" : "Save to Obsidian"
                const parsedError = response.message as ErrorObject
                alert(parsedError?.errorCode === "018"
                    ? "Add your Obsidian API key in the settings below first."
                    : (parsedError?.errorMessage ?? "Could not save to Obsidian."))
            })
        }
    }

    if (downloadButton) {
        const wantsSummary = (focus === "summary") && !!meeting.summary
        downloadButton.textContent = wantsSummary ? "Download summary" : "Download"
        downloadButton.onclick = () => {
            const message: ExtensionMessage = {
                type: wantsSummary ? "download_summary_at_index" : "download_transcript_at_index",
                index: index,
                format: document.querySelector<HTMLSelectElement>("#transcript-dialog-format")?.value === "html" ? "html" : "txt"
            }
            chrome.runtime.sendMessage(message, (responseUntyped) => {
                const response = responseUntyped as ExtensionResponse
                if (!response.success) {
                    alert(wantsSummary ? "Could not download summary" : "Could not download transcript")
                }
            })
        }
    }

    dialog.showModal()
}

/** One titled group of speaker blocks, or a placeholder when there is nothing to show. */
function buildSection(heading: string, blocks: { personName: string, timestamp: string, text: string }[], emptyMessage: string): DocumentFragment {
    const fragment = document.createDocumentFragment()

    const headingElement = document.createElement("h3")
    headingElement.textContent = heading
    fragment.appendChild(headingElement)

    if (blocks.length === 0) {
        const empty = document.createElement("p")
        empty.className = "sub-text"
        empty.textContent = emptyMessage
        fragment.appendChild(empty)
        return fragment
    }

    blocks.forEach(block => {
        const wrapper = document.createElement("div")
        wrapper.className = "dialog-block"

        const speaker = document.createElement("span")
        speaker.className = "dialog-block-speaker"
        speaker.textContent = block.personName

        const time = document.createElement("span")
        time.className = "dialog-block-time"
        time.textContent = new Date(block.timestamp).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })

        const text = document.createElement("p")
        text.className = "dialog-block-text"
        text.textContent = block.text

        wrapper.appendChild(speaker)
        wrapper.appendChild(time)
        wrapper.appendChild(text)
        fragment.appendChild(wrapper)
    })

    return fragment
}

/** The same plain text layout as the downloaded file, for the clipboard. */
function getPlainText(meeting: Meeting): string {
    let content = ""
    meeting.transcript.forEach(block => {
        content += `${block.personName} (${new Date(block.timestamp).toLocaleString()})\n${block.transcriptText}\n\n`
    })
    content += `\n---------------\nCHAT MESSAGES\n---------------\n\n`
    meeting.chatMessages.forEach(message => {
        content += `${message.personName} (${new Date(message.timestamp).toLocaleString()})\n${message.chatMessageText}\n\n`
    })
    return content
}

/**
 * Brings exported .txt transcripts back in as meetings, so files from another computer
 * (or from before a reinstall) can be summarised and filed to Obsidian like any other.
 */
function setupTranscriptImport(): void {
    const importButton = document.querySelector<HTMLButtonElement>("#import-transcript")
    const fileInput = document.querySelector<HTMLInputElement>("#import-transcript-file")
    const result = document.querySelector<HTMLElement>("#import-transcript-result")

    if (!importButton || !fileInput) {
        return
    }

    const report = (message: string, isError = false) => {
        if (!result) return
        result.textContent = message
        result.style.color = isError ? "var(--danger)" : "var(--success)"
    }

    importButton.addEventListener("click", () => fileInput.click())

    fileInput.addEventListener("change", async function () {
        const files = [...(fileInput.files ?? [])]
        if (files.length === 0) return

        const imported: Meeting[] = []
        const failures: string[] = []

        for (const file of files) {
            try {
                imported.push(parseTranscriptFile(file.name, await file.text()))
            }
            catch (error) {
                failures.push(`${file.name}: ${error instanceof Error ? error.message : "could not be read"}`)
            }
        }

        fileInput.value = ""

        if (imported.length === 0) {
            report(failures[0] ?? "Nothing could be imported.", true)
            return
        }

        chrome.storage.local.get(["meetings"], function (resultLocalUntyped) {
            const resultLocal = resultLocalUntyped as Partial<ResultLocal>
            const existing = resultLocal.meetings ?? []

            // A file already in the list is an update, not a second copy
            const merged = [...existing]
            let added = 0
            let updated = 0
            for (const meeting of imported) {
                const match = merged.findIndex(m => m.meetingStartTimestamp === meeting.meetingStartTimestamp
                    && (m.meetingTitle || m.title) === meeting.meetingTitle)
                if (match >= 0) {
                    // Keep whatever the existing entry already earned
                    merged[match] = { ...meeting, summary: merged[match].summary, summaryStatus: merged[match].summaryStatus }
                    updated++
                }
                else {
                    merged.push(meeting)
                    added++
                }
            }

            merged.sort((a, b) => a.meetingStartTimestamp.localeCompare(b.meetingStartTimestamp))

            // Only the newest 10 are kept, so say plainly when that pushes something out
            // rather than letting meetings disappear silently.
            const dropped = Math.max(0, merged.length - 10)
            const kept = merged.slice(-10)

            chrome.storage.local.set({ meetings: kept }, function () {
                loadMeetings()
                const parts: string[] = []
                if (added > 0) parts.push(`${added} imported`)
                if (updated > 0) parts.push(`${updated} updated`)
                if (failures.length > 0) parts.push(`${failures.length} skipped`)
                let message = parts.join(", ") + "."
                if (dropped > 0) {
                    message += ` ${dropped} older meeting${dropped === 1 ? "" : "s"} dropped to keep the most recent 10.`
                }
                report(message, failures.length > 0 && added === 0)
            })
        })
    })
}

/**
 * Export and import of the whole configuration, so a second computer can be set up
 * without repeating every field by hand.
 */
function setupSettingsTransfer(): void {
    const exportButton = document.querySelector<HTMLButtonElement>("#export-settings")
    const importButton = document.querySelector<HTMLButtonElement>("#import-settings")
    const fileInput = document.querySelector<HTMLInputElement>("#import-settings-file")
    const includeKeyToggle = document.querySelector<HTMLInputElement>("#export-include-key")
    const result = document.querySelector<HTMLElement>("#transfer-result")

    if (!exportButton || !importButton || !fileInput) {
        return
    }

    const report = (message: string, isError = false) => {
        if (!result) return
        result.textContent = message
        result.style.color = isError ? "var(--danger)" : "var(--success)"
    }

    exportButton.addEventListener("click", function () {
        chrome.storage.sync.get([...TRANSFERABLE_SYNC_KEYS], function (resultSyncUntyped) {
            const settings = resultSyncUntyped as Partial<ResultSync>

            chrome.storage.local.get(["obsidianApiKey"], function (resultLocalUntyped) {
                const resultLocal = resultLocalUntyped as Partial<ResultLocal>
                const includeKey = includeKeyToggle?.checked !== false
                const file = buildSettingsFile(settings, includeKey ? resultLocal.obsidianApiKey : undefined)

                // Built in the page rather than the service worker so the file lands
                // through the browser's own save flow
                const blob = new Blob([JSON.stringify(file, null, 2)], { type: "application/json" })
                const url = URL.createObjectURL(blob)
                const link = document.createElement("a")
                link.href = url
                link.download = `meetingtranscripttoobsidian-settings-${new Date().toISOString().slice(0, 10)}.json`
                link.click()
                URL.revokeObjectURL(url)

                report(file.obsidianApiKey
                    ? "Exported, including the Obsidian API key. Keep the file private."
                    : "Exported. The Obsidian API key was left out, so you will need to paste it on the other computer.")
            })
        })
    })

    importButton.addEventListener("click", () => fileInput.click())

    fileInput.addEventListener("change", function () {
        const file = fileInput.files?.[0]
        if (!file) return

        file.text().then((raw) => {
            const parsed = parseSettingsFile(raw)
            // A transferred setup is being set up to file notes, so Obsidian comes on
            const settings = withObsidianEnabled(parsed.settings)

            const origins = [`${new URL(settings.obsidianBaseUrl || DEFAULT_OBSIDIAN_BASE_URL).origin}/*`]
            const platforms = getWantedPlatforms(settings)

            // One prompt covering everything the imported setup needs. This runs inside
            // the file-picker's change event, which still counts as a user gesture.
            chrome.permissions.request({ origins }).catch(() => false).then(() => {
                chrome.storage.sync.set(settings, function () {
                    const finish = () => {
                        // Re-register content scripts for the platforms the file asked for
                        if (platforms.length > 0) {
                            chrome.runtime.sendMessage({ type: "enable_platform", platform: platforms } satisfies ExtensionMessage, () => {
                                report("Settings imported and Obsidian switched on. Reloading...")
                                setTimeout(() => location.reload(), 1200)
                            })
                        }
                        else {
                            report("Settings imported and Obsidian switched on. Reloading...")
                            setTimeout(() => location.reload(), 1200)
                        }
                    }

                    if (parsed.obsidianApiKey) {
                        chrome.storage.local.set({ obsidianApiKey: parsed.obsidianApiKey }, finish)
                    }
                    else {
                        finish()
                    }
                })
            })
        }).catch((error) => {
            report(error instanceof Error ? error.message : "Could not read that file.", true)
        }).finally(() => {
            // Allow re-importing the same file
            fileInput.value = ""
        })
    })
}

/** The name written into transcripts wherever captions say "You". */
function setupParticipantNameSetting(): void {
    const form = document.querySelector<HTMLFormElement>("#participant-name-form")
    const input = document.querySelector<HTMLInputElement>("#participant-name")
    if (!form || !input) {
        return
    }

    chrome.storage.sync.get(["participantName"], function (resultSyncUntyped) {
        const resultSync = resultSyncUntyped as Partial<ResultSync>
        input.value = resultSync.participantName ?? ""
    })

    form.addEventListener("submit", function (e) {
        e.preventDefault()
        const participantName = input.value.trim()
        chrome.storage.sync.set({ participantName: participantName }, function () {
            input.value = participantName
            alert(participantName
                ? `Your speech will be attributed to "${participantName}" from the next meeting`
                : `Your speech will be attributed to "You" again`)
        })
    })
}

/**
 * Obsidian settings: enable toggle, API key, vault folder and the plugin address.
 *
 * The API key lives in chrome.storage.local rather than sync, so the credential never
 * leaves the device — the rest of the settings sync normally.
 */
function setupObsidianSettings(): void {
    const enabledToggle = document.querySelector<HTMLInputElement>("#obsidian-enabled")
    const apiKeyForm = document.querySelector<HTMLFormElement>("#obsidian-api-key-form")
    const apiKeyInput = document.querySelector<HTMLInputElement>("#obsidian-api-key")
    const folderForm = document.querySelector<HTMLFormElement>("#obsidian-folder-form")
    const folderInput = document.querySelector<HTMLInputElement>("#obsidian-folder")
    const summaryFolderForm = document.querySelector<HTMLFormElement>("#obsidian-summary-folder-form")
    const summaryFolderInput = document.querySelector<HTMLInputElement>("#obsidian-summary-folder")
    const baseUrlForm = document.querySelector<HTMLFormElement>("#obsidian-base-url-form")
    const baseUrlInput = document.querySelector<HTMLInputElement>("#obsidian-base-url")
    const testButton = document.querySelector<HTMLButtonElement>("#test-obsidian")
    const syncAllButton = document.querySelector<HTMLButtonElement>("#sync-all-obsidian")
    const testResult = document.querySelector<HTMLElement>("#obsidian-test-result")

    chrome.storage.sync.get(["obsidianEnabled", "obsidianFolder", "obsidianSummaryFolder", "obsidianBaseUrl"], function (resultSyncUntyped) {
        const resultSync = resultSyncUntyped as Partial<ResultSync>

        if (enabledToggle) {
            enabledToggle.checked = resultSync.obsidianEnabled === true
        }
        if (folderInput) {
            // An empty string is a real choice ("vault root"), so only fall back to the
            // default when nothing has ever been stored.
            folderInput.value = typeof resultSync.obsidianFolder === "string"
                ? resultSync.obsidianFolder
                : DEFAULT_OBSIDIAN_FOLDER
        }
        if (summaryFolderInput) {
            summaryFolderInput.value = typeof resultSync.obsidianSummaryFolder === "string"
                ? resultSync.obsidianSummaryFolder
                : DEFAULT_OBSIDIAN_SUMMARY_FOLDER
        }
        if (baseUrlInput) {
            baseUrlInput.value = resultSync.obsidianBaseUrl || DEFAULT_OBSIDIAN_BASE_URL
        }
    })

    chrome.storage.local.get(["obsidianApiKey"], function (resultLocalUntyped) {
        const resultLocal = resultLocalUntyped as Partial<ResultLocal>
        if (apiKeyInput && resultLocal.obsidianApiKey) {
            apiKeyInput.value = resultLocal.obsidianApiKey
        }
    })

    enabledToggle?.addEventListener("change", function () {
        if (!enabledToggle.checked) {
            chrome.storage.sync.set({ obsidianEnabled: false }, function () { })
            return
        }

        // Turning it on needs host permission for the plugin's address. Requesting it
        // here keeps the prompt inside a user gesture, which Chrome requires.
        requestObsidianPermission(baseUrlInput?.value || DEFAULT_OBSIDIAN_BASE_URL).then(() => {
            chrome.storage.sync.set({ obsidianEnabled: true }, function () { })
        }).catch((error) => {
            enabledToggle.checked = false
            alert("MeetingTranscriptToObsidian needs permission to talk to Obsidian on this machine.")
            console.error("Obsidian permission error:", error)
        })
    })

    apiKeyForm?.addEventListener("submit", function (e) {
        e.preventDefault()
        if (!apiKeyInput) return
        chrome.storage.local.set({ obsidianApiKey: apiKeyInput.value.trim() }, function () {
            alert(apiKeyInput.value.trim() ? "Obsidian API key saved!" : "Obsidian API key cleared")
        })
    })

    folderForm?.addEventListener("submit", function (e) {
        e.preventDefault()
        if (!folderInput) return
        // Store the cleaned value so the field always shows what will actually be used
        const obsidianFolder = getFolderPrefix(folderInput.value, DEFAULT_OBSIDIAN_FOLDER).replace(/\/$/, "")
        chrome.storage.sync.set({ obsidianFolder: obsidianFolder }, function () {
            folderInput.value = obsidianFolder
            alert(obsidianFolder
                ? `Meeting notes will be saved to ${obsidianFolder}/ in your vault`
                : "Meeting notes will be saved to your vault root")
        })
    })

    summaryFolderForm?.addEventListener("submit", function (e) {
        e.preventDefault()
        if (!summaryFolderInput) return
        const obsidianSummaryFolder = getFolderPrefix(summaryFolderInput.value, DEFAULT_OBSIDIAN_SUMMARY_FOLDER).replace(/\/$/, "")
        chrome.storage.sync.set({ obsidianSummaryFolder: obsidianSummaryFolder }, function () {
            summaryFolderInput.value = obsidianSummaryFolder
            alert(obsidianSummaryFolder
                ? `Summaries will be saved to ${obsidianSummaryFolder}/ in your vault`
                : "Summaries will be saved to your vault root")
        })
    })

    baseUrlForm?.addEventListener("submit", function (e) {
        e.preventDefault()
        if (!baseUrlInput) return
        const obsidianBaseUrl = baseUrlInput.value.trim() || DEFAULT_OBSIDIAN_BASE_URL
        requestObsidianPermission(obsidianBaseUrl).then(() => {
            chrome.storage.sync.set({ obsidianBaseUrl: obsidianBaseUrl }, function () {
                baseUrlInput.value = obsidianBaseUrl
                alert("Obsidian address saved!")
            })
        }).catch((error) => {
            alert("MeetingTranscriptToObsidian needs permission to talk to Obsidian at that address.")
            console.error("Obsidian permission error:", error)
        })
    })

    syncAllButton?.addEventListener("click", function () {
        if (!testResult) return
        const originalLabel = syncAllButton.textContent
        syncAllButton.disabled = true
        syncAllButton.textContent = "Sending..."
        testResult.textContent = "Writing meetings into your vault..."

        chrome.runtime.sendMessage({ type: "sync_all_to_obsidian" } satisfies ExtensionMessage, (responseUntyped) => {
            const response = responseUntyped as ExtensionResponse
            syncAllButton.disabled = false
            syncAllButton.textContent = originalLabel

            if (!response.success) {
                const parsedError = response.message as ErrorObject
                testResult.textContent = parsedError?.errorCode === "018"
                    ? "Add your Obsidian API key above first."
                    : (parsedError?.errorMessage ?? "Could not write to your vault.")
                return
            }

            const result = JSON.parse(String(response.message)) as {
                transcripts: number, summaries: number, skipped: number, failed: number, firstError?: string
            }
            const parts = [`${result.transcripts} transcript${result.transcripts === 1 ? "" : "s"}`]
            if (result.summaries > 0) {
                parts.push(`${result.summaries} summar${result.summaries === 1 ? "y" : "ies"}`)
            }
            if (result.skipped > 0) {
                parts.push(`${result.skipped} already in the vault, left untouched`)
            }
            testResult.textContent = result.failed > 0
                ? `Sent ${parts.join(", ")}. ${result.failed} could not be written: ${result.firstError ?? ""}`
                : (result.transcripts === 0 && result.summaries === 0
                    ? `Everything is already in your vault. Existing notes were not overwritten, so any edits you made are safe.`
                    : `Sent ${parts.join(", ")} to your vault.`)
            loadMeetings()
        })
    })

    testButton?.addEventListener("click", function () {
        if (!testResult) return
        const baseUrl = (baseUrlInput?.value.trim() || DEFAULT_OBSIDIAN_BASE_URL).replace(/\/+$/, "")
        const apiKey = apiKeyInput?.value.trim() || ""

        testResult.textContent = "Checking..."
        requestObsidianPermission(baseUrl)
            .then(() => fetch(`${baseUrl}/`, { headers: { "Authorization": `Bearer ${apiKey}` } }))
            .then(response => response.json())
            .then(result => {
                testResult.textContent = result.authenticated
                    ? `Connected to Obsidian (Local REST API ${result.versions?.self ?? ""}).`
                    : "Obsidian is reachable, but the API key was not accepted. Check the key and save it."
            })
            .catch(() => {
                testResult.textContent = "Could not reach Obsidian. Make sure the app is running and the plugin's non-encrypted (HTTP) server is enabled."
            })
    })
}

/** Request host permission for ChatGPT, needed to drive its page. */
function requestChatGptPermission(): Promise<string> {
    return new Promise((resolve, reject) => {
        chrome.permissions.request({ origins: ["https://chatgpt.com/*"] })
            .then((granted) => {
                granted ? resolve("Permission granted") : reject(new Error("Permission denied"))
            })
            .catch(reject)
    })
}

/** Request host permission for the Obsidian plugin's origin. */
function requestObsidianPermission(baseUrl: string): Promise<string> {
    return new Promise((resolve, reject) => {
        try {
            const urlObj = new URL(baseUrl)
            chrome.permissions.request({
                origins: [`${urlObj.protocol}//${urlObj.hostname}/*`]
            }).then((granted) => {
                granted ? resolve("Permission granted") : reject(new Error("Permission denied"))
            }).catch(reject)
        } catch (error) {
            reject(error)
        }
    })
}

/** The instructions sent to ChatGPT, editable so tone, sections and language can change. */
function setupSummaryPrompt(): void {
    const form = document.querySelector<HTMLFormElement>("#summary-prompt-form")
    const field = document.querySelector<HTMLTextAreaElement>("#summary-prompt")
    const reset = document.querySelector<HTMLButtonElement>("#reset-summary-prompt")
    if (!form || !field) return

    chrome.storage.sync.get(["summaryPrompt"], function (resultSyncUntyped) {
        const resultSync = resultSyncUntyped as Partial<ResultSync>
        field.value = resultSync.summaryPrompt || DEFAULT_SUMMARY_INSTRUCTIONS
    })

    form.addEventListener("submit", function (e) {
        e.preventDefault()
        chrome.storage.sync.set({ summaryPrompt: field.value.trim() }, function () {
            alert("Summary instructions saved. They apply to the next summary.")
        })
    })

    reset?.addEventListener("click", function () {
        field.value = DEFAULT_SUMMARY_INSTRUCTIONS
        chrome.storage.sync.set({ summaryPrompt: "" }, function () {
            alert("Back to the default instructions.")
        })
    })

    setupAutoSummarize()
}

/**
 * The switch that turns every meeting's end into a summary, with no button press.
 *
 * Turning it on is treated as the moment of consent, because it is: it is the last
 * point at which the user is asked before transcripts start leaving the device on
 * their own. The host permission is requested from inside the change handler so it
 * still counts as a user gesture, and the box goes back to unchecked if either the
 * confirmation or the permission is refused — a checkbox that stays ticked after a
 * refusal would promise something that cannot happen.
 */
function setupAutoSummarize(): void {
    const toggle = document.querySelector<HTMLInputElement>("#auto-summarize")
    if (!toggle) return

    chrome.storage.sync.get(["autoSummarizeAfterMeeting"], function (resultSyncUntyped) {
        const resultSync = resultSyncUntyped as Partial<ResultSync>
        toggle.checked = resultSync.autoSummarizeAfterMeeting === true
    })

    toggle.addEventListener("change", function () {
        if (!toggle.checked) {
            chrome.storage.sync.set({ autoSummarizeAfterMeeting: false })
            return
        }

        if (!confirm("Every meeting will be sent to ChatGPT automatically when it ends, where OpenAI's terms and data handling apply.\n\nA ChatGPT tab will come to the front for a few seconds each time, then hand your place back.\n\nTurn this on?")) {
            toggle.checked = false
            return
        }

        requestChatGptPermission()
            .then(() => {
                chrome.storage.sync.set({ autoSummarizeAfterMeeting: true })
            })
            .catch(() => {
                toggle.checked = false
                alert("MeetingTranscriptToObsidian needs access to chatgpt.com to summarise meetings.")
            })
    })
}

/**
 * Surfaces capture failures recorded on this device.
 *
 * Without this the only record is a service worker console nobody opens, which leaves a
 * user whose transcript broke with nothing to look at or report.
 */
function showRecentErrors(): void {
    const panel = document.querySelector<HTMLDetailsElement>("#recent-errors")
    const count = document.querySelector<HTMLElement>("#recent-errors-count")
    const list = document.querySelector<HTMLElement>("#recent-errors-list")
    if (!panel || !count || !list) return

    chrome.storage.local.get(["recentErrors"], function (resultUntyped) {
        const errors = (resultUntyped as Partial<ResultLocal>).recentErrors ?? []
        if (errors.length === 0) {
            panel.hidden = true
            return
        }
        panel.hidden = false
        count.textContent = `${errors.length} capture problem${errors.length === 1 ? "" : "s"} recorded on this device`
        list.textContent = errors
            .map(e => `${new Date(e.at).toLocaleString()}  [${e.code}]  ${e.meetingSoftware || "unknown"}\n    ${e.message}`)
            .join("\n")
    })
}

/** How many meetings are kept before the oldest are deleted. */
function setupRetentionSetting(): void {
    const form = document.querySelector<HTMLFormElement>("#meetings-to-keep-form")
    const input = document.querySelector<HTMLInputElement>("#meetings-to-keep")
    if (!form || !input) return

    chrome.storage.sync.get(["meetingsToKeep"], function (resultSyncUntyped) {
        const resultSync = resultSyncUntyped as Partial<ResultSync>
        input.value = String(typeof resultSync.meetingsToKeep === "number"
            ? resultSync.meetingsToKeep
            : DEFAULT_MEETINGS_TO_KEEP)
    })

    form.addEventListener("submit", function (e) {
        e.preventDefault()
        const requested = Math.max(0, Math.min(500, Math.floor(Number(input.value) || 0)))

        chrome.storage.local.get(["meetings"], function (resultLocalUntyped) {
            const stored = (resultLocalUntyped as Partial<ResultLocal>).meetings ?? []

            // Lowering the limit deletes meetings immediately, so say what will go and
            // let the user back out. Silent deletion is what this setting exists to fix.
            const wouldDelete = meetingsDroppedBy(stored, requested)
            if (wouldDelete > 0 && !confirm(
                `This removes ${wouldDelete} older meeting${wouldDelete === 1 ? "" : "s"} from the extension's list now.\n\n` +
                `Obsidian notes and downloaded files are not affected — only the copy held in the browser.\n\nContinue?`)) {
                return
            }

            chrome.storage.sync.set({ meetingsToKeep: requested }, function () {
                if (wouldDelete > 0) {
                    chrome.storage.local.set({ meetings: trimMeetings(stored, requested) }, loadMeetings)
                }
                input.value = String(requested)
                alert(requested === 0
                    ? "Every meeting will be kept in the extension."
                    : `Keeping the most recent ${requested} meetings in the extension. Obsidian notes and downloaded files are unaffected.`)
            })
        })
    })
}

/** Filters the table by title, platform or date, without touching stored data. */
function setupMeetingsFilter(): void {
    const filter = document.querySelector<HTMLInputElement>("#meetings-filter")
    if (!filter) return

    filter.addEventListener("input", function () {
        const term = filter.value.trim().toLowerCase()
        const rows = document.querySelectorAll<HTMLTableRowElement>("#meetings-table tr")
        let shown = 0

        rows.forEach(row => {
            const match = !term || (row.innerText || "").toLowerCase().includes(term)
            row.hidden = !match
            if (match) shown++
        })

        const empty = document.querySelector<HTMLElement>("#meetings-filter-empty")
        if (empty) empty.hidden = !(term && shown === 0)
    })
}

/**
 * Escapes a value for interpolation into markup.
 *
 * Meeting titles are user-controlled three ways — renamed inline, taken from an imported
 * file name, and read from the meeting page's own title — so they cannot be trusted into
 * innerHTML. MV3's CSP blocks inline script, but unescaped markup can still spoof the
 * table's controls and fetch a remote URL through an injected tag.
 */
function escapeHtml(value: string): string {
    return value.replace(/[&<>"']/g, (ch) => (
        { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[ch] as string
    ))
}

/** Load and display recent transcripts */
function loadMeetings(): void {
    const meetingsTable = document.querySelector("#meetings-table")

    chrome.storage.local.get(["meetings"], function (resultLocalUntyped) {
        const resultLocal = resultLocalUntyped as Partial<ResultLocal>
        // Clear existing content
        if (meetingsTable) {
            meetingsTable.innerHTML = ""


            if (resultLocal.meetings && resultLocal.meetings.length > 0) {
                const meetings = resultLocal.meetings
                // Loop through the array in reverse order to list latest meeting first
                for (let i = meetings.length - 1; i >= 0; i--) {
                    const meeting = meetings[i]
                    const durationString = getDuration(meeting.meetingStartTimestamp, meeting.meetingEndTimestamp)

                    const row = document.createElement("tr")
                    row.innerHTML = `
                    <td>
                        <div contenteditable="true" class="meeting-title" data-index="${i}" title="Rename">${escapeHtml(meeting.meetingTitle || meeting.title || "Google Meet call")}</div>
                        <div class="meeting-meta">${[meeting.meetingSoftware ?? "", formatWhen(meeting.meetingStartTimestamp), durationString]
                            .filter(part => part).map(escapeHtml).join(" &middot; ")}${
                            meeting.transcript.length === 0 && meeting.captionsSeen === false
                                ? `<span class="status status-failed" title="The captions panel never appeared, so nothing could be captured">Captions never started</span>`
                                : meeting.transcript.length === 0
                                    ? `<span class="status status-new" title="Capture ran, but nobody spoke with captions on">No speech captured</span>`
                                    : ``
                        }</div>
                    </td>
                    <td>
                        <div class="stack">
                            ${meeting.summary ? `<button class="view-summary-button" data-index="${i}" title="${
                                // A failed attempt leaves the previous summary in place rather than
                                // destroying one that may have been good. That is only safe if the
                                // row says so — otherwise a stale, half-written summary reads as
                                // the current one, which is exactly how a truncated summary went
                                // unnoticed.
                                meeting.summaryStatus === "failed"
                                    ? `The last attempt to summarise failed — this is from an earlier one`
                                    : `View summary`
                            }" aria-label="View this meeting summary">
                                View${meeting.summaryStatus === "failed" ? ` <span class="status status-failed">stale</span>` : ``}
                            </button>
                            <button class="download-summary-html-button" data-index="${i}" title="Download summary as HTML" aria-label="Download this meeting summary as HTML">Download summary HTML</button>` : ``}
                            <button class="summarize-button${meeting.summary ? ` is-redo` : ``}" data-index="${i}" title="${meeting.summary ? `Summarise again with ChatGPT` : `Summarise with ChatGPT`}" aria-label="${meeting.summary ? `Summarise this meeting again with ChatGPT` : `Summarise this meeting with ChatGPT`}">
                                ${meeting.summary ? `Redo` : `Summarise`}
                            </button>
                        </div>
                    </td>
                    <td>
                        <div class="stack">
                            ${(
                                () => {
                                    switch (meeting.webhookPostStatus) {
                                        case "successful":
                                            return `<span class="status status-success" title="Posted to your webhook">Webhook</span>`
                                        case "failed":
                                            return `<span class="status status-failed" title="Posting to your webhook failed — use the webhook button to retry">Webhook failed</span>`
                                        default:
                                            return ``
                                    }
                                }
                            )()}
                            ${(
                                () => {
                                    switch (meeting.obsidianPostStatus) {
                                        case "successful":
                                            return `<span class="status status-success" title="Saved as a note in your vault">Obsidian</span>`
                                        case "failed":
                                            return `<span class="status status-failed">Obsidian failed</span> <button class="obsidian-button" data-index="${i}" title="Retry saving to Obsidian">Retry</button>`
                                        default:
                                            return `<button class="obsidian-button" data-index="${i}" title="Save this meeting to Obsidian">Save to Obsidian</button>`
                                    }
                                }
                            )()}
                        </div>
                    </td>
                    <td>
                        <div class="row-actions">
                            <button class="view-button" data-index="${i}" title="View transcript" aria-label="View this meeting transcript">
                                <span class="icon icon-view"></span>
                            </button>
                            <button class="download-button" data-index="${i}" title="Download" aria-label="Download this meeting transcript">
                                <span class="icon icon-download"></span>
                            </button>
                            <button class="post-button" data-index="${i}" title="${meeting.webhookPostStatus === "new" ? `Post to webhook` : `Post to webhook again`}" aria-label="${meeting.webhookPostStatus === "new" ? `Post this meeting to the webhook` : `Post this meeting to the webhook again`}">
                                <span class="icon icon-webhook"></span>
                            </button>
                            <button class="delete-button" data-index="${i}" title="Delete" aria-label="Delete this meeting">
                                <span class="icon icon-delete"></span>
                            </button>
                        </div>
                    </td>
                `
                    meetingsTable.appendChild(row)

                    // Add event listener to meeting title input
                    const meetingTitleInput = row.querySelector(".meeting-title")
                    if (meetingTitleInput instanceof HTMLDivElement) {
                        meetingTitleInput.addEventListener("blur", function () {
                            const updatedMeeting: Meeting = {
                                ...meeting,
                                meetingTitle: meetingTitleInput.innerText
                            }
                            meetings[i] = updatedMeeting
                            chrome.storage.local.set({ meetings: meetings }, function () {
                                console.log("Meeting title updated")
                            })
                        })
                    }

                    // Add event listener to the view button
                    const viewButton = row.querySelector(".view-button")
                    if (viewButton instanceof HTMLButtonElement) {
                        viewButton.addEventListener("click", function () {
                            openTranscriptDialog(meeting, i)
                        })
                    }

                    // Add event listener to the view summary button
                    const viewSummaryButton = row.querySelector(".view-summary-button")
                    if (viewSummaryButton instanceof HTMLButtonElement) {
                        viewSummaryButton.addEventListener("click", function () {
                            openTranscriptDialog(meeting, i, "summary")
                        })
                    }

                    const downloadSummaryHtmlButton = row.querySelector<HTMLButtonElement>(".download-summary-html-button")
                    if (downloadSummaryHtmlButton) {
                        downloadSummaryHtmlButton.addEventListener("click", function () {
                            downloadSummaryHtmlButton.disabled = true
                            const message: ExtensionMessage = {
                                type: "download_summary_at_index",
                                index: i,
                                format: "html"
                            }
                            chrome.runtime.sendMessage(message, (responseUntyped) => {
                                downloadSummaryHtmlButton.disabled = false
                                const response = responseUntyped as ExtensionResponse | undefined
                                if (chrome.runtime.lastError || !response?.success) {
                                    const error = response?.message as ErrorObject | undefined
                                    alert(error?.errorMessage || chrome.runtime.lastError?.message || "Could not download summary as HTML")
                                }
                            })
                        })
                    }

                    // Add event listener to the summarise button
                    const summarizeButton = row.querySelector(".summarize-button")
                    if (summarizeButton instanceof HTMLButtonElement) {
                        summarizeButton.addEventListener("click", function () {
                            // Said here rather than only in settings: this is the moment
                            // the transcript leaves the device, and it is the only feature
                            // that sends it to a third party.
                            if (!sessionStorage.getItem("chatgpt-disclosure-accepted")) {
                                if (!confirm("Summarising sends this meeting's full transcript to ChatGPT, where OpenAI's terms and data handling apply.\n\nContinue?")) {
                                    return
                                }
                                sessionStorage.setItem("chatgpt-disclosure-accepted", "1")
                            }

                            // Requested from the click so it counts as a user gesture,
                            // which Chrome requires for an optional host permission.
                            requestChatGptPermission().then(() => {
                                summarizeButton.disabled = true
                                summarizeButton.textContent = "Summarising..."

                                const index = parseInt(summarizeButton.getAttribute("data-index") ?? "-1")
                                const message: ExtensionMessage = { type: "summarize_at_index", index: index }
                                chrome.runtime.sendMessage(message, (responseUntyped) => {
                                    const response = responseUntyped as ExtensionResponse
                                    loadMeetings()
                                    if (response.success) {
                                        chrome.storage.local.get(["meetings"], function (resultUntyped) {
                                            const result = resultUntyped as Partial<ResultLocal>
                                            const updated = result.meetings?.[index]
                                            if (updated) {
                                                openTranscriptDialog(updated, index)
                                            }
                                        })
                                    }
                                    else {
                                        const parsedError = response.message as ErrorObject
                                        alert(typeof parsedError === "object" ? parsedError.errorMessage : "Could not summarise this meeting")
                                    }
                                })
                            }).catch(() => {
                                alert("MeetingTranscriptToObsidian needs access to chatgpt.com to summarise meetings.")
                            })
                        })
                    }

                    // Add event listener to the webhook post button
                    const downloadButton = row.querySelector(".download-button")
                    if (downloadButton instanceof HTMLButtonElement) {
                        downloadButton.addEventListener("click", function () {
                            // Send message to background script to download text file
                            const index = parseInt(downloadButton.getAttribute("data-index") ?? "-1")
                            const message: ExtensionMessage = {
                                type: "download_transcript_at_index",
                                index: index
                            }
                            chrome.runtime.sendMessage(message, (responseUntyped) => {
                                const response = responseUntyped as ExtensionResponse
                                if (!response.success) {
                                    alert("Could not download transcript")
                                    const parsedError = response.message as ErrorObject
                                    if (typeof parsedError === 'object') {
                                        console.error(parsedError.errorMessage)
                                    }
                                }
                            })
                        })
                    }

                    // Add event listener to the webhook post button
                    const webhookPostButton = row.querySelector(".post-button")
                    if (webhookPostButton instanceof HTMLButtonElement) {
                        webhookPostButton.addEventListener("click", function () {
                            chrome.storage.sync.get(["webhookUrl"], function (resultSyncUntyped) {
                                const resultSync = resultSyncUntyped as Partial<ResultSync>
                                if (resultSync.webhookUrl) {
                                    // Request runtime permission for the webhook URL. Needed for cases when user signs on a new browser—webhook URL and other sync variables are available, but runtime permissions will be missing.
                                    requestWebhookAndNotificationPermission(resultSync.webhookUrl).then(() => {
                                        // Disable button and update text
                                        webhookPostButton.disabled = true
                                        webhookPostButton.textContent = meeting.webhookPostStatus === "new" ? "Posting..." : "Reposting..."

                                        // Send message to background script to post webhook
                                        const index = parseInt(webhookPostButton.getAttribute("data-index") ?? "-1")
                                        const message: ExtensionMessage = {
                                            type: "post_webhook_at_index",
                                            index: index
                                        }
                                        chrome.runtime.sendMessage(message, (responseUntyped) => {
                                            const response = responseUntyped as ExtensionResponse
                                            loadMeetings()
                                            if (response.success) {
                                                alert("Posted successfully!")
                                            }
                                            else {
                                                const parsedError = response.message as ErrorObject
                                                if (typeof parsedError === 'object') {
                                                    console.error(parsedError.errorMessage)
                                                }
                                            }
                                        })
                                    }).catch((error) => {
                                        alert("Fine! No webhooks for you!")
                                        console.error("Webhook permission error:", error)
                                    })
                                }
                                else {
                                    alert("Please provide a webhook URL")
                                }
                            })
                        })
                    }

                    // Add event listener to the Obsidian save/retry button
                    const obsidianButton = row.querySelector(".obsidian-button")
                    if (obsidianButton instanceof HTMLButtonElement) {
                        obsidianButton.addEventListener("click", function () {
                            obsidianButton.disabled = true
                            obsidianButton.textContent = "Saving..."

                            const index = parseInt(obsidianButton.getAttribute("data-index") ?? "-1")
                            const message: ExtensionMessage = {
                                type: "post_obsidian_at_index",
                                index: index
                            }
                            chrome.runtime.sendMessage(message, (responseUntyped) => {
                                const response = responseUntyped as ExtensionResponse
                                loadMeetings()
                                if (!response.success) {
                                    const parsedError = response.message as ErrorObject
                                    if (typeof parsedError === "object") {
                                        alert(parsedError.errorCode === "018"
                                            ? "Add your Obsidian API key below first."
                                            : parsedError.errorMessage)
                                        console.error(parsedError.errorMessage)
                                    }
                                }
                            })
                        })
                    }

                    // Add event listener to the meeting delete button
                    const deleteButton = row.querySelector(".delete-button")
                    if (deleteButton instanceof HTMLButtonElement) {
                        deleteButton.addEventListener("click", function () {
                            // Only the extension's copy. Obsidian notes and downloaded
                            // files are the permanent record and are never touched here.
                            if (!confirm("Remove this meeting from the list?\n\nObsidian notes and downloaded files are not affected.")) {
                                return
                            }
                            meetings.splice(i, 1)
                            chrome.storage.local.set({ meetings: meetings }, loadMeetings)
                        })
                    }
                }
            }
            else {
                meetingsTable.innerHTML = `<tr><td colspan="4"><div class="empty-state"><strong>No meetings yet</strong>Join a meeting with captions on and the transcript appears here automatically.<br><br>Worth setting up first: <a href="#settings-you">your name</a>, so your own speech is not labelled &ldquo;You&rdquo;, and <a href="#settings-obsidian">Obsidian</a> or a <a href="#settings-webhook">webhook</a> if you want transcripts to land somewhere else.</div></td></tr>`
            }
        }
    })
}

// Kept for reference. Currently not wired up—see the commented call in the DOMContentLoaded handler.
export function handlePermissions(): void {
    chrome.scripting
        .getRegisteredContentScripts()
        .then((scripts) => {
            console.log(scripts)
        })

    chrome.storage.sync.get(["wantGoogleMeet", "wantTeams", "wantZoom"], function (resultSyncUntyped) {
        const resultSync = resultSyncUntyped as Partial<ResultSync>

        console.log(resultSync)

        const wantedPlatforms: Platform[] = []
        //  Consider enabled if user has not explicitly opted out
        if (resultSync.wantGoogleMeet) {
            wantedPlatforms.push("google_meet")
        }
        if (resultSync.wantTeams) {
            wantedPlatforms.push("teams")
        }
        if (resultSync.wantZoom) {
            wantedPlatforms.push("zoom")
        }

        console.log(`Wanted platforms: ${wantedPlatforms}`)
        const message: ExtensionMessage = {
            type: "get_platform_permission_status",
            platform: wantedPlatforms
        }
        chrome.runtime.sendMessage(message, (responseUntyped) => {
            const response = responseUntyped as ExtensionResponse
            if (response.success) {
                console.log(response.message)

                const permissionMissingPlatforms: Platform[] = []

                for (let i = 0; i < wantedPlatforms.length; i++) {
                    if (Array.isArray(response.message) && response.message[i] === "Disabled") {
                        permissionMissingPlatforms.push(wantedPlatforms[i])
                    }
                }

                if (permissionMissingPlatforms.length > 0) {
                    const dialog = document.querySelector<HTMLDialogElement>('#permission-dialog')
                    const reviewButton = document.querySelector<HTMLButtonElement>('#review-button')
                    const laterButton = document.querySelector<HTMLButtonElement>('#later-button')

                    // Show the modal
                    dialog?.showModal()
                    reviewButton?.focus()

                    // Handle the "Confirm" action
                    reviewButton?.addEventListener("click", () => {
                        dialog?.close()

                        // This execution context is now considered a "User Gesture"
                        const disableMsg: ExtensionMessage = {
                            type: "disable_platform",
                            platform: permissionMissingPlatforms
                        }

                        chrome.runtime.sendMessage(disableMsg, (responseUntyped) => {
                            const disableResponse = responseUntyped as ExtensionResponse
                            if (disableResponse.success) {
                                const enableMsg: ExtensionMessage = {
                                    type: "enable_platform",
                                    platform: permissionMissingPlatforms
                                }
                                chrome.runtime.sendMessage(enableMsg)
                            }
                        })
                    })

                    // Handle the "Cancel" action
                    laterButton?.addEventListener("click", () => {
                        dialog?.close()
                    })

                }
            }
        })
    })
}

/** Format duration between two ISO timestamps */
function getDuration(meetingStartTimestamp: string, meetingEndTimestamp: string): string {
    const duration = new Date(meetingEndTimestamp).getTime() - new Date(meetingStartTimestamp).getTime()
    const durationMinutes = Math.round(duration / (1000 * 60))
    const durationHours = Math.floor(durationMinutes / 60)
    const remainingMinutes = durationMinutes % 60
    return durationHours > 0
        ? `${durationHours}h ${remainingMinutes}m`
        : `${durationMinutes}m`
}

/** "Thu 2 Oct, 14:30" in the user's locale — the weekday is what people scan a meeting list by. */
function formatWhen(isoTimestamp: string): string {
    return new Date(isoTimestamp).toLocaleString(undefined, {
        weekday: "short",
        day: "numeric",
        month: "short",
        hour: "2-digit",
        minute: "2-digit"
    })
}

/**
 * Switches between the Meetings and Settings views from the address hash, so the popup
 * can link straight to a settings group and the back button behaves. `#webhooks` is the
 * address older builds of the popup linked to, and still lands on Settings.
 */
function setupViews(): void {
    const meetingsView = document.querySelector<HTMLElement>("#view-meetings")
    const settingsView = document.querySelector<HTMLElement>("#view-settings")
    const tabs = document.querySelectorAll<HTMLAnchorElement>(".tab")
    const groupLinks = document.querySelectorAll<HTMLAnchorElement>(".settings-nav a")
    if (!meetingsView || !settingsView) return

    function show() {
        const hash = location.hash.slice(1)
        const isSettings = hash === "settings" || hash === "webhooks" || hash.startsWith("settings-")
        meetingsView!.hidden = isSettings
        settingsView!.hidden = !isSettings
        tabs.forEach(tab => {
            if ((tab.dataset.view === "settings") === isSettings) tab.setAttribute("aria-current", "page")
            else tab.removeAttribute("aria-current")
        })

        const target = hash === "webhooks" ? "settings-webhook" : hash
        const group = target.startsWith("settings-") ? document.getElementById(target) : null
        if (group) group.scrollIntoView({ block: "start" })
        else window.scrollTo({ top: 0 })
    }

    window.addEventListener("hashchange", show)
    show()

    // Marks the settings group currently at the top of the viewport in the side nav
    const groups = Array.from(document.querySelectorAll<HTMLElement>(".settings-group"))
    const observer = new IntersectionObserver(() => {
        const offset = 120
        let current = groups[0]
        groups.forEach(group => {
            if (group.getBoundingClientRect().top - offset <= 0) current = group
        })
        groupLinks.forEach(link => {
            if (link.getAttribute("href") === `#${current?.id}`) link.setAttribute("aria-current", "true")
            else link.removeAttribute("aria-current")
        })
    }, { rootMargin: "-120px 0px -60% 0px", threshold: [0, 1] })
    groups.forEach(group => observer.observe(group))
}
