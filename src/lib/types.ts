/** Human friendly meeting software name. */
export type MeetingSoftware = "Google Meet" | "Zoom" | "Teams" | ""

/** Meeting platform. */
export type Platform = "google_meet" | "teams" | "zoom"

/**
 * Tab id of the meeting tab, captured when the meeting starts. A valid value or
 * "processing" indicates that a meeting is in progress. Set to null once the meeting
 * ends and associated processing is complete.
 */
export type MeetingTabId = number | "processing" | null

/** Mode of the extension, deciding whether transcripts are captured automatically. */
export type OperationMode = "auto" | "manual"

/** Type of webhook body to use. */
export type WebhookBodyType = "simple" | "advanced"

/** A chunk of transcript. */
export interface TranscriptBlock {
  /** Name of the person who spoke. */
  personName: string
  /** ISO timestamp of when the words were spoken. */
  timestamp: string
  /** Actual transcript text. */
  transcriptText: string
}

/** A chat message. */
export interface ChatMessage {
  /** Name of the person who sent the message. */
  personName: string
  /** ISO timestamp of when the message was sent. */
  timestamp: string
  /** Actual message text. */
  chatMessageText: string
}

export interface WebhookBody {
  webhookBodyType: WebhookBodyType
  meetingSoftware: MeetingSoftware
  meetingTitle: string
  /** ISO timestamp of when the meeting started. */
  meetingStartTimestamp: string
  /** ISO timestamp of when the meeting ended. */
  meetingEndTimestamp: string
  /** Transcript as a formatted string, or the raw transcript blocks. */
  transcript: TranscriptBlock[] | string
  /** Chat messages as a formatted string, or the raw chat messages. */
  chatMessages: ChatMessage[] | string
}

export interface ExtensionStatusJSON {
  status: number
  message: string
  showBetaMessage?: boolean
}

export interface Meeting {
  meetingSoftware?: MeetingSoftware
  meetingTitle?: string | undefined
  /** Older key for meetingTitle, as written by v3.1.0. */
  title?: string | undefined
  /** ISO timestamp of when the meeting started. */
  meetingStartTimestamp: string
  /** ISO timestamp of when the meeting ended. */
  meetingEndTimestamp: string
  transcript: TranscriptBlock[]
  chatMessages: ChatMessage[]
  webhookPostStatus: "new" | "failed" | "successful"
  /** Undefined for meetings recorded before the Obsidian integration existed. */
  obsidianPostStatus?: "new" | "failed" | "successful"
  /**
   * Account the meeting was attended from, used to file notes per account.
   * Undefined on platforms where it cannot be determined, and on meetings recorded
   * before this existed.
   */
  accountEmail?: string
  /**
   * Whether the captions region was ever found during the meeting. Distinguishes
   * "nobody spoke" from "capture never started", which otherwise look identical.
   */
  captionsSeen?: boolean
  /** Summary produced by ChatGPT, as markdown. */
  summary?: string
  summaryStatus?: "new" | "failed" | "successful"
  /**
   * Address identifying the meeting, with account-selection queries removed for Meet.
   * A reload or rejoin of this address extends the saved transcript.
   */
  meetingUrl?: string
  /**
   * Start times of later pieces merged into this meeting after a rejoin. The meeting
   * keeps its first start time, so recovery checks these too.
   */
  mergedStartTimestamps?: string[]
}

/**
 * Buffer holding the utterance currently being spoken. Pushed into the transcript
 * array as a transcript block when the speaker or the captions node changes.
 */
export interface StateTranscriptBlock {
  timestamp: string
  mutationTargetElement: Element | null
  personName: string
  transcriptTextBuffer: string
}

export interface ContentScriptState {
  meetingSoftware: MeetingSoftware
  platform: Platform
  userName: string
  transcript: TranscriptBlock[]
  chatMessages: ChatMessage[]
  stateTranscriptBlock: StateTranscriptBlock
  /** ISO timestamp of when the most recent meeting started. */
  meetingStartTimestamp: string
  meetingTitle: string
  transcriptTargetNode: Element | null
  transcriptObserver: MutationObserver | null
  chatMessagesTargetNode: Element | null
  chatMessagesObserver: MutationObserver | null
  /**
   * True once userName came from the user's own setting, which must not then be
   * overwritten by whatever the meeting UI reports.
   */
  isUserNameLocked: boolean
  /** data-message-id values already captured, so a re-sweep cannot duplicate them. */
  capturedChatMessageIds: string[]
  isTranscriptDomErrorCaptured: boolean
  isChatMessagesDomErrorCaptured: boolean
  hasMeetingStarted: boolean
  hasMeetingEnded: boolean
  extensionStatusJSON: ExtensionStatusJSON
}

export type ExtensionMessageType =
  | "new_meeting_started"
  | "meeting_ended"
  | "download_transcript_at_index"
  | "post_webhook_at_index"
  | "post_obsidian_at_index"
  | "post_summary_obsidian_at_index"
  | "summarize_at_index"
  | "download_summary_at_index"
  | "sync_all_to_obsidian"
  | "recover_last_meeting"
  | "get_platform_enablement_status"
  | "get_platform_permission_status"
  | "enable_platform"
  | "disable_platform"
  | "open_popup"
  | "open_side_panel"
  | "broadcast_live_buffer"
  | "meet_embedded_chat_messages"
  | "summarize_live_meeting"
  | "download_live_transcript"
  | "download_live_summary"

/** Message sent by the calling script. */
export interface ExtensionMessage {
  type: ExtensionMessageType
  /** Index of the meeting to process. */
  index?: number
  /** File format for manual downloads. Defaults to plain text. */
  format?: "txt" | "html"
  platform?: Platform | Platform[]
  stateTranscriptBlock?: StateTranscriptBlock
  /** Messages read from Meet's embedded Google Chat iframe, relayed to the meeting frame. */
  embeddedChatMessages?: EmbeddedChatMessage[]
}

/** A chat message read inside Meet's chat.google.com iframe, keyed for de-duplication. */
export interface EmbeddedChatMessage extends ChatMessage {
  /** Google Chat message id (data-id), stable across re-renders. */
  messageId: string
}

/** Response sent by the called script. */
export interface ExtensionResponse {
  success: boolean
  /** Message explaining success or failure. */
  message?: string | string[] | ErrorObject
}

export interface RecentError {
  /** ISO timestamp of when it happened. */
  at: string
  /** Code from the table at the bottom of this file. */
  code: string
  meetingSoftware: MeetingSoftware
  message: string
}

export interface ErrorObject {
  errorCode: string
  errorMessage: string
}

/** Shape of chrome.storage.local. */
export interface ResultLocal {
  extensionStatusJSON: ExtensionStatusJSON
  meetingTabId: MeetingTabId
  meetingSoftware: MeetingSoftware
  /** Title of the most recent meeting, dumped by the content script. */
  meetingTitle: string
  /** ISO timestamp of when the most recent meeting started, dumped by the content script. */
  meetingStartTimestamp: string
  /** Transcript of the most recent meeting, dumped by the content script. */
  transcript: TranscriptBlock[]
  /** Chat messages captured during the most recent meeting, dumped by the content script. */
  chatMessages: ChatMessage[]
  /**
   * The utterance currently being spoken, persisted periodically so it survives the
   * meeting tab being closed mid-sentence. Merged into the transcript by
   * pickupLastMeetingFromStorage() if the content script never got to flush it.
   */
  liveTranscriptBlock: TranscriptBlock | null | undefined
  /**
   * Recent capture errors, newest last. Kept locally so a user whose transcript failed
   * has something to look at and attach to a report — the removed upstream telemetry
   * used to be the only record, and a service worker console is not one.
   */
  recentErrors: RecentError[] | undefined
  /** Account the in-progress meeting is being attended from. */
  accountEmail: string | undefined
  /** Whether the captions region was seen during the in-progress meeting. */
  captionsSeen: boolean | undefined
  /** Address of the in-progress meeting, recorded by the background script when it starts. */
  meetingUrl: string | undefined
  /** Whether the extension has a deferred update waiting to be applied. */
  isDeferredUpdatedAvailable: boolean | undefined
  meetings: Meeting[] | undefined
  /**
   * Obsidian Local REST API key. Deliberately kept in local (not sync) storage so the
   * credential never leaves the device, matching the extension's privacy promise.
   */
  obsidianApiKey: string | undefined
  /**
   * Summary of the meeting still in progress, made from the transcript so far. Belongs
   * to the meeting whose start timestamp it carries; a different meeting ignores it.
   */
  liveSummary: LiveSummary | null | undefined
}

/** A summary taken part-way through a meeting. */
export interface LiveSummary {
  /** Start timestamp of the meeting it belongs to. */
  meetingStartTimestamp: string
  summary: string
  /** ISO timestamp of when it was made. */
  createdAt: string
  /** How many transcript blocks it covers, so the panel can say how far it got. */
  blockCount: number
}

/** Shape of chrome.storage.sync. */
export interface ResultSync {
  /** Whether to automatically post the webhook after each meeting. */
  autoPostWebhookAfterMeeting: boolean
  /** Whether to automatically download the transcript file after each meeting. */
  autoDownloadFileAfterMeeting: boolean
  /**
   * Whether to summarise each meeting through ChatGPT as soon as it ends, without
   * waiting for the Summarise button. Off unless the user turns it on: it sends the
   * transcript to a third party and briefly brings a tab to the front, neither of
   * which may start happening on its own.
   */
  autoSummarizeAfterMeeting: boolean
  /**
   * Folder the transcript text file is saved into, relative to the browser's Downloads
   * directory. Empty string saves straight into Downloads. Chrome cannot write outside
   * the Downloads directory, so this is always a relative path.
   */
  downloadSubfolder: string
  /** Whether to save each meeting as a note in Obsidian via the Local REST API plugin. */
  obsidianEnabled: boolean
  /** Vault-relative folder the meeting notes are written into. */
  obsidianFolder: string
  /** Vault-relative folder the meeting summaries are written into. */
  obsidianSummaryFolder: string
  /**
   * How many meetings to keep on the device. Older ones are removed when a new meeting
   * is filed. 0 means keep everything.
   */
  meetingsToKeep: number
  /** Instructions sent to ChatGPT ahead of the transcript. Empty means the default. */
  summaryPrompt: string
  /** Base URL of the Obsidian Local REST API, e.g. http://127.0.0.1:27123 */
  obsidianBaseUrl: string
  operationMode: OperationMode
  /** Hide the captions on the UI by changing height and opacity. */
  hideCaptions: boolean
  /**
   * The user's own display name, used wherever captions attribute speech to "You".
   * Empty means fall back to auto-detection from the meeting UI, which only Google Meet
   * supports, so Teams and Zoom rely on this being set.
   */
  participantName: string
  webhookBodyType: WebhookBodyType
  webhookUrl: string
  /**
   * Indicates whether the user explicitly opted in for Google Meet. Does not necessarily
   * mean Google Meet is enabled for them — only an indicator to re-inject content scripts
   * between reloads.
   */
  wantGoogleMeet: boolean
  /** Indicates whether the user explicitly opted in for Teams. */
  wantTeams: boolean
  /** Indicates whether the user explicitly opted in for Zoom. */
  wantZoom: boolean
}

// CONTENT SCRIPT ERRORS
// | Error Code | Error Message |
// | :--- | :--- |
// | **001** | "Transcript element not found in DOM" |
// | **002** | "Chat messages element not found in DOM" (currently not in use) |
// | **003** | "Chat button element not found in DOM" |
// | **004** | "Call end button element not found in DOM" |
// | **005** | "Transcript mutation failed to process" |
// | **006** | "Chat messages mutation failed to process" |
// | **007** | "Meeting title element not found in DOM" (currently not in use) |
// | **008** | "Failed to fetch extension status" |
// | **016** | "Recovery timed out" |

// BACKGROUND SCRIPT ERRORS
// | Error Code | Error Message |
// | :--- | :--- |
// | **009** | "Failed to read blob" |
// | **010** | "Meeting at specified index not found" |
// | **011** | "Webhook request failed with HTTP status code [number] [statusText]" |
// | **012** | "No webhook URL configured" |
// | **013** | "No meetings found. May be attend one?" |
// | **014** | "Empty transcript and empty chatMessages" |
// | **015** | "Invalid index" |
// | **017** | "Obsidian request failed with HTTP status code [number] [statusText]" |
// | **018** | "Obsidian is not configured" |
// | **019** | "Could not reach Obsidian. Is the app running with the Local REST API plugin enabled?" |
// | **020** | "Could not open ChatGPT. Grant access to chatgpt.com and make sure you are signed in." |
// |         | also "Could not send the prompt to ChatGPT — <reason>" when the tab opened but the composer or send button was not usable |
// | **021** | "ChatGPT did not return a summary in time" |
// | **022** | "Nothing to summarise for this meeting" |
// | **023** | "No meeting is in progress" |
