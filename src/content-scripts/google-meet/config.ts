export const SELECTORS_GOOGLE_MEET = {
    USER_NAME: ".awLEm",
    MEETING_TITLE: ".u6vdEc",
    GOOGLE_SYMBOLS: ".google-symbols",
    TEXT_CALL_END: "call_end",
    // Meet swaps the ligature to reflect state: "_off" means captions are currently
    // off (clicking turns them on), the bare name means they are already on.
    TEXT_CAPTIONS_OFF: "closed_caption_off",
    TEXT_CAPTIONS_ON: "closed_caption",
    // Meetings with translated speech replace the CC toggle with a "Captions and
    // translation" button whose icon never changes; its state is in aria-pressed.
    TEXT_CAPTIONS_TRANSLATE: "translate_cc",
    TEXT_CHAT: "chat",
    TRANSCRIPT_REGION: 'div[role="region"][tabindex="0"]',
    CHAT_LIVE_REGION: 'div[aria-live="polite"].Ge9Kpc',
    // Every chat message carries a stable id like "spaces/<id>/messages/<id>"
    CHAT_MESSAGE: "[data-message-id]",
    // Sender name + timestamp header, once per run of consecutive messages
    CHAT_GROUP_HEADER: ".HNucUd",
    // Newer Meet renders chat in this iframe instead of CHAT_LIVE_REGION
    EMBEDDED_CHAT_IFRAME: 'iframe[src*="chat.google.com"]'
}

/**
 * Since late Sep 2026 Meet renders chat inside a cross-origin Google Chat iframe
 * (chat.google.com/.../embed/space/...), so none of the CHAT_* selectors above exist in
 * the Meet document any more. These are read by the content script running inside that
 * iframe (see embedded-chat.ts).
 */
export const SELECTORS_GOOGLE_MEET_EMBEDDED_CHAT = {
    // One per message; data-id is the Google Chat message id
    MESSAGE: 'div[role="group"][data-id]',
    // Sender chip, only rendered on the first message of a run from the same person
    SENDER: "[data-message-id][data-name]",
    // Epoch milliseconds the message was sent
    TIMESTAMP: "[data-absolute-timestamp]",
    // Message body
    BODY: ".DTp27d"
}
