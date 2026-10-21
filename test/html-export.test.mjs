import { test } from "node:test"
import assert from "node:assert/strict"
import { createRequire } from "node:module"
const require = createRequire(import.meta.url)
const { buildMeetingHtml } = require("../.test-build/lib/html-export.js")
const meeting = {
  meetingTitle: 'Planning <script>alert("title")</script>',
  meetingSoftware: "Google Meet",
  meetingStartTimestamp: "2026-10-21T10:00:00Z",
  meetingEndTimestamp: "2026-10-21T11:00:00Z",
  transcript: [{ personName: "Alice & Bob", timestamp: "2026-10-21T10:01:00Z", transcriptText: '<img src=x onerror=alert(1)>\nNext line' }],
  chatMessages: [{ personName: "Charlie", timestamp: "2026-10-21T10:02:00Z", chatMessageText: "Chat only text" }],
  summary: "## Decisions\n- **Ship** tomorrow\n- [x] Review\n\n<script>alert('summary')</script>",
  webhookPostStatus: "new",
}

test("transcript HTML is standalone and escapes all captured content", () => {
  const html = buildMeetingHtml(meeting, "transcript")
  assert.match(html, /^<!doctype html>/i)
  assert.match(html, /<meta charset="utf-8">/)
  assert.match(html, /<meta name="viewport"/)
  assert.match(html, /<style>/)
  assert.match(html, /Alice &amp; Bob/)
  assert.match(html, /&lt;img src=x onerror=alert\(1\)&gt;\nNext line/)
  assert.match(html, /<time datetime="2026-10-21T10:01:00Z">/)
  assert.match(html, /Chat only text/)
  assert.doesNotMatch(html, /<script|<img|Ship/)
})

test("summary HTML preserves headings, bold and checked tasks without raw markup", () => {
  const html = buildMeetingHtml(meeting, "summary")
  assert.match(html, /<h3>Decisions<\/h3>/)
  assert.match(html, /<strong>Ship<\/strong>/)
  assert.match(html, /type="checkbox" disabled checked/)
  assert.match(html, /&lt;script&gt;/)
  assert.doesNotMatch(html, /<script|Chat only text|Alice &amp; Bob/)
})

test("live exports are labelled and older titles and empty sections are supported", () => {
  const html = buildMeetingHtml({ ...meeting, meetingTitle: undefined, title: "Older meeting", transcript: [], chatMessages: [] }, "transcript", true)
  assert.match(html, /Older meeting/)
  assert.match(html, /Transcript so far/)
  assert.match(html, /No transcript was captured/)
  assert.match(html, /No chat messages were captured/)
})

// Run the real exporter with Chrome storage/downloads and FileReader boundaries supplied by Node.
import { readFileSync } from "node:fs"
import { runInNewContext } from "node:vm"
function exporterHarness(failures = 0) {
  const downloads = []
  const exported = {}
  const backgroundRequire = createRequire(new URL("../.test-build/background-script/exporters.js", import.meta.url))
  runInNewContext(readFileSync(new URL("../.test-build/background-script/exporters.js", import.meta.url), "utf8"), {
    exports: exported,
    require: name => name.startsWith("@lib/") ? require(`../.test-build/lib/${name.slice(5)}.js`) : backgroundRequire(name),
    Blob,
    console: { log() {}, error() {} },
    FileReader: class {
      readAsDataURL(blob) {
        blob.arrayBuffer().then(buffer => this.onload({ target: { result: `data:${blob.type};base64,${Buffer.from(buffer).toString("base64")}` } }))
      }
    },
    chrome: {
      storage: {
        local: { get: (_, callback) => callback({ meetings: [meeting] }) },
        sync: { get: (_, callback) => callback({ downloadSubfolder: "Work/Calls" }) },
      },
      downloads: { download: async options => { downloads.push(options); if (failures-- > 0) throw new Error("Download rejected"); return downloads.length } },
    },
  })
  return { ...exported, downloads }
}

for (const kind of ["transcript", "summary"]) {
  test(`${kind} downloads HTML to the configured folder and retains text as the default`, async () => {
    const harness = exporterHarness()
    const download = format => kind === "transcript" ? harness.downloadTranscript(0, false, format) : harness.downloadSummary(0, format)
    await download("html")
    await download(undefined)
    const [html, text] = harness.downloads
    assert.match(html.filename, /^Work\/Calls\//)
    assert.match(html.filename, /\.html$/)
    assert.match(html.url, /^data:text\/html;charset=utf-8;base64,/)
    assert.match(Buffer.from(html.url.split(",")[1], "base64").toString(), /<!doctype html>/)
    assert.match(text.filename, /\.txt$/)
    assert.match(text.url, /^data:text\/plain;base64,/)
  })
}

test("live HTML downloads and filename fallbacks retain the selected format", async () => {
  const harness = exporterHarness(1)
  await harness.downloadTranscript(meeting, false, "html")
  assert.match(harness.downloads[0].filename, /transcript so far/)
  assert.equal(harness.downloads[1].filename, "Work/Calls/Transcript.html")
  assert.match(Buffer.from(harness.downloads[1].url.split(",")[1], "base64").toString(), /Transcript so far/)
})

test("a failed fallback download is reported as a failure", async () => {
  const harness = exporterHarness(2)
  await assert.rejects(harness.downloadSummary(0, "html"), /Download rejected/)
})
