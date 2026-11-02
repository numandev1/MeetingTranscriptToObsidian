import { test } from 'node:test'
import assert from 'node:assert/strict'
import { createRequire } from 'node:module'
const require = createRequire(import.meta.url)
const { createCaptureStore } = require('../.test-build/background-script/capture-sessions.js')

test('simultaneous tabs keep captions, title and account separate when either ends', async () => {
  let storage = {}
  const store = createCaptureStore(async () => structuredClone(storage), async patch => {
    await Promise.resolve()
    storage = { ...storage, ...structuredClone(patch) }
  })
  await Promise.all([
    store.update(11, 'a', { meetingTitle: 'Alpha', transcript: ['alpha'], accountEmail: 'a@example.com' }),
    store.update(22, 'b', { meetingTitle: 'Beta', transcript: ['beta'], liveTranscriptBlock: 'still beta' }),
  ])
  await store.update(11, 'a', { liveTranscriptBlock: 'still alpha' })
  const a = await store.finish(11, 'a')
  assert.equal(a.meetingTitle, 'Alpha')
  assert.deepEqual(a.transcript, ['alpha'])
  assert.equal(a.liveTranscriptBlock, 'still alpha')
  assert.equal(a.accountEmail, 'a@example.com')
  assert.equal(storage.meetingTitle, 'Beta')
  assert.equal(storage.liveTranscriptBlock, 'still beta')
  assert.equal(storage.meetingTabId, 22)
  assert.equal(await store.finish(11, 'a'), undefined)
  const b = await store.finish(22)
  assert.deepEqual(b.transcript, ['beta'])
  assert.equal(b.accountEmail, undefined)
})

test('late updates from a previous session cannot change a new meeting in the same tab', async () => {
  let storage = {}
  const store = createCaptureStore(async () => structuredClone(storage), async patch => { storage = { ...storage, ...structuredClone(patch) } })
  await store.update(11, 'old', { meetingTitle: 'Old' })
  await store.finish(11, 'old')
  await store.update(11, 'new', { meetingTitle: 'New' })
  await store.update(11, 'old', { accountEmail: 'wrong@example.com' }, false)
  assert.equal((await store.finish(11, 'new')).accountEmail, undefined)
})

test('selection and worker restart preserve ownership; completed captures release their transcript', async () => {
  let storage = {}
  const read = async () => structuredClone(storage)
  const write = async patch => { storage = { ...storage, ...structuredClone(patch) } }
  const firstWorker = createCaptureStore(read, write)
  await firstWorker.update(1, 'first', { meetingTitle: 'First', transcript: ['private first'] })
  await firstWorker.update(2, 'second', { meetingTitle: 'Second', transcript: ['private second'] })
  await firstWorker.select(1)
  assert.equal(storage.meetingTitle, 'First')
  await firstWorker.update(2, 'second', { transcript: ['updated second'] })
  assert.deepEqual(storage.transcript, ['private first'])
  const restartedWorker = createCaptureStore(read, write)
  const snapshot = await restartedWorker.finish(1)
  assert.deepEqual(snapshot.transcript, ['private first'])
  assert.deepEqual(storage.transcript, ['updated second'])
  assert.equal((await restartedWorker.pending(1)).captureId, 'first')
  await restartedWorker.complete(1, 'first')
  assert.equal(await restartedWorker.pending(1), undefined)
  assert.equal(storage.captureSessions[1].transcript, undefined)
  await restartedWorker.update(1, 'first', { transcript: ['late'] })
  assert.equal(storage.captureSessions[1].transcript, undefined)
})

import { readFileSync } from 'node:fs'
import { runInNewContext } from 'node:vm'
test('concurrent finalization files each snapshot and never clears another live buffer', async () => {
  let storage = { liveTranscriptBlock: { personName: 'Live', transcriptText: 'another call is speaking' } }
  const deliveries = []
  const exported = {}
  const backgroundRequire = createRequire(new URL('../.test-build/background-script/meetings.js', import.meta.url))
  runInNewContext(readFileSync(new URL('../.test-build/background-script/meetings.js', import.meta.url), 'utf8'), {
    exports: exported,
    require: name => {
      if (name.startsWith('@lib/')) return require(`../.test-build/lib/${name.slice(5)}.js`)
      if (name === './exporters') return {
        downloadTranscript: async index => { await Promise.resolve(); deliveries.push(structuredClone(storage.meetings[index])) },
      }
      if (name === './obsidian' || name === './summarizer') return {}
      return backgroundRequire(name)
    },
    console: { log() {}, error() {} },
    chrome: { storage: {
      local: {
        get: (_, callback) => queueMicrotask(() => callback(structuredClone(storage))),
        set: (patch, callback) => queueMicrotask(() => { storage = { ...storage, ...structuredClone(patch) }; callback?.() }),
      },
      sync: { get: (_, callback) => queueMicrotask(() => callback({ autoDownloadFileAfterMeeting: true })) },
    } },
  })
  const snapshot = (id, title) => ({
    captureId: id, meetingTitle: title, meetingSoftware: 'Google Meet',
    meetingUrl: `https://meet.google.com/${id}`,
    meetingStartTimestamp: '2026-11-02T10:00:00.000Z',
    meetingEndTimestamp: '2026-11-02T11:00:00.000Z',
    transcript: [{ personName: title, timestamp: '2026-11-02T10:01:00.000Z', transcriptText: `Only ${title}` }],
    chatMessages: [],
  })
  const a = snapshot('a', 'Alpha')
  const b = snapshot('b', 'Beta')
  await Promise.all([exported.processLastMeeting(a), exported.processLastMeeting(b)])
  assert.equal(storage.meetings.length, 2)
  assert.deepEqual(deliveries.map(item => item.meetingTitle), ['Alpha', 'Beta'])
  assert.deepEqual(deliveries.map(item => item.transcript[0].transcriptText), ['Only Alpha', 'Only Beta'])
  assert.equal(storage.liveTranscriptBlock.transcriptText, 'another call is speaking')
  // Retrying an earlier export must address Alpha, even though Beta is last now.
  await exported.processLastMeeting(a)
  assert.equal(deliveries[2].meetingTitle, 'Alpha')
})

test('two notes with identical titles and start times have different vault paths', () => {
  const exported = {}
  const backgroundRequire = createRequire(new URL('../.test-build/background-script/obsidian.js', import.meta.url))
  runInNewContext(readFileSync(new URL('../.test-build/background-script/obsidian.js', import.meta.url), 'utf8'), {
    exports: exported,
    require: name => name.startsWith('@lib/') ? require(`../.test-build/lib/${name.slice(5)}.js`) : backgroundRequire(name),
  })
  const common = { meetingTitle: 'Standup', meetingStartTimestamp: '2026-11-02T10:00:00.000Z' }
  assert.notEqual(exported.getObsidianNoteName({ ...common, captureId: 'alpha' }), exported.getObsidianNoteName({ ...common, captureId: 'beta' }))
})
