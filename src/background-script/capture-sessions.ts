import type { ResultLocal } from '../lib/types'

export interface CaptureSession extends Partial<ResultLocal> {
    captureId: string
    tabId: number
    ended?: boolean
    archived?: boolean
    meetingEndTimestamp?: string
}
interface CaptureStorage extends Partial<ResultLocal> {
    captureSessions?: Record<string, CaptureSession>
}

/** All read/modify/write operations share a queue; Chrome storage has no transactions. */
export function createCaptureStore(read: () => Promise<CaptureStorage>, write: (patch: CaptureStorage) => Promise<void>) {
    let queue: Promise<unknown> = Promise.resolve()
    function serial<T>(operation: () => Promise<T>): Promise<T> {
        const run = queue.then(operation)
        queue = run.catch(() => {})
        return run
    }
    function mirror(session: CaptureSession | undefined): Partial<ResultLocal> {
        return {
            activeCaptureId: session?.captureId,
            meetingTabId: session?.tabId ?? null,
            meetingStartTimestamp: session?.meetingStartTimestamp ?? '',
            meetingTitle: session?.meetingTitle ?? '',
            meetingSoftware: session?.meetingSoftware ?? '',
            transcript: session?.transcript ?? [],
            chatMessages: session?.chatMessages ?? [],
            liveTranscriptBlock: session?.liveTranscriptBlock ?? null,
            accountEmail: session?.accountEmail,
            meetingUrl: session?.meetingUrl,
            captionsSeen: session?.captionsSeen ?? false,
        }
    }
    return {
        update(tabId: number, captureId: string, patch: Partial<ResultLocal>, create = true) {
            return serial(async () => {
                const stored = await read()
                const sessions = stored.captureSessions ?? {}
                const previous = sessions[tabId]
                if ((!previous || previous.captureId !== captureId) && !create) return
                if (previous?.captureId === captureId && previous.ended) return
                const isNew = previous?.captureId !== captureId
                const session = { ...(isNew ? {} : previous), ...patch, tabId, captureId } as CaptureSession
                sessions[tabId] = session
                const selected = isNew || stored.meetingTabId === tabId
                await write({ captureSessions: sessions, ...(selected ? mirror(session) : {}), ...(isNew ? { liveSummary: null } : {}) })
            })
        },
        complete(tabId: number, captureId: string) {
            return serial(async () => {
                const stored = await read()
                const sessions = stored.captureSessions ?? {}
                const session = sessions[tabId]
                if (session?.captureId !== captureId || !session.ended) return
                // Keep a small tombstone so delayed caption messages cannot reopen it.
                sessions[tabId] = { tabId, captureId, ended: true, archived: true }
                await write({ captureSessions: sessions })
            })
        },
        pending(tabId: number) {
            return serial(async () => {
                const session = (await read()).captureSessions?.[tabId]
                return session?.ended && !session.archived ? session : undefined
            })
        },
        select(tabId: number) {
            return serial(async () => {
                const stored = await read()
                const session = stored.captureSessions?.[tabId]
                if (session && !session.ended) await write({ ...mirror(session), liveSummary: null })
            })
        },
        finish(tabId: number, captureId?: string) {
            return serial(async () => {
                const stored = await read()
                const sessions = stored.captureSessions ?? {}
                const session = sessions[tabId]
                if (!session || session.ended || (captureId && captureId !== session.captureId)) return undefined
                session.ended = true
                session.meetingEndTimestamp = new Date().toISOString()
                const remaining = Object.values(sessions).find(item => !item.ended)
                await write({ captureSessions: sessions, ...(stored.meetingTabId === tabId ? mirror(remaining) : {}) })
                return session
            })
        },
    }
}
