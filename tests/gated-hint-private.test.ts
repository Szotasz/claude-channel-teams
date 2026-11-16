import { describe, expect, test } from 'bun:test'
import { makePairingDmSender } from '../src/teams/pairing-dm'

// Owner decision (Teams #21 review, 2026-11-16): the gated pairing hint is
// delivered to the triggering member as a 1:1 message and NEVER posted into
// a group/channel conversation -- not even as a fallback when the private
// delivery fails. This file is the binding contract for that decision.

const HINT = 'This bot is gated. To pair, send me a direct (1:1) message.'

type Ref = {
  conversation: { id: string; conversationType?: string }
  user?: { id: string }
  [k: string]: any
}

const personalRef: Ref = { conversation: { id: 'a:1xyz', conversationType: 'personal' }, user: { id: '29:u1' } }
const channelRef: Ref = { conversation: { id: '19:chan1@thread.tacv2', conversationType: 'channel' }, user: { id: '29:u2' } }

type Sent = { raw: Array<[Ref, string]>; dmTexts: string[]; createdWith: Ref[] }

function harness(fail?: { create?: Error; send?: Error }) {
  const sent: Sent = { raw: [], dmTexts: [], createdWith: [] }
  const sender = makePairingDmSender({
    sendActivityRaw: async (ref, text) => {
      sent.raw.push([ref as Ref, text])
    },
    isPersonal: ref => (ref as Ref).conversation.conversationType === 'personal',
    createConversation: async (ref, logic) => {
      sent.createdWith.push(ref as Ref)
      if (fail?.create) throw fail.create
      await logic({
        sendActivity: async (text: string) => {
          if (fail?.send) throw fail.send
          sent.dmTexts.push(text)
        },
      })
    },
  })
  return { sender, sent }
}

describe('gated pairing hint delivery (owner decision: 1:1 only, never the channel)', () => {
  test('a personal reference is answered in place, exactly as before', async () => {
    const { sender, sent } = harness()
    await sender(personalRef, HINT)
    expect(sent.raw).toEqual([[personalRef, HINT]])
    expect(sent.createdWith).toEqual([])
    expect(sent.dmTexts).toEqual([])
  })

  test('a channel reference: the hint goes to the member via a bot-initiated 1:1, not the channel', async () => {
    const { sender, sent } = harness()
    await sender(channelRef, HINT)
    expect(sent.raw).toEqual([])
    expect(sent.createdWith).toEqual([channelRef])
    expect(sent.dmTexts).toEqual([HINT])
  })

  test('a channel reference whose 1:1 conversation cannot be created: the same error object propagates, nothing goes public', async () => {
    const err = new Error('cannot create conversation')
    const { sender, sent } = harness({ create: err })
    await expect(sender(channelRef, HINT)).rejects.toBe(err)
    expect(sent.raw).toEqual([])
    expect(sent.dmTexts).toEqual([])
  })

  test('a channel reference whose 1:1 send fails inside the conversation: the same error object propagates, nothing goes public', async () => {
    const err = new Error('send failed')
    const { sender, sent } = harness({ send: err })
    await expect(sender(channelRef, HINT)).rejects.toBe(err)
    expect(sent.raw).toEqual([])
  })
})
