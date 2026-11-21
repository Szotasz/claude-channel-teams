import { expect, test } from 'bun:test'
import { validateAdaptiveCard } from '../src/teams/cards'

const card = (type?: string) => ({
  type: 'AdaptiveCard',
  version: '1.5',
  body: [{ type: 'TextBlock', text: 'Hi <at>Ann</at>' }],
  ...(type === undefined
    ? {}
    : { msteams: { entities: [{ type, text: '<at>Ann</at>', mentioned: { id: '29:x', name: 'Ann' } }] } }),
})

test('a card with a mention entity is refused', () => {
  const error = "card.msteams.entities[0] is a mention entity; mention people and team tags with reply's mentions instead"
  expect(validateAdaptiveCard(card('mention'))).toEqual({ ok: false, error })
  expect(validateAdaptiveCard(card('Mention'))).toEqual({ ok: false, error })
})

test('a mention entity anywhere in the card is refused', () => {
  const nested = {
    type: 'AdaptiveCard',
    version: '1.5',
    body: [{ type: 'Container', items: [{ type: 'mention', text: '<at>Ann</at>', mentioned: { id: '29:x', name: 'Ann' } }] }],
  }
  expect(validateAdaptiveCard(nested)).toEqual({
    ok: false,
    error: "card.body[0].items[0] is a mention entity; mention people and team tags with reply's mentions instead",
  })
})

test('the same card without msteams is fine', () => {
  expect(validateAdaptiveCard(card())).toEqual({ ok: true })
})
