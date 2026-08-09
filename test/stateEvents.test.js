const test = require("node:test")
const assert = require("node:assert/strict")

const { getStateEventOccurrencesInRange } = require("../src/stateEventOccurrenceCalculation")
const { buildStateEventDeliveryClaims } = require("../src/stateEventDeliveryGeneration")
const { formatStateEventDelivery } = require("../src/stateEventDeliveryFormatting")
const { buildStateRoundupOccurrences } = require("../src/stateEventRoundupCalculation")
const { validateStateEventDraft } = require("../src/stateEventValidation")
const { formatWeeklyRoundup } = require("../src/weeklyRoundupFormatting")

function stateEvent(overrides = {}) {
  return {
    id: "11111111-1111-4111-8111-111111111111",
    game_profile: "wos",
    state_guild_id: "state-guild",
    state_number: "689",
    event_name: "SvS",
    first_occurrence_date: "2026-08-10",
    recurrence_days: 2,
    schedule_version: 3,
    status: "active",
    phases: [
      {
        id: "22222222-2222-4222-8222-222222222222",
        phase_name: "Borders open",
        phase_time_utc: "10:00",
        pre_alert_minutes: 30,
        announce_exact: true,
        sort_order: 0
      },
      {
        id: "33333333-3333-4333-8333-333333333333",
        phase_name: "Battle starts",
        phase_time_utc: "12:00",
        pre_alert_minutes: 5,
        announce_exact: false,
        sort_order: 1
      }
    ],
    ...overrides
  }
}

test("state event recurrence remains anchored and supports true every 2 days", () => {
  const occurrences = getStateEventOccurrencesInRange(
    stateEvent({ phases: [stateEvent().phases[0]] }),
    new Date("2026-08-10T00:00:00Z"),
    new Date("2026-08-17T00:00:00Z")
  )
  assert.deepEqual(occurrences.map(item => item.occurrenceAt.toISOString()), [
    "2026-08-10T10:00:00.000Z",
    "2026-08-12T10:00:00.000Z",
    "2026-08-14T10:00:00.000Z",
    "2026-08-16T10:00:00.000Z"
  ])
})

test("state event recurrence preserves genuine 3-day and 42-day intervals", () => {
  for (const recurrenceDays of [3, 42]) {
    const occurrences = getStateEventOccurrencesInRange(
      stateEvent({ recurrence_days: recurrenceDays, phases: [stateEvent().phases[0]] }),
      new Date("2026-08-10T00:00:00Z"),
      new Date("2026-11-10T00:00:00Z")
    )
    assert.equal(
      occurrences[1].occurrenceAt.getTime() - occurrences[0].occurrenceAt.getTime(),
      recurrenceDays * 86400000
    )
  }
})

test("state delivery generation creates phase claims for unique target guilds", async () => {
  const targets = [
    { target_kind: "state", target_guild_id: "state-guild", target_channel_id: "state-channel" },
    { target_kind: "alliance", target_guild_id: "alliance-guild", target_channel_id: "alliance-channel" },
    { target_kind: "alliance", target_guild_id: "alliance-guild", target_channel_id: "alliance-channel" }
  ]
  const repository = {
    async listTargetsForStateGuild() { return targets }
  }
  const claims = await buildStateEventDeliveryClaims([stateEvent()], {
    repository,
    gameProfile: "wos",
    windowStart: new Date("2026-08-10T09:29:00Z"),
    windowEnd: new Date("2026-08-10T12:01:00Z")
  })
  const unique = new Set(claims.map(claim => [
    claim.stateEventId,
    claim.phaseId,
    claim.occurrenceAt.toISOString(),
    claim.deliveryKind,
    claim.targetGuildId
  ].join(":")))
  assert.equal(unique.size, claims.length)
  assert.equal(claims.filter(claim => claim.targetGuildId === "alliance-guild").length, 3)
  assert.ok(claims.some(claim => claim.deliveryKind === "exact"
    && claim.phaseId === "22222222-2222-4222-8222-222222222222"))
  assert.ok(!claims.some(claim => claim.deliveryKind === "exact"
    && claim.phaseId === "33333333-3333-4333-8333-333333333333"))
})

test("state event formatting uses state number, phase time and optional custom text", () => {
  const message = formatStateEventDelivery({
    claim: {
      deliveryKind: "pre_alert",
      occurrenceAt: new Date("2026-08-10T10:00:00Z"),
      deliverAt: new Date("2026-08-10T09:30:00Z")
    },
    stateEvent: { stateNumber: "689", eventName: "SvS" },
    phase: {
      name: "Borders open",
      preAlertMinutes: 30,
      preAlertMessage: "Get ready."
    }
  })
  const description = message.embeds[0].toJSON().description
  assert.match(description, /\*\*689\*\*/)
  assert.match(description, /\*\*SvS\*\*/)
  assert.match(description, /10:00 UTC/)
  assert.match(description, /Local time: <t:1786356000:t>/)
  assert.match(description, /Borders open in 30 minutes/)
  assert.match(description, /Get ready\./)
})

test("state event validation rejects duplicate phase names and accepts one-time recurrence", () => {
  assert.equal(validateStateEventDraft({
    eventName: "SvS",
    firstOccurrenceDate: "2026-08-10",
    recurrenceDays: "none",
    phases: [{ phaseName: "Battle", phaseTimeUtc: "10am", announceExact: true }]
  }).recurrenceDays, null)

  assert.throws(() => validateStateEventDraft({
    eventName: "SvS",
    firstOccurrenceDate: "2026-08-10",
    recurrenceDays: 42,
    phases: [
      { phaseName: "Battle", phaseTimeUtc: "10:00" },
      { phaseName: "battle", phaseTimeUtc: "11:00" }
    ]
  }), /Duplicate phase name/)
})

test("weekly roundups contain every state phase once and never add pre-alert milestones", () => {
  const occurrences = buildStateRoundupOccurrences(
    [stateEvent()],
    new Date("2026-08-10T00:00:00Z"),
    new Date("2026-08-11T00:00:00Z")
  )
  assert.deepEqual(occurrences.map(item => item.phaseName), ["Borders open", "Battle starts"])
  assert.ok(occurrences.some(item => item.phaseName === "Battle starts" && item.announceExact === false))
  assert.ok(occurrences.every(item => item.deliveryKind === undefined))

  const messages = formatWeeklyRoundup({
    claim: {
      targetKind: "alliance",
      weekStart: new Date("2026-08-10T00:00:00Z"),
      weekEnd: new Date("2026-08-17T00:00:00Z"),
      postWhenEmpty: false
    },
    allianceName: "YOU",
    occurrences: [],
    stateOccurrences: [...occurrences, ...occurrences]
  })
  const description = messages[0].embeds[0].toJSON().description
  assert.equal((description.match(/STATE EVENTS/g) || []).length, 1)
  assert.equal((description.match(/Borders open/g) || []).length, 1)
  assert.doesNotMatch(description, /30 minutes|5 minutes/)
})
