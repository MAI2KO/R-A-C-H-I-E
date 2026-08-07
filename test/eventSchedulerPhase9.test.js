const test = require("node:test")
const assert = require("node:assert/strict")

const {
  assertMigrationHasNoTransactionControl,
  withoutDollarQuotedBodies
} = require("../src/migrate")
const {
  ALLOWED_ADVANCE_REMINDERS,
  EventValidationError,
  validateEventDraft
} = require("../src/eventValidation")
const { buildDeliveryClaims } = require("../src/eventDeliveryGeneration")
const { formatAllianceEventDelivery } = require("../src/eventDeliveryFormatting")
const { deliveryUsesImage } = require("../src/discordEventDelivery")
const {
  buildAllianceSelectionView,
  buildMessagesModal,
  buildTimingView
} = require("../src/eventCreationInteractions")
const { eventDraft } = require("../src/eventManagementInteractions")
const { allianceListView } = require("../src/allianceManagementInteractions")
const { formatEventPreview } = require("../src/eventSchedulerFormatting")
const { formatWeeklyRoundup } = require("../src/weeklyRoundupFormatting")

const occurrence = new Date("2026-08-10T18:00:00Z")

function draft(overrides = {}) {
  return {
    allianceId: "11",
    allianceName: "YOU",
    eventName: "Foundry",
    firstOccurrenceDate: "2026-08-10",
    eventTimeUtc: "18:00",
    groups: [],
    grouped: false,
    recurrenceDays: 7,
    advanceReminderMinutes: 30,
    advanceReminderMessage: null,
    reminderAtStart: true,
    finalReminderMessage: null,
    publishToAlliance: true,
    publishToState: true,
    includeInWeeklyRoundup: true,
    image: null,
    ...overrides
  }
}

function generationEvent(overrides = {}) {
  return {
    id: "41",
    guild_id: "guild-wos",
    game_profile: "wos",
    schedule_version: 1,
    alliance_id: "11",
    alliance_name: "YOU",
    event_name: "Foundry",
    first_occurrence_date: "2026-08-10",
    event_time_utc: "18:00",
    recurrence_days: 7,
    advance_reminder_minutes: 30,
    reminder_at_start: false,
    publish_to_alliance: true,
    publish_to_state: true,
    status: "active",
    event_channel_id: "channel-wos",
    groups: [],
    ...overrides
  }
}

function claimsFor(event, start = "2026-08-10T16:00:00Z", end = "2026-08-10T19:00:00Z") {
  return buildDeliveryClaims([event], {
    gameProfile: "wos",
    windowStart: new Date(start),
    windowEnd: new Date(end)
  })
}

function deliveryPayload(deliveryKind, custom = {}) {
  const deliverAt = deliveryKind === "final_reminder"
    ? new Date("2026-08-10T17:59:00Z")
    : new Date("2026-08-10T17:45:00Z")
  return {
    claim: {
      gameProfile: "wos",
      deliveryKind,
      targetKind: "alliance",
      targetGuildId: "guild-wos",
      targetChannelId: "channel-wos",
      occurrenceAt: new Date(occurrence),
      deliverAt
    },
    event: {
      eventName: "Foundry",
      guildId: "guild-wos",
      recurrenceDays: 7,
      advanceReminderMessage: custom.advance,
      finalReminderMessage: custom.final
    },
    alliance: { name: "YOU", guildId: "guild-wos" },
    group: { name: "Alpha" },
    image: null
  }
}

test("all six advance choices validate and generate exactly one correctly timed claim", () => {
  assert.deepEqual([...ALLOWED_ADVANCE_REMINDERS], [null, 5, 10, 15, 20, 30])
  assert.equal(claimsFor(generationEvent({
    advance_reminder_minutes: null
  })).length, 0)

  for (const minutes of [5, 10, 15, 20, 30]) {
    const validated = validateEventDraft(draft({ advanceReminderMinutes: minutes }), {
      stateLinkEnabled: true
    })
    assert.equal(validated.advanceReminderMinutes, minutes)
    const claims = claimsFor(generationEvent({ advance_reminder_minutes: minutes }))
    assert.equal(claims.length, 1, `${minutes} minutes`)
    assert.equal(claims[0].deliveryKind, "advance_reminder")
    assert.equal(
      claims[0].deliverAt.toISOString(),
      new Date(occurrence.getTime() - minutes * 60000).toISOString()
    )
  }

  for (const value of [-1, 1, 25, 60, "five"]) {
    assert.throws(
      () => validateEventDraft(draft({ advanceReminderMinutes: value }), {
        stateLinkEnabled: true
      }),
      EventValidationError
    )
  }
})

test("the optional final announcement is one minute early across UTC boundaries", () => {
  const disabled = claimsFor(generationEvent({
    advance_reminder_minutes: null,
    reminder_at_start: false
  }))
  assert.deepEqual(disabled, [])

  const boundaries = [
    ["2026-09-01", "2026-08-31T23:59:00.000Z"],
    ["2026-01-01", "2025-12-31T23:59:00.000Z"],
    ["2024-03-01", "2024-02-29T23:59:00.000Z"]
  ]
  for (const [date, expected] of boundaries) {
    const start = `${date}T00:00:00Z`
    const startDate = new Date(start)
    const claims = claimsFor(generationEvent({
      first_occurrence_date: date,
      event_time_utc: "00:00",
      advance_reminder_minutes: null,
      reminder_at_start: true
    }), new Date(startDate.getTime() - 3600000).toISOString(), new Date(startDate.getTime() + 3600000).toISOString())
    assert.equal(claims.length, 1)
    assert.equal(claims[0].deliveryKind, "final_reminder")
    assert.equal(claims[0].deliverAt.toISOString(), expected)
    assert.notEqual(claims[0].deliverAt.getTime(), claims[0].occurrenceAt.getTime())
  }
})

test("custom messages are trimmed, bounded and reject unsafe schemes", () => {
  const validated = validateEventDraft(draft({
    advanceReminderMessage: "  Gather at the gate  ",
    finalReminderMessage: "   "
  }), { stateLinkEnabled: true })
  assert.equal(validated.advanceReminderMessage, "Gather at the gate")
  assert.equal(validated.finalReminderMessage, null)
  assert.throws(() => validateEventDraft(draft({
    advanceReminderMessage: "x".repeat(501)
  }), { stateLinkEnabled: true }), /500/)
  assert.throws(() => validateEventDraft(draft({
    finalReminderMessage: "[click](javascript:alert(1))"
  }), { stateLinkEnabled: true }), /URL scheme/)
  assert.throws(() => validateEventDraft(draft({ allianceId: null }), {
    stateLinkEnabled: true
  }), /valid alliance/)
})

test("advance and final custom messages stay isolated in concise reminders", () => {
  const advance = formatAllianceEventDelivery(deliveryPayload("advance_reminder", {
    advance: "Advance @everyone",
    final: "Final only"
  }))
  const advanceJson = advance.embeds[0].toJSON()
  assert.match(advanceJson.title, /Foundry/)
  assert.match(advanceJson.description, /YOU/)
  assert.match(advanceJson.description, /Alpha/)
  assert.match(advanceJson.description, /Starts in 15 minutes/)
  assert.match(advanceJson.description, /Advance @​everyone/)
  assert.equal(advanceJson.fields, undefined)
  assert.doesNotMatch(JSON.stringify(advanceJson), /When|UTC|<t:|Recurrence|Status:|Alliance message:/i)
  assert.doesNotMatch(JSON.stringify(advanceJson), /Final only/)
  assert.deepEqual(advance.allowedMentions, { parse: [], repliedUser: false })

  const final = formatAllianceEventDelivery(deliveryPayload("final_reminder", {
    advance: "Advance only",
    final: "Final message"
  }))
  const finalJson = final.embeds[0].toJSON()
  assert.match(finalJson.title, /About to start/)
  assert.match(finalJson.description, /approximately 1 minute/)
  assert.match(finalJson.description, /Final message/)
  assert.equal(finalJson.fields, undefined)
  assert.doesNotMatch(JSON.stringify(finalJson), /Advance only/)

  const defaults = formatAllianceEventDelivery(deliveryPayload("advance_reminder", {
    advance: "   "
  })).embeds[0].toJSON()
  assert.equal(defaults.description, "YOU\nFoundry\nAlpha\n\nStarts in 15 minutes")
})

test("images are eligible only for alliance advance reminders at every valid offset", () => {
  for (const minutes of [5, 10, 15, 20, 30]) {
    const payload = deliveryPayload("advance_reminder")
    payload.claim.deliverAt = new Date(payload.claim.occurrenceAt.getTime() - minutes * 60000)
    assert.equal(deliveryUsesImage(payload), true)
  }
  assert.equal(deliveryUsesImage(deliveryPayload("final_reminder")), false)
  assert.equal(deliveryUsesImage({
    ...deliveryPayload("advance_reminder"),
    claim: { ...deliveryPayload("advance_reminder").claim, targetKind: "state" }
  }), false)
  assert.deepEqual(claimsFor(generationEvent({
    advance_reminder_minutes: null,
    reminder_at_start: false,
    image_filename: "stored.png"
  })), [])
})

test("roundups contain distinct alliance names and never contain image payloads", () => {
  const base = {
    claim: {
      targetKind: "alliance",
      weekStart: new Date("2026-08-10T00:00:00Z"),
      weekEnd: new Date("2026-08-17T00:00:00Z"),
      postWhenEmpty: false
    },
    allianceName: "YOU",
    occurrences: [
      { occurrenceAt: new Date("2026-08-10T18:00:00Z"), allianceName: "YOU", eventName: "Foundry" },
      { occurrenceAt: new Date("2026-08-10T19:00:00Z"), allianceName: "YOU2", eventName: "Bear Hunt" }
    ]
  }
  const alliance = formatWeeklyRoundup(base)
  const state = formatWeeklyRoundup({
    ...base,
    claim: { ...base.claim, targetKind: "state" }
  })
  for (const messages of [alliance, state]) {
    const json = JSON.stringify(messages)
    assert.match(json, /YOU/)
    assert.match(json, /YOU2/)
    assert.ok(messages.every(message => message.files === undefined))
    assert.ok(messages.every(message => message.allowedMentions.parse.length === 0))
  }
})

test("alliance selectors and management controls expose names through opaque values", () => {
  const alliances = [
    { id: "101", alliance_name: "YOU", is_default: true, managed_event_count: 1 },
    { id: "102", alliance_name: "YOU2", is_default: false, managed_event_count: 2 }
  ]
  const selection = buildAllianceSelectionView("session", alliances, 0, 2, "101")
  const selectJson = selection.view.components[0].toJSON()
  assert.deepEqual(Object.values(selection.tokenMap).sort(), ["101", "102"])
  assert.ok(selectJson.components[0].options.every(option => !["101", "102"].includes(option.value)))
  const currentToken = Object.entries(selection.tokenMap)
    .find(([, allianceId]) => allianceId === "101")[0]
  assert.equal(
    selectJson.components[0].options.find(option => option.value === currentToken).default,
    true
  )
  assert.doesNotMatch(selection.view.content, /101|102/)

  const management = allianceListView(alliances, 0, 2, "session", "101")
  assert.doesNotMatch(management.view.content, /101|102/)
  assert.deepEqual(Object.values(management.tokenMap).sort(), ["101", "102"])

  const edit = eventDraft({
    id: "41",
    alliance_id: "102",
    alliance_name: "YOU2",
    event_name: "Foundry",
    first_occurrence_date: "2026-08-10",
    event_time_utc: "18:00",
    groups: [],
    recurrence_days: 7,
    advance_reminder_minutes: 15,
    advance_reminder_message: "Prepare",
    reminder_at_start: true,
    final_reminder_message: "Go",
    publish_to_alliance: true,
    publish_to_state: true,
    include_in_weekly_roundup: true
  })
  assert.equal(edit.allianceId, "102")
  assert.equal(edit.advanceReminderMessage, "Prepare")
  assert.equal(edit.finalReminderMessage, "Go")
})

test("creation controls expose all offsets and optional custom-message fields", () => {
  const timing = buildTimingView("session", draft())
  const reminderOptions = timing.components[1].toJSON().components[0].options
    .map(option => option.value)
  assert.deepEqual(reminderOptions, ["none", "5", "10", "15", "20", "30"])
  const modal = buildMessagesModal("session", draft()).toJSON()
  assert.equal(modal.components.length, 2)
  assert.ok(modal.components.every(row => row.components[0].required === false))

  const preview = formatEventPreview(draft({
    advanceReminderMinutes: null,
    advanceReminderMessage: "Advance custom",
    finalReminderMessage: "Final custom",
    image: { originalFilename: "event.png" }
  }), { now: new Date("2026-08-01T00:00:00Z") })
  assert.match(preview, /Alliance: YOU/)
  assert.match(preview, /Advance custom/)
  assert.match(preview, /Final custom/)
  assert.match(preview, /not posted while the advance reminder is disabled/)
  assert.match(preview, /Weekly roundup: Yes/)
  assert.match(preview, /State roundup: Automatic/)
  assert.doesNotMatch(preview, /State weekly roundup|State-roundup eligibility/)
})

test("migration transaction guard ignores function bodies but rejects top-level control", () => {
  const functionSql = `CREATE FUNCTION demo() RETURNS void AS $$\nBEGIN\n  RETURN;\nEND;\n$$ LANGUAGE plpgsql;`
  assert.doesNotThrow(() => assertMigrationHasNoTransactionControl("function.sql", functionSql))
  assert.equal(withoutDollarQuotedBodies(functionSql).includes("BEGIN"), false)
  assert.throws(
    () => assertMigrationHasNoTransactionControl("unsafe.sql", "BEGIN;\nSELECT 1;\nCOMMIT;"),
    /must not contain transaction control/
  )
})
