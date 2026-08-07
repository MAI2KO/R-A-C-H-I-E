const test = require("node:test")
const assert = require("node:assert/strict")

const {
  roundupPeriod,
  buildRoundupOccurrences
} = require("../src/weeklyRoundupCalculation")
const { claimsForConfigurations } = require("../src/weeklyRoundupGeneration")
const {
  ROUNDUP_DESCRIPTION_LIMIT,
  formatWeeklyRoundup
} = require("../src/weeklyRoundupFormatting")
const { createWeeklyRoundupProcessor } = require("../src/weeklyRoundupProcessor")
const {
  PermanentDeliveryError,
  createEventDeliveryWorker
} = require("../src/eventDeliveryWorker")
const { buildDeliveryClaims } = require("../src/eventDeliveryGeneration")
const {
  eventDraft,
  listView,
  imageChoiceView,
  confirmationView,
  managementSessions
} = require("../src/eventManagementInteractions")
const { InteractionSessionError } = require("../src/interactionSessions")

function event(overrides = {}) {
  return {
    id: "1",
    guild_id: "guild-a",
    game_profile: "wos",
    alliance_name: "North",
    event_name: "Bear Hunt",
    first_occurrence_date: "2028-02-28",
    event_time_utc: "09:00:00",
    recurrence_days: 3,
    advance_reminder_minutes: 10,
    reminder_at_start: true,
    publish_to_alliance: true,
    publish_to_state: true,
    include_in_weekly_roundup: true,
    status: "active",
    schedule_version: 4,
    event_channel_id: "alliance-channel",
    sharing_enabled: true,
    state_guild_id: "state-guild",
    state_event_channel_id: "state-channel",
    groups: [],
    ...overrides
  }
}

test("roundup due calculation is deterministic across UTC boundaries", () => {
  const exact = roundupPeriod(new Date("2028-02-28T09:00:00Z"), 1, "09:00", 60)
  assert.equal(exact.weekStartDate, "2028-02-28")
  assert.equal(exact.weekEnd.toISOString(), "2028-03-06T00:00:00.000Z")
  assert.equal(roundupPeriod(new Date("2028-02-28T08:59:59Z"), 1, "09:00", 60), null)
  assert.equal(
    roundupPeriod(new Date("2028-02-28T09:59:59Z"), 1, "09:00", 60).scheduledFor.toISOString(),
    "2028-02-28T09:00:00.000Z"
  )
  assert.equal(roundupPeriod(new Date("2028-02-28T10:00:01Z"), 1, "09:00", 60), null)

  const custom = roundupPeriod(new Date("2026-12-30T18:30:00Z"), 3, "18:30", 15)
  assert.equal(custom.weekStartDate, "2026-12-30")
  assert.equal(custom.weekEnd.toISOString(), "2027-01-06T00:00:00.000Z")
  const year = roundupPeriod(new Date("2027-01-01T00:05:00Z"), 5, "00:00", 10)
  assert.equal(year.weekStartDate, "2027-01-01")
})

test("weekly selection is half-open, recurring, grouped, filtered and ordered", () => {
  const start = new Date("2028-02-28T00:00:00Z")
  const end = new Date("2028-03-06T00:00:00Z")
  const selected = buildRoundupOccurrences([
    event(),
    event({ id: "2", event_name: "Alpha", first_occurrence_date: "2028-03-06" }),
    event({ id: "3", status: "paused" }),
    event({ id: "4", status: "deleted" }),
    event({ id: "5", include_in_weekly_roundup: false }),
    event({
      id: "6",
      event_name: "Grouped",
      event_time_utc: null,
      recurrence_days: 7,
      groups: [
        { group_id: "61", group_name: "Beta", event_time_utc: "08:00", sort_order: 1 },
        { group_id: "60", group_name: "Alpha", event_time_utc: "08:00", sort_order: 0 }
      ]
    })
  ], start, end)
  assert.equal(selected.filter(item => item.eventId === "1").length, 3)
  assert.equal(selected.some(item => item.eventId === "2"), false)
  assert.equal(selected.some(item => ["3", "4", "5"].includes(String(item.eventId))), false)
  assert.deepEqual(selected.slice(0, 2).map(item => item.groupName), ["Alpha", "Beta"])
  assert.ok(selected.every((item, index) => index === 0
    || selected[index - 1].occurrenceAt <= item.occurrenceAt))
})

test("roundup occurrences group the main alliance first and sort within each alliance", () => {
  const start = new Date("2028-02-28T00:00:00Z")
  const end = new Date("2028-03-06T00:00:00Z")
  const selected = buildRoundupOccurrences([
    event({
      id: "main-later",
      alliance_id: "main",
      alliance_name: "Zulu Main",
      is_default_alliance: true,
      event_name: "Foundry",
      first_occurrence_date: "2028-03-05",
      event_time_utc: "12:00",
      recurrence_days: 7
    }),
    event({
      id: "alpha-later",
      alliance_id: "alpha",
      alliance_name: "Alpha Sub",
      is_default_alliance: false,
      event_name: "SvS",
      first_occurrence_date: "2028-03-04",
      event_time_utc: "10:00",
      recurrence_days: 7
    }),
    event({
      id: "alpha-earlier",
      alliance_id: "alpha",
      alliance_name: "Alpha Sub",
      is_default_alliance: false,
      event_name: "Bear Trap",
      first_occurrence_date: "2028-03-03",
      event_time_utc: "18:00",
      recurrence_days: 7
    }),
    event({
      id: "beta-earlier",
      alliance_id: "beta",
      alliance_name: "Beta Sub",
      is_default_alliance: false,
      event_name: "Arena",
      first_occurrence_date: "2028-03-01",
      event_time_utc: "08:00",
      recurrence_days: 7
    })
  ], start, end)

  assert.deepEqual(selected.map(item => item.eventId), [
    "main-later",
    "alpha-earlier",
    "alpha-later",
    "beta-earlier"
  ])
  assert.deepEqual(selected.map(item => item.allianceName), [
    "Zulu Main",
    "Alpha Sub",
    "Alpha Sub",
    "Beta Sub"
  ])
})

function configuration(overrides = {}) {
  return {
    source_guild_id: "guild-a",
    game_profile: "wos",
    weekly_roundup_day: 1,
    weekly_roundup_time_utc: "09:00:00",
    weekly_roundup_channel_id: "alliance-channel",
    weekly_roundup_enabled: true,
    state_roundup_enabled: true,
    weekly_roundup_not_before: "2028-02-27T00:00:00Z",
    roundup_when_empty: false,
    sharing_enabled: true,
    state_guild_id: "state-guild",
    state_event_channel_id: "state-channel",
    ...overrides
  }
}

test("roundup generation is profile scoped and deduplicates a shared state target", () => {
  const now = new Date("2028-02-28T09:10:00Z")
  const claims = claimsForConfigurations([
    configuration(),
    configuration({ source_guild_id: "guild-b", weekly_roundup_channel_id: "channel-b" }),
    configuration({ source_guild_id: "guild-k", game_profile: "kingshot" })
  ], { gameProfile: "wos", now, graceMinutes: 60 })
  assert.equal(claims.filter(claim => claim.targetKind === "alliance").length, 2)
  assert.equal(claims.filter(claim => claim.targetKind === "state").length, 1)
  assert.ok(claims.every(claim => claim.gameProfile === "wos"))
  assert.deepEqual(claimsForConfigurations([configuration()], {
    gameProfile: "wos",
    now: new Date("2028-02-28T11:00:00Z"),
    graceMinutes: 60
  }), [])
})

test("alliance and state roundup enablement are independent", () => {
  const now = new Date("2028-02-28T09:10:00Z")
  const stateOnly = claimsForConfigurations([configuration({
    weekly_roundup_enabled: false,
    state_roundup_enabled: true
  })], { gameProfile: "wos", now, graceMinutes: 60 })
  assert.deepEqual(stateOnly.map(claim => claim.targetKind), ["state"])

  const allianceOnly = claimsForConfigurations([configuration({
    weekly_roundup_enabled: true,
    state_roundup_enabled: false
  })], { gameProfile: "wos", now, graceMinutes: 60 })
  assert.deepEqual(allianceOnly.map(claim => claim.targetKind), ["alliance"])
})

test("a changed or re-enabled schedule does not replay the elapsed roundup", () => {
  const currentWeek = claimsForConfigurations([configuration({
    weekly_roundup_not_before: "2028-02-28T09:05:00Z"
  })], {
    gameProfile: "wos",
    now: new Date("2028-02-28T09:10:00Z"),
    graceMinutes: 60
  })
  assert.deepEqual(currentWeek, [])

  const nextWeek = claimsForConfigurations([configuration({
    weekly_roundup_not_before: "2028-02-28T09:05:00Z"
  })], {
    gameProfile: "wos",
    now: new Date("2028-03-06T09:10:00Z"),
    graceMinutes: 60
  })
  assert.equal(nextWeek.length, 2)
})

function roundupPayload(overrides = {}) {
  return {
    claim: {
      gameProfile: "wos",
      targetKind: "alliance",
      weekStart: new Date("2028-02-28T00:00:00Z"),
      weekEnd: new Date("2028-03-06T00:00:00Z"),
      postWhenEmpty: false,
      ...overrides.claim
    },
    allianceName: overrides.allianceName || "North",
    occurrences: overrides.occurrences || []
  }
}

test("roundup formatting skips empty output unless configured", () => {
  assert.deepEqual(formatWeeklyRoundup(roundupPayload()), [])
  const messages = formatWeeklyRoundup(roundupPayload({ claim: { postWhenEmpty: true } }))
  assert.equal(messages.length, 1)
  assert.match(messages[0].embeds[0].toJSON().description, /No scheduled events/)
  assert.equal(messages[0].files, undefined)
})

test("state roundups show alliances, group days, split safely and suppress mentions", () => {
  const occurrences = Array.from({ length: 90 }, (_, index) => ({
    eventId: String(index),
    groupId: null,
    groupName: index % 2 ? "Alpha" : null,
    groupSortOrder: 0,
    allianceName: index % 2 ? "@everyone North" : "South",
    eventName: `Event ${index} ${"x".repeat(45)}`,
    occurrenceAt: new Date(Date.UTC(2028, 1, 28 + (index % 6), 8 + (index % 10)))
  })).sort((a, b) => a.occurrenceAt - b.occurrenceAt)
  const messages = formatWeeklyRoundup(roundupPayload({
    claim: { targetKind: "state" },
    occurrences
  }))
  assert.ok(messages.length > 1)
  messages.forEach((message, index) => {
    const embed = message.embeds[0].toJSON()
    assert.match(embed.title, new RegExp(`Part ${index + 1} of ${messages.length}`))
    assert.ok(embed.description.length <= ROUNDUP_DESCRIPTION_LIMIT)
    assert.doesNotMatch(JSON.stringify(message), /@everyone/)
    assert.deepEqual(message.allowedMentions, { parse: [], repliedUser: false })
    assert.equal(message.files, undefined)
  })
  assert.match(
    messages.map(message => message.embeds[0].toJSON().description).join("\n"),
    /\*\*South\*\*/
  )
})

test("alliance and state roundup presentation keep alliance sections together", () => {
  const occurrences = [
    {
      eventId: "sub-early",
      allianceId: "sub",
      sourceGuildId: "guild-a",
      allianceName: "Alpha Sub",
      isMainAlliance: false,
      eventName: "Bear Trap",
      occurrenceAt: new Date("2028-03-03T18:00:00Z")
    },
    {
      eventId: "main-late",
      allianceId: "main",
      sourceGuildId: "guild-a",
      allianceName: "Zulu Main",
      isMainAlliance: true,
      eventName: "Foundry",
      occurrenceAt: new Date("2028-03-05T12:00:00Z")
    },
    {
      eventId: "sub-late",
      allianceId: "sub",
      sourceGuildId: "guild-a",
      allianceName: "Alpha Sub",
      isMainAlliance: false,
      eventName: "SvS",
      occurrenceAt: new Date("2028-03-04T10:00:00Z")
    }
  ]
  for (const targetKind of ["alliance", "state"]) {
    const messages = formatWeeklyRoundup(roundupPayload({
      claim: { targetKind },
      occurrences
    }))
    const description = messages.map(message => message.embeds[0].toJSON().description).join("\n")
    assert.ok(description.indexOf("Zulu Main") < description.indexOf("Alpha Sub"))
    assert.ok(description.indexOf("Bear Trap") < description.indexOf("SvS"))
    assert.equal((description.match(/\*\*Alpha Sub\*\*/g) || []).length, 1)
  }
})

test("event reminder claims carry the current schedule version", () => {
  const claims = buildDeliveryClaims([event()], {
    gameProfile: "wos",
    windowStart: new Date("2028-02-28T08:00:00Z"),
    windowEnd: new Date("2028-02-28T10:00:00Z")
  })
  assert.ok(claims.length > 0)
  assert.ok(claims.every(claim => claim.scheduleVersion === 4))
})

function processorFixture() {
  const sent = new Map()
  let status = "pending"
  let attempt = 0
  let failMiddle = true
  const repository = {
    gameProfile: "wos",
    async listRoundupConfigurations() { return [] },
    async insertMissingClaims() { return 0 },
    async claimDue() {
      if (!['pending', 'failed'].includes(status)) return []
      status = "claimed"
      attempt += 1
      return [{ id: "claim", attempt_count: attempt }]
    },
    async getClaimPayload() {
      return { ...roundupPayload(), sentMessages: new Map(sent) }
    },
    async setPartCount() { return true },
    async renewLease() { return true },
    async recordSentMessage({ messageIndex, sentMessageId, payloadHash }) {
      sent.set(messageIndex, { sentMessageId, payloadHash })
      return true
    },
    async markSent() { status = "sent"; return true },
    async markFailed() { status = "failed"; return true },
    async markPermanentlyFailed() { status = "permanent"; return true }
  }
  const sendCounts = [0, 0, 0]
  const delivery = async () => ({
    messages: [{}, {}, {}],
    async sendPart(index) {
      sendCounts[index] += 1
      if (index === 1 && failMiddle) {
        failMiddle = false
        throw new Error("temporary")
      }
      return `message-${index}`
    }
  })
  const processor = createWeeklyRoundupProcessor({
    repository,
    gameProfile: "wos",
    botInstanceName: "rachie-wos",
    delivery,
    now: () => new Date("2028-02-28T09:10:00Z"),
    workerId: "roundup-worker",
    logger: { error() {} },
    config: {
      graceMinutes: 60,
      batchSize: 10,
      claimLeaseSeconds: 60,
      handlerTimeoutMs: 1000
    }
  })
  return { processor, sent, sendCounts, status: () => status }
}

test("multipart retry records successes, skips sent parts and completes only at the end", async () => {
  const fixture = processorFixture()
  await fixture.processor.tick()
  assert.equal(fixture.status(), "failed")
  assert.deepEqual([...fixture.sent.keys()], [0])
  await fixture.processor.tick()
  assert.equal(fixture.status(), "sent")
  assert.deepEqual(fixture.sendCounts, [1, 2, 1])
  assert.deepEqual([...fixture.sent.keys()], [0, 1, 2])
})

test("roundup permanent failures and maximum retries terminate the claim", async () => {
  for (const failure of [new PermanentDeliveryError("bad target"), new Error("fifth failure")]) {
    let permanentlyFailed = false
    const repository = {
      gameProfile: "wos",
      async listRoundupConfigurations() { return [] },
      async insertMissingClaims() { return 0 },
      async claimDue() { return [{ id: "claim", attempt_count: failure.name === "Error" ? 5 : 1 }] },
      async getClaimPayload() { return { ...roundupPayload(), sentMessages: new Map() } },
      async renewLease() { return true },
      async markPermanentlyFailed() { permanentlyFailed = true; return true },
      async markFailed() { throw new Error("must not retry") }
    }
    const processor = createWeeklyRoundupProcessor({
      repository,
      gameProfile: "wos",
      botInstanceName: "rachie-wos",
      delivery: async () => { throw failure },
      now: () => new Date("2028-02-28T09:10:00Z"),
      workerId: "terminal-worker",
      logger: { error() {} },
      config: { graceMinutes: 60, batchSize: 1, claimLeaseSeconds: 60, handlerTimeoutMs: 1000 }
    })
    await processor.tick()
    assert.equal(permanentlyFailed, true)
  }
})

test("the production polling worker runs roundup ticks independently", async () => {
  let roundupTicks = 0
  const repository = {
    gameProfile: "wos",
    async listActiveEventDefinitions() { return [] },
    async insertMissingDeliveryClaims() { return 0 },
    async claimDueDeliveries() { return [] }
  }
  const worker = createEventDeliveryWorker({
    env: { EVENT_SCHEDULER_ENABLED: "true" },
    health: { available: true, gameProfile: "wos", botInstanceName: "rachie-wos" },
    repository,
    gameProfile: "wos",
    botInstanceName: "rachie-wos",
    deliveryHandler: async () => ({}),
    additionalTick: async () => { roundupTicks += 1 },
    logger: { error() {} },
    workerId: "combined-worker",
    config: {
      lookaheadMinutes: 60, graceMinutes: 60, pollIntervalMs: 5000,
      batchSize: 10, claimLeaseSeconds: 60, handlerTimeoutMs: 1000
    }
  })
  await worker.tick()
  assert.equal(roundupTicks, 1)

  repository.listActiveEventDefinitions = async () => { throw new Error("event generation failed") }
  await worker.tick()
  assert.equal(roundupTicks, 2)
})

test("management controls use opaque session tokens and expose confirmations", () => {
  const source = event({
    image_filename: "old.png",
    image_content_type: "image/png",
    image_byte_size: 8,
    groups: []
  })
  const draft = eventDraft(source)
  assert.equal(draft.imageAction, "retain")
  assert.equal(draft.image.originalFilename, "old.png")
  const eventMap = {}
  const view = listView([source], 0, 1, "opaque-session", eventMap)
  assert.equal(Object.values(eventMap)[0], "1")
  assert.doesNotMatch(JSON.stringify(view.components), /preview:1|edit:1|delete:1/)
  assert.doesNotMatch(view.content, /image/i)
  assert.match(confirmationView("session", "delete", "opaque", source).content, /confirmation/)
  assert.equal(imageChoiceView("edit", draft).components[0].components.length, 3)
})

test("management session ownership rejects wrong user, guild, profile and expiry", () => {
  const context = { userId: "u1", guildId: "g1", gameProfile: "wos" }
  const sessionId = managementSessions.create(context, {})
  assert.throws(() => managementSessions.get(sessionId, { ...context, userId: "u2" }), InteractionSessionError)
  assert.throws(() => managementSessions.get(sessionId, { ...context, guildId: "g2" }), InteractionSessionError)
  assert.throws(() => managementSessions.get(sessionId, { ...context, gameProfile: "kingshot" }), InteractionSessionError)
  managementSessions.cancel(sessionId, context)
  assert.throws(() => managementSessions.get(sessionId, context), InteractionSessionError)
})
