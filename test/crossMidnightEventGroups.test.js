const test = require("node:test")
const assert = require("node:assert/strict")
const { ComponentType, MessageFlags, TextInputStyle } = require("discord.js")

const {
  CREATION_IDS,
  buildGroupModal,
  buildGroupManagerView,
  creationSessions,
  handleEventCreationInteraction,
  handleEventCreationModalOpeningInteraction,
  normalizeGroupInput
} = require("../src/eventCreationInteractions")
const { eventDraft } = require("../src/eventManagementInteractions")
const { buildDeliveryClaims } = require("../src/eventDeliveryGeneration")
const { validateEventDraft } = require("../src/eventValidation")
const {
  getOccurrenceAtIndex,
  getOccurrencesInRange
} = require("../src/occurrenceCalculation")
const { buildRoundupOccurrences } = require("../src/weeklyRoundupCalculation")
const { formatWeeklyRoundup } = require("../src/weeklyRoundupFormatting")

function groupedEvent(overrides = {}) {
  return {
    id: "bear-trap",
    guild_id: "guild",
    game_profile: "wos",
    alliance_id: "alliance",
    alliance_name: "HwC",
    is_default_alliance: true,
    event_name: "Bear Trap",
    first_occurrence_date: "2026-08-08",
    event_time_utc: null,
    recurrence_days: 2,
    advance_reminder_minutes: 30,
    reminder_at_start: true,
    publish_to_alliance: true,
    include_in_weekly_roundup: true,
    status: "active",
    event_channel_id: "channel",
    groups: [
      {
        group_id: "europe",
        group_name: "European Bear",
        first_occurrence_date: "2026-08-08",
        event_time_utc: "17:30:00",
        sort_order: 0
      },
      {
        group_id: "america",
        group_name: "American Bear",
        first_occurrence_date: "2026-08-09",
        event_time_utc: "00:15:00",
        sort_order: 1
      }
    ],
    ...overrides
  }
}

function modalInputs(modal) {
  const payload = modal.toJSON()
  assert.equal(payload.components.length, 3)
  for (const row of payload.components) {
    assert.equal(row.type, ComponentType.ActionRow)
    assert.equal(row.components.length, 1)
    assert.equal(row.components[0].type, ComponentType.TextInput)
    assert.equal(row.components[0].style, TextInputStyle.Short)
  }
  return payload.components.map(row => row.components[0])
}

test("group modal pre-fills parent date and existing group date without premature components", () => {
  const added = modalInputs(buildGroupModal("session", "add", {}, "2026-08-08"))
  assert.deepEqual(added.map(input => input.custom_id), ["name", "date", "time"])
  assert.equal(new Set(added.map(input => input.custom_id)).size, 3)
  assert.equal(added[1].label, "First date (YYYY-MM-DD)")
  assert.equal(added[1].value, "2026-08-08")

  const edited = modalInputs(buildGroupModal("session", "edit", {
    groupName: "American Bear",
    firstOccurrenceDate: "2026-08-09",
    eventTimeUtc: "00:15"
  }, "2026-08-08"))
  assert.deepEqual(edited.map(input => input.value), [
    "American Bear",
    "2026-08-09",
    "00:15"
  ])
})

test("add-group button shows its modal immediately with the parent date", async () => {
  const health = { gameProfile: "wos" }
  const context = { userId: "admin", guildId: "guild-modal", gameProfile: "wos" }
  const sessionId = creationSessions.create(context, {
    firstOccurrenceDate: "2026-08-08",
    groups: []
  })
  const interaction = {
    customId: `${CREATION_IDS.groupAddPrefix}${sessionId}`,
    guildId: context.guildId,
    user: { id: context.userId },
    deferred: false,
    replied: false,
    isButton: () => true,
    isStringSelectMenu: () => false,
    async showModal(modal) {
      assert.equal(this.deferred, false)
      assert.equal(this.replied, false)
      this.modal = modal
    }
  }

  assert.equal(await handleEventCreationModalOpeningInteraction(interaction, { health }), true)
  const inputs = modalInputs(interaction.modal)
  assert.equal(inputs[1].value, "2026-08-08")
  creationSessions.cancel(sessionId, context)
})

test("group modal submission defers before storing normalized date and time", async () => {
  const health = { gameProfile: "wos" }
  const context = { userId: "admin", guildId: "guild-submit", gameProfile: "wos" }
  const sessionId = creationSessions.create(context, {
    firstOccurrenceDate: "2026-08-08",
    groups: []
  })
  const values = {
    name: "American Bear",
    date: "2026-08-09",
    time: "12:15am"
  }
  const interaction = {
    customId: `${CREATION_IDS.groupModalPrefix}${sessionId}:add`,
    guildId: context.guildId,
    user: { id: context.userId },
    deferred: false,
    replied: false,
    fields: { getTextInputValue: customId => values[customId] },
    isButton: () => false,
    isModalSubmit: () => true,
    isStringSelectMenu: () => false,
    isChannelSelectMenu: () => false,
    isChatInputCommand: () => false,
    async deferReply(options) {
      assert.equal(options.flags, MessageFlags.Ephemeral)
      this.deferred = true
    },
    async editReply(payload) {
      assert.equal(this.deferred, true)
      this.edited = payload
    }
  }

  assert.equal(await handleEventCreationInteraction(interaction, {
    repository: {},
    health,
    loadHome: async () => ({})
  }), true)
  assert.deepEqual(creationSessions.get(sessionId, context).data.groups, [{
    groupName: "American Bear",
    firstOccurrenceDate: "2026-08-09",
    eventTimeUtc: "00:15",
    sortOrder: 0
  }])
  creationSessions.cancel(sessionId, context)
})

test("group input accepts same-day and next-day free-form UTC times", () => {
  const sameDay = normalizeGroupInput(
    "European Bear", "5:30pm", [], null, "2026-08-08", "2026-08-08"
  )
  const nextDay = normalizeGroupInput(
    "American Bear", "0015", [sameDay], null, "2026-08-09", "2026-08-08"
  )
  assert.deepEqual(sameDay, {
    groupName: "European Bear",
    firstOccurrenceDate: "2026-08-08",
    eventTimeUtc: "17:30"
  })
  assert.equal(nextDay.firstOccurrenceDate, "2026-08-09")
  assert.equal(nextDay.eventTimeUtc, "00:15")
  assert.throws(
    () => normalizeGroupInput("Bad", "00:15", [], null, "08/09/2026", "2026-08-08"),
    /YYYY-MM-DD/
  )
  assert.throws(
    () => normalizeGroupInput("Bad", "midnight", [], null, "2026-08-09", "2026-08-08"),
    /UTC time/
  )
  assert.throws(
    () => normalizeGroupInput("Bad", "00:15", [], null, "2026-08-07", "2026-08-08"),
    /must not be before/
  )
})

test("two-day recurrence advances same-day and next-day groups without drift", () => {
  const event = groupedEvent()
  assert.deepEqual(
    [0, 1, 2].flatMap(index =>
      getOccurrenceAtIndex(event, index).map(item => item.occurrenceAt.toISOString())
    ),
    [
      "2026-08-08T17:30:00.000Z",
      "2026-08-09T00:15:00.000Z",
      "2026-08-10T17:30:00.000Z",
      "2026-08-11T00:15:00.000Z",
      "2026-08-12T17:30:00.000Z",
      "2026-08-13T00:15:00.000Z"
    ]
  )
})

test("legacy groups without dates retain the parent-date schedule", () => {
  const legacy = groupedEvent({
    groups: [{
      group_id: "legacy",
      group_name: "Legacy Bear",
      event_time_utc: "00:15:00",
      sort_order: 0
    }]
  })
  assert.equal(
    getOccurrenceAtIndex(legacy, 0)[0].occurrenceAt.toISOString(),
    "2026-08-08T00:15:00.000Z"
  )
  assert.equal(
    getOccurrenceAtIndex(legacy, 1)[0].occurrenceAt.toISOString(),
    "2026-08-10T00:15:00.000Z"
  )

  const validated = validateEventDraft({
    allianceId: "1",
    allianceName: "HwC",
    eventName: "Legacy Bear",
    firstOccurrenceDate: "2026-08-08",
    eventTimeUtc: null,
    recurrenceDays: 2,
    advanceReminderMinutes: null,
    groups: [{ groupName: "Legacy", eventTimeUtc: "00:15" }],
    grouped: true
  })
  assert.equal(validated.groups[0].firstOccurrenceDate, "2026-08-08")
})

test("cross-midnight advance and final reminders use the actual group timestamp", () => {
  const claims = buildDeliveryClaims([groupedEvent()], {
    gameProfile: "wos",
    windowStart: new Date("2026-08-08T23:40:00Z"),
    windowEnd: new Date("2026-08-09T00:20:00Z")
  })
  assert.deepEqual(claims.map(claim => ({
    groupId: claim.groupId,
    kind: claim.deliveryKind,
    occurrenceAt: claim.occurrenceAt.toISOString(),
    deliverAt: claim.deliverAt.toISOString()
  })), [
    {
      groupId: "america",
      kind: "advance_reminder",
      occurrenceAt: "2026-08-09T00:15:00.000Z",
      deliverAt: "2026-08-08T23:45:00.000Z"
    },
    {
      groupId: "america",
      kind: "final_reminder",
      occurrenceAt: "2026-08-09T00:15:00.000Z",
      deliverAt: "2026-08-09T00:14:00.000Z"
    }
  ])
})

test("roundup splits one logical grouped event across actual UTC weekdays", () => {
  const event = groupedEvent()
  const weekStart = new Date("2026-08-03T00:00:00Z")
  const weekEnd = new Date("2026-08-10T00:00:00Z")
  const occurrences = buildRoundupOccurrences([event], weekStart, weekEnd)
  const description = formatWeeklyRoundup({
    claim: { targetKind: "alliance", weekStart, weekEnd, postWhenEmpty: false },
    allianceName: "HwC",
    occurrences,
    stateOccurrences: []
  })[0].embeds[0].toJSON().description

  assert.match(description, /\*\*Saturday\*\*[\s\S]*\*\*Bear Trap\*\*[\s\S]*European Bear — 17:30 UTC · <t:\d+:t> local/)
  assert.match(description, /\*\*Sunday\*\*[\s\S]*\*\*Bear Trap\*\*[\s\S]*American Bear — 00:15 UTC · <t:\d+:t> local/)
  assert.ok(description.indexOf("Saturday") < description.indexOf("Sunday"))
})

test("Monday half-open roundup boundary follows the group timestamp", () => {
  const event = groupedEvent({
    first_occurrence_date: "2026-08-09",
    recurrence_days: 7,
    groups: [
      {
        group_id: "sunday",
        group_name: "Sunday Group",
        first_occurrence_date: "2026-08-09",
        event_time_utc: "23:00:00",
        sort_order: 0
      },
      {
        group_id: "monday",
        group_name: "Monday Group",
        first_occurrence_date: "2026-08-10",
        event_time_utc: "00:15:00",
        sort_order: 1
      }
    ]
  })
  const previous = getOccurrencesInRange(
    event,
    new Date("2026-08-03T00:00:00Z"),
    new Date("2026-08-10T00:00:00Z")
  )
  const next = getOccurrencesInRange(
    event,
    new Date("2026-08-10T00:00:00Z"),
    new Date("2026-08-17T00:00:00Z")
  )
  assert.deepEqual(previous.map(item => item.groupName), ["Sunday Group"])
  assert.deepEqual(next.map(item => item.groupName), ["Monday Group", "Sunday Group"])
  assert.equal(next[0].occurrenceIndex, 0)
})

test("management editing preserves explicit and legacy group dates", () => {
  const stored = groupedEvent({
    alliance_id: "1",
    image_filename: null,
    advance_reminder_message: null,
    final_reminder_message: null,
    publish_to_state: false,
    groups: [
      groupedEvent().groups[1],
      {
        group_name: "Legacy Bear",
        event_time_utc: "08:00:00",
        sort_order: 2
      }
    ]
  })
  const draft = eventDraft(stored)
  assert.equal(draft.groups[0].firstOccurrenceDate, "2026-08-09")
  assert.equal(draft.groups[1].firstOccurrenceDate, "2026-08-08")
  assert.equal(draft.recurrenceDays, 2)
  assert.equal(draft.advanceReminderMinutes, 30)

  const manager = buildGroupManagerView("session", draft)
  assert.match(manager.content, /American Bear - 2026-08-09 - 00:15 UTC/)
  assert.match(manager.content, /Legacy Bear - 2026-08-08 - 08:00 UTC/)
})
