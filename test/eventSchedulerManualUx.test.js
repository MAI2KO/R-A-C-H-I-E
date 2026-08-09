const test = require("node:test")
const assert = require("node:assert/strict")

const {
  buildCoreModal,
  buildGroupManagerView,
  buildImageChoiceView,
  buildSingleTimeModal,
  buildTimingChoiceView,
  normalizeGroupInput,
  removeGroup,
  upsertGroup
} = require("../src/eventCreationInteractions")
const {
  actionOptions,
  eventDraft
} = require("../src/eventManagementInteractions")
const {
  formatEventPreview,
  formatUpcomingOccurrencePreview
} = require("../src/eventSchedulerFormatting")
const { formatWeeklyRoundup } = require("../src/weeklyRoundupFormatting")
const { EventValidationError } = require("../src/eventValidation")

function storedEvent(overrides = {}) {
  return {
    id: "41",
    alliance_id: "11",
    alliance_name: "HnC",
    event_name: "Bear Trap",
    first_occurrence_date: "2026-08-10",
    event_time_utc: null,
    recurrence_days: 7,
    advance_reminder_minutes: 30,
    advance_reminder_message: "bear is coming",
    reminder_at_start: true,
    final_reminder_message: "attack!!!!",
    publish_to_alliance: true,
    publish_to_state: true,
    include_in_weekly_roundup: true,
    image_filename: "bear.png",
    image_content_type: "image/png",
    image_byte_size: 8,
    status: "active",
    groups: [
      { group_name: "Best Bear", event_time_utc: "12:15:00", sort_order: 0 },
      { group_name: "Smelly Bear", event_time_utc: "12:15:00", sort_order: 1 }
    ],
    ...overrides
  }
}

test("guided timing separates a single UTC time from structured groups", () => {
  const timing = buildTimingChoiceView("session", {})
  assert.match(timing.content, /Event timing/)
  assert.deepEqual(
    timing.components[0].toJSON().components.slice(0, 2).map(button => button.label),
    ["Single time", "Multiple groups"]
  )
  const single = buildSingleTimeModal("session").toJSON()
  assert.equal(single.components[0].components[0].label, "Event time (UTC)")

  const core = buildCoreModal("session", { mode: "edit", eventName: "Bear Trap" }).toJSON()
  const coreText = JSON.stringify(core)
  assert.equal(core.components.length, 2)
  assert.doesNotMatch(coreText, /alliance|image|group|event time/i)
  assert.doesNotMatch(JSON.stringify([timing, single]), /Group\s*=|group per line/i)
})

test("group manager adds, edits and removes groups while allowing shared times", () => {
  const first = normalizeGroupInput("  Best Bear  ", "1215", [])
  let groups = upsertGroup([], first)
  const second = normalizeGroupInput("Smelly Bear", "12:15", groups)
  groups = upsertGroup(groups, second)
  assert.deepEqual(groups.map(group => [group.groupName, group.eventTimeUtc]), [
    ["Best Bear", "12:15"],
    ["Smelly Bear", "12:15"]
  ])

  const edited = normalizeGroupInput("Smelliest Bear", "6:30pm", groups, 1)
  groups = upsertGroup(groups, edited, 1)
  assert.equal(groups[1].eventTimeUtc, "18:30")
  groups = removeGroup(groups, 0)
  assert.deepEqual(groups, [{ groupName: "Smelliest Bear", eventTimeUtc: "18:30", sortOrder: 0 }])

  const view = buildGroupManagerView("session", { groups, selectedGroupIndex: 0 })
  assert.match(view.content, /Smelliest Bear - 18:30 UTC/)
  assert.deepEqual(
    view.components[1].toJSON().components.map(button => button.label),
    ["Add group", "Edit group", "Remove group", "Continue"]
  )
})

test("group validation trims names and rejects empty or duplicate names case-insensitively", () => {
  const groups = [{ groupName: "Best Bear", eventTimeUtc: "12:15", sortOrder: 0 }]
  assert.throws(() => normalizeGroupInput("   ", "12:15", groups), EventValidationError)
  assert.throws(() => normalizeGroupInput("best bear", "18:30", groups), /Duplicate group name/)
  assert.doesNotThrow(() => normalizeGroupInput("BEST BEAR", "18:30", groups, 0))
})

test("ordinary edits retain alliance and image while exposing separate controls", () => {
  const draft = eventDraft(storedEvent())
  assert.equal(draft.allianceId, "11")
  assert.equal(draft.imageAction, "retain")
  assert.equal(draft.image.originalFilename, "bear.png")
  assert.equal(draft.advanceReminderMessage, "bear is coming")
  assert.equal(draft.finalReminderMessage, "attack!!!!")
  assert.equal(draft.groups.length, 2)

  const actions = actionOptions(storedEvent(), "opaque").map(option => option.label)
  assert.deepEqual(actions.slice(0, 4), [
    "Preview", "Edit details", "Change alliance", "Manage image"
  ])
  const image = buildImageChoiceView("session", draft)
  assert.match(image.content, /currently stored/)
  assert.deepEqual(
    image.components[0].toJSON().components.map(button => button.label),
    ["Keep current image", "Replace image", "Remove image"]
  )
})

test("previews and roundups label UTC and Discord-local times clearly", () => {
  const previewDraft = {
    ...eventDraft(storedEvent({
      event_time_utc: "13:00:00",
      groups: [],
      image_filename: null
    })),
    firstDateIsPast: false
  }
  const preview = formatEventPreview(previewDraft, { now: new Date("2026-08-01T00:00:00Z") })
  assert.match(preview, /13:00 UTC\nLocal time: <t:/)

  const upcoming = formatUpcomingOccurrencePreview(storedEvent({
    event_time_utc: "13:00:00",
    groups: []
  }), [{ occurrenceAt: new Date("2026-08-10T13:00:00Z"), groupName: null }])
  assert.match(upcoming, /13:00 UTC\nLocal time: <t:/)

  const roundup = formatWeeklyRoundup({
    claim: {
      targetKind: "alliance",
      weekStart: new Date("2026-08-10T00:00:00Z"),
      weekEnd: new Date("2026-08-17T00:00:00Z"),
      postWhenEmpty: false
    },
    allianceName: "HnC",
    occurrences: [{
      occurrenceAt: new Date("2026-08-10T13:00:00Z"),
      allianceName: "HnC",
      eventName: "Bear Trap",
      groupName: null
    }]
  })[0].embeds[0].toJSON().description
  assert.match(roundup, /13:00 UTC · <t:\d+:t> local/)
})
