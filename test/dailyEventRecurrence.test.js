const test = require("node:test")
const assert = require("node:assert/strict")
const fs = require("node:fs")
const path = require("node:path")

const { validateEventDraft, ALLOWED_RECURRENCES } = require("../src/eventValidation")
const { getNextOccurrences } = require("../src/occurrenceCalculation")
const { recurrenceLabel, formatEventEntry } = require("../src/eventSchedulerFormatting")
const { publicAllianceEventsReadModel } = require("../src/publicAllianceEvents")
const { validateStateEventDraft } = require("../src/stateEventValidation")
const { getStateEventOccurrencesInRange } = require("../src/stateEventOccurrenceCalculation")

function allianceEvent(overrides = {}) {
  return {
    allianceId: "1", allianceName: "Daily Alliance", eventName: "Daily Event",
    firstOccurrenceDate: "2026-09-01", eventTimeUtc: "10:00", groups: [], grouped: false,
    recurrenceDays: 1, advanceReminderMinutes: null, reminderAtStart: false,
    publishToAlliance: true, publishToState: false, includeInWeeklyRoundup: true,
    image: null, ...overrides
  }
}

test("daily alliance events validate for ungrouped and grouped creation", () => {
  assert.equal(validateEventDraft(allianceEvent()).recurrenceDays, 1)
  const grouped = validateEventDraft(allianceEvent({
    eventTimeUtc: null,
    grouped: true,
    groups: [
      { groupName: "Group A", firstOccurrenceDate: "2026-09-01", eventTimeUtc: "10:00" },
      { groupName: "Group B", firstOccurrenceDate: "2026-09-01", eventTimeUtc: "14:00" }
    ]
  }))
  assert.equal(grouped.groups.length, 2)
  assert.deepEqual(getNextOccurrences(grouped, new Date("2026-09-01T00:00:00Z"), 4)
    .map(({ groupName, occurrenceAt }) => [groupName, occurrenceAt.toISOString()]), [
    ["Group A", "2026-09-01T10:00:00.000Z"],
    ["Group B", "2026-09-01T14:00:00.000Z"],
    ["Group A", "2026-09-02T10:00:00.000Z"],
    ["Group B", "2026-09-02T14:00:00.000Z"]
  ])
  assert.deepEqual([...ALLOWED_RECURRENCES], [1, 2, 3, 7, 14, 21, 28, 35, 42])
})

test("daily next occurrences are consecutive anchored UTC days", () => {
  const occurrences = getNextOccurrences(allianceEvent(), new Date("2026-09-01T10:00:00Z"), 3)
  assert.deepEqual(occurrences.map(({ occurrenceAt }) => occurrenceAt.toISOString()), [
    "2026-09-01T10:00:00.000Z",
    "2026-09-02T10:00:00.000Z",
    "2026-09-03T10:00:00.000Z"
  ])
})

test("daily labels are human-readable in listings and public Alliance Events", () => {
  assert.equal(recurrenceLabel(1), "Every day")
  assert.equal(recurrenceLabel(2), "Every 2 days")
  assert.equal(recurrenceLabel(7), "Every week")
  assert.equal(recurrenceLabel(14), "Every 2 weeks")
  assert.equal(recurrenceLabel(28), "Every 4 weeks")
  assert.match(formatEventEntry({
    ...allianceEvent(), alliance_name: "Daily Alliance", event_name: "Daily Event",
    first_occurrence_date: "2026-09-01", event_time_utc: "10:00", recurrence_days: 1,
    status: "paused", advance_reminder_minutes: null, reminder_at_start: false,
    publish_to_alliance: true, include_in_weekly_roundup: true
  }), /Recurrence: Every day/)
  const publicModel = publicAllianceEventsReadModel({
    profile: "wos", now: new Date("2026-09-01T00:00:00Z"), events: [{
      alliance_id: "1", alliance_name: "Daily Alliance", event_id: "2", event_name: "Daily Event",
      first_occurrence_date: "2026-09-01", event_time_utc: "10:00", recurrence_days: 1, groups: []
    }]
  })
  assert.equal(publicModel.alliances[0].events[0].recurrence.summary, "Every day")
  assert.deepEqual(publicModel.alliances[0].events[0].upcoming.map(({ at }) => at), [
    "2026-09-01T10:00:00.000Z", "2026-09-02T10:00:00.000Z", "2026-09-03T10:00:00.000Z"
  ])
})

test("daily state events use the existing generic recurrence calculation", () => {
  const event = validateStateEventDraft({
    eventName: "Daily State Event", firstOccurrenceDate: "2026-09-01", recurrenceDays: 1,
    phases: [{ phaseName: "Reset", phaseTimeUtc: "00:00" }]
  })
  const occurrences = getStateEventOccurrencesInRange(event,
    new Date("2026-09-01T00:00:00Z"), new Date("2026-09-04T00:00:00Z"))
  assert.deepEqual(occurrences.map(({ occurrenceAt }) => occurrenceAt.toISOString()), [
    "2026-09-01T00:00:00.000Z", "2026-09-02T00:00:00.000Z", "2026-09-03T00:00:00.000Z"
  ])
})

test("creation controls, help, and additive constraints expose daily recurrence", () => {
  const creation = fs.readFileSync(path.join(__dirname, "../src/eventCreationInteractions.js"), "utf8")
  const state = fs.readFileSync(path.join(__dirname, "../src/stateEventInteractions.js"), "utf8")
  const help = fs.readFileSync(path.join(__dirname, "../src/eventSchedulerHelp.js"), "utf8")
  const migration = fs.readFileSync(path.join(__dirname, "../migrations/019_daily_event_recurrence.sql"), "utf8")
  assert.match(creation, /selectOption\("Every day", "1"/)
  assert.match(state, /\["Every day", "1"\]/)
  assert.match(help, /Daily recurrence is shown as \*\*Every day\*\*/)
  assert.match(migration, /scheduled_events_recurrence_check/)
  assert.match(migration, /state_events_recurrence_check/)
  assert.match(migration, /IN \(1, 2, 3, 7, 14, 21, 28, 35, 42\)/)
})
