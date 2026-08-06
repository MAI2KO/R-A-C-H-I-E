const test = require("node:test")
const assert = require("node:assert/strict")

const {
  DAY_MS,
  OccurrenceValidationError,
  getOccurrenceAtIndex,
  getNextOccurrence,
  getPreviousOccurrence,
  getOccurrencesInRange,
  getNextOccurrences
} = require("../src/occurrenceCalculation")
const {
  formatEventPreview,
  formatUpcomingOccurrencePreview
} = require("../src/eventSchedulerFormatting")
const { buildListView } = require("../src/eventCreationInteractions")
const { createEventSchedulerRepository } = require("../src/eventSchedulerRepository")

function event(overrides = {}) {
  return {
    id: "1",
    alliance_name: "North",
    event_name: "Bear Hunt",
    first_occurrence_date: "2026-01-01",
    event_time_utc: "18:00:00",
    recurrence_days: 7,
    status: "active",
    groups: [],
    ...overrides
  }
}

test("occurrence index zero and positive indexes derive from the anchor", () => {
  assert.equal(
    getOccurrenceAtIndex(event(), 0)[0].occurrenceAt.toISOString(),
    "2026-01-01T18:00:00.000Z"
  )
  assert.equal(
    getOccurrenceAtIndex(event(), 3)[0].occurrenceAt.toISOString(),
    "2026-01-22T18:00:00.000Z"
  )
  const distant = getOccurrenceAtIndex(event({
    first_occurrence_date: "1900-01-01",
    recurrence_days: 3
  }), 50000)[0]
  assert.equal(distant.occurrenceIndex, 50000)
  assert.equal(
    distant.occurrenceAt.getTime(),
    Date.UTC(1900, 0, 1, 18) + 50000 * 3 * DAY_MS
  )
})

test("next occurrence uses an inclusive exact boundary", () => {
  const definition = event()
  assert.equal(
    getNextOccurrence(definition, new Date("2026-01-01T18:00:00Z")).occurrenceIndex,
    0
  )
  assert.equal(
    getNextOccurrence(definition, new Date("2026-01-01T17:59:59Z")).occurrenceIndex,
    0
  )
  assert.equal(
    getNextOccurrence(definition, new Date("2026-01-01T18:00:01Z")).occurrenceIndex,
    1
  )
  assert.ok(
    getNextOccurrence(definition, new Date("2036-01-01T00:00:00Z")).occurrenceIndex > 500
  )
})

test("previous occurrence is strictly before the supplied instant", () => {
  const definition = event()
  assert.equal(getPreviousOccurrence(definition, new Date("2026-01-01T18:00:00Z")), null)
  assert.equal(getPreviousOccurrence(definition, new Date("2026-01-01T18:00:01Z")).occurrenceIndex, 0)
  assert.equal(getPreviousOccurrence(definition, new Date("2026-01-08T18:00:00Z")).occurrenceIndex, 0)
  assert.equal(getPreviousOccurrence(definition, new Date("2026-01-08T18:00:01Z")).occurrenceIndex, 1)
})

test("range generation consistently uses half-open [start, end) boundaries", () => {
  const definition = event()
  const occurrences = getOccurrencesInRange(
    definition,
    new Date("2026-01-01T18:00:00Z"),
    new Date("2026-01-15T18:00:00Z")
  )
  assert.deepEqual(
    occurrences.map(item => item.occurrenceAt.toISOString()),
    ["2026-01-01T18:00:00.000Z", "2026-01-08T18:00:00.000Z"]
  )
  assert.deepEqual(
    getOccurrencesInRange(
      definition,
      new Date("2026-01-02T00:00:00Z"),
      new Date("2026-01-02T00:00:00Z")
    ),
    []
  )
  assert.deepEqual(
    getOccurrencesInRange(
      definition,
      new Date("2026-02-01T00:00:00Z"),
      new Date("2026-02-02T00:00:00Z")
    ),
    []
  )
  assert.throws(
    () => getOccurrencesInRange(
      definition,
      new Date("2026-01-03T00:00:00Z"),
      new Date("2026-01-02T00:00:00Z")
    ),
    /before range start/
  )
})

test("3, 7, 14 and 28 day recurrences do not accumulate drift", () => {
  for (const days of [3, 7, 14, 28]) {
    const definition = event({ recurrence_days: days })
    const atIndex = getOccurrenceAtIndex(definition, 100)[0]
    assert.equal(
      atIndex.occurrenceAt.getTime(),
      Date.UTC(2026, 0, 1, 18) + 100 * days * DAY_MS
    )
  }
})

test("UTC midnight, end-of-day, month, year and leap boundaries are exact", () => {
  assert.equal(
    getOccurrenceAtIndex(event({ event_time_utc: "00:00:00" }), 0)[0]
      .occurrenceAt.toISOString(),
    "2026-01-01T00:00:00.000Z"
  )
  assert.equal(
    getOccurrenceAtIndex(event({ event_time_utc: "23:59:00" }), 0)[0]
      .occurrenceAt.toISOString(),
    "2026-01-01T23:59:00.000Z"
  )
  assert.equal(
    getOccurrenceAtIndex(event({
      first_occurrence_date: "2025-12-31",
      recurrence_days: 3
    }), 1)[0].occurrenceAt.toISOString(),
    "2026-01-03T18:00:00.000Z"
  )
  assert.equal(
    getOccurrenceAtIndex(event({
      first_occurrence_date: "2024-02-29",
      recurrence_days: 3
    }), 1)[0].occurrenceAt.toISOString(),
    "2024-03-03T18:00:00.000Z"
  )
  assert.equal(
    getOccurrenceAtIndex(event({
      first_occurrence_date: "2026-01-31",
      recurrence_days: 28
    }), 1)[0].occurrenceAt.toISOString(),
    "2026-02-28T18:00:00.000Z"
  )
})

test("historical anchors jump arithmetically to far-future queries", () => {
  const historical = event({
    first_occurrence_date: "1000-01-01",
    recurrence_days: 3
  })
  const next = getNextOccurrence(historical, new Date("2500-01-01T00:00:00Z"))
  assert.ok(next.occurrenceIndex > 180000)
  assert.ok(next.occurrenceAt >= new Date("2500-01-01T00:00:00Z"))
  assert.ok(next.occurrenceAt < new Date("2500-01-04T00:00:00Z"))
})

test("group streams retain metadata and sort deterministically", () => {
  const grouped = event({
    event_time_utc: null,
    groups: [
      { group_id: "30", group_name: "Zulu", event_time_utc: "20:00:00", sort_order: 2 },
      { group_id: "20", group_name: "Beta", event_time_utc: "18:00:00", sort_order: 1 },
      { group_id: "10", group_name: "Alpha", event_time_utc: "18:00:00", sort_order: 1 }
    ]
  })
  const atAnchor = getOccurrenceAtIndex(grouped, 0)
  assert.deepEqual(atAnchor.map(item => item.groupId), ["10", "20", "30"])
  assert.deepEqual(atAnchor.map(item => item.groupName), ["Alpha", "Beta", "Zulu"])
  assert.deepEqual(atAnchor.map(item => item.occurrenceIndex), [0, 0, 0])

  const upcoming = getNextOccurrences(grouped, new Date("2026-01-01T18:00:00Z"), 5)
  assert.deepEqual(
    upcoming.map(item => `${item.groupName}:${item.occurrenceIndex}`),
    ["Alpha:0", "Beta:0", "Zulu:0", "Alpha:1", "Beta:1"]
  )
})

test("occurrence API validates inputs and bounds result generation", () => {
  assert.throws(() => getOccurrenceAtIndex(event(), -1), OccurrenceValidationError)
  assert.throws(() => getNextOccurrence(event(), new Date("invalid")), OccurrenceValidationError)
  assert.throws(() => getNextOccurrences(event(), new Date(), 101), OccurrenceValidationError)
  assert.throws(
    () => getOccurrencesInRange(
      event({ recurrence_days: 3 }),
      new Date("2026-01-01T00:00:00Z"),
      new Date("2027-01-01T00:00:00Z"),
      { maximumResults: 2 }
    ),
    /result limit/
  )
})

test("creation preview uses calculated anchors and historical next occurrence", () => {
  const draft = {
    allianceName: "North",
    eventName: "Foundry",
    firstOccurrenceDate: "2026-01-01",
    firstDateIsPast: true,
    eventTimeUtc: null,
    recurrenceDays: 7,
    groups: [
      { groupName: "Alpha", eventTimeUtc: "18:00", sortOrder: 0 },
      { groupName: "Beta", eventTimeUtc: "20:00", sortOrder: 1 }
    ],
    advanceReminderMinutes: null,
    reminderAtStart: false,
    publishToAlliance: true,
    publishToState: false,
    includeInWeeklyRoundup: false,
    image: null
  }
  const preview = formatEventPreview(draft, { now: new Date("2026-02-01T00:00:00Z") })
  assert.match(preview, /historical anchor/)
  assert.match(preview, /Calculated first occurrences/)
  assert.match(preview, /Alpha/)
  assert.match(preview, /Beta/)
  assert.match(preview, /Next upcoming/)
  assert.ok(preview.length <= 1950)
})

test("upcoming preview formats grouped, ungrouped and paused events within limits", () => {
  const paused = event({ status: "paused" })
  const ungrouped = getNextOccurrences(paused, new Date("2026-01-01T00:00:00Z"), 5)
  const ungroupedText = formatUpcomingOccurrencePreview(paused, ungrouped)
  assert.match(ungroupedText, /\[paused\]/)
  assert.match(ungroupedText, /UTC/)
  assert.match(ungroupedText, /<t:/)

  const grouped = event({
    event_time_utc: null,
    groups: [{ group_id: "1", group_name: "Alpha", event_time_utc: "18:00:00", sort_order: 0 }]
  })
  assert.match(
    formatUpcomingOccurrencePreview(
      grouped,
      getNextOccurrences(grouped, new Date("2026-01-01T00:00:00Z"), 5)
    ),
    /Alpha/
  )
  assert.ok(ungroupedText.length <= 1950)
})

test("event listing exposes bounded preview controls", () => {
  const events = [1, 2, 3].map(id => ({
    ...event({ id: String(id) }),
    total_count: 3,
    has_image: false,
    publish_to_alliance: true,
    publish_to_state: false,
    include_in_weekly_roundup: false,
    reminder_at_start: false,
    advance_reminder_minutes: null
  }))
  const view = buildListView(events, 0, 3)
  assert.equal(view.components.length, 2)
  assert.deepEqual(
    view.components[0].components.map(button => button.data.custom_id),
    ["ep:1:0", "ep:2:0", "ep:3:0"]
  )
  assert.ok(view.content.length <= 1950)
})

test("preview repository read excludes deleted events and preserves profile scope", async () => {
  const calls = []
  const pool = {
    async query(text, values) {
      calls.push({ text, values })
      return { rows: [] }
    }
  }
  await createEventSchedulerRepository(pool, "wos").getEvent("same-guild", "7")
  await createEventSchedulerRepository(pool, "kingshot").getEvent("same-guild", "7")
  assert.match(calls[0].text, /e\.status IN \('active', 'paused'\)/)
  assert.deepEqual(calls[0].values, ["7", "same-guild", "wos"])
  assert.deepEqual(calls[1].values, ["7", "same-guild", "kingshot"])
})
