const test = require("node:test")
const assert = require("node:assert/strict")

const { formatWeeklyRoundup } = require("../src/weeklyRoundupFormatting")

function payload(overrides = {}) {
  return {
    claim: {
      targetKind: "alliance",
      weekStart: new Date("2026-08-10T00:00:00Z"),
      weekEnd: new Date("2026-08-17T00:00:00Z"),
      postWhenEmpty: false
    },
    allianceName: "HwC",
    occurrences: [],
    stateOccurrences: [],
    ...overrides
  }
}

function occurrence(overrides = {}) {
  return {
    eventId: "event",
    sourceGuildId: "guild",
    allianceId: "main",
    allianceName: "HwC",
    isMainAlliance: true,
    eventName: "SvS",
    groupId: null,
    groupName: null,
    groupSortOrder: 0,
    occurrenceAt: new Date("2026-08-13T10:00:00Z"),
    ...overrides
  }
}

function descriptionFor(input) {
  return formatWeeklyRoundup(input)
    .map(message => message.embeds[0].toJSON().description)
    .join("\n")
}

test("roundups use ordered weekday headings with compact single and grouped event lines", () => {
  const description = descriptionFor(payload({
    occurrences: [
      occurrence({
        eventId: "sub-monday",
        allianceId: "alpha",
        allianceName: "Alpha Sub",
        isMainAlliance: false,
        eventName: "Arena",
        occurrenceAt: new Date("2026-08-10T09:00:00Z")
      }),
      occurrence({
        eventId: "foundry",
        eventName: "Foundry",
        groupId: "l2",
        groupName: "L2",
        groupSortOrder: 1,
        occurrenceAt: new Date("2026-08-14T20:00:00Z")
      }),
      occurrence({
        eventId: "bear",
        eventName: "Bear Trap",
        groupId: "gay",
        groupName: "Gay Bear",
        groupSortOrder: 1,
        occurrenceAt: new Date("2026-08-13T13:40:00Z")
      }),
      occurrence(),
      occurrence({
        eventId: "bear",
        eventName: "Bear Trap",
        groupId: "fat",
        groupName: "Fat Bear",
        groupSortOrder: 0,
        occurrenceAt: new Date("2026-08-13T13:40:00Z")
      }),
      occurrence({
        eventId: "foundry",
        eventName: "Foundry",
        groupId: "l1",
        groupName: "L1",
        groupSortOrder: 0,
        occurrenceAt: new Date("2026-08-14T19:00:00Z")
      }),
      occurrence({
        eventId: "main-monday",
        eventName: "Monday Event",
        occurrenceAt: new Date("2026-08-10T18:00:00Z")
      })
    ]
  }))

  assert.doesNotMatch(description, /2026-08-\d{2}/)
  assert.match(description, /\*\*Monday\*\*/)
  assert.match(description, /\*\*Thursday\*\*/)
  assert.match(description, /\*\*Friday\*\*/)
  assert.doesNotMatch(description, /\*\*Tuesday\*\*|\*\*Wednesday\*\*|\*\*Saturday\*\*|\*\*Sunday\*\*/)
  assert.ok(description.indexOf("Monday") < description.indexOf("Thursday"))
  assert.ok(description.indexOf("Thursday") < description.indexOf("Friday"))
  assert.ok(description.indexOf("HwC") < description.indexOf("Alpha Sub"))
  assert.ok(description.indexOf("SvS") < description.indexOf("Bear Trap"))
  assert.ok(description.indexOf("Fat Bear") < description.indexOf("Gay Bear"))
  assert.ok(description.indexOf("L1") < description.indexOf("L2"))

  assert.match(description, /SvS — 10:00 UTC · <t:\d+:t> local/)
  assert.match(description, /\*\*Bear Trap\*\*\nFat Bear — 13:40 UTC · <t:\d+:t> local\nGay Bear — 13:40 UTC · <t:\d+:t> local/)
  assert.match(description, /\*\*Foundry\*\*\nL1 — 19:00 UTC · <t:\d+:t> local\nL2 — 20:00 UTC · <t:\d+:t> local/)
})

test("sub-alliances remain alphabetical after the main alliance", () => {
  const description = descriptionFor(payload({
    occurrences: [
      occurrence({ allianceId: "zulu", allianceName: "Zulu Sub", isMainAlliance: false }),
      occurrence({ allianceId: "beta", allianceName: "Beta Sub", isMainAlliance: false }),
      occurrence({ allianceId: "main", allianceName: "Main", isMainAlliance: true })
    ]
  }))
  assert.ok(description.indexOf("Main") < description.indexOf("Beta Sub"))
  assert.ok(description.indexOf("Beta Sub") < description.indexOf("Zulu Sub"))
})

test("STATE EVENTS uses weekday and phase formatting without pre-alert rows", () => {
  const description = descriptionFor(payload({
    stateOccurrences: [
      {
        stateEventId: "state-event",
        phaseId: "close",
        stateNumber: "689",
        eventName: "SvS",
        phaseName: "Borders close",
        preAlertMinutes: 30,
        occurrenceAt: new Date("2026-08-15T16:00:00Z")
      },
      {
        stateEventId: "state-event",
        phaseId: "open",
        stateNumber: "689",
        eventName: "SvS",
        phaseName: "Borders open",
        preAlertMinutes: 30,
        occurrenceAt: new Date("2026-08-15T10:00:00Z")
      },
      {
        stateEventId: "state-event",
        phaseId: "battle",
        stateNumber: "689",
        eventName: "SvS",
        phaseName: "Battle starts",
        preAlertMinutes: 5,
        announceExact: false,
        occurrenceAt: new Date("2026-08-15T12:00:00Z")
      }
    ]
  }))

  assert.match(description, /STATE EVENTS\n\n\*\*Saturday\*\*/)
  assert.match(description, /\*\*689 - SvS\*\*/)
  assert.ok(description.indexOf("Borders open") < description.indexOf("Battle starts"))
  assert.ok(description.indexOf("Battle starts") < description.indexOf("Borders close"))
  assert.match(description, /Borders open — 10:00 UTC · <t:\d+:t> local/)
  assert.match(description, /Battle starts — 12:00 UTC · <t:\d+:t> local/)
  assert.doesNotMatch(description, /5 minutes|30 minutes|pre-alert/i)
  assert.doesNotMatch(description, /2026-08-15/)
})
