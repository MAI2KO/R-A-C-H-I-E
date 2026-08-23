const assert = require("node:assert/strict")
const test = require("node:test")

const { discordTimestamp, utcAppointmentInstant, validInstant } = require("../src/discordTimeFormatting")
const { occurrenceLine } = require("../src/eventSchedulerFormatting")

test("booking and scheduler formatting share safe Discord-native timestamp rendering", () => {
  const instant = new Date("2030-03-31T00:30:00.000Z")
  const expected = `<t:${Math.floor(instant.getTime() / 1000)}:F>`
  assert.equal(discordTimestamp(instant, "F"), expected)
  assert.match(occurrenceLine({ occurrenceAt: instant, groupName: null }),
    new RegExp(`Local time: ${expected.replace(/[<>:]/g, "\\$&")}$`))
})

test("UTC appointment fallback validates calendar dates and times before creating markup", () => {
  assert.equal(utcAppointmentInstant("2030-08-21", "05:30").toISOString(),
    "2030-08-21T05:30:00.000Z")
  assert.equal(utcAppointmentInstant("2030-02-30", "05:30"), null)
  assert.equal(utcAppointmentInstant("2030-08-21", "25:00"), null)
  assert.equal(discordTimestamp("not-an-instant", "F"), null)
  assert.equal(discordTimestamp("2030-08-21T05:30:00Z", "unsupported"), null)
  assert.equal(validInstant("not-an-instant"), null)
})
