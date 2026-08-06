const { getOccurrencesInRange } = require("./occurrenceCalculation")

const DAY_MS = 24 * 60 * 60 * 1000
const WEEK_MS = 7 * DAY_MS
const DEFAULT_ROUNDUP_GRACE_MINUTES = 60

function parseRoundupTime(value) {
  const match = String(value || "").match(/^(\d{2}):(\d{2})(?::00)?$/)
  if (!match) throw new Error("Roundup time must be normalized HH:MM UTC")
  const hour = Number(match[1])
  const minute = Number(match[2])
  if (hour > 23 || minute > 59) throw new Error("Roundup time is invalid")
  return { hour, minute }
}

function roundupPeriod(now, weekday, timeUtc, graceMinutes = DEFAULT_ROUNDUP_GRACE_MINUTES) {
  if (!(now instanceof Date) || !Number.isFinite(now.getTime())) {
    throw new Error("Roundup calculation requires a valid date")
  }
  const day = Number(weekday)
  if (!Number.isInteger(day) || day < 0 || day > 6) {
    throw new Error("Roundup weekday must be between 0 and 6")
  }
  const grace = Number(graceMinutes)
  if (!Number.isInteger(grace) || grace < 0) throw new Error("Roundup grace is invalid")
  const { hour, minute } = parseRoundupTime(timeUtc)
  const midnight = Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate())
  const weekStartMs = midnight - ((now.getUTCDay() - day + 7) % 7) * DAY_MS
  const scheduledMs = weekStartMs + (hour * 60 + minute) * 60000
  const ageMs = now.getTime() - scheduledMs
  if (ageMs < 0 || ageMs > grace * 60000) return null

  const weekStart = new Date(weekStartMs)
  return Object.freeze({
    weekStart,
    weekEnd: new Date(weekStartMs + WEEK_MS),
    scheduledFor: new Date(scheduledMs),
    weekStartDate: weekStart.toISOString().slice(0, 10)
  })
}

function occurrenceKey(item) {
  return [
    item.eventId,
    item.groupId || "event",
    item.occurrenceAt.toISOString()
  ].join(":")
}

function compareRoundupOccurrences(left, right) {
  return left.occurrenceAt.getTime() - right.occurrenceAt.getTime()
    || String(left.allianceName).localeCompare(String(right.allianceName))
    || String(left.eventName).localeCompare(String(right.eventName))
    || Number(left.groupSortOrder || 0) - Number(right.groupSortOrder || 0)
    || String(left.groupName || "").localeCompare(String(right.groupName || ""))
    || String(left.eventId).localeCompare(String(right.eventId))
}

function buildRoundupOccurrences(events, weekStart, weekEnd) {
  const unique = new Map()
  for (const event of events || []) {
    if (
      event.status !== "active"
      || event.include_in_weekly_roundup !== true
    ) continue
    for (const occurrence of getOccurrencesInRange(event, weekStart, weekEnd)) {
      const item = {
        ...occurrence,
        allianceName: event.alliance_name,
        eventName: event.event_name
      }
      unique.set(occurrenceKey(item), item)
    }
  }
  return [...unique.values()].sort(compareRoundupOccurrences)
}

module.exports = {
  DAY_MS,
  WEEK_MS,
  DEFAULT_ROUNDUP_GRACE_MINUTES,
  parseRoundupTime,
  roundupPeriod,
  compareRoundupOccurrences,
  buildRoundupOccurrences
}
