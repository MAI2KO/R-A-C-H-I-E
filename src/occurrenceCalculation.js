const DAY_MS = 24 * 60 * 60 * 1000
const ALLOWED_RECURRENCE_DAYS = new Set([2, 3, 7, 14, 21, 28, 35, 42])
const DEFAULT_MAXIMUM_RESULTS = 1000

class OccurrenceValidationError extends Error {
  constructor(message) {
    super(message)
    this.name = "OccurrenceValidationError"
  }
}

function requireUtcInstant(value, label) {
  if (!(value instanceof Date) || !Number.isFinite(value.getTime())) {
    throw new OccurrenceValidationError(`${label} must be a valid Date.`)
  }
  return value.getTime()
}

function parseAnchorDate(value) {
  if (value instanceof Date && !Number.isFinite(value.getTime())) {
    throw new OccurrenceValidationError("First occurrence date is invalid.")
  }
  const normalized = value instanceof Date
    ? value.toISOString().slice(0, 10)
    : String(value || "").slice(0, 10)
  const match = normalized.match(/^(\d{4})-(\d{2})-(\d{2})$/)
  if (!match) throw new OccurrenceValidationError("First occurrence date must use YYYY-MM-DD.")

  const year = Number(match[1])
  const month = Number(match[2])
  const day = Number(match[3])
  const timestamp = Date.UTC(year, month - 1, day)
  const parsed = new Date(timestamp)
  if (
    parsed.getUTCFullYear() !== year
    || parsed.getUTCMonth() !== month - 1
    || parsed.getUTCDate() !== day
  ) {
    throw new OccurrenceValidationError("First occurrence date is invalid.")
  }
  return { value: normalized, timestamp }
}

function parseStoredUtcTime(value) {
  const match = String(value || "").match(/^(\d{2}):(\d{2})(?::00)?$/)
  if (!match) throw new OccurrenceValidationError("Event time must be normalized HH:MM UTC.")
  const hour = Number(match[1])
  const minute = Number(match[2])
  if (hour > 23 || minute > 59) {
    throw new OccurrenceValidationError("Event time is outside the UTC clock range.")
  }
  return { value: `${match[1]}:${match[2]}`, milliseconds: (hour * 60 + minute) * 60000 }
}

function readField(object, camelName, snakeName) {
  return object?.[camelName] ?? object?.[snakeName]
}

function normalizeEventDefinition(event) {
  if (!event || typeof event !== "object") {
    throw new OccurrenceValidationError("Event definition is required.")
  }
  const recurrenceDays = Number(readField(event, "recurrenceDays", "recurrence_days"))
  if (!ALLOWED_RECURRENCE_DAYS.has(recurrenceDays)) {
    throw new OccurrenceValidationError("Recurrence must be 2, 3, 7, 14, 21, 28, 35 or 42 days.")
  }
  const anchorDate = parseAnchorDate(
    readField(event, "firstOccurrenceDate", "first_occurrence_date")
  )
  const groups = Array.isArray(event.groups) ? event.groups : []
  const eventTime = readField(event, "eventTimeUtc", "event_time_utc")

  let streams
  if (groups.length > 0) {
    if (eventTime !== null && eventTime !== undefined && eventTime !== "") {
      throw new OccurrenceValidationError("Grouped events must not have an event-level time.")
    }
    streams = groups.map((group, index) => {
      const groupName = String(readField(group, "groupName", "group_name") || "").trim()
      if (!groupName) throw new OccurrenceValidationError("Every group requires a name.")
      const groupId = group.groupId ?? group.group_id ?? group.id ?? null
      const sortOrder = Number(readField(group, "sortOrder", "sort_order") ?? index)
      if (!Number.isInteger(sortOrder)) {
        throw new OccurrenceValidationError("Group sort order must be an integer.")
      }
      const groupDateValue = readField(
        group,
        "firstOccurrenceDate",
        "first_occurrence_date"
      )
      const groupDate = groupDateValue
        ? parseAnchorDate(groupDateValue)
        : anchorDate
      if (groupDate.timestamp < anchorDate.timestamp) {
        throw new OccurrenceValidationError(
          `Group ${groupName} first occurrence date must not precede the event anchor.`
        )
      }
      const time = parseStoredUtcTime(readField(group, "eventTimeUtc", "event_time_utc"))
      return {
        groupId,
        groupName,
        groupSortOrder: sortOrder,
        stableIdentifier: String(groupId ?? index),
        dateOffsetMs: groupDate.timestamp - anchorDate.timestamp,
        time
      }
    })
  } else {
    const time = parseStoredUtcTime(eventTime)
    streams = [{
      groupId: null,
      groupName: null,
      groupSortOrder: 0,
      stableIdentifier: "event",
      dateOffsetMs: 0,
      time
    }]
  }

  return {
    eventId: readField(event, "eventId", "id") ?? null,
    allianceName: String(readField(event, "allianceName", "alliance_name") || "").trim(),
    eventName: String(readField(event, "eventName", "event_name") || "").trim(),
    status: String(event.status || "active"),
    recurrenceDays,
    intervalMs: recurrenceDays * DAY_MS,
    anchorDate,
    streams
  }
}

function compareText(left, right) {
  if (left === right) return 0
  return left < right ? -1 : 1
}

function compareOccurrences(left, right) {
  return left.occurrenceAt.getTime() - right.occurrenceAt.getTime()
    || left.groupSortOrder - right.groupSortOrder
    || compareText(left.groupName || "", right.groupName || "")
    || compareText(left.stableIdentifier, right.stableIdentifier)
}

function occurrenceForStream(definition, stream, occurrenceIndex) {
  if (!Number.isSafeInteger(occurrenceIndex) || occurrenceIndex < 0) {
    throw new OccurrenceValidationError("Occurrence index must be a non-negative safe integer.")
  }
  const timestamp = definition.anchorDate.timestamp
    + stream.dateOffsetMs
    + stream.time.milliseconds
    + occurrenceIndex * definition.intervalMs
  if (!Number.isFinite(timestamp)) {
    throw new OccurrenceValidationError("Occurrence is outside the supported date range.")
  }
  const occurrenceAt = new Date(timestamp)
  if (!Number.isFinite(occurrenceAt.getTime())) {
    throw new OccurrenceValidationError("Occurrence is outside the supported date range.")
  }
  return {
    eventId: definition.eventId,
    allianceName: definition.allianceName,
    eventName: definition.eventName,
    status: definition.status,
    recurrenceDays: definition.recurrenceDays,
    occurrenceIndex,
    occurrenceAt,
    groupId: stream.groupId,
    groupName: stream.groupName,
    groupSortOrder: stream.groupSortOrder,
    stableIdentifier: stream.stableIdentifier
  }
}

function getOccurrenceAtIndex(event, occurrenceIndex) {
  const definition = normalizeEventDefinition(event)
  return definition.streams
    .map(stream => occurrenceForStream(definition, stream, occurrenceIndex))
    .sort(compareOccurrences)
}

function streamAnchorMs(definition, stream) {
  return definition.anchorDate.timestamp + stream.dateOffsetMs + stream.time.milliseconds
}

function nextIndexAtOrAfter(definition, stream, instantMs) {
  const anchorMs = streamAnchorMs(definition, stream)
  if (instantMs <= anchorMs) return 0
  return Math.ceil((instantMs - anchorMs) / definition.intervalMs)
}

function previousIndexBefore(definition, stream, instantMs) {
  const anchorMs = streamAnchorMs(definition, stream)
  if (instantMs <= anchorMs) return null
  return Math.ceil((instantMs - anchorMs) / definition.intervalMs) - 1
}

function getNextOccurrences(event, atOrAfter, count = 5) {
  if (!Number.isSafeInteger(count) || count < 0 || count > 100) {
    throw new OccurrenceValidationError("Next occurrence count must be from 0 to 100.")
  }
  if (count === 0) return []
  const instantMs = requireUtcInstant(atOrAfter, "Next-occurrence instant")
  const definition = normalizeEventDefinition(event)
  const candidates = definition.streams.map(stream => ({
    stream,
    index: nextIndexAtOrAfter(definition, stream, instantMs)
  }))
  const results = []

  while (results.length < count) {
    const occurrences = candidates.map(candidate =>
      occurrenceForStream(definition, candidate.stream, candidate.index)
    )
    occurrences.sort(compareOccurrences)
    const selected = occurrences[0]
    results.push(selected)
    const candidate = candidates.find(item =>
      item.stream.stableIdentifier === selected.stableIdentifier
    )
    candidate.index += 1
  }
  return results
}

function getNextOccurrence(event, atOrAfter) {
  return getNextOccurrences(event, atOrAfter, 1)[0] || null
}

function getPreviousOccurrence(event, before) {
  const instantMs = requireUtcInstant(before, "Previous-occurrence instant")
  const definition = normalizeEventDefinition(event)
  const occurrences = definition.streams
    .map(stream => {
      const index = previousIndexBefore(definition, stream, instantMs)
      return index === null ? null : occurrenceForStream(definition, stream, index)
    })
    .filter(Boolean)
    .sort(compareOccurrences)
  return occurrences.at(-1) || null
}

// Range generation is half-open: start is inclusive and end is exclusive.
function getOccurrencesInRange(
  event,
  rangeStart,
  rangeEnd,
  { maximumResults = DEFAULT_MAXIMUM_RESULTS } = {}
) {
  const startMs = requireUtcInstant(rangeStart, "Range start")
  const endMs = requireUtcInstant(rangeEnd, "Range end")
  if (!Number.isSafeInteger(maximumResults) || maximumResults < 1) {
    throw new OccurrenceValidationError("Maximum results must be a positive safe integer.")
  }
  if (endMs < startMs) {
    throw new OccurrenceValidationError("Range end must not be before range start.")
  }
  if (endMs === startMs) return []

  const definition = normalizeEventDefinition(event)
  const results = []
  for (const stream of definition.streams) {
    let index = nextIndexAtOrAfter(definition, stream, startMs)
    while (true) {
      const occurrence = occurrenceForStream(definition, stream, index)
      if (occurrence.occurrenceAt.getTime() >= endMs) break
      results.push(occurrence)
      if (results.length > maximumResults) {
        throw new OccurrenceValidationError("Occurrence range exceeds the configured result limit.")
      }
      index += 1
    }
  }
  return results.sort(compareOccurrences)
}

module.exports = {
  DAY_MS,
  ALLOWED_RECURRENCE_DAYS,
  DEFAULT_MAXIMUM_RESULTS,
  OccurrenceValidationError,
  requireUtcInstant,
  parseAnchorDate,
  parseStoredUtcTime,
  normalizeEventDefinition,
  compareOccurrences,
  getOccurrenceAtIndex,
  getNextOccurrence,
  getPreviousOccurrence,
  getOccurrencesInRange,
  getNextOccurrences
}
