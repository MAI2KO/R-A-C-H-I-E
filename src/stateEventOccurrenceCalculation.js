const {
  DAY_MS,
  ALLOWED_RECURRENCE_DAYS,
  OccurrenceValidationError,
  parseAnchorDate,
  parseStoredUtcTime,
  requireUtcInstant
} = require("./occurrenceCalculation")

const DEFAULT_MAXIMUM_RESULTS = 1000

function readField(object, camelName, snakeName) {
  return object?.[camelName] ?? object?.[snakeName]
}

function normalizeStateEventDefinition(event) {
  if (!event || typeof event !== "object") {
    throw new OccurrenceValidationError("State event definition is required.")
  }
  const recurrenceValue = readField(event, "recurrenceDays", "recurrence_days")
  const recurrenceDays = recurrenceValue === null || recurrenceValue === undefined || recurrenceValue === ""
    ? null
    : Number(recurrenceValue)
  if (recurrenceDays !== null && !ALLOWED_RECURRENCE_DAYS.has(recurrenceDays)) {
    throw new OccurrenceValidationError(
      "Recurrence must be one-time, 1, 2, 3, 7, 14, 21, 28, 35 or 42 days."
    )
  }
  const phases = Array.isArray(event.phases) ? event.phases : []
  if (!phases.length) throw new OccurrenceValidationError("State event requires at least one phase.")
  return {
    stateEventId: readField(event, "stateEventId", "id") ?? null,
    stateNumber: readField(event, "stateNumber", "state_number"),
    eventName: String(readField(event, "eventName", "event_name") || "").trim(),
    status: String(event.status || "active"),
    recurrenceDays,
    intervalMs: recurrenceDays === null ? null : recurrenceDays * DAY_MS,
    anchorDate: parseAnchorDate(readField(event, "firstOccurrenceDate", "first_occurrence_date")),
    phases: phases.map((phase, index) => {
      const phaseName = String(readField(phase, "phaseName", "phase_name") || "").trim()
      if (!phaseName) throw new OccurrenceValidationError("Every phase requires a name.")
      return {
        phaseId: phase.phaseId ?? phase.phase_id ?? phase.id,
        phaseName,
        phaseSortOrder: Number(readField(phase, "sortOrder", "sort_order") ?? index),
        preAlertMinutes: readField(phase, "preAlertMinutes", "pre_alert_minutes"),
        announceExact: readField(phase, "announceExact", "announce_exact") !== false,
        time: parseStoredUtcTime(readField(phase, "phaseTimeUtc", "phase_time_utc"))
      }
    })
  }
}

function compareStateOccurrences(left, right) {
  return left.occurrenceAt.getTime() - right.occurrenceAt.getTime()
    || Number(left.phaseSortOrder || 0) - Number(right.phaseSortOrder || 0)
    || String(left.phaseName).localeCompare(String(right.phaseName))
    || String(left.phaseId || "").localeCompare(String(right.phaseId || ""))
}

function occurrenceForPhase(definition, phase, occurrenceIndex) {
  if (!Number.isSafeInteger(occurrenceIndex) || occurrenceIndex < 0) {
    throw new OccurrenceValidationError("Occurrence index must be a non-negative safe integer.")
  }
  if (definition.recurrenceDays === null && occurrenceIndex > 0) {
    throw new OccurrenceValidationError("One-time state events have only one occurrence.")
  }
  const intervalMs = definition.intervalMs || 0
  const timestamp = definition.anchorDate.timestamp
    + phase.time.milliseconds
    + occurrenceIndex * intervalMs
  const occurrenceAt = new Date(timestamp)
  if (!Number.isFinite(occurrenceAt.getTime())) {
    throw new OccurrenceValidationError("Occurrence is outside the supported date range.")
  }
  return {
    stateEventId: definition.stateEventId,
    stateNumber: definition.stateNumber,
    eventName: definition.eventName,
    status: definition.status,
    recurrenceDays: definition.recurrenceDays,
    occurrenceIndex,
    occurrenceAt,
    phaseId: phase.phaseId,
    phaseName: phase.phaseName,
    phaseSortOrder: phase.phaseSortOrder,
    preAlertMinutes: phase.preAlertMinutes,
    announceExact: phase.announceExact
  }
}

function nextIndexAtOrAfter(definition, phase, instantMs) {
  const anchorMs = definition.anchorDate.timestamp + phase.time.milliseconds
  if (instantMs <= anchorMs || definition.recurrenceDays === null) return 0
  return Math.ceil((instantMs - anchorMs) / definition.intervalMs)
}

function getStateEventOccurrencesInRange(
  event,
  rangeStart,
  rangeEnd,
  { maximumResults = DEFAULT_MAXIMUM_RESULTS } = {}
) {
  const startMs = requireUtcInstant(rangeStart, "Range start")
  const endMs = requireUtcInstant(rangeEnd, "Range end")
  if (endMs < startMs) throw new OccurrenceValidationError("Range end must not be before range start.")
  if (endMs === startMs) return []
  const definition = normalizeStateEventDefinition(event)
  const results = []
  for (const phase of definition.phases) {
    let index = nextIndexAtOrAfter(definition, phase, startMs)
    while (true) {
      const occurrence = occurrenceForPhase(definition, phase, index)
      if (occurrence.occurrenceAt.getTime() >= endMs) break
      if (occurrence.occurrenceAt.getTime() >= startMs) results.push(occurrence)
      if (results.length > maximumResults) {
        throw new OccurrenceValidationError("Occurrence range exceeds the configured result limit.")
      }
      if (definition.recurrenceDays === null) break
      index += 1
    }
  }
  return results.sort(compareStateOccurrences)
}

module.exports = {
  DEFAULT_MAXIMUM_RESULTS,
  normalizeStateEventDefinition,
  compareStateOccurrences,
  getStateEventOccurrencesInRange
}
