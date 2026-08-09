const { parseAnchorDate } = require("./occurrenceCalculation")
const { parseUtcTime } = require("./timeParsing")
const {
  ALLOWED_ADVANCE_REMINDERS,
  CUSTOM_MESSAGE_MAX_LENGTH,
  EventValidationError,
  normalizeCustomMessage
} = require("./eventValidation")

const ALLOWED_STATE_RECURRENCES = new Set([null, 2, 3, 7, 14, 21, 28, 35, 42])

function normalizeStateNumber(value) {
  const normalized = String(value || "").trim()
  if (!/^\d{1,10}$/.test(normalized)) {
    throw new EventValidationError("State number must contain 1 to 10 digits.")
  }
  return normalized
}

function normalizeStateEventName(value) {
  const normalized = String(value || "").trim()
  if (!normalized || normalized.length > 100) {
    throw new EventValidationError("Event name must be 1 to 100 characters.")
  }
  return normalized
}

function normalizeRecurrenceDays(value) {
  if (value === null || value === undefined || value === "" || value === "none") return null
  const recurrenceDays = Number(value)
  if (!ALLOWED_STATE_RECURRENCES.has(recurrenceDays)) {
    throw new EventValidationError(
      "Recurrence must be one-time, 2, 3, 7, 14, 21, 28, 35 or 42 days."
    )
  }
  return recurrenceDays
}

function normalizeStatePhase(phase, index = 0) {
  const phaseName = String(phase?.phaseName ?? phase?.phase_name ?? "").trim()
  if (!phaseName || phaseName.length > 100) {
    throw new EventValidationError("Phase name must be 1 to 100 characters.")
  }
  const rawPhaseTime = String(phase?.phaseTimeUtc ?? phase?.phase_time_utc ?? "")
  const phaseTimeUtc = parseUtcTime(/^\d{2}:\d{2}:\d{2}(?:\.\d+)?$/.test(rawPhaseTime)
    ? rawPhaseTime.slice(0, 5)
    : rawPhaseTime)
  const preAlertMinutes = phase?.preAlertMinutes ?? phase?.pre_alert_minutes ?? null
  const normalizedPreAlert = preAlertMinutes === null || preAlertMinutes === undefined || preAlertMinutes === ""
    ? null
    : Number(preAlertMinutes)
  if (!ALLOWED_ADVANCE_REMINDERS.has(normalizedPreAlert)) {
    throw new EventValidationError("Pre-alert must be none, 5, 10, 15, 20 or 30 minutes.")
  }
  return {
    id: phase?.id || null,
    phaseName,
    phaseTimeUtc,
    preAlertMinutes: normalizedPreAlert,
    preAlertMessage: normalizeCustomMessage(
      phase?.preAlertMessage ?? phase?.pre_alert_message,
      "Pre-alert message"
    ),
    announceExact: phase?.announceExact ?? phase?.announce_exact ?? true,
    exactMessage: normalizeCustomMessage(
      phase?.exactMessage ?? phase?.exact_message,
      "Exact-time message"
    ),
    sortOrder: Number.isInteger(Number(phase?.sortOrder ?? phase?.sort_order))
      ? Number(phase?.sortOrder ?? phase?.sort_order)
      : index,
    preAlertMedia: phase?.preAlertMedia ?? phase?.pre_alert_media ?? null,
    exactMedia: phase?.exactMedia ?? phase?.exact_media ?? null
  }
}

function validateStateEventDraft(draft) {
  const eventName = normalizeStateEventName(draft?.eventName ?? draft?.event_name)
  const firstOccurrenceDate = parseAnchorDate(
    draft?.firstOccurrenceDate ?? draft?.first_occurrence_date
  ).value
  const recurrenceDays = normalizeRecurrenceDays(
    draft?.recurrenceDays ?? draft?.recurrence_days
  )
  const phases = Array.isArray(draft?.phases) ? draft.phases : []
  if (phases.length === 0) {
    throw new EventValidationError("State event requires at least one phase.")
  }
  const names = new Set()
  const normalizedPhases = phases.map((phase, index) => {
    const normalized = normalizeStatePhase(phase, index)
    const key = normalized.phaseName.toLowerCase()
    if (names.has(key)) throw new EventValidationError(`Duplicate phase name: ${normalized.phaseName}.`)
    names.add(key)
    return normalized
  })
  return {
    ...draft,
    eventName,
    firstOccurrenceDate,
    recurrenceDays,
    phases: normalizedPhases
  }
}

module.exports = {
  ALLOWED_STATE_RECURRENCES,
  CUSTOM_MESSAGE_MAX_LENGTH,
  normalizeStateNumber,
  normalizeStateEventName,
  normalizeRecurrenceDays,
  normalizeStatePhase,
  validateStateEventDraft
}
