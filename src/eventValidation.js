const ALLOWED_RECURRENCES = new Set([3, 7, 14, 28])
const ALLOWED_ADVANCE_REMINDERS = new Set([null, 5, 10, 15, 20, 30])
const CUSTOM_MESSAGE_MAX_LENGTH = 500

class EventValidationError extends Error {
  constructor(message) {
    super(message)
    this.name = "EventValidationError"
  }
}

function validateText(value, label, maximum = 100) {
  const normalized = String(value || "").trim()
  if (!normalized || normalized.length > maximum) {
    throw new EventValidationError(`${label} must be 1 to ${maximum} characters.`)
  }
  return normalized
}

function normalizeCustomMessage(value, label) {
  const normalized = String(value || "").trim()
  if (!normalized) return null
  if (normalized.length > CUSTOM_MESSAGE_MAX_LENGTH) {
    throw new EventValidationError(`${label} must be 500 characters or fewer.`)
  }
  if (
    /\b(?:javascript|data|vbscript|file):/i.test(normalized)
    || /\]\(\s*(?!https?:\/\/)[a-z][a-z0-9+.-]*:/i.test(normalized)
  ) {
    throw new EventValidationError(`${label} contains an unsupported URL scheme.`)
  }
  return normalized
}

function validateEventDraft(draft) {
  const allianceId = String(draft.allianceId || "").trim()
  if (!/^\d+$/.test(allianceId) || allianceId === "0") {
    throw new EventValidationError("Select a valid alliance or sub-alliance.")
  }
  const allianceName = validateText(draft.allianceName, "Alliance name")
  const eventName = validateText(draft.eventName, "Event name")
  const groups = Array.isArray(draft.groups) ? draft.groups : []
  const recurrenceDays = Number(draft.recurrenceDays)
  const advanceReminderMinutes = draft.advanceReminderMinutes === null
    ? null
    : Number(draft.advanceReminderMinutes)

  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(draft.firstOccurrenceDate || ""))) {
    throw new EventValidationError("A validated first occurrence date is required.")
  }
  if (!ALLOWED_RECURRENCES.has(recurrenceDays)) {
    throw new EventValidationError("Recurrence must be 3, 7, 14 or 28 days.")
  }
  if (!ALLOWED_ADVANCE_REMINDERS.has(advanceReminderMinutes)) {
    throw new EventValidationError("Advance reminder must be none, 5, 10, 15, 20 or 30 minutes.")
  }

  if (groups.length === 0 && !draft.eventTimeUtc) {
    throw new EventValidationError("An ungrouped event requires one UTC event time.")
  }
  if (groups.length > 0 && draft.eventTimeUtc !== null) {
    throw new EventValidationError("A grouped event must not have a server-wide event time.")
  }
  if (draft.grouped === true && groups.length === 0) {
    throw new EventValidationError("A grouped event requires at least one group.")
  }

  const names = new Set()
  const normalizedGroups = groups.map((group, index) => {
    const groupName = validateText(group.groupName, "Group name")
    if (!group.eventTimeUtc) {
      throw new EventValidationError(`Group ${groupName} requires a UTC time.`)
    }
    const key = groupName.toLowerCase()
    if (names.has(key)) throw new EventValidationError(`Duplicate group name: ${groupName}.`)
    names.add(key)
    return { ...group, groupName, sortOrder: group.sortOrder ?? index }
  })

  return {
    ...draft,
    allianceId,
    allianceName,
    eventName,
    recurrenceDays,
    advanceReminderMinutes,
    advanceReminderMessage: normalizeCustomMessage(
      draft.advanceReminderMessage,
      "Advance reminder message"
    ),
    finalReminderMessage: normalizeCustomMessage(
      draft.finalReminderMessage,
      "Final announcement message"
    ),
    groups: normalizedGroups,
    reminderAtStart: Boolean(draft.reminderAtStart),
    publishToAlliance: Boolean(draft.publishToAlliance),
    publishToState: Boolean(draft.publishToState),
    includeInWeeklyRoundup: Boolean(draft.includeInWeeklyRoundup)
  }
}

module.exports = {
  ALLOWED_RECURRENCES,
  ALLOWED_ADVANCE_REMINDERS,
  CUSTOM_MESSAGE_MAX_LENGTH,
  EventValidationError,
  normalizeCustomMessage,
  validateEventDraft
}
