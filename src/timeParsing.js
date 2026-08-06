class DateTimeValidationError extends Error {
  constructor(message) {
    super(message)
    this.name = "DateTimeValidationError"
  }
}

function formatTime(hour, minute) {
  if (!Number.isInteger(hour) || hour < 0 || hour > 23) {
    throw new DateTimeValidationError("Hour must be between 0 and 23.")
  }
  if (!Number.isInteger(minute) || minute < 0 || minute > 59) {
    throw new DateTimeValidationError("Minutes must be between 00 and 59.")
  }
  return `${String(hour).padStart(2, "0")}:${String(minute).padStart(2, "0")}`
}

function parseUtcTime(value) {
  const compact = String(value || "").trim().toLowerCase().replace(/\s+/g, "")
  if (!compact) throw new DateTimeValidationError("UTC time is required.")

  const meridiem = compact.match(/^(\d{1,2})(?:([:.])(\d{2}))?(am|pm)$/)
  if (meridiem) {
    let hour = Number(meridiem[1])
    const minute = meridiem[3] === undefined ? 0 : Number(meridiem[3])
    if (hour < 1 || hour > 12) {
      throw new DateTimeValidationError("AM/PM times must use an hour from 1 to 12.")
    }
    if (meridiem[4] === "am") hour = hour === 12 ? 0 : hour
    if (meridiem[4] === "pm") hour = hour === 12 ? 12 : hour + 12
    return formatTime(hour, minute)
  }

  let hour
  let minute
  if (/^\d{1,2}$/.test(compact)) {
    hour = Number(compact)
    minute = 0
  } else if (/^\d{4}$/.test(compact)) {
    hour = Number(compact.slice(0, 2))
    minute = Number(compact.slice(2))
  } else {
    const separated = compact.match(/^(\d{1,2})[:.](\d{2})$/)
    if (!separated) {
      throw new DateTimeValidationError("Use a UTC time such as 18:30, 1830 or 6:30pm.")
    }
    hour = Number(separated[1])
    minute = Number(separated[2])
  }

  return formatTime(hour, minute)
}

function parseIsoDate(value, now = new Date()) {
  const normalized = String(value || "").trim()
  const match = normalized.match(/^(\d{4})-(\d{2})-(\d{2})$/)
  if (!match) {
    throw new DateTimeValidationError("Date must use YYYY-MM-DD.")
  }

  const year = Number(match[1])
  const month = Number(match[2])
  const day = Number(match[3])
  const milliseconds = Date.UTC(year, month - 1, day)
  const parsed = new Date(milliseconds)

  if (
    parsed.getUTCFullYear() !== year
    || parsed.getUTCMonth() !== month - 1
    || parsed.getUTCDate() !== day
  ) {
    throw new DateTimeValidationError("Date is not a real calendar date.")
  }

  const today = Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate())
  return {
    value: normalized,
    date: parsed,
    isPast: milliseconds < today
  }
}

function parseTimeOrGroups(value) {
  const lines = String(value || "")
    .split(/\r?\n/)
    .map(line => line.trim())
    .filter(Boolean)

  if (lines.length === 0) {
    throw new DateTimeValidationError("Enter one UTC time or one group per line.")
  }

  const hasGroups = lines.some(line => line.includes("="))
  if (!hasGroups) {
    if (lines.length !== 1) {
      throw new DateTimeValidationError("An ungrouped event must contain exactly one UTC time.")
    }
    return { eventTimeUtc: parseUtcTime(lines[0]), groups: [] }
  }

  if (lines.some(line => !line.includes("="))) {
    throw new DateTimeValidationError("Every grouped line must use Group name = UTC time.")
  }
  if (lines.length > 20) {
    throw new DateTimeValidationError("A maximum of 20 event groups is supported.")
  }

  const names = new Set()
  const groups = lines.map((line, index) => {
    const separator = line.indexOf("=")
    const groupName = line.slice(0, separator).trim()
    const time = line.slice(separator + 1).trim()
    if (!groupName || groupName.length > 100) {
      throw new DateTimeValidationError("Each group name must be 1 to 100 characters.")
    }
    const key = groupName.toLowerCase()
    if (names.has(key)) {
      throw new DateTimeValidationError(`Duplicate group name: ${groupName}.`)
    }
    names.add(key)
    return {
      groupName,
      eventTimeUtc: parseUtcTime(time),
      sortOrder: index
    }
  })

  return { eventTimeUtc: null, groups }
}

module.exports = {
  DateTimeValidationError,
  parseUtcTime,
  parseIsoDate,
  parseTimeOrGroups
}
