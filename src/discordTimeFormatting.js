const ALLOWED_STYLES = new Set(["t", "T", "d", "D", "f", "F", "R"])

function validInstant(value) {
  if (value === null || value === undefined || String(value).trim() === "") return null
  const instant = value instanceof Date ? value : new Date(value)
  if (Number.isNaN(instant.getTime())) return null
  return instant
}

function discordTimestamp(value, style = "F") {
  if (!ALLOWED_STYLES.has(style)) return null
  const instant = validInstant(value)
  if (!instant) return null
  return `<t:${Math.floor(instant.getTime() / 1000)}:${style}>`
}

function utcAppointmentInstant(date, time) {
  const dateText = String(date || "").slice(0, 10)
  const timeText = String(time || "").slice(0, 5)
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateText) || !/^\d{2}:\d{2}$/.test(timeText)) return null
  const instant = new Date(`${dateText}T${timeText}:00.000Z`)
  if (Number.isNaN(instant.getTime()) || instant.toISOString().slice(0, 16) !== `${dateText}T${timeText}`) return null
  return instant
}

module.exports = { discordTimestamp, utcAppointmentInstant, validInstant }
