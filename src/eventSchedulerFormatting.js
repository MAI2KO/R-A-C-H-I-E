const EVENTS_PER_PAGE = 3
const {
  getOccurrenceAtIndex,
  getNextOccurrence
} = require("./occurrenceCalculation")

function displayDate(value) {
  if (value instanceof Date) return value.toISOString().slice(0, 10)
  return String(value || "").slice(0, 10)
}

function displayTime(value) {
  return String(value || "").slice(0, 5)
}

function recurrenceLabel(days) {
  if (Number(days) === 7) return "Every week"
  if (Number(days) === 14) return "Every 2 weeks"
  if (Number(days) === 28) return "Every 4 weeks"
  return `Every ${days} days`
}

function reminderLabel(minutes) {
  return minutes === null || minutes === undefined
    ? "None"
    : `${minutes} minutes before`
}

function previewText(value) {
  const normalized = String(value || "").trim()
  return normalized ? normalized.replace(/@/g, "@\u200b") : "Default wording"
}

function imageBehaviour(event) {
  const hasImage = Boolean(event.image || event.has_image || event.image_filename)
  if (!hasImage) return "No image stored"
  return event.advanceReminderMinutes ?? event.advance_reminder_minutes
    ? "Stored; attached only to the alliance advance reminder"
    : "Stored; not posted while the advance reminder is disabled"
}

function publishingLabel(event) {
  const targets = []
  if (event.publish_to_alliance ?? event.publishToAlliance) targets.push("alliance reminders")
  return targets.length ? targets.join(" and ") : "none"
}

function groupLines(groups, maximum = 12, parentFirstOccurrenceDate = null) {
  const visible = groups.slice(0, maximum).map(group => {
    const name = group.group_name ?? group.groupName
    const date = group.first_occurrence_date
      ?? group.firstOccurrenceDate
      ?? parentFirstOccurrenceDate
    const time = group.event_time_utc ?? group.eventTimeUtc
    const datePrefix = date ? `${displayDate(date)} - ` : ""
    return `  ${name}: ${datePrefix}${displayTime(time)} UTC`
  })
  if (groups.length > maximum) visible.push(`  ...and ${groups.length - maximum} more groups`)
  return visible
}

function occurrenceLine(occurrence, { numbered = false, index = 0 } = {}) {
  const timestamp = Math.floor(occurrence.occurrenceAt.getTime() / 1000)
  const utc = occurrence.occurrenceAt.toISOString().slice(0, 16).replace("T", " ")
  const group = occurrence.groupName ? ` - ${occurrence.groupName}` : ""
  return `${numbered ? `${index + 1}. ` : ""}${utc} UTC${group}\nLocal time: <t:${timestamp}:F>`
}

function formatEventPreview(event, { now = new Date() } = {}) {
  const groups = event.groups || []
  const timeBlock = groups.length
    ? groupLines(groups, 12, event.firstOccurrenceDate).join("\n")
    : `${displayTime(event.eventTimeUtc)} UTC`
  const pastNote = event.firstDateIsPast ? " (historical anchor)" : ""
  const anchorOccurrences = getOccurrenceAtIndex(event, 0)
  const anchorLines = anchorOccurrences.slice(0, 12).map(occurrence => occurrenceLine(occurrence))
  if (anchorOccurrences.length > 12) {
    anchorLines.push(`...and ${anchorOccurrences.length - 12} more group anchors`)
  }
  const nextOccurrence = event.firstDateIsPast
    ? getNextOccurrence(event, now)
    : null
  const nextLine = nextOccurrence
    ? `\nNext upcoming: ${occurrenceLine(nextOccurrence)}`
    : ""

  return (
    `${event.mode === "edit" ? "Edit" : "Create"} event preview\n\n` +
    `Alliance: ${event.allianceName}\n` +
    `Event: ${event.eventName}\n` +
    `First date: ${displayDate(event.firstOccurrenceDate)}${pastNote}\n` +
    `Calculated first occurrence${anchorOccurrences.length === 1 ? "" : "s"}:\n` +
    `${anchorLines.join("\n")}${nextLine}\n` +
    `Time${groups.length ? "s" : ""}:\n${timeBlock}\n` +
    `Recurrence: ${recurrenceLabel(event.recurrenceDays)}\n` +
    `Advance reminder: ${reminderLabel(event.advanceReminderMinutes)}\n` +
    `Advance message: ${previewText(event.advanceReminderMessage)}\n` +
    `Final announcement (1 minute before): ${event.reminderAtStart ? "Yes" : "No"}\n` +
    `Final message: ${previewText(event.finalReminderMessage)}\n` +
    `Image delivery: ${imageBehaviour(event)}\n` +
    `Alliance reminders: ${event.publishToAlliance ? "Yes" : "No"}\n` +
    `Weekly roundup: ${event.includeInWeeklyRoundup ? "Yes" : "No"}\n` +
    "State roundup: Automatic when Weekly roundup is enabled and state sharing is enabled"
  ).slice(0, 1950)
}

function formatUpcomingOccurrencePreview(event, occurrences) {
  const status = event.status === "paused" ? " [paused]" : ""
  const lines = occurrences.map((occurrence, index) =>
    occurrenceLine(occurrence, { numbered: true, index })
  )
  return (
    `Upcoming occurrences - next ${occurrences.length}\n\n` +
    `Alliance: ${event.alliance_name}\n` +
    `Event: ${event.event_name}${status}\n` +
    `Recurrence: ${recurrenceLabel(event.recurrence_days)}\n` +
    `Advance: ${reminderLabel(event.advance_reminder_minutes)}\n` +
    `Advance message: ${previewText(event.advance_reminder_message)}\n` +
    `Final announcement: ${event.reminder_at_start ? "1 minute before" : "Disabled"}\n` +
    `Final message: ${previewText(event.final_reminder_message)}\n` +
    `Image delivery: ${imageBehaviour(event)}\n\n` +
    `${lines.join("\n")}`
  ).slice(0, 1950)
}

function formatEventEntry(event) {
  const groups = event.groups || []
  const timeBlock = groups.length
    ? groupLines(groups, 8, event.first_occurrence_date).join("\n")
    : `  ${displayTime(event.event_time_utc)} UTC`

  return (
    `**${event.event_name}** [${event.status}]\n` +
    `Alliance: ${event.alliance_name}\n` +
    `First date: ${displayDate(event.first_occurrence_date)}\n` +
    `Time${groups.length ? "s" : ""}:\n${timeBlock}\n` +
    `Recurrence: ${recurrenceLabel(event.recurrence_days)}\n` +
    `Advance: ${reminderLabel(event.advance_reminder_minutes)}; final: ${event.reminder_at_start ? "Yes" : "No"}\n` +
    `Custom messages: advance ${event.advance_reminder_message ? "Yes" : "No"}; final ${event.final_reminder_message ? "Yes" : "No"}\n` +
    `Publish: ${publishingLabel(event)}; roundup: ${event.include_in_weekly_roundup ? "Yes" : "No"}`
  )
}

function formatEventListPage(events, page, total) {
  if (events.length === 0) return "No active or paused events are configured."
  const totalPages = Math.max(1, Math.ceil(total / EVENTS_PER_PAGE))
  const header = `Scheduled events - page ${page + 1} of ${totalPages}\n\n`
  const body = events.map(formatEventEntry).join("\n\n")
  return `${header}${body}`.slice(0, 1950)
}

module.exports = {
  EVENTS_PER_PAGE,
  displayDate,
  displayTime,
  recurrenceLabel,
  reminderLabel,
  previewText,
  imageBehaviour,
  publishingLabel,
  occurrenceLine,
  formatEventPreview,
  formatUpcomingOccurrencePreview,
  formatEventEntry,
  formatEventListPage
}
