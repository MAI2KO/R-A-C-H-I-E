const EVENTS_PER_PAGE = 3

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

function publishingLabel(event) {
  const targets = []
  if (event.publish_to_alliance ?? event.publishToAlliance) targets.push("alliance")
  if (event.publish_to_state ?? event.publishToState) targets.push("state")
  return targets.length ? targets.join(" and ") : "none"
}

function groupLines(groups, maximum = 12) {
  const visible = groups.slice(0, maximum).map(group => {
    const name = group.group_name ?? group.groupName
    const time = group.event_time_utc ?? group.eventTimeUtc
    return `  ${name}: ${displayTime(time)} UTC`
  })
  if (groups.length > maximum) visible.push(`  ...and ${groups.length - maximum} more groups`)
  return visible
}

function formatEventPreview(event, { image = event.image } = {}) {
  const groups = event.groups || []
  const timeBlock = groups.length
    ? groupLines(groups).join("\n")
    : `${displayTime(event.eventTimeUtc)} UTC`
  const pastNote = event.firstDateIsPast ? " (historical anchor)" : ""

  return (
    `Create event preview\n\n` +
    `Alliance: ${event.allianceName}\n` +
    `Event: ${event.eventName}\n` +
    `First date: ${displayDate(event.firstOccurrenceDate)}${pastNote}\n` +
    `Time${groups.length ? "s" : ""}:\n${timeBlock}\n` +
    `Recurrence: ${recurrenceLabel(event.recurrenceDays)}\n` +
    `Advance reminder: ${reminderLabel(event.advanceReminderMinutes)}\n` +
    `At event start: ${event.reminderAtStart ? "Yes" : "No"}\n` +
    `Publish to: ${publishingLabel(event)}\n` +
    `Monday roundup: ${event.includeInWeeklyRoundup ? "Yes" : "No"}\n` +
    `Image: ${image ? `${image.originalFilename} (${image.byteSize} bytes)` : "None"}`
  ).slice(0, 1950)
}

function formatEventEntry(event) {
  const groups = event.groups || []
  const timeBlock = groups.length
    ? groupLines(groups, 8).join("\n")
    : `  ${displayTime(event.event_time_utc)} UTC`

  return (
    `**${event.event_name}** [${event.status}]\n` +
    `Alliance: ${event.alliance_name}\n` +
    `First date: ${displayDate(event.first_occurrence_date)}\n` +
    `Time${groups.length ? "s" : ""}:\n${timeBlock}\n` +
    `Recurrence: ${recurrenceLabel(event.recurrence_days)}\n` +
    `Advance: ${reminderLabel(event.advance_reminder_minutes)}; start: ${event.reminder_at_start ? "Yes" : "No"}\n` +
    `Publish: ${publishingLabel(event)}; roundup: ${event.include_in_weekly_roundup ? "Yes" : "No"}; image: ${event.has_image ? "Yes" : "No"}`
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
  publishingLabel,
  formatEventPreview,
  formatEventEntry,
  formatEventListPage
}
