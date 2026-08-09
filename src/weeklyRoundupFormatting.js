const { EmbedBuilder } = require("discord.js")
const { PermanentDeliveryError } = require("./eventDeliveryWorker")
const { compareRoundupOccurrences } = require("./weeklyRoundupCalculation")
const { compareStateRoundupOccurrences } = require("./stateEventRoundupCalculation")

const ROUNDUP_DESCRIPTION_LIMIT = 3800
const WEEKDAY_NAMES = Object.freeze([
  "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"
])

function safeText(value) {
  return String(value ?? "").replace(/@/g, "@\u200b").replace(/[\r\n]+/g, " ").trim()
}

function dateLabel(date) {
  return date.toISOString().slice(0, 10)
}

function weekdayLabel(date) {
  return WEEKDAY_NAMES[date.getUTCDay()]
}

function weekdayOrder(date) {
  return (date.getUTCDay() + 6) % 7
}

function timingText(occurrence) {
  const timestamp = Math.floor(occurrence.occurrenceAt.getTime() / 1000)
  const time = occurrence.occurrenceAt.toISOString().slice(11, 16)
  return `${time} UTC · <t:${timestamp}:t> local`
}

function occurrenceLine(occurrence, targetKind) {
  const label = occurrence.groupName || occurrence.eventName
  return `${safeText(label)} — ${timingText(occurrence)}`
}

function compareWithinEvent(left, right) {
  return left.occurrenceAt.getTime() - right.occurrenceAt.getTime()
    || Number(left.groupSortOrder || 0) - Number(right.groupSortOrder || 0)
    || String(left.groupName || "").localeCompare(String(right.groupName || ""))
    || String(left.groupId || "").localeCompare(String(right.groupId || ""))
}

function eventBlock(eventOccurrences, targetKind) {
  const sorted = [...eventOccurrences].sort(compareWithinEvent)
  if (sorted.some(occurrence => occurrence.groupName)) {
    return `**${safeText(sorted[0].eventName)}**\n` +
      sorted.map(occurrence => occurrenceLine(occurrence, targetKind)).join("\n")
  }
  return sorted.map(occurrence => occurrenceLine(occurrence, targetKind)).join("\n")
}

function allianceEntries(occurrences, targetKind) {
  const alliances = new Map()
  for (const occurrence of [...occurrences].sort(compareRoundupOccurrences)) {
    const allianceKey = [
      occurrence.sourceGuildId || "",
      occurrence.allianceId || occurrence.allianceName
    ].join(":")
    if (!alliances.has(allianceKey)) {
      alliances.set(allianceKey, { allianceKey, allianceName: occurrence.allianceName, occurrences: [] })
    }
    alliances.get(allianceKey).occurrences.push(occurrence)
  }

  const entries = []
  for (const alliance of alliances.values()) {
    const days = new Map()
    for (const occurrence of alliance.occurrences) {
      const dayOrder = weekdayOrder(occurrence.occurrenceAt)
      if (!days.has(dayOrder)) days.set(dayOrder, [])
      days.get(dayOrder).push(occurrence)
    }
    for (const [dayOrder, dayOccurrences] of [...days].sort(([left], [right]) => left - right)) {
      const events = new Map()
      for (const occurrence of dayOccurrences) {
        const eventKey = String(occurrence.eventId || occurrence.eventName)
        if (!events.has(eventKey)) events.set(eventKey, [])
        events.get(eventKey).push(occurrence)
      }
      const orderedEvents = [...events.entries()].sort((left, right) => {
        const leftFirst = [...left[1]].sort(compareWithinEvent)[0]
        const rightFirst = [...right[1]].sort(compareWithinEvent)[0]
        return compareWithinEvent(leftFirst, rightFirst)
          || String(leftFirst.eventName).localeCompare(String(rightFirst.eventName))
          || left[0].localeCompare(right[0])
      })
      for (const [, eventOccurrences] of orderedEvents) {
        entries.push({
          allianceKey: alliance.allianceKey,
          allianceName: alliance.allianceName,
          dayOrder,
          weekday: weekdayLabel(eventOccurrences[0].occurrenceAt),
          body: eventBlock(eventOccurrences, targetKind)
        })
      }
    }
  }
  return entries
}

function packEntries(entries, {
  maximumLength = ROUNDUP_DESCRIPTION_LIMIT,
  rootHeading = entry => `**${safeText(entry.allianceName)}**`,
  rootKey = entry => entry.allianceKey
} = {}) {
  const parts = []
  let current = ""
  let currentRoot = null
  let currentDay = null
  for (const entry of entries) {
    const entryRoot = rootKey(entry)
    const rootChanged = entryRoot !== currentRoot
    const dayChanged = rootChanged || entry.dayOrder !== currentDay
    const blocks = []
    if (rootChanged) blocks.push(rootHeading(entry))
    if (dayChanged) blocks.push(`**${entry.weekday}**`)
    blocks.push(entry.body)
    let addition = blocks.join("\n\n")
    let candidate = current ? `${current}\n\n${addition}` : addition
    if (candidate.length > maximumLength) {
      if (!current) throw new PermanentDeliveryError("A roundup entry exceeds Discord limits.")
      parts.push(current)
      addition = [rootHeading(entry), `**${entry.weekday}**`, entry.body].join("\n\n")
      if (addition.length > maximumLength) {
        throw new PermanentDeliveryError("A roundup entry exceeds Discord limits.")
      }
      current = addition
      candidate = addition
    }
    current = candidate
    currentRoot = entryRoot
    currentDay = entry.dayOrder
  }
  if (current) parts.push(current)
  return parts
}

function splitRoundupLines(occurrences, targetKind) {
  return packEntries(allianceEntries(occurrences, targetKind))
}

function stateOccurrenceLine(occurrence) {
  return `${safeText(occurrence.phaseName)} — ${timingText(occurrence)}`
}

function splitStateRoundupLines(occurrences) {
  const days = new Map()
  for (const occurrence of [...occurrences].sort(compareStateRoundupOccurrences)) {
    const dayOrder = weekdayOrder(occurrence.occurrenceAt)
    if (!days.has(dayOrder)) days.set(dayOrder, [])
    days.get(dayOrder).push(occurrence)
  }
  const entries = []
  for (const [dayOrder, dayOccurrences] of [...days].sort(([left], [right]) => left - right)) {
    const events = new Map()
    for (const occurrence of dayOccurrences) {
      const eventKey = [
        occurrence.stateNumber || "",
        occurrence.stateEventId || occurrence.eventName || ""
      ].join(":")
      if (!events.has(eventKey)) events.set(eventKey, [])
      events.get(eventKey).push(occurrence)
    }
    const orderedEvents = [...events.entries()].sort((left, right) =>
      left[1][0].occurrenceAt.getTime() - right[1][0].occurrenceAt.getTime()
      || String(left[1][0].eventName).localeCompare(String(right[1][0].eventName))
      || left[0].localeCompare(right[0]))
    for (const [eventKey, eventOccurrences] of orderedEvents) {
      const sorted = [...eventOccurrences].sort(compareStateRoundupOccurrences)
      entries.push({
        eventKey,
        dayOrder,
        weekday: weekdayLabel(sorted[0].occurrenceAt),
        body: `**${safeText(sorted[0].stateNumber)} - ${safeText(sorted[0].eventName)}**\n` +
          sorted.map(stateOccurrenceLine).join("\n")
      })
    }
  }
  return packEntries(entries, {
    maximumLength: ROUNDUP_DESCRIPTION_LIMIT - "STATE EVENTS\n\n".length,
    rootHeading: () => "",
    rootKey: () => "state-events"
  }).map(part => part.replace(/^\n+/, ""))
}

function formatWeeklyRoundup(payload) {
  const occurrences = [...(payload?.occurrences || [])].sort(compareRoundupOccurrences)
  const uniqueStateOccurrences = new Map()
  for (const occurrence of payload?.stateOccurrences || []) {
    const key = `${occurrence.stateEventId}:${occurrence.phaseId}:${occurrence.occurrenceAt.toISOString()}`
    if (!uniqueStateOccurrences.has(key)) uniqueStateOccurrences.set(key, occurrence)
  }
  const stateOccurrences = [...uniqueStateOccurrences.values()].sort(compareStateRoundupOccurrences)
  if (occurrences.length === 0 && stateOccurrences.length === 0 && !payload?.claim?.postWhenEmpty) return []
  const targetKind = payload.claim.targetKind
  if (!["alliance", "state"].includes(targetKind)) {
    throw new PermanentDeliveryError("Roundup target type is unsupported.")
  }
  const titleBase = targetKind === "state"
    ? "State weekly roundup"
    : `${safeText(payload.allianceName)} weekly roundup`
  const testLabel = payload.claim.isTest ? " — TEST" : ""
  const range = `${dateLabel(payload.claim.weekStart)} to ${dateLabel(payload.claim.weekEnd)}`
  const parts = []
  if (occurrences.length) parts.push(...splitRoundupLines(occurrences, targetKind))
  if (stateOccurrences.length) {
    for (const section of splitStateRoundupLines(stateOccurrences)) {
      parts.push(`STATE EVENTS\n\n${section}`)
    }
  }
  if (!parts.length) parts.push("No scheduled events this week.")

  return parts.map((description, index) => {
    const partLabel = parts.length > 1 ? ` · Part ${index + 1} of ${parts.length}` : ""
    const embed = new EmbedBuilder()
      .setColor(targetKind === "state" ? 0x3498db : 0x2ecc71)
      .setTitle(`${titleBase}${testLabel}${partLabel}`.slice(0, 256))
      .setDescription(description)
      .setFooter({ text: `UTC week: ${range}` })
    return {
      embeds: [embed],
      allowedMentions: { parse: [], repliedUser: false }
    }
  })
}

module.exports = {
  ROUNDUP_DESCRIPTION_LIMIT,
  WEEKDAY_NAMES,
  safeText,
  weekdayLabel,
  timingText,
  occurrenceLine,
  eventBlock,
  allianceEntries,
  packEntries,
  splitRoundupLines,
  stateOccurrenceLine,
  splitStateRoundupLines,
  formatWeeklyRoundup
}
