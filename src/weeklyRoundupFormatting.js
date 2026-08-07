const { EmbedBuilder } = require("discord.js")
const { PermanentDeliveryError } = require("./eventDeliveryWorker")
const { compareRoundupOccurrences } = require("./weeklyRoundupCalculation")

const ROUNDUP_DESCRIPTION_LIMIT = 3800

function safeText(value) {
  return String(value ?? "").replace(/@/g, "@\u200b").replace(/[\r\n]+/g, " ").trim()
}

function dateLabel(date) {
  return date.toISOString().slice(0, 10)
}

function occurrenceLine(occurrence, targetKind) {
  const timestamp = Math.floor(occurrence.occurrenceAt.getTime() / 1000)
  const time = occurrence.occurrenceAt.toISOString().slice(11, 16)
  const group = occurrence.groupName ? ` (${safeText(occurrence.groupName)})` : ""
  return `${time} UTC · Local time: <t:${timestamp}:t> · ${safeText(occurrence.eventName)}${group}`
}

function splitRoundupLines(occurrences, targetKind) {
  const parts = []
  let current = ""
  let currentAlliance = null
  let currentDay = null
  for (const occurrence of occurrences) {
    const allianceKey = [
      occurrence.sourceGuildId || "",
      occurrence.allianceId || occurrence.allianceName
    ].join(":")
    const day = dateLabel(occurrence.occurrenceAt)
    const line = occurrenceLine(occurrence, targetKind)
    const allianceChanged = allianceKey !== currentAlliance
    const heading = allianceChanged
      ? `${current ? "\n" : ""}**${safeText(occurrence.allianceName)}**\n`
      : ""
    const dayHeading = allianceChanged || day !== currentDay ? `**${day}**\n` : ""
    const block = `${heading}${dayHeading}${line}`
    if (block.length > ROUNDUP_DESCRIPTION_LIMIT) {
      throw new PermanentDeliveryError("A roundup entry exceeds Discord limits.")
    }
    const separator = current ? "\n" : ""
    if ((current + separator + block).length > ROUNDUP_DESCRIPTION_LIMIT) {
      parts.push(current)
      current = `**${safeText(occurrence.allianceName)}**\n**${day}**\n${line}`
    } else {
      current += separator + block
    }
    currentAlliance = allianceKey
    currentDay = day
  }
  if (current) parts.push(current)
  return parts
}

function formatWeeklyRoundup(payload) {
  const occurrences = [...(payload?.occurrences || [])].sort(compareRoundupOccurrences)
  if (occurrences.length === 0 && !payload?.claim?.postWhenEmpty) return []
  const targetKind = payload.claim.targetKind
  if (!["alliance", "state"].includes(targetKind)) {
    throw new PermanentDeliveryError("Roundup target type is unsupported.")
  }
  const titleBase = targetKind === "state"
    ? "State weekly roundup"
    : `${safeText(payload.allianceName)} weekly roundup`
  const testLabel = payload.claim.isTest ? " — TEST" : ""
  const range = `${dateLabel(payload.claim.weekStart)} to ${dateLabel(payload.claim.weekEnd)}`
  const parts = occurrences.length
    ? splitRoundupLines(occurrences, targetKind)
    : ["No scheduled events this week."]

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
  safeText,
  occurrenceLine,
  splitRoundupLines,
  formatWeeklyRoundup
}
