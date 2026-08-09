const { EmbedBuilder } = require("discord.js")

const { PermanentDeliveryError } = require("./eventDeliveryWorker")
const { boundedText } = require("./eventDeliveryFormatting")

function requireOccurrence(value) {
  if (!(value instanceof Date) || !Number.isFinite(value.getTime())) {
    throw new PermanentDeliveryError("State event occurrence time is invalid.")
  }
  return value
}

function stateEventTiming(payload) {
  const occurrence = requireOccurrence(payload?.claim?.occurrenceAt)
  const utcTime = occurrence.toISOString().slice(11, 16)
  const timestamp = Math.floor(occurrence.getTime() / 1000)
  return { utcTime, timestamp }
}

function formatStateEventDelivery(payload, { imageFilename = null, test = false } = {}) {
  const timing = stateEventTiming(payload)
  const deliveryKind = payload?.claim?.deliveryKind
  const phaseName = boundedText(payload?.phase?.name, 100)
  const stateNumber = boundedText(payload?.stateEvent?.stateNumber, 32, "State")
  const eventName = boundedText(payload?.stateEvent?.eventName, 100)
  const parts = [
    `**${test ? "TEST - " : ""}${stateNumber}**`,
    `**${eventName}**`,
    "",
    `${timing.utcTime} UTC`,
    `Local time: <t:${timing.timestamp}:t>`,
    ""
  ]

  if (deliveryKind === "pre_alert") {
    const minutes = Number(payload?.phase?.preAlertMinutes)
    if (![5, 10, 15, 20, 30].includes(minutes)) {
      throw new PermanentDeliveryError("State event pre-alert interval is invalid.")
    }
    parts.push(`${phaseName} in ${minutes} minutes`)
    const message = payload?.phase?.preAlertMessage
    if (String(message || "").trim()) parts.push("", boundedText(message, 500))
  } else if (deliveryKind === "exact") {
    parts.push(phaseName)
    const message = payload?.phase?.exactMessage
    if (String(message || "").trim()) parts.push("", boundedText(message, 500))
  } else {
    throw new PermanentDeliveryError("State event delivery type is unsupported.")
  }

  const embed = new EmbedBuilder()
    .setColor(deliveryKind === "exact" ? 0x3498db : 0xf1c40f)
    .setTitle(boundedText(
      `${test ? "TEST - " : ""}${eventName}: ${phaseName}`,
      256
    ))
    .setDescription(parts.join("\n").slice(0, 4096))

  if (imageFilename) embed.setImage(`attachment://${imageFilename}`)

  return {
    embeds: [embed],
    allowedMentions: { parse: [], repliedUser: false }
  }
}

module.exports = {
  stateEventTiming,
  formatStateEventDelivery
}
