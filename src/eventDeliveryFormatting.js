const { EmbedBuilder } = require("discord.js")

const { PermanentDeliveryError } = require("./eventDeliveryWorker")

const EMBED_TITLE_LIMIT = 256
const EMBED_DESCRIPTION_LIMIT = 4096

function neutralizeMentions(value) {
  return String(value ?? "").replace(/@/g, "@\u200b")
}

function boundedText(value, maximum, fallback = "Not specified") {
  const normalized = neutralizeMentions(value).trim() || fallback
  return normalized.length <= maximum
    ? normalized
    : `${normalized.slice(0, Math.max(0, maximum - 3))}...`
}

function requireOccurrence(value) {
  if (!(value instanceof Date) || !Number.isFinite(value.getTime())) {
    throw new PermanentDeliveryError("Delivery occurrence time is invalid.")
  }
  return value
}

function timingDetails(payload) {
  const occurrence = requireOccurrence(payload?.claim?.occurrenceAt)
  if (payload.claim.deliveryKind === "final_reminder") {
    const deliverAt = payload.claim.deliverAt
    if (
      !(deliverAt instanceof Date)
      || !Number.isFinite(deliverAt.getTime())
      || occurrence.getTime() - deliverAt.getTime() !== 60000
    ) {
      throw new PermanentDeliveryError("Final reminder time is invalid.")
    }
    return { status: "About to start\nStarts in approximately 1 minute" }
  }
  if (payload.claim.deliveryKind === "event_start") {
    throw new PermanentDeliveryError("Exact-start reminders are disabled.")
  }
  if (payload.claim.deliveryKind !== "advance_reminder") {
    throw new PermanentDeliveryError("Delivery type is unsupported.")
  }
  const deliverAt = payload.claim.deliverAt
  if (!(deliverAt instanceof Date) || !Number.isFinite(deliverAt.getTime())) {
    throw new PermanentDeliveryError("Delivery reminder time is invalid.")
  }
  const differenceMs = occurrence.getTime() - deliverAt.getTime()
  const minutes = differenceMs / 60000
  if (!Number.isInteger(minutes) || ![5, 10, 15, 20, 30].includes(minutes)) {
    throw new PermanentDeliveryError("Advance reminder interval is invalid.")
  }
  return { status: `Starts in ${minutes} minutes` }
}

function formatAllianceEventDelivery(payload, { imageFilename = null } = {}) {
  const timing = timingDetails(payload)
  const eventName = boundedText(payload?.event?.eventName, 220)
  const allianceName = boundedText(payload?.alliance?.name, 1000)
  const groupName = payload?.group?.name
    ? boundedText(payload.group.name, 1000)
    : null
  const descriptionParts = [allianceName, eventName]
  if (groupName) descriptionParts.push(groupName)
  descriptionParts.push("", timing.status)
  const customMessage = payload.claim.deliveryKind === "final_reminder"
    ? payload?.event?.finalReminderMessage
    : payload?.event?.advanceReminderMessage
  if (String(customMessage || "").trim()) {
    descriptionParts.push("", boundedText(customMessage, 500))
  }

  const titlePrefix = payload.claim.deliveryKind === "final_reminder"
    ? "About to start: "
    : "Event reminder: "
  const embed = new EmbedBuilder()
    .setColor(payload.claim.deliveryKind === "final_reminder" ? 0x2ecc71 : 0xf1c40f)
    .setTitle(boundedText(
      `${titlePrefix}${payload?.event?.eventName || "Not specified"}`,
      EMBED_TITLE_LIMIT
    ))
    .setDescription(descriptionParts.join("\n").slice(0, EMBED_DESCRIPTION_LIMIT))

  if (imageFilename) embed.setImage(`attachment://${imageFilename}`)

  return {
    embeds: [embed],
    allowedMentions: { parse: [], repliedUser: false }
  }
}

function formatEventDelivery(payload, options) {
  if (payload?.claim?.targetKind === "alliance") {
    return formatAllianceEventDelivery(payload, options)
  }
  throw new PermanentDeliveryError("Individual state reminders are disabled.")
}

module.exports = {
  EMBED_TITLE_LIMIT,
  EMBED_DESCRIPTION_LIMIT,
  neutralizeMentions,
  boundedText,
  timingDetails,
  formatAllianceEventDelivery,
  formatEventDelivery
}
