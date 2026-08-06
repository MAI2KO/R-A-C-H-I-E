const { EmbedBuilder } = require("discord.js")

const { recurrenceLabel } = require("./eventSchedulerFormatting")
const { PermanentDeliveryError } = require("./eventDeliveryWorker")

const EMBED_TITLE_LIMIT = 256
const EMBED_DESCRIPTION_LIMIT = 4096
const EMBED_FIELD_NAME_LIMIT = 256
const EMBED_FIELD_VALUE_LIMIT = 1024

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
  const timestamp = Math.floor(occurrence.getTime() / 1000)
  const utc = occurrence.toISOString().slice(0, 16).replace("T", " ")
  if (payload.claim.deliveryKind === "final_reminder") {
    const deliverAt = payload.claim.deliverAt
    if (
      !(deliverAt instanceof Date)
      || !Number.isFinite(deliverAt.getTime())
      || occurrence.getTime() - deliverAt.getTime() !== 60000
    ) {
      throw new PermanentDeliveryError("Final reminder time is invalid.")
    }
    return {
      status: "About to start\nStarts in approximately 1 minute",
      timestamp,
      utc
    }
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
  if (!Number.isInteger(minutes) || ![10, 30].includes(minutes)) {
    throw new PermanentDeliveryError("Advance reminder interval is invalid.")
  }
  return { status: `Starts in ${minutes} minutes`, timestamp, utc }
}

function formatAllianceEventDelivery(payload, { imageFilename = null } = {}) {
  const timing = timingDetails(payload)
  const eventName = boundedText(payload?.event?.eventName, 220)
  const allianceName = boundedText(payload?.alliance?.name, 900)
  const groupName = payload?.group?.name
    ? boundedText(payload.group.name, 900)
    : null
  const descriptionParts = [`Alliance: **${allianceName}**`]
  if (groupName) descriptionParts.push(`Group: **${groupName}**`)

  const titlePrefix = payload.claim.deliveryKind === "final_reminder"
    ? "About to start: "
    : "Event reminder: "
  const embed = new EmbedBuilder()
    .setColor(payload.claim.deliveryKind === "final_reminder" ? 0x2ecc71 : 0xf1c40f)
    .setTitle(boundedText(`${titlePrefix}${eventName}`, EMBED_TITLE_LIMIT))
    .setDescription(boundedText(
      descriptionParts.join("\n"),
      EMBED_DESCRIPTION_LIMIT
    ))
    .addFields(
      {
        name: boundedText("When", EMBED_FIELD_NAME_LIMIT),
        value: boundedText(
          `${timing.utc} UTC\n<t:${timing.timestamp}:F>`,
          EMBED_FIELD_VALUE_LIMIT
        ),
        inline: false
      },
      {
        name: boundedText("Status", EMBED_FIELD_NAME_LIMIT),
        value: boundedText(timing.status, EMBED_FIELD_VALUE_LIMIT),
        inline: true
      },
      {
        name: boundedText("Recurrence", EMBED_FIELD_NAME_LIMIT),
        value: boundedText(
          recurrenceLabel(payload?.event?.recurrenceDays),
          EMBED_FIELD_VALUE_LIMIT
        ),
        inline: true
      }
    )

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
  EMBED_FIELD_NAME_LIMIT,
  EMBED_FIELD_VALUE_LIMIT,
  neutralizeMentions,
  boundedText,
  timingDetails,
  formatAllianceEventDelivery,
  formatEventDelivery
}
