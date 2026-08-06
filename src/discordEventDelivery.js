const {
  ChannelType,
  PermissionFlagsBits
} = require("discord.js")

const {
  MAX_IMAGE_BYTES,
  SUPPORTED_IMAGE_TYPES,
  normalizeContentType,
  validateImageSignature
} = require("./eventImage")
const {
  PermanentDeliveryError,
  RetryableDeliveryError,
  sanitizeDeliveryError
} = require("./eventDeliveryWorker")
const { formatAllianceEventDelivery } = require("./eventDeliveryFormatting")

const IMAGE_EXTENSIONS = Object.freeze({
  "image/png": "png",
  "image/jpeg": "jpg",
  "image/gif": "gif",
  "image/webp": "webp"
})
const PERMANENT_DISCORD_CODES = new Set([10003, 10004, 50001, 50013, 50035])
const RETRYABLE_NETWORK_CODES = new Set([
  "ECONNRESET",
  "ECONNREFUSED",
  "EHOSTUNREACH",
  "ENETUNREACH",
  "ETIMEDOUT",
  "EAI_AGAIN",
  "UND_ERR_CONNECT_TIMEOUT",
  "UND_ERR_SOCKET"
])

class DiscordPermanentDeliveryError extends PermanentDeliveryError {
  constructor(message) {
    super(message)
    this.name = "DiscordPermanentDeliveryError"
  }
}

class DiscordRetryableDeliveryError extends RetryableDeliveryError {
  constructor(message) {
    super(message)
    this.name = "DiscordRetryableDeliveryError"
  }
}

function permanent(message) {
  return new DiscordPermanentDeliveryError(message)
}

function retryable(message) {
  return new DiscordRetryableDeliveryError(message)
}

function normalizeDiscordDeliveryError(error) {
  if (error instanceof PermanentDeliveryError) {
    return permanent(sanitizeDeliveryError(error))
  }
  if (error instanceof RetryableDeliveryError) {
    return retryable(sanitizeDeliveryError(error))
  }
  const code = typeof error?.code === "string" && /^\d+$/.test(error.code)
    ? Number(error.code)
    : error?.code
  const status = Number(error?.status)
  if (PERMANENT_DISCORD_CODES.has(code)) {
    return permanent("Discord rejected the configured target or message.")
  }
  if (RETRYABLE_NETWORK_CODES.has(code) || error?.name === "AbortError") {
    return retryable("Discord could not be reached temporarily.")
  }
  if (error?.name === "RateLimitError" || status === 429 || status >= 500) {
    return retryable("Discord temporarily rejected the delivery request.")
  }
  if (status >= 400 && status < 500) {
    return permanent("Discord permanently rejected the delivery request.")
  }
  return retryable("Discord delivery failed temporarily.")
}

function prepareStoredEventImage(image) {
  if (!image) return null
  const contentType = normalizeContentType(image.contentType)
  if (!SUPPORTED_IMAGE_TYPES.has(contentType)) {
    throw permanent("Stored event image type is unsupported.")
  }
  if (!Buffer.isBuffer(image.imageData)) {
    throw permanent("Stored event image data is invalid.")
  }
  if (
    !Number.isInteger(image.byteSize)
    || image.byteSize <= 0
    || image.byteSize > MAX_IMAGE_BYTES
    || image.imageData.length !== image.byteSize
  ) {
    throw permanent("Stored event image size is invalid.")
  }
  if (!validateImageSignature(contentType, image.imageData)) {
    throw permanent("Stored event image content is invalid.")
  }
  const filename = `event-image.${IMAGE_EXTENSIONS[contentType]}`
  return Object.freeze({
    filename,
    file: Object.freeze({
      attachment: Buffer.from(image.imageData),
      name: filename
    })
  })
}

async function resolveAllianceTarget(client, payload, { hasImage }) {
  if (payload?.claim?.targetKind !== "alliance") {
    throw permanent("Delivery target is not an alliance channel.")
  }
  const guildId = String(payload.claim.targetGuildId || "")
  const channelId = String(payload.claim.targetChannelId || "")
  if (!guildId || payload?.alliance?.guildId !== guildId || payload?.event?.guildId !== guildId) {
    throw permanent("Delivery guild ownership is invalid.")
  }

  let guild = client.guilds?.cache?.get?.(guildId)
  if (!guild) guild = await client.guilds?.fetch?.(guildId)
  if (!guild) throw permanent("Configured alliance guild is unavailable.")

  const botMember = guild.members?.me || await guild.members?.fetchMe?.()
  if (!botMember) throw permanent("Bot is not a member of the configured alliance guild.")

  let channel = guild.channels?.cache?.get?.(channelId)
  if (!channel) channel = await guild.channels?.fetch?.(channelId)
  if (!channel) throw permanent("Configured alliance event channel is unavailable.")
  if (channel.guildId !== guildId) {
    throw permanent("Configured event channel belongs to another guild.")
  }
  const supportedType = [ChannelType.GuildText, ChannelType.GuildAnnouncement]
    .includes(channel.type)
  if (
    !supportedType
    || !channel.isTextBased?.()
    || !channel.isSendable?.()
    || typeof channel.send !== "function"
  ) {
    throw permanent("Configured event channel is not a sendable guild text channel.")
  }

  const permissions = channel.permissionsFor?.(botMember)
  const required = [
    [PermissionFlagsBits.ViewChannel, "View Channel"],
    [PermissionFlagsBits.SendMessages, "Send Messages"],
    [PermissionFlagsBits.EmbedLinks, "Embed Links"]
  ]
  if (hasImage) required.push([PermissionFlagsBits.AttachFiles, "Attach Files"])
  for (const [permission, label] of required) {
    if (!permissions?.has?.(permission)) {
      throw permanent(`Bot is missing ${label} permission in the event channel.`)
    }
  }
  return channel
}

function createDiscordEventDeliveryHandler({ client, gameProfile }) {
  if (!client) throw new Error("Discord client is required")
  const expectedProfile = String(gameProfile || "").trim()
  if (!expectedProfile) throw new Error("Game profile is required")

  return async function deliverEventClaim(payload) {
    try {
      if (!client.isReady?.()) throw retryable("Discord client is not ready.")
      if (payload?.claim?.gameProfile !== expectedProfile) {
        throw permanent("Delivery game profile does not match this bot.")
      }
      const image = prepareStoredEventImage(payload.image)
      const channel = await resolveAllianceTarget(client, payload, { hasImage: Boolean(image) })
      const message = formatAllianceEventDelivery(payload, {
        imageFilename: image?.filename || null
      })
      if (image) message.files = [image.file]
      const sent = await channel.send(message)
      if (!String(sent?.id || "").trim()) {
        throw retryable("Discord did not return a message identifier.")
      }
      return Object.freeze({ sentMessageId: String(sent.id) })
    } catch (error) {
      throw normalizeDiscordDeliveryError(error)
    }
  }
}

module.exports = {
  IMAGE_EXTENSIONS,
  PERMANENT_DISCORD_CODES,
  RETRYABLE_NETWORK_CODES,
  DiscordPermanentDeliveryError,
  DiscordRetryableDeliveryError,
  normalizeDiscordDeliveryError,
  prepareStoredEventImage,
  resolveAllianceTarget,
  createDiscordEventDeliveryHandler
}
