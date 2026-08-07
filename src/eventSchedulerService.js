const { ChannelType, PermissionFlagsBits } = require("discord.js")

class SchedulerValidationError extends Error {
  constructor(message) {
    super(message)
    this.name = "SchedulerValidationError"
  }
}

function normalizeSnowflake(value, label) {
  const normalized = String(value || "").trim()
  if (!/^\d{16,22}$/.test(normalized)) {
    throw new SchedulerValidationError(`${label} must be a valid Discord ID.`)
  }
  return normalized
}

function normalizeAllianceName(value) {
  const normalized = String(value || "").trim()
  if (!normalized || normalized.length > 100) {
    throw new SchedulerValidationError("Alliance name must be 1 to 100 characters.")
  }
  return normalized
}

async function resolveSendableChannel(
  client,
  guildIdInput,
  channelIdInput,
  { requireAttachments = true } = {}
) {
  const guildId = normalizeSnowflake(guildIdInput, "Guild ID")
  const channelId = normalizeSnowflake(channelIdInput, "Channel ID")

  let guild
  let channel
  try {
    guild = await client.guilds.fetch(guildId)
    channel = await guild.channels.fetch(channelId)
  } catch {
    throw new SchedulerValidationError(
      "The bot cannot access that Discord guild or channel."
    )
  }

  if (
    !channel
    || channel.guildId !== guildId
    || ![ChannelType.GuildText, ChannelType.GuildAnnouncement].includes(channel.type)
    || !channel.isTextBased?.()
    || !channel.isSendable?.()
  ) {
    throw new SchedulerValidationError("Select an accessible text channel in that guild.")
  }

  const botMember = guild.members.me || await guild.members.fetchMe()
  const permissions = channel.permissionsFor?.(botMember)
  const required = [
    [PermissionFlagsBits.ViewChannel, "View Channel"],
    [PermissionFlagsBits.SendMessages, "Send Messages"],
    [PermissionFlagsBits.EmbedLinks, "Embed Links"]
  ]
  if (requireAttachments) {
    required.push([PermissionFlagsBits.AttachFiles, "Attach Files"])
  }

  const missing = required
    .filter(([permission]) => !permissions?.has(permission))
    .map(([, label]) => label)
  if (missing.length) {
    throw new SchedulerValidationError(
      `The bot is missing ${missing.join(", ")} in that channel.`
    )
  }

  return { guild, channel, guildId, channelId }
}

module.exports = {
  SchedulerValidationError,
  normalizeSnowflake,
  normalizeAllianceName,
  resolveSendableChannel
}
