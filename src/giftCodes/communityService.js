const { randomUUID } = require("node:crypto")
const { ChannelType, PermissionFlagsBits } = require("discord.js")
const { profileTerminology } = require("./terminology")

const PUBLIC_MENTIONS = userId => ({ parse: [], users: userId ? [userId] : [] })
const TEXT_CHANNEL_TYPES = new Set([ChannelType.GuildText, ChannelType.GuildAnnouncement])
const CONTRIBUTOR_ROLE_NAME = "🍭"
const CONTRIBUTOR_RETRY_MS = 5 * 60 * 1000
const CONTRIBUTOR_CONFIRMATION = "Nice find. Have a 🍭"
const CONTRIBUTOR_FALLBACK = "Nice find. I owe you a 🍭."

class ContributorRoleError extends Error {
  constructor(code, message) {
    super(message)
    this.name = "ContributorRoleError"
    this.code = code
  }
}

function safeDiscordError(error) {
  return String(error?.code || error?.name || "discord_delivery_failed").slice(0, 100)
}

function codeProgressMessage(payload, progress, terms, { initial = false } = {}) {
  const heading = initial ? "**New Gift Code Verified**" : `**Gift Code: ${payload.code}**`
  return [
    heading,
    initial ? `\n${payload.code}` : "",
    `Submitted by: <@${payload.discord_user_id}>`,
    "Status: Active",
    initial ? `Accounts queued: ${payload.metadata?.queuedCount ?? progress.total}` : "",
    "",
    `Redeemed: ${progress.successful}`,
    `Already claimed: ${progress.already_redeemed}`,
    `Account issue: ${progress.account_issues}`,
    `Restricted: ${progress.restricted}`,
    `Remaining: ${progress.remaining}`,
    `Game: ${terms.gameName}`
  ].filter(Boolean).join("\n")
}

function joinMessage(payload, stats, accountStats, terms, maximumEnabled) {
  return [
    "**Gift Code Auto-Redeem Activated**",
    "",
    `<@${payload.discord_user_id}> is now set up for automatic gift-code redemption.`,
    "",
    `Game: ${terms.gameName}`,
    `${terms.locationLabel}: ${payload.state_or_kingdom_number}`,
    `Characters covered: ${accountStats.enabledCount} / ${maximumEnabled}`,
    `Their successful redemptions: ${accountStats.successfulRedemptions}`,
    "",
    "**Community**",
    `Players using Auto-Redeem: ${stats.auto_redeem_players}`,
    `Characters covered: ${stats.enabled_accounts}`,
    `Successful redemptions: ${stats.successful_redemptions}`
  ].join("\n")
}

function eventNonce(eventId) {
  return BigInt(`0x${eventId.replaceAll("-", "").slice(0, 16)}`).toString()
}

function createGiftCodeCommunityService({
  repository,
  client,
  gameProfile,
  maximumEnabledAccounts = 2,
  logger = console,
  now = () => new Date(),
  workerId = `community-${process.pid}-${randomUUID()}`
}) {
  const terms = profileTerminology(gameProfile)

  async function resolveGuild(guildId) {
    return client.guilds.fetch(guildId)
  }

  async function resolveChannel(guild, channelId) {
    const channel = await guild.channels.fetch(channelId)
    if (!channel || !TEXT_CHANNEL_TYPES.has(channel.type) || !channel.isSendable?.()) {
      throw new Error("configured gift-code channel is unavailable")
    }
    const permissions = channel.permissionsFor?.(guild.members.me)
    if (permissions && !permissions.has(PermissionFlagsBits.ViewChannel | PermissionFlagsBits.SendMessages)) {
      throw new Error("configured gift-code channel is not sendable")
    }
    return channel
  }

  function botCanManageRoles(guild) {
    return Boolean(guild.members.me?.permissions?.has?.(PermissionFlagsBits.ManageRoles))
  }

  function roleHasNoPermissions(role) {
    if (role?.permissions?.bitfield !== undefined) return role.permissions.bitfield === 0n
    if (typeof role?.permissions?.equals === "function") return role.permissions.equals(0n)
    return false
  }

  function reusableContributorRole(role) {
    return Boolean(
      role
      && role.name === CONTRIBUTOR_ROLE_NAME
      && role.managed !== true
      && role.mentionable === false
      && roleHasNoPermissions(role)
      && role.editable !== false
    )
  }

  async function fetchPersistedContributorRole(guild, roleId) {
    if (!roleId) return null
    return guild.roles.fetch(roleId).catch(() => null)
  }

  async function getOrCreateContributorRole(guild, payload) {
    const settings = await repository.getSettings(payload.guild_id)
    if (settings?.contributor_role_status === "error"
      && settings.contributor_role_claimed_until_utc
      && new Date(settings.contributor_role_claimed_until_utc) > now()) {
      throw new ContributorRoleError(
        "CONTRIBUTOR_ROLE_RETRY_PENDING",
        "Contributor reward role setup is waiting for its next retry."
      )
    }
    const persisted = await fetchPersistedContributorRole(guild, settings?.contributor_role_id)
    if (reusableContributorRole(persisted)) return persisted

    const claim = await repository.claimContributorRoleProvision(
      payload.guild_id,
      workerId,
      now()
    )
    if (!claim) {
      throw new ContributorRoleError(
        "CONTRIBUTOR_ROLE_PROVISION_BUSY",
        "Contributor reward role provisioning is already in progress."
      )
    }

    try {
      if (!botCanManageRoles(guild)) {
        throw new ContributorRoleError(
          "CONTRIBUTOR_ROLE_MANAGE_PERMISSION",
          "The bot needs Manage Roles to create or assign the contributor reward role."
        )
      }
      const roles = await guild.roles.fetch()
      const reusable = [...roles.values()].find(reusableContributorRole)
      const role = reusable || await guild.roles.create({
        name: CONTRIBUTOR_ROLE_NAME,
        permissions: [],
        mentionable: false,
        reason: "Gift-code contributor reward"
      })
      const completed = await repository.completeContributorRoleProvision(
        payload.guild_id,
        workerId,
        role.id,
        now()
      )
      if (!completed) {
        throw new ContributorRoleError(
          "CONTRIBUTOR_ROLE_PROVISION_LOST",
          "Contributor reward role provisioning could not be finalized."
        )
      }
      return role
    } catch (error) {
      await repository.failContributorRoleProvision(
        payload.guild_id,
        workerId,
        safeDiscordError(error),
        now()
      ).catch(() => {})
      throw error
    }
  }

  async function assignContributorRole(payload) {
    const guild = await resolveGuild(payload.guild_id)
    const role = await getOrCreateContributorRole(guild, payload)
    if (!botCanManageRoles(guild)) {
      throw new ContributorRoleError(
        "CONTRIBUTOR_ROLE_MANAGE_PERMISSION",
        "The bot needs Manage Roles to assign the contributor reward role."
      )
    }
    if (!role.editable || role.managed) {
      throw new ContributorRoleError(
        "CONTRIBUTOR_ROLE_HIERARCHY",
        "The contributor reward role must be below the bot's highest role."
      )
    }
    const member = await guild.members.fetch(payload.discord_user_id)
    if (!member.roles.cache.has(role.id)) await member.roles.add(role)
  }

  async function sendContributorConfirmation(payload, eventId, awarded) {
    try {
      const user = await client.users.fetch(payload.discord_user_id)
      await user.send({
        content: awarded ? CONTRIBUTOR_CONFIRMATION : CONTRIBUTOR_FALLBACK,
        nonce: eventNonce(eventId),
        enforceNonce: true
      })
      return true
    } catch (error) {
      logger.warn(`[Gift codes] Contributor confirmation failed: ${safeDiscordError(error)}`)
      return false
    }
  }

  async function deliverClaimedEvent(event) {
    const payload = await repository.getEventPayload(event.id)
    if (!payload) throw new Error("engagement payload unavailable")
    if (payload.event_type === "contributor_role") {
      try {
        await assignContributorRole(payload)
      } catch (error) {
        await sendContributorConfirmation(payload, event.id, false)
        throw error
      }
      await sendContributorConfirmation(payload, event.id, true)
      return repository.completeEvent(event.id, workerId, { now: now() })
    }
    const guild = await resolveGuild(payload.guild_id)
    const channel = await resolveChannel(guild, payload.gift_code_channel_id)
    if (payload.event_type === "auto_redeem_join") {
      const [stats, accountRows] = await Promise.all([
        repository.communityStats(payload.guild_id),
        repository.accountOwnerStats(payload.discord_user_id)
      ])
      const message = await channel.send({
        content: joinMessage(payload, stats, accountRows, terms, maximumEnabledAccounts),
        allowedMentions: PUBLIC_MENTIONS(payload.discord_user_id),
        nonce: eventNonce(event.id),
        enforceNonce: true
      })
      return repository.completeEvent(event.id, workerId, {
        channelId: channel.id,
        messageId: message.id,
        finalized: true,
        now: now()
      })
    }
    const progress = await repository.codeProgress(payload.gift_code_id)
    const message = await channel.send({
      content: codeProgressMessage(payload, progress, terms, { initial: true }),
      allowedMentions: PUBLIC_MENTIONS(payload.discord_user_id),
      nonce: eventNonce(event.id),
      enforceNonce: true
    })
    const completed = progress.successful + progress.already_redeemed +
      progress.account_issues + progress.restricted
    return repository.completeEvent(event.id, workerId, {
      channelId: channel.id,
      messageId: message.id,
      progressCount: completed,
      progress,
      finalized: progress.remaining === 0,
      now: now()
    })
  }

  async function deliverEvent(event) {
    if (!event) return null
    const claim = await repository.claimEvent(event.id, workerId, now())
    if (!claim) return null
    try {
      return await deliverClaimedEvent(claim)
    } catch (error) {
      const errorCode = safeDiscordError(error)
      const failedAt = now()
      if (claim.event_type === "contributor_role") {
        await repository.markContributorRoleUnavailable(
          claim.guild_id,
          errorCode,
          failedAt
        ).catch(() => {})
      }
      const retryAt = claim.event_type === "contributor_role"
        ? new Date(failedAt.getTime() + CONTRIBUTOR_RETRY_MS)
        : null
      await repository.failEvent(
        claim.id,
        workerId,
        errorCode,
        failedAt,
        { retryAt }
      ).catch(() => {})
      logger.warn(`[Gift codes] Community event failed: ${errorCode}`)
      return null
    }
  }

  return Object.freeze({
    terms,

    async configureChannel(guildId, channelId) {
      const guild = await resolveGuild(guildId)
      const channel = await resolveChannel(guild, channelId)
      return repository.setChannel(guildId, channel.id)
    },

    async configuration(guildId) {
      const settings = await repository.getSettings(guildId)
      if (!settings) return { settings: null, channelAvailable: false, roleAvailable: false }
      const guild = await resolveGuild(guildId).catch(() => null)
      const channelAvailable = guild && settings.gift_code_channel_id
        ? Boolean(await resolveChannel(guild, settings.gift_code_channel_id).catch(() => null))
        : false
      const role = guild && settings.contributor_role_id
        ? await guild.roles.fetch(settings.contributor_role_id).catch(() => null)
        : null
      const roleAvailable = Boolean(
        guild && botCanManageRoles(guild) && reusableContributorRole(role)
      )
      return {
        settings,
        channelAvailable,
        roleAvailable,
        roleStatus: roleAvailable ? "ready" : settings.contributor_role_status || "unconfigured"
      }
    },

    async onAutoRedemptionEnabled(event) {
      return deliverEvent(event)
    },

    async onCodeActivated(giftCodeId, queuedCount) {
      const events = await repository.prepareCodeEngagement(giftCodeId, queuedCount)
      for (const event of events) await deliverEvent(event)
      return events.length
    },

    async onRedemptionUpdated(giftCodeId, resultStatus) {
      const refresh = await repository.claimProgressRefresh(giftCodeId, resultStatus, workerId, now())
      if (!refresh) return false
      try {
        const payload = await repository.getEventPayload(refresh.event.id)
        const guild = await resolveGuild(payload.guild_id)
        const channel = await resolveChannel(guild, refresh.event.channel_id || payload.gift_code_channel_id)
        const message = await channel.messages.fetch(refresh.event.message_id)
        await message.edit({
          content: codeProgressMessage(payload, refresh.progress, terms),
          allowedMentions: PUBLIC_MENTIONS(payload.discord_user_id)
        })
        await repository.completeEvent(refresh.event.id, workerId, {
          progressCount: refresh.completed,
          progress: refresh.progress,
          finalized: refresh.progress.remaining === 0,
          now: now()
        })
        return true
      } catch (error) {
        const errorCode = safeDiscordError(error)
        await repository.failEvent(refresh.event.id, workerId, errorCode, now()).catch(() => {})
        logger.warn(`[Gift codes] Progress update failed: ${errorCode}`)
        return false
      }
    },

    async recoverOne() {
      const event = await repository.claimNextPending(workerId, now())
      if (!event) return 0
      try {
        await deliverClaimedEvent(event)
      } catch (error) {
        const errorCode = safeDiscordError(error)
        const failedAt = now()
        if (event.event_type === "contributor_role") {
          await repository.markContributorRoleUnavailable(
            event.guild_id,
            errorCode,
            failedAt
          ).catch(() => {})
        }
        const retryAt = event.event_type === "contributor_role"
          ? new Date(failedAt.getTime() + CONTRIBUTOR_RETRY_MS)
          : null
        await repository.failEvent(
          event.id,
          workerId,
          errorCode,
          failedAt,
          { retryAt }
        ).catch(() => {})
        logger.warn(`[Gift codes] Community recovery failed: ${errorCode}`)
      }
      return 1
    },

    communityStats: guildId => repository.communityStats(guildId)
  })
}

module.exports = {
  PUBLIC_MENTIONS,
  CONTRIBUTOR_ROLE_NAME,
  CONTRIBUTOR_CONFIRMATION,
  CONTRIBUTOR_FALLBACK,
  ContributorRoleError,
  safeDiscordError,
  eventNonce,
  codeProgressMessage,
  joinMessage,
  createGiftCodeCommunityService
}
