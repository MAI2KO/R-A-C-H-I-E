const { randomUUID } = require("node:crypto")
const { ChannelType, PermissionFlagsBits } = require("discord.js")
const { profileTerminology } = require("./terminology")

const PUBLIC_MENTIONS = userId => ({ parse: [], users: userId ? [userId] : [] })
const TEXT_CHANNEL_TYPES = new Set([ChannelType.GuildText, ChannelType.GuildAnnouncement])
const CONTRIBUTOR_ROLE_NAME = "🍭"
const CONTRIBUTOR_RETRY_MS = 5 * 60 * 1000
const CONTRIBUTOR_CONFIRMATION = "Nice find. Have a 🍭"
const CONTRIBUTOR_FALLBACK = "Nice find. I owe you a 🍭."
const VERIFICATION_RESULT_MESSAGES = Object.freeze({
  invalid_code: "That code doesn't look valid.",
  expired: "That code has expired.",
  eligibility_restriction: "I couldn't confirm that code, so I haven't added it.",
  redemption_limit: "I couldn't confirm that code, so I haven't added it.",
  claim_limit: "I couldn't confirm that code, so I haven't added it.",
  level_restriction: "I couldn't confirm that code, so I haven't added it.",
  account_restriction: "I couldn't confirm that code, so I haven't added it.",
  account_age_restriction: "I couldn't confirm that code, so I haven't added it.",
  verification_error: "I couldn't verify that code, so I've left it for review.",
  upstream_rejection: "I couldn't verify that code right now. I'll leave it for review.",
  temporary_error: "I couldn't verify that code right now. I'll leave it for review.",
  rate_limited: "I couldn't verify that code right now. I'll leave it for review.",
  verification_throttle: "I couldn't verify that code right now. I'll leave it for review.",
  simultaneous_action_throttle: "I couldn't verify that code right now. I'll leave it for review.",
  invalid_player: "I couldn't verify that code right now. I'll leave it for review.",
  unknown_response: "I couldn't verify that code, so I've left it for review."
})

class ContributorRoleError extends Error {
  constructor(code, message) {
    super(message)
    this.name = "ContributorRoleError"
    this.code = code
  }
}

class GiftCodeCommunityError extends Error {
  constructor(code, message) {
    super(message)
    this.name = "GiftCodeCommunityError"
    this.code = code
  }
}

function safeDiscordError(error) {
  return String(error?.code || error?.name || "discord_delivery_failed").slice(0, 100)
}

function staleDiscordReference(error) {
  return new Set(["10003", "10008", "50001", "50013"]).has(String(error?.code || ""))
    || error?.code === "GIFT_CODE_CHANNEL_UNAVAILABLE"
    || error?.code === "GIFT_CODE_CHANNEL_NOT_SENDABLE"
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
    `Successful redemptions: ${progress.successful}`,
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
    `Their successful redemptions: ${accountStats.successfulRedemptions ?? 0}`,
    `Their already claimed: ${accountStats.alreadyRedeemed ?? 0}`,
    "",
    "**Community**",
    `Players using Auto-Redeem: ${stats.auto_redeem_players}`,
    `Characters covered: ${stats.enabled_accounts}`,
    `Successful redemptions: ${stats.successful_redemptions}`,
    `Already claimed: ${stats.already_redeemed ?? 0}`
  ].join("\n")
}

function eventNonce(eventId) {
  return BigInt(`0x${eventId.replaceAll("-", "").slice(0, 16)}`).toString()
}

function statusCardNonce(event, generationOffset = 0) {
  const generation = Number(event.metadata?.statusCardGeneration || 0) + generationOffset
  return (BigInt(eventNonce(event.id)) + BigInt(generation)).toString()
}

function verificationResultMessage(classification) {
  return VERIFICATION_RESULT_MESSAGES[classification]
    || VERIFICATION_RESULT_MESSAGES.unknown_response
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
      const error = new GiftCodeCommunityError(
        "GIFT_CODE_CHANNEL_UNAVAILABLE",
        "That channel is unavailable to the bot. Choose a visible text channel."
      )
      error.giftCodeHandler = "configure_channel_resolve"
      throw error
    }
    const permissions = channel.permissionsFor?.(guild.members.me)
    if (permissions && !permissions.has(PermissionFlagsBits.ViewChannel | PermissionFlagsBits.SendMessages)) {
      const error = new GiftCodeCommunityError(
        "GIFT_CODE_CHANNEL_NOT_SENDABLE",
        "The bot needs View Channel and Send Messages in that channel."
      )
      error.giftCodeHandler = "configure_channel_permissions"
      throw error
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
    if (payload.event_type === "auto_redeem_join") {
      return deliverStatusCard(event, payload, guild)
    }
    const channel = await resolveChannel(guild, payload.gift_code_channel_id)
    const progress = await repository.codeProgress(payload.gift_code_id, payload.guild_id)
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

  async function statusCardContent(payload) {
    const [stats, accountStats] = await Promise.all([
      repository.communityStats(payload.guild_id),
      repository.accountOwnerStats(payload.discord_user_id, payload.guild_id)
    ])
    return joinMessage({
      ...payload,
      state_or_kingdom_number: accountStats.locationNumber
        || payload.state_or_kingdom_number
    }, stats, accountStats, terms, maximumEnabledAccounts)
  }

  async function deliverStatusCard(event, payload, resolvedGuild = null) {
    const guild = resolvedGuild || await resolveGuild(payload.guild_id)
    const content = await statusCardContent(payload)
    const allowedMentions = PUBLIC_MENTIONS(payload.discord_user_id)
    let staleError = null

    if (event.message_id && event.channel_id) {
      try {
        const channel = await resolveChannel(guild, event.channel_id)
        const message = await channel.messages.fetch(event.message_id)
        await message.edit({ content, allowedMentions })
        return repository.completeEvent(event.id, workerId, {
          channelId: channel.id,
          messageId: message.id,
          finalized: true,
          now: now()
        })
      } catch (error) {
        if (!staleDiscordReference(error)) throw error
        staleError = error
      }
    }

    try {
      const channel = await resolveChannel(guild, payload.gift_code_channel_id)
      const sendOptions = {
        content,
        allowedMentions,
        nonce: statusCardNonce(event, staleError ? 1 : 0),
        enforceNonce: true
      }
      const message = await channel.send(sendOptions)
      return repository.completeEvent(event.id, workerId, {
        channelId: channel.id,
        messageId: message.id,
        finalized: true,
        now: now()
      })
    } catch (error) {
      const finalError = staleError || error
      if (staleDiscordReference(finalError) || staleDiscordReference(error)) {
        await repository.clearStatusCardReference(
          event.id,
          workerId,
          safeDiscordError(finalError),
          now()
        ).catch(() => {})
      }
      throw error
    }
  }

  async function refreshStatusCard(guildId, discordUserId) {
    let event
    try {
      event = await repository.claimStatusCard(
        guildId,
        discordUserId,
        workerId,
        now()
      )
      if (!event) return false
      const payload = await repository.getEventPayload(event.id)
      if (!payload) throw new Error("status card payload unavailable")
      await deliverStatusCard(event, payload)
      return true
    } catch (error) {
      const errorCode = safeDiscordError(error)
      if (event) {
        await repository.failEvent(event.id, workerId, errorCode, now()).catch(() => {})
      }
      logger.warn(`[Gift codes] Status card refresh failed: ${errorCode}`)
      return false
    }
  }

  async function refreshStatusCardsForAccount(playerAccountId) {
    const targets = await repository.statusCardTargetsForAccount(playerAccountId)
    for (const target of targets) {
      await refreshStatusCard(target.guild_id, target.discord_user_id)
    }
    return targets.length
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

  async function deliverVerificationResult(claim) {
    if (!claim) return false
    try {
      const user = await client.users.fetch(claim.submitted_by_discord_user_id)
      await user.send({
        content: verificationResultMessage(claim.classification),
        nonce: eventNonce(claim.id),
        enforceNonce: true
      })
      await repository.finishVerificationResultNotification(claim.id, workerId, {
        sent: true,
        now: now()
      })
      return true
    } catch (error) {
      const errorCode = safeDiscordError(error)
      await repository.finishVerificationResultNotification(claim.id, workerId, {
        sent: false,
        now: now(),
        errorCode
      }).catch(() => {})
      logger.warn(`[Gift codes] Verification result notification failed: ${errorCode}`)
      return false
    }
  }

  return Object.freeze({
    terms,

    async configureChannel(guildId, channelId) {
      let guild
      try {
        guild = await resolveGuild(guildId)
      } catch (error) {
        error.giftCodeHandler = "configure_channel_guild"
        throw error
      }
      let channel
      try {
        channel = await resolveChannel(guild, channelId)
      } catch (error) {
        error.giftCodeHandler ||= "configure_channel_resolve"
        throw error
      }
      try {
        return await repository.setChannel(guildId, channel.id)
      } catch (error) {
        error.giftCodeHandler = "configure_channel_persist"
        throw error
      }
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

    async onAutoRedemptionEnabled(event, { guildId = null, discordUserId = null } = {}) {
      if (event) return deliverEvent(event)
      if (!guildId || !discordUserId) return null
      return refreshStatusCard(guildId, discordUserId)
    },

    refreshStatusCard,

    async onCodeActivated(giftCodeId, queuedCount) {
      const events = await repository.prepareCodeEngagement(giftCodeId, queuedCount)
      for (const event of events) await deliverEvent(event)
      return events.length
    },

    async onVerificationResult(giftCodeId) {
      const claim = await repository.claimVerificationResultNotification(
        workerId,
        now(),
        60,
        giftCodeId
      )
      return deliverVerificationResult(claim)
    },

    async onRedemptionUpdated(giftCodeId, playerAccountId, resultStatus) {
      if (["success", "already_redeemed"].includes(resultStatus)) {
        try {
          await refreshStatusCardsForAccount(playerAccountId)
        } catch (error) {
          logger.warn(`[Gift codes] Status card target lookup failed: ${safeDiscordError(error)}`)
        }
      }
      const refresh = await repository.claimProgressRefresh(
        giftCodeId,
        playerAccountId,
        resultStatus,
        workerId,
        now()
      )
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
      const notification = await repository.claimVerificationResultNotification(workerId, now())
      if (notification) {
        await deliverVerificationResult(notification)
        return 1
      }
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
  VERIFICATION_RESULT_MESSAGES,
  ContributorRoleError,
  GiftCodeCommunityError,
  safeDiscordError,
  staleDiscordReference,
  eventNonce,
  verificationResultMessage,
  codeProgressMessage,
  joinMessage,
  createGiftCodeCommunityService
}
