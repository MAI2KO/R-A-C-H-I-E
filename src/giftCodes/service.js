const crypto = require("node:crypto")
const {
  normalizeDiscordUserId,
  normalizeGuildId,
  normalizeGiftCode,
  normalizePlayerId
} = require("./validation")
const { profileTerminology } = require("./terminology")
const { giftAccountConfig, giftWorkerConfig } = require("./config")

class GiftCodeError extends Error {
  constructor(code, message) {
    super(message)
    this.name = "GiftCodeError"
    this.code = code
  }
}

const SUBMISSION_LABELS = Object.freeze({
  active: "already known active",
  expired: "already known expired",
  invalid: "already known invalid",
  verifying: "currently being verified",
  restricted: "previously restricted",
  unknown: "previously unknown and awaiting review",
  candidate: "currently awaiting verification",
  disabled: "known but disabled"
})

function createGiftCodeService({ repository, gameProfile, env = process.env, ingestion = null }) {
  const terms = profileTerminology(gameProfile)
  const accountConfig = giftAccountConfig(env)
  const workerConfig = giftWorkerConfig(env)

  async function selectOwnedAccount(discordUserId, playerId = null, guildId = null) {
    const owner = normalizeDiscordUserId(discordUserId)
    const selectedPlayer = playerId ? normalizePlayerId(playerId, terms.playerLabel) : null
    const accounts = await repository.activeAccountStatuses(owner, selectedPlayer, guildId)
    if (!accounts.length) {
      throw new GiftCodeError("PLAYER_NOT_FOUND", `No matching ${terms.playerLabel} was found.`)
    }
    if (selectedPlayer) return accounts[0]
    return accounts.find(account => account.is_active && account.is_primary)
      || accounts.find(account => account.is_active)
      || accounts[0]
  }

  return Object.freeze({
    terms,

    async submit({ discordUserId, guildId = null, code, isAdmin = false }) {
      const owner = normalizeDiscordUserId(discordUserId)
      const exactCode = normalizeGiftCode(code)
      const guild = guildId ? normalizeGuildId(guildId) : null
      const recorded = ingestion
        ? await ingestion.ingest({
          source: {
            sourceType: isAdmin ? "manual_admin" : "manual_user",
            sourceName: isAdmin ? "Discord admin submissions" : "Discord user submissions",
            trusted: false
          },
          code: exactCode,
          observationKey: `discord-submit:${crypto.randomUUID()}`,
          submittedByDiscordUserId: owner,
          submissionKind: isAdmin ? "admin" : "user",
          provenance: { transport: "discord", guildId: guild }
        })
        : await repository.recordSubmission({
          code: exactCode,
          submittedByDiscordUserId: owner,
          metadata: {
            submissionKind: isAdmin ? "admin" : "user",
            transport: "discord",
            guildId: guild
          }
        })
      return {
        ...recorded,
        outcome: recorded.duplicate
          ? SUBMISSION_LABELS[recorded.giftCode.status] || "already known"
          : "new candidate"
      }
    },

    async setAutomaticRedemption({
      discordUserId,
      guildId = null,
      playerId = null,
      enabled,
      preferenceSource = "user"
    }) {
      const guild = guildId ? normalizeGuildId(guildId) : null
      const account = await selectOwnedAccount(discordUserId, playerId, guild)
      if (!account.is_active) {
        throw new GiftCodeError("PLAYER_INACTIVE", `That ${terms.playerLabel} is inactive.`)
      }
      const result = await repository.setAutoRedemption({
        discordUserId: normalizeDiscordUserId(discordUserId),
        playerId: account.player_id,
        enabled: Boolean(enabled),
        guildId: guild,
        maximumEnabledAccounts: accountConfig.maximumAutoRedeemAccountsPerUser,
        preferenceSource
      })
      if (!result?.account) throw new GiftCodeError("PLAYER_NOT_FOUND", `No active ${terms.playerLabel} was found.`)
      if (result.limitReached) {
        throw new GiftCodeError(
          "AUTO_REDEEM_ACCOUNT_LIMIT",
          `You can currently enable automatic gift-code redemption for up to ` +
          `${accountConfig.maximumAutoRedeemAccountsPerUser} ${terms.gameName} accounts. ` +
          "Disable one before enabling another."
        )
      }
      return {
        ...result.account,
        guild_gift_code_enrolled: result.guildEnrolled,
        enabled_count: result.enabledCount,
        engagement_event: result.engagementEvent,
        maximum_enabled: accountConfig.maximumAutoRedeemAccountsPerUser
      }
    },

    async status({ discordUserId, playerId = null, guildId = null }) {
      const owner = normalizeDiscordUserId(discordUserId)
      const selectedPlayer = playerId ? normalizePlayerId(playerId, terms.playerLabel) : null
      const guild = guildId ? normalizeGuildId(guildId) : null
      return repository.activeAccountStatuses(owner, selectedPlayer, guild)
    },

    async history({ discordUserId, playerId = null, limit = 10 }) {
      const account = await selectOwnedAccount(discordUserId, playerId)
      return {
        account,
        redemptions: await repository.redemptionHistory(
          account.id,
          normalizeDiscordUserId(discordUserId),
          limit
        )
      }
    },

    async adminStatus() {
      return repository.diagnostics()
    },

    async activeCodes({ page = 0, pageSize = 15 } = {}) {
      return repository.activeCodeVisibility({ page, pageSize })
    },

    async adminCode(code) {
      const diagnostics = await repository.codeDiagnostics(normalizeGiftCode(code))
      return diagnostics
        ? { ...diagnostics, maximum_verification_attempts: workerConfig.maximumAttempts }
        : null
    }
  })
}

module.exports = { GiftCodeError, SUBMISSION_LABELS, createGiftCodeService }
