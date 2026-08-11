const { normalizeDiscordUserId, normalizeGiftCode, normalizePlayerId } = require("./validation")
const { profileTerminology } = require("./terminology")

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

function createGiftCodeService({ repository, gameProfile }) {
  const terms = profileTerminology(gameProfile)

  async function selectOwnedAccount(discordUserId, playerId = null) {
    const owner = normalizeDiscordUserId(discordUserId)
    const selectedPlayer = playerId ? normalizePlayerId(playerId, terms.playerLabel) : null
    const accounts = await repository.accountStatuses(owner, selectedPlayer)
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

    async submit({ discordUserId, code, isAdmin = false }) {
      const owner = normalizeDiscordUserId(discordUserId)
      const exactCode = normalizeGiftCode(code)
      const recorded = await repository.recordSubmission({
        code: exactCode,
        submittedByDiscordUserId: owner,
        metadata: { submissionKind: isAdmin ? "admin" : "user", transport: "discord" }
      })
      return {
        ...recorded,
        outcome: recorded.duplicate
          ? SUBMISSION_LABELS[recorded.giftCode.status] || "already known"
          : "new candidate"
      }
    },

    async setAutomaticRedemption({ discordUserId, playerId = null, enabled }) {
      const account = await selectOwnedAccount(discordUserId, playerId)
      if (!account.is_active) {
        throw new GiftCodeError("PLAYER_INACTIVE", `That ${terms.playerLabel} is inactive.`)
      }
      const updated = await repository.setAutoRedemption({
        discordUserId: normalizeDiscordUserId(discordUserId),
        playerId: account.player_id,
        enabled: Boolean(enabled)
      })
      if (!updated) throw new GiftCodeError("PLAYER_NOT_FOUND", `No active ${terms.playerLabel} was found.`)
      return updated
    },

    async status({ discordUserId, playerId = null }) {
      const owner = normalizeDiscordUserId(discordUserId)
      const selectedPlayer = playerId ? normalizePlayerId(playerId, terms.playerLabel) : null
      return repository.accountStatuses(owner, selectedPlayer)
    },

    async adminStatus() {
      return repository.diagnostics()
    },

    async adminCode(code) {
      return repository.codeDiagnostics(normalizeGiftCode(code))
    }
  })
}

module.exports = { GiftCodeError, SUBMISSION_LABELS, createGiftCodeService }
