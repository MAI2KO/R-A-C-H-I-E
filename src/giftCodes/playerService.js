const { profileTerminology } = require("./terminology")
const {
  PlayerValidationError,
  normalizeDiscordUserId,
  normalizePlayerId,
  normalizeLocationNumber
} = require("./validation")
const { createNoopPlayerMirror } = require("./playerMirror")

class PlayerAccountError extends Error {
  constructor(code, message) {
    super(message)
    this.name = "PlayerAccountError"
    this.code = code
  }
}

function duplicatePlayerError(error) {
  return error?.code === "23505"
    && ["player_accounts_game_profile_player_id_key", "player_accounts_one_active_primary_idx"]
      .includes(error.constraint)
}

function createPlayerService({ repository, gameProfile, mirror = createNoopPlayerMirror(), logger = console }) {
  const terms = profileTerminology(gameProfile)

  return {
    terms,

    async register({ discordUserId, playerId, locationNumber }) {
      const owner = normalizeDiscordUserId(discordUserId)
      const player = normalizePlayerId(playerId, terms.playerLabel)
      const location = normalizeLocationNumber(locationNumber, terms.locationLabel)
      let account
      try {
        account = await repository.registerAccount({
          discordUserId: owner,
          playerId: player,
          locationNumber: location
        })
      } catch (error) {
        if (duplicatePlayerError(error)) {
          throw new PlayerAccountError(
            "PLAYER_ALREADY_REGISTERED",
            `That ${terms.playerLabel} is already registered for ${terms.gameName}.`
          )
        }
        throw error
      }
      try {
        await mirror.mirrorRegistration(account)
      } catch (error) {
        logger.warn(`[Player accounts] Optional mirror failed for ${gameProfile}: ${error?.code || "error"}`)
      }
      return account
    },

    async view({ discordUserId, playerId = null }) {
      const owner = normalizeDiscordUserId(discordUserId)
      if (playerId) {
        const player = normalizePlayerId(playerId, terms.playerLabel)
        const account = await repository.getOwnedAccount(owner, player)
        return account ? [account] : []
      }
      return repository.listOwnedAccounts(owner)
    },

    async changeLocation({ discordUserId, playerId, locationNumber }) {
      const owner = normalizeDiscordUserId(discordUserId)
      const player = normalizePlayerId(playerId, terms.playerLabel)
      const location = normalizeLocationNumber(locationNumber, terms.locationLabel)
      const result = await repository.updateLocation({
        discordUserId: owner,
        playerId: player,
        newNumber: location,
        changeSource: "user_command"
      })
      if (!result) {
        throw new PlayerAccountError("PLAYER_NOT_FOUND", `No active ${terms.playerLabel} was found.`)
      }
      return result
    },

    async remove({ discordUserId, playerId }) {
      const owner = normalizeDiscordUserId(discordUserId)
      const player = normalizePlayerId(playerId, terms.playerLabel)
      const result = await repository.deactivateAccount({ discordUserId: owner, playerId: player })
      if (!result) {
        throw new PlayerAccountError("PLAYER_NOT_FOUND", `No active ${terms.playerLabel} was found.`)
      }
      return result
    }
  }
}

module.exports = {
  PlayerValidationError,
  PlayerAccountError,
  duplicatePlayerError,
  createPlayerService
}
