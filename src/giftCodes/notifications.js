const { profileTerminology } = require("./terminology")

function maskedPlayerId(playerId) {
  const value = String(playerId || "")
  return `***${value.slice(-4)}`
}

function notificationMessage(claim, status, gameProfile) {
  const terms = profileTerminology(gameProfile)
  const identity = Number(claim.owner_account_count) > 1
    ? ` for ${terms.playerLabel} ${maskedPlayerId(claim.player_id_snapshot)}`
    : ""
  if (status === "success") {
    return `${claim.code} redeemed${identity}. Check your in-game mail.`
  }
  if (status === "already_redeemed") {
    const character = identity || " on this character"
    return `${claim.code} was already claimed${character}. You're good.`
  }
  if (status === "invalid_player") {
    return `I couldn't redeem ${claim.code}${identity}. Check that this character is still in ` +
      `${terms.locationLabel} ${claim.location_number_snapshot}.`
  }
  if (status === "redemption_limit") {
    return `${claim.code} can't be claimed on this character - you've already used another code of this type.`
  }
  if (status === "restricted") {
    return `${claim.code} could not be redeemed on this character.`
  }
  if (status === "unknown") {
    return `${claim.code} returned an unrecognised result. An administrator can review it.`
  }
  return null
}

function createDiscordGiftNotifier({ client, gameProfile, logger = console }) {
  return async function notify(claim, status) {
    const content = notificationMessage(claim, status, gameProfile)
    if (!content) return { sent: false, suppressed: true }
    try {
      const user = await client.users.fetch(claim.discord_user_id)
      await user.send(content)
      return { sent: true }
    } catch (error) {
      const errorCode = String(error?.code || error?.name || "dm_failed").slice(0, 100)
      logger.warn(`[Gift codes] Player notification failed: ${errorCode}`)
      return { sent: false, errorCode }
    }
  }
}

module.exports = { maskedPlayerId, notificationMessage, createDiscordGiftNotifier }
