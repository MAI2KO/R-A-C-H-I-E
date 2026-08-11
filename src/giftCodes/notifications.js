const { profileTerminology } = require("./terminology")

function notificationMessage(claim, status, gameProfile) {
  const terms = profileTerminology(gameProfile)
  const identity = `${terms.playerLabel} ${claim.player_id_snapshot} in ` +
    `${terms.locationLabel} ${claim.location_number_snapshot}`
  if (status === "success") {
    return `Gift code ${claim.code} was redeemed successfully for ${identity}.`
  }
  if (status === "already_redeemed") {
    return `Gift code ${claim.code} had already been claimed for ${identity}.`
  }
  if (status === "invalid_player") {
    return `Gift code ${claim.code} could not be redeemed for ${identity}. Please confirm your ` +
      `${terms.locationLabel} is current with /player location.`
  }
  if (status === "restricted") {
    return `Gift code ${claim.code} is not available for ${identity}.`
  }
  if (status === "unknown") {
    return `Gift code ${claim.code} returned an unrecognised result for ${identity}. An administrator can review it.`
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

module.exports = { notificationMessage, createDiscordGiftNotifier }
