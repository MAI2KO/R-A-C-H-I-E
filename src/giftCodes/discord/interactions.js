const { MessageFlags } = require("discord.js")
const { getPool } = require("../../db")
const { createPlayerRepository } = require("../playerRepository")
const {
  PlayerValidationError,
  PlayerAccountError,
  createPlayerService
} = require("../playerService")
const { getPlayerGiftCodesHealth } = require("../runtime")

function formatAccount(account, terms) {
  return [
    `**${terms.gameName}**`,
    `${terms.playerLabel}: ${account.player_id}`,
    `${terms.locationLabel}: ${account.state_or_kingdom_number}`,
    `Account: ${account.is_primary ? "Primary" : "Secondary"}`,
    `Gift redemption: ${account.gift_redemption_enabled ? "Enabled" : "Disabled"}`,
    `Verification: ${account.verification_status}`
  ].join("\n")
}

async function handlePlayerInteraction(
  interaction,
  {
    healthProvider = getPlayerGiftCodesHealth,
    poolProvider = () => getPool(),
    repositoryFactory = createPlayerRepository,
    serviceFactory = createPlayerService,
    logger = console
  } = {}
) {
  if (!interaction.isChatInputCommand?.() || interaction.commandName !== "player") return false

  await interaction.deferReply({ flags: MessageFlags.Ephemeral })
  const health = healthProvider()
  if (!health.available) {
    await interaction.editReply("Player accounts are temporarily unavailable. Please try again later.")
    return true
  }

  try {
    const repository = repositoryFactory(poolProvider(), health.gameProfile)
    const service = serviceFactory({ repository, gameProfile: health.gameProfile, logger })
    const subcommand = interaction.options.getSubcommand()
    const playerId = interaction.options.getString("player_id")
    const location = ["register", "location"].includes(subcommand)
      ? interaction.options.getString(service.terms.locationLabelLower)
      : null

    if (subcommand === "register") {
      const account = await service.register({
        discordUserId: interaction.user.id,
        playerId,
        locationNumber: location
      })
      await interaction.editReply(
        `Registered ${service.terms.playerLabel} ${account.player_id} in ` +
        `${service.terms.locationLabel} ${account.state_or_kingdom_number} as ` +
        `${account.is_primary ? "your primary account" : "a secondary account"}.`
      )
      return true
    }

    if (subcommand === "view") {
      const accounts = await service.view({ discordUserId: interaction.user.id, playerId })
      await interaction.editReply(accounts.length
        ? accounts.map(account => formatAccount(account, service.terms)).join("\n\n")
        : "You have no matching active player account.")
      return true
    }

    if (subcommand === "location") {
      const result = await service.changeLocation({
        discordUserId: interaction.user.id,
        playerId,
        locationNumber: location
      })
      await interaction.editReply(result.changed
        ? `Your ${service.terms.locationLabel} has been changed from ` +
          `${result.previousNumber} to ${result.account.state_or_kingdom_number}.`
        : `Your ${service.terms.locationLabel} is already ${result.account.state_or_kingdom_number}.`)
      return true
    }

    if (subcommand === "remove") {
      const result = await service.remove({ discordUserId: interaction.user.id, playerId })
      const replacement = result.replacement
        ? ` ${service.terms.playerLabel} ${result.replacement.player_id} is now primary.`
        : ""
      await interaction.editReply(`Player account ${result.account.player_id} was deactivated.${replacement}`)
      return true
    }

    throw new PlayerAccountError("UNKNOWN_PLAYER_ACTION", "That player action is not supported.")
  } catch (error) {
    if (error instanceof PlayerValidationError || error instanceof PlayerAccountError) {
      await interaction.editReply(error.message)
      return true
    }
    logger.error(`[Player accounts] Interaction failed: ${error?.code || error?.name || "error"}`)
    await interaction.editReply("Player accounts are temporarily unavailable. Please try again later.")
    return true
  }
}

module.exports = { formatAccount, handlePlayerInteraction }
