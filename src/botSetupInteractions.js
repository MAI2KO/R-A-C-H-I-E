const { MessageFlags, PermissionFlagsBits, SlashCommandBuilder } = require("discord.js")
const { getPool } = require("./db")
const { getPlayerGiftCodesHealth } = require("./giftCodes/runtime")
const { createBotSetupRepository } = require("./botSetupRepository")
const { BotSetupError, createBotSetupService } = require("./botSetupService")
const { BOT_SETUP_IDS } = require("./botSetupService")

function buildBotSetupCommand() {
  return new SlashCommandBuilder()
    .setName("bot-setup")
    .setDescription("Create or reconcile the bot-managed server channels")
    .setDefaultMemberPermissions(PermissionFlagsBits.ManageGuild)
}

async function handleBotSetupInteraction(interaction, {
  userCanManageServer,
  healthProvider = getPlayerGiftCodesHealth,
  poolProvider = () => getPool(),
  repositoryFactory = createBotSetupRepository,
  serviceFactory = createBotSetupService,
  logger = console
} = {}) {
  if (!interaction.isChatInputCommand?.() || interaction.commandName !== "bot-setup") return false
  await interaction.deferReply({ flags: MessageFlags.Ephemeral })
  if (!await userCanManageServer(interaction)) {
    await interaction.editReply("You do not have permission to run bot setup.")
    return true
  }
  const health = healthProvider()
  if (!health.available) {
    await interaction.editReply("Bot setup is unavailable until PostgreSQL-backed services are ready.")
    return true
  }
  try {
    const repository = repositoryFactory(poolProvider(), health.gameProfile)
    const service = serviceFactory({
      repository,
      client: interaction.client,
      gameProfile: health.gameProfile,
      logger
    })
    const result = await service.reconcile(interaction.guildId)
    await interaction.editReply(result.content)
  } catch (error) {
    if (error instanceof BotSetupError) {
      await interaction.editReply(error.message)
    } else {
      logger.error(`[Bot setup] Failed: ${String(error?.code || error?.name || "error").slice(0, 100)}`)
      await interaction.editReply("Bot setup could not be completed. No unrelated channels were changed.")
    }
  }
  return true
}

async function handlePersistentOnboardingInteraction(interaction, { ministerModalBuilder }) {
  if (!interaction.isButton?.() || interaction.customId !== BOT_SETUP_IDS.registerMinisters) {
    return false
  }
  await interaction.showModal(ministerModalBuilder())
  return true
}

module.exports = {
  buildBotSetupCommand,
  handleBotSetupInteraction,
  handlePersistentOnboardingInteraction
}
