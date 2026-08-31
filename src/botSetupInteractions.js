const {
  ActionRowBuilder, ButtonBuilder, ButtonStyle, MessageFlags, ModalBuilder,
  PermissionFlagsBits, SlashCommandBuilder, TextInputBuilder, TextInputStyle
} = require("discord.js")
const crypto = require("node:crypto")
const { getPool } = require("./db")
const { getPlayerGiftCodesHealth } = require("./giftCodes/runtime")
const { createBotSetupRepository } = require("./botSetupRepository")
const { BotSetupError, createBotSetupService } = require("./botSetupService")
const { BOT_SETUP_IDS } = require("./botSetupService")
const {
  interactionIsDiscordAdministrator,
  interactionIsGuildOwner
} = require("./botManagerAuthorization")

const SETUP_MODAL_ID = "botsetup:community"
const SETUP_APPLY_PREFIX = "botsetup:apply:"
const SETUP_SESSION_TTL_MS = 15 * 60 * 1000
const setupSessions = new Map()

function nativeBookingStatus(result) {
  return `${result.status}; bookings ${result.bookingsOpen ? "open" : "closed"}`
}

function setupPublicError(error, gameProfile) {
  const location = gameProfile === "kingshot" ? "Kingdom" : "State"
  if (error?.code === "kingshot_defaults_unavailable") {
    return "This Kingdom is not yet configured for automatic native booking cycles. No setup was changed."
  }
  if (["community_claim_conflict", "guild_conflict"].includes(error?.code)) {
    return `That ${location} or Discord server is already linked elsewhere. Platform approval is required; no mapping was changed.`
  }
  if (error?.code === "community_inactive") {
    return `That native ${location} is inactive. No setup was changed.`
  }
  return null
}

function pruneSetupSessions(now = Date.now()) {
  for (const [token, session] of setupSessions) {
    if (now - session.createdAt > SETUP_SESSION_TTL_MS) setupSessions.delete(token)
  }
}

function buildCommunitySetupModal(gameProfile) {
  const location = gameProfile === "kingshot" ? "Kingdom" : "State"
  return new ModalBuilder().setCustomId(SETUP_MODAL_ID).setTitle("Community setup")
    .addComponents(
      new ActionRowBuilder().addComponents(
        new TextInputBuilder().setCustomId("community_number").setLabel(`${location} number`)
          .setStyle(TextInputStyle.Short).setRequired(true).setMaxLength(10)
      ),
      new ActionRowBuilder().addComponents(
        new TextInputBuilder().setCustomId("alliance").setLabel("Alliance abbreviation")
          .setStyle(TextInputStyle.Short).setRequired(true).setMinLength(3).setMaxLength(3)
      )
    )
}

function applyRow(token) {
  return new ActionRowBuilder().addComponents(
    new ButtonBuilder().setCustomId(`${SETUP_APPLY_PREFIX}${token}`)
      .setLabel("Apply setup").setStyle(ButtonStyle.Primary)
  )
}

function buildBotSetupCommand() {
  return new SlashCommandBuilder()
    .setName("setup")
    .setDescription("Create or reconcile the bot-managed server channels")
    .setDefaultMemberPermissions(PermissionFlagsBits.Administrator)
}

async function handleBotSetupInteraction(interaction, {
  bookingApi = null,
  healthProvider = getPlayerGiftCodesHealth,
  poolProvider = () => getPool(),
  repositoryFactory = createBotSetupRepository,
  serviceFactory = createBotSetupService,
  logger = console
} = {}) {
  const isCommand = interaction.isChatInputCommand?.() && interaction.commandName === "setup"
  const isModal = interaction.isModalSubmit?.() && interaction.customId === SETUP_MODAL_ID
  const isApply = interaction.isButton?.() && String(interaction.customId || "").startsWith(SETUP_APPLY_PREFIX)
  if (!isCommand && !isModal && !isApply) return false
  if (!interactionIsGuildOwner(interaction) && !interactionIsDiscordAdministrator(interaction)) {
    await interaction.reply({ content: "You do not have permission to run bot setup.", flags: MessageFlags.Ephemeral })
    return true
  }
  const health = healthProvider()
  if (!health.available) {
    await interaction.reply({ content: "Bot setup is unavailable until PostgreSQL-backed services are ready.", flags: MessageFlags.Ephemeral })
    return true
  }
  if (!bookingApi?.communitySetup) {
    await interaction.reply({ content: "Bot setup is unavailable until the native booking integration is ready.", flags: MessageFlags.Ephemeral })
    return true
  }
  if (isCommand) {
    await interaction.showModal(buildCommunitySetupModal(health.gameProfile))
    return true
  }
  if (isModal) await interaction.deferReply({ flags: MessageFlags.Ephemeral })
  else await interaction.deferUpdate()
  try {
    const repository = repositoryFactory(poolProvider(), health.gameProfile)
    const service = serviceFactory({
      repository,
      client: interaction.client,
      gameProfile: health.gameProfile,
      botInstanceName: health.botInstanceName,
      logger
    })
    if (isModal) {
      pruneSetupSessions()
      const community = {
        communityNumber: interaction.fields.getTextInputValue("community_number"),
        allianceAbbreviation: interaction.fields.getTextInputValue("alliance")
      }
      const native = await bookingApi.communitySetup({
        guildId: interaction.guildId, guildName: interaction.guild.name,
        communityCode: community.communityNumber, discordUserId: interaction.user.id,
        allianceAbbreviation: community.allianceAbbreviation, dryRun: true
      })
      const preview = await service.preview(
        interaction.guildId, community, nativeBookingStatus(native)
      )
      const token = crypto.randomUUID()
      setupSessions.set(token, {
        createdAt: Date.now(), guildId: interaction.guildId, userId: interaction.user.id,
        community: preview.community
      })
      await interaction.editReply({ content: preview.content, components: [applyRow(token)] })
      return true
    }
    const token = String(interaction.customId).slice(SETUP_APPLY_PREFIX.length)
    const session = setupSessions.get(token)
    setupSessions.delete(token)
    if (!session || Date.now() - session.createdAt > SETUP_SESSION_TTL_MS
        || session.guildId !== interaction.guildId || session.userId !== interaction.user.id) {
      await interaction.editReply({ content: "This setup preview expired. Run `/setup` again.", components: [] })
      return true
    }
    const native = await bookingApi.communitySetup({
      guildId: interaction.guildId, guildName: interaction.guild.name,
      communityCode: session.community.communityNumber, discordUserId: interaction.user.id,
      allianceAbbreviation: session.community.allianceAbbreviation, dryRun: false
    })
    const result = await service.reconcile(
      interaction.guildId, session.community, nativeBookingStatus(native), native
    )
    await interaction.editReply({ content: result.content, components: [] })
  } catch (error) {
    if (error instanceof BotSetupError) {
      await interaction.editReply({ content: error.message, components: [] })
    } else {
      logger.error(`[Bot setup] Failed: ${String(error?.code || error?.name || "error").slice(0, 100)}`)
      await interaction.editReply({ content: setupPublicError(error, health.gameProfile)
        || "Bot setup could not be completed. No unrelated channels were changed.", components: [] })
    }
  }
  return true
}

async function handlePersistentOnboardingInteraction(interaction) {
  if (!interaction.isButton?.() || interaction.customId !== BOT_SETUP_IDS.registerMinisters) {
    return false
  }
  await interaction.reply({
    content: "Minister registration and bookings have moved to the authenticated community website. Use `/register` here to manage your game character.",
    flags: MessageFlags.Ephemeral
  })
  return true
}

module.exports = {
  buildBotSetupCommand,
  buildCommunitySetupModal,
  SETUP_MODAL_ID,
  SETUP_APPLY_PREFIX,
  setupPublicError,
  handleBotSetupInteraction,
  handlePersistentOnboardingInteraction
}
