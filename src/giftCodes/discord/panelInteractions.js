const {
  ActionRowBuilder,
  ButtonBuilder,
  ButtonStyle,
  ChannelSelectMenuBuilder,
  ChannelType,
  MessageFlags,
  ModalBuilder,
  StringSelectMenuBuilder,
  TextInputBuilder,
  TextInputStyle
} = require("discord.js")

const { getPool } = require("../../db")
const { InteractionSessionError, InteractionSessionStore } = require("../../interactionSessions")
const { createPlayerRepository } = require("../playerRepository")
const { createPlayerService, PlayerAccountError } = require("../playerService")
const { createGiftCodeRepository } = require("../repository")
const { createGiftCodeService, GiftCodeError } = require("../service")
const { createGiftCodeCommunityRepository } = require("../communityRepository")
const { createGiftCodeCommunityService, GiftCodeCommunityError } = require("../communityService")
const { createGiftCodeSourceRepository } = require("../sourceRepository")
const { createGiftCodeSourceIngestionService } = require("../sourceIngestion")
const { giftAccountConfig } = require("../config")
const { PlayerValidationError } = require("../validation")
const { getPlayerGiftCodesHealth } = require("../runtime")
const { getGiftCodeRuntime } = require("../workflowRuntime")
const { formatCodeDiagnostics } = require("./diagnostics")

const PREFIX = "gcux:"
const IDS = Object.freeze({
  playerAdd: `${PREFIX}pa:`,
  playerLocation: `${PREFIX}pl:`,
  playerGift: `${PREFIX}pg:`,
  playerRemove: `${PREFIX}pr:`,
  playerSelect: `${PREFIX}ps:`,
  giftSubmit: `${PREFIX}gs:`,
  giftToggle: `${PREFIX}gt:`,
  giftHistory: `${PREFIX}gh:`,
  giftChange: `${PREFIX}gc:`,
  giftActive: `${PREFIX}gal:`,
  giftActivePrevious: `${PREFIX}gap:`,
  giftActiveNext: `${PREFIX}gan:`,
  giftRegister: `${PREFIX}gr:`,
  registerModal: `${PREFIX}rm:`,
  locationModal: `${PREFIX}lm:`,
  submitModal: `${PREFIX}sm:`,
  adminChannel: `${PREFIX}ac:`,
  adminVerify: `${PREFIX}av:`,
  adminInspect: `${PREFIX}ai:`,
  adminQueue: `${PREFIX}aq:`,
  adminStats: `${PREFIX}as:`,
  adminRefresh: `${PREFIX}af:`,
  adminChannelSelect: `${PREFIX}acs:`,
  adminSourceChannel: `${PREFIX}asc:`,
  adminSourceChannelSelect: `${PREFIX}ascs:`,
  verifyModal: `${PREFIX}vm:`,
  inspectModal: `${PREFIX}im:`
})
const COMMANDS = new Set(["player-register", "gift-code-add", "gift-codes", "gift-codes-admin"])
const sessionStore = new InteractionSessionStore({ maximumSessions: 250 })

function suffix(customId, prefix) {
  return String(customId || "").slice(prefix.length)
}

function interactionContext(interaction, health) {
  return {
    userId: interaction.user.id,
    guildId: interaction.guildId,
    gameProfile: health.gameProfile
  }
}

function selectedAccount(accounts, playerId) {
  const activeAccounts = accounts.filter(account => account.is_active)
  return activeAccounts.find(account => account.player_id === playerId)
    || activeAccounts.find(account => account.is_primary)
    || activeAccounts[0]
    || null
}

function accountMenu(sessionId, accounts, selected) {
  const activeAccounts = accounts.filter(account => account.is_active)
  if (activeAccounts.length < 2) return null
  return new ActionRowBuilder().addComponents(
    new StringSelectMenuBuilder()
      .setCustomId(`${IDS.playerSelect}${sessionId}`)
      .setPlaceholder("Choose a player account")
      .addOptions(activeAccounts.slice(0, 25).map(account => ({
        label: `Player ID ${account.player_id}`,
        description: `${account.is_active ? "Active" : "Inactive"}${account.is_primary ? " · Primary" : ""}`,
        value: account.player_id,
        default: account.player_id === selected?.player_id
      })))
  )
}

function playerPanel({ sessionId, accounts, selected, terms }) {
  if (!selected) {
    return {
      content: "Register your game account to use supported player features such as automatic gift-code redemption.",
      components: [new ActionRowBuilder().addComponents(
        new ButtonBuilder().setCustomId(`${IDS.playerAdd}${sessionId}`)
          .setLabel("Register Player").setStyle(ButtonStyle.Primary)
      )]
    }
  }
  const components = []
  const menu = accountMenu(sessionId, accounts, selected)
  if (menu) components.push(menu)
  components.push(new ActionRowBuilder().addComponents(
    new ButtonBuilder().setCustomId(`${IDS.playerAdd}${sessionId}`)
      .setLabel("Add Account").setStyle(ButtonStyle.Primary),
    new ButtonBuilder().setCustomId(`${IDS.playerLocation}${sessionId}`)
      .setLabel(`Change ${terms.locationLabel}`).setStyle(ButtonStyle.Secondary),
    new ButtonBuilder().setCustomId(`${IDS.playerGift}${sessionId}`)
      .setLabel("Gift Code Settings").setStyle(ButtonStyle.Secondary),
    new ButtonBuilder().setCustomId(`${IDS.playerRemove}${sessionId}`)
      .setLabel("Remove Account").setStyle(ButtonStyle.Danger)
  ))
  return {
    content: [
      `**${terms.gameName} Player Accounts**`,
      `Selected: Player ID ${selected.player_id}`,
      `${terms.locationLabel}: ${selected.state_or_kingdom_number}`,
      `Primary: ${selected.is_primary ? "Yes" : "No"}`,
      `Active: ${selected.is_active ? "Yes" : "No"}`,
      `Automatic gift-code redemption: ${selected.gift_redemption_enabled ? "Enabled" : "Disabled"}`
    ].join("\n"),
    components
  }
}

function giftPanel({ sessionId, accounts, selected, terms, maximumEnabled }) {
  const enabledCount = accounts.filter(account => account.is_active && account.gift_redemption_enabled).length
  if (!selected) {
    return {
      content: [
        `**${terms.gameName} Gift Codes**`,
        "No registered player account is available."
      ].join("\n"),
      components: [new ActionRowBuilder().addComponents(
        new ButtonBuilder().setCustomId(`${IDS.giftSubmit}${sessionId}`)
          .setLabel("Submit Gift Code").setStyle(ButtonStyle.Primary),
        new ButtonBuilder().setCustomId(`${IDS.giftActive}${sessionId}`)
          .setLabel("Active Codes").setStyle(ButtonStyle.Secondary),
        new ButtonBuilder().setCustomId(`${IDS.giftRegister}${sessionId}`)
          .setLabel("Register Player").setStyle(ButtonStyle.Secondary)
      )]
    }
  }
  const components = []
  const menu = accountMenu(sessionId, accounts, selected)
  if (menu) components.push(menu)
  components.push(new ActionRowBuilder().addComponents(
    new ButtonBuilder().setCustomId(`${IDS.giftSubmit}${sessionId}`)
      .setLabel("Submit Gift Code").setStyle(ButtonStyle.Primary),
    new ButtonBuilder().setCustomId(`${IDS.giftToggle}${sessionId}`)
      .setLabel(selected.gift_redemption_enabled
        ? (selected.guild_gift_code_enrolled ? "Disable Auto-Redeem" : "Join This Server")
        : "Enable Auto-Redeem")
      .setStyle(selected.gift_redemption_enabled && selected.guild_gift_code_enrolled
        ? ButtonStyle.Danger
        : ButtonStyle.Success),
    new ButtonBuilder().setCustomId(`${IDS.giftHistory}${sessionId}`)
      .setLabel("Redemption History").setStyle(ButtonStyle.Secondary),
    new ButtonBuilder().setCustomId(`${IDS.giftActive}${sessionId}`)
      .setLabel("Active Codes").setStyle(ButtonStyle.Secondary),
    new ButtonBuilder().setCustomId(`${IDS.giftChange}${sessionId}`)
      .setLabel("Change Player").setStyle(ButtonStyle.Secondary)
  ))
  return {
    content: [
      `**${terms.gameName} Gift Codes**`,
      `Selected player: ${selected.player_id}`,
      `${terms.locationLabel}: ${selected.state_or_kingdom_number}`,
      `Auto-redemption: ${selected.gift_redemption_enabled ? "Enabled" : "Disabled"}`,
      `This server: ${selected.guild_gift_code_enrolled ? "Enrolled" : "Not enrolled"}`,
      `Accounts enabled: ${enabledCount} / ${maximumEnabled}`,
      `Recent result: ${selected.last_redemption_status || "None"}`
    ].join("\n"),
    components
  }
}

function activeCodesPanel({ sessionId, visibility }) {
  const totalPages = Math.max(1, Math.ceil(visibility.activeCount / visibility.pageSize))
  const page = Math.min(visibility.page, totalPages - 1)
  const codeLines = visibility.codes.length
    ? visibility.codes.map(row => row.code)
    : ["No active gift codes."]
  return {
    content: `**Active Gift Codes**\n\n${codeLines.join("\n")}\n\n` +
      `Active codes: ${visibility.activeCount}\n` +
      `Expired codes recorded: ${visibility.expiredCount}` +
      (totalPages > 1 ? `\nPage ${page + 1} of ${totalPages}` : ""),
    components: [new ActionRowBuilder().addComponents(
      new ButtonBuilder().setCustomId(`${IDS.giftActivePrevious}${sessionId}`)
        .setLabel("Previous").setStyle(ButtonStyle.Secondary).setDisabled(page === 0),
      new ButtonBuilder().setCustomId(`${IDS.giftChange}${sessionId}`)
        .setLabel("Back").setStyle(ButtonStyle.Secondary),
      new ButtonBuilder().setCustomId(`${IDS.giftActiveNext}${sessionId}`)
        .setLabel("Next").setStyle(ButtonStyle.Secondary).setDisabled(page >= totalPages - 1)
    )]
  }
}

function registrationModal(sessionId, terms) {
  return new ModalBuilder().setCustomId(`${IDS.registerModal}${sessionId}`)
    .setTitle(`Register ${terms.gameName} player`)
    .addComponents(
      new ActionRowBuilder().addComponents(
        new TextInputBuilder().setCustomId("player_id").setLabel("Player ID")
          .setStyle(TextInputStyle.Short).setRequired(true).setMaxLength(32)
      ),
      new ActionRowBuilder().addComponents(
        new TextInputBuilder().setCustomId("location").setLabel(terms.locationLabel)
          .setStyle(TextInputStyle.Short).setRequired(true).setMaxLength(10)
      )
    )
}

function oneFieldModal(customId, title, fieldId, label, maximumLength = 128) {
  return new ModalBuilder().setCustomId(customId).setTitle(title).addComponents(
    new ActionRowBuilder().addComponents(
      new TextInputBuilder().setCustomId(fieldId).setLabel(label)
        .setStyle(TextInputStyle.Short).setRequired(true).setMaxLength(maximumLength)
    )
  )
}

function sourceTime(value) {
  return value ? `<t:${Math.floor(new Date(value).getTime() / 1000)}:R>` : "Never"
}

function latestSourceValue(sources, field) {
  return sources.reduce((latest, source) => {
    if (!source[field]) return latest
    return !latest || new Date(source[field]) > new Date(latest) ? source[field] : latest
  }, null)
}

function adminPanel({ sessionId, runtime, diagnostics, configuration, sourceStatus, terms }) {
  const settings = configuration.settings
  const channel = settings?.gift_code_channel_id
    ? (configuration.channelAvailable ? `<#${settings.gift_code_channel_id}>` : "Configured but unavailable")
    : "Not configured"
  const role = configuration.roleAvailable
    ? "🍭 Ready"
    : configuration.roleStatus === "error"
      ? "Unable to create/assign"
      : "Created automatically when first earned"
  const mirrorChannels = sourceStatus?.channels?.filter(channel => channel.enabled) || []
  const catalogue = sourceStatus?.sources?.find(source => source.source_type === "public_catalogue")
  const mirrors = sourceStatus?.sources?.filter(source => source.source_type === "discord_mirror") || []
  return {
    content: [
      "**Gift Code Administration**",
      `Game: ${terms.gameName}`,
      `Verification worker: ${runtime.verificationEnabled ? "Enabled" : "Disabled"}`,
      `Redemption worker: ${runtime.redemptionEnabled ? "Enabled" : "Disabled"}`,
      `Verifier: ${runtime.verifierConfigured ? "Configured" : "Not configured"}`,
      `Gift-code channel: ${channel}`,
      `Contributor reward role: ${role}`,
      "",
      `Pending verification: ${diagnostics.pending_candidates}`,
      `Active codes: ${diagnostics.active_codes}`,
      `Expired codes: ${diagnostics.expired_codes}`,
      `Invalid codes: ${diagnostics.invalid_codes}`,
      `Restricted/review: ${diagnostics.restricted_review_codes}`,
      `Pending redemptions: ${diagnostics.pending_redemptions}`,
      `Retry queue: ${diagnostics.retry_count}`,
      "",
      "**Sources**",
      `Official Discord mirror: ${mirrorChannels.length ? `Enabled (${mirrorChannels.length})` : "Not configured"}`,
      `Last mirror observation: ${sourceTime(latestSourceValue(mirrors, "last_observation_at_utc"))}`,
      `Last mirror candidate: ${sourceTime(latestSourceValue(mirrors, "last_candidate_at_utc"))}`,
      `Public catalogue: ${runtime.sourcePollingEnabled ? "Enabled" : "Disabled"}`,
      `Last poll: ${sourceTime(catalogue?.last_poll_at_utc)}`,
      `Last successful poll: ${sourceTime(catalogue?.last_successful_poll_at_utc)}`,
      `Codes observed: ${catalogue?.observations_count || 0}`,
      `New candidates: ${catalogue?.candidates_count || 0}`,
      `Last source error: ${catalogue?.last_error || "None"}`
    ].join("\n"),
    components: [
      new ActionRowBuilder().addComponents(
        new ButtonBuilder().setCustomId(`${IDS.adminChannel}${sessionId}`).setLabel("Configure Channel").setStyle(ButtonStyle.Secondary),
        new ButtonBuilder().setCustomId(`${IDS.adminVerify}${sessionId}`).setLabel("Verify Code").setStyle(ButtonStyle.Primary),
        new ButtonBuilder().setCustomId(`${IDS.adminInspect}${sessionId}`).setLabel("Inspect Code").setStyle(ButtonStyle.Secondary)
      ),
      new ActionRowBuilder().addComponents(
        new ButtonBuilder().setCustomId(`${IDS.adminSourceChannel}${sessionId}`).setLabel("Configure Source Channel").setStyle(ButtonStyle.Secondary),
        new ButtonBuilder().setCustomId(`${IDS.adminQueue}${sessionId}`).setLabel("Queue Status").setStyle(ButtonStyle.Secondary),
        new ButtonBuilder().setCustomId(`${IDS.adminStats}${sessionId}`).setLabel("Community Stats").setStyle(ButtonStyle.Secondary),
        new ButtonBuilder().setCustomId(`${IDS.adminRefresh}${sessionId}`).setLabel("Refresh").setStyle(ButtonStyle.Secondary)
      )
    ]
  }
}

function formatCommunityStats(stats) {
  return [
    "**Community Statistics**",
    `Players using Auto-Redeem: ${stats.auto_redeem_players}`,
    `Characters covered: ${stats.enabled_accounts}`,
    `Successful redemptions: ${stats.successful_redemptions}`,
    `Verified gift codes: ${stats.verified_codes}`,
    `Already claimed: ${stats.already_redeemed}`,
    `Successful this month: ${stats.successful_this_month}`,
    `Latest verified code: ${stats.latest_verified_code || "None"}`,
    `Unique contributors: ${stats.unique_contributors}`
  ].join("\n")
}

async function acknowledge(interaction) {
  if (interaction.isModalSubmit?.() || interaction.isChatInputCommand?.()) {
    await interaction.deferReply({ flags: MessageFlags.Ephemeral })
  } else {
    await interaction.deferUpdate()
  }
}

function isPanelInteraction(interaction) {
  return COMMANDS.has(interaction.commandName) || String(interaction.customId || "").startsWith(PREFIX)
}

function giftCodePanelHandler(interaction) {
  const customId = String(interaction.customId || "")
  if (interaction.commandName) return `command:${interaction.commandName}`
  const handlers = [
    [IDS.adminChannelSelect, "configure_channel_select"],
    [IDS.adminChannel, "configure_channel"],
    [IDS.adminSourceChannelSelect, "configure_source_channel_select"],
    [IDS.adminSourceChannel, "configure_source_channel"],
    [IDS.adminVerify, "verify_code"],
    [IDS.adminInspect, "inspect_code"],
    [IDS.adminQueue, "queue_status"],
    [IDS.adminStats, "community_stats"],
    [IDS.adminRefresh, "admin_refresh"]
  ]
  return handlers.find(([prefix]) => customId.startsWith(prefix))?.[1] || "panel_interaction"
}

function safePanelErrorMessage(error) {
  return String(error?.message || "Gift-code panel failed")
    .replace(/https?:\/\/\S+/gi, "[url]")
    .replace(/(admin_api_key|authorization|token)\s*[=:]\s*\S+/gi, "$1=[redacted]")
    .slice(0, 300)
}

function giftCodePanelFailureDiagnostics(interaction, health, error) {
  const guildId = String(interaction.guildId || "")
  return {
    event: "gift_code_panel_failure",
    game_profile: health?.gameProfile || "unknown",
    guild_id: /^[0-9]{1,32}$/.test(guildId) ? guildId : undefined,
    interaction_category: interaction.isChatInputCommand?.()
      ? "command"
      : interaction.isModalSubmit?.() ? "modal" : "component",
    handler: String(error?.giftCodeHandler || giftCodePanelHandler(interaction)).slice(0, 100),
    error_name: String(error?.name || "Error").slice(0, 100),
    error_code: String(error?.code || "unknown").slice(0, 100),
    error_message: safePanelErrorMessage(error)
  }
}

async function handleGiftCodePanelInteraction(interaction, {
  userCanManageServer,
  healthProvider = getPlayerGiftCodesHealth,
  runtimeProvider = getGiftCodeRuntime,
  poolProvider = () => getPool(),
  playerRepositoryFactory = createPlayerRepository,
  playerServiceFactory = createPlayerService,
  giftRepositoryFactory = createGiftCodeRepository,
  giftServiceFactory = createGiftCodeService,
  sourceRepositoryFactory = createGiftCodeSourceRepository,
  sourceIngestionFactory = createGiftCodeSourceIngestionService,
  communityRepositoryFactory = createGiftCodeCommunityRepository,
  communityServiceFactory = createGiftCodeCommunityService,
  sessions = sessionStore,
  env = process.env,
  logger = console
} = {}) {
  if (!isPanelInteraction(interaction)) return false
  const health = healthProvider()

  try {
    const customId = String(interaction.customId || "")
    const modalPrefix = [IDS.playerAdd, IDS.giftRegister].find(prefix => customId.startsWith(prefix))
    if (modalPrefix) {
      const sessionId = suffix(customId, modalPrefix)
      sessions.get(sessionId, interactionContext(interaction, health))
      await interaction.showModal(registrationModal(sessionId, require("../terminology").profileTerminology(health.gameProfile)))
      return true
    }
    if (customId.startsWith(IDS.playerLocation)) {
      const sessionId = suffix(customId, IDS.playerLocation)
      sessions.get(sessionId, interactionContext(interaction, health))
      const terms = require("../terminology").profileTerminology(health.gameProfile)
      await interaction.showModal(oneFieldModal(
        `${IDS.locationModal}${sessionId}`,
        `Change ${terms.locationLabel}`,
        "location",
        `${terms.locationLabel} number`,
        10
      ))
      return true
    }
    if (customId.startsWith(IDS.giftSubmit)) {
      const sessionId = suffix(customId, IDS.giftSubmit)
      sessions.get(sessionId, interactionContext(interaction, health))
      await interaction.showModal(oneFieldModal(
        `${IDS.submitModal}${sessionId}`, "Submit Gift Code", "code", "Gift code"
      ))
      return true
    }
    if (customId.startsWith(IDS.adminVerify) || customId.startsWith(IDS.adminInspect)) {
      const prefix = customId.startsWith(IDS.adminVerify) ? IDS.adminVerify : IDS.adminInspect
      const sessionId = suffix(customId, prefix)
      sessions.get(sessionId, interactionContext(interaction, health))
      await interaction.showModal(oneFieldModal(
        `${prefix === IDS.adminVerify ? IDS.verifyModal : IDS.inspectModal}${sessionId}`,
        prefix === IDS.adminVerify ? "Controlled Code Verification" : "Inspect Gift Code",
        "code",
        "Gift code"
      ))
      return true
    }

    await acknowledge(interaction)
    if (!health.available) {
      await interaction.editReply({ content: "Player and gift-code services are temporarily unavailable.", components: [] })
      return true
    }

    const pool = poolProvider()
    const playerRepository = playerRepositoryFactory(pool, health.gameProfile)
    const players = playerServiceFactory({ repository: playerRepository, gameProfile: health.gameProfile, logger })
    const giftRepository = giftRepositoryFactory(pool, health.gameProfile)
    const sourceRepository = sourceRepositoryFactory(pool, health.gameProfile)
    const ingestion = sourceIngestionFactory({
      giftRepository,
      sourceRepository,
      gameProfile: health.gameProfile,
      logger
    })
    const gifts = giftServiceFactory({
      repository: giftRepository,
      gameProfile: health.gameProfile,
      env,
      ingestion
    })
    const communityRepository = communityRepositoryFactory(pool, health.gameProfile)
    const community = communityServiceFactory({
      repository: communityRepository,
      client: interaction.client,
      gameProfile: health.gameProfile,
      maximumEnabledAccounts: giftAccountConfig(env).maximumAutoRedeemAccountsPerUser,
      logger
    })
    const context = interactionContext(interaction, health)
    const isAdmin = interaction.commandName === "gift-codes-admin"
      || Object.values(IDS).some(prefix => prefix.startsWith(`${PREFIX}a`) && customId.startsWith(prefix))
      || customId.startsWith(IDS.verifyModal)
      || customId.startsWith(IDS.inspectModal)
    if (isAdmin && !(await userCanManageServer(interaction))) {
      await interaction.editReply({ content: "You do not have permission to use this panel.", components: [] })
      return true
    }

    async function loadAccounts() {
      return gifts.status({ discordUserId: interaction.user.id, guildId: interaction.guildId })
    }

    async function renderPlayer(sessionId, preferredPlayerId = null) {
      const accounts = await loadAccounts()
      const selected = selectedAccount(accounts, preferredPlayerId)
      sessions.update(sessionId, context, { selectedPlayerId: selected?.player_id || null, panel: "player" })
      await interaction.editReply(playerPanel({ sessionId, accounts, selected, terms: gifts.terms }))
    }

    async function renderGift(sessionId, preferredPlayerId = null) {
      const accounts = await loadAccounts()
      const selected = selectedAccount(accounts, preferredPlayerId)
      sessions.update(sessionId, context, { selectedPlayerId: selected?.player_id || null, panel: "gift" })
      await interaction.editReply(giftPanel({
        sessionId,
        accounts,
        selected,
        terms: gifts.terms,
        maximumEnabled: giftAccountConfig(env).maximumAutoRedeemAccountsPerUser
      }))
    }

    async function renderActiveCodes(sessionId, page = 0) {
      const visibility = await gifts.activeCodes({ page, pageSize: 15 })
      sessions.update(sessionId, context, { panel: "active-codes", activeCodesPage: visibility.page })
      await interaction.editReply(activeCodesPanel({ sessionId, visibility }))
    }

    async function renderAdmin(sessionId) {
      const [diagnostics, configuration, sourceStatus] = await Promise.all([
        gifts.adminStatus(),
        community.configuration(interaction.guildId),
        sourceRepository.sourceStatus(interaction.guildId)
      ])
      const runtime = runtimeProvider()?.status() || {
        verificationEnabled: false,
        redemptionEnabled: false,
        verifierConfigured: false
      }
      await interaction.editReply(adminPanel({
        sessionId, runtime, diagnostics, configuration, sourceStatus, terms: gifts.terms
      }))
    }

    if (interaction.isChatInputCommand?.()) {
      if (interaction.commandName === "gift-code-add") {
        const result = await gifts.submit({
          discordUserId: interaction.user.id,
          guildId: interaction.guildId,
          code: interaction.options.getString("code")
        })
        const messages = {
          active: "I've already got that one.",
          verifying: "Already checking that one.",
          candidate: "Already checking that one.",
          expired: "That one's already marked expired.",
          invalid: "I've already checked that one and it wasn't valid.",
          restricted: "That one's already waiting for review.",
          unknown: "That one's already waiting for review."
        }
        await interaction.editReply({
          content: result.duplicate ? messages[result.giftCode.status] || "I've already got that one." : "Got it. I'll check that one.",
          components: []
        })
        return true
      }
      const panel = interaction.commandName === "player-register"
        ? "player"
        : interaction.commandName === "gift-codes" ? "gift" : "admin"
      const sessionId = sessions.create(context, { panel, selectedPlayerId: null })
      if (panel === "player") await renderPlayer(sessionId)
      else if (panel === "gift") await renderGift(sessionId)
      else await renderAdmin(sessionId)
      return true
    }

    const matchingPrefix = Object.values(IDS).find(prefix => customId.startsWith(prefix))
    const sessionId = matchingPrefix ? suffix(customId, matchingPrefix) : ""
    const session = sessions.get(sessionId, context)

    if (customId.startsWith(IDS.playerSelect)) {
      const playerId = interaction.values[0]
      if (session.data.panel === "gift") await renderGift(sessionId, playerId)
      else await renderPlayer(sessionId, playerId)
      return true
    }
    if (customId.startsWith(IDS.registerModal)) {
      const account = await players.register({
        discordUserId: interaction.user.id,
        guildId: interaction.guildId,
        playerId: interaction.fields.getTextInputValue("player_id"),
        locationNumber: interaction.fields.getTextInputValue("location")
      })
      if (session.data.panel === "gift") await renderGift(sessionId, account.player_id)
      else await renderPlayer(sessionId, account.player_id)
      return true
    }
    if (customId.startsWith(IDS.locationModal)) {
      await players.changeLocation({
        discordUserId: interaction.user.id,
        playerId: session.data.selectedPlayerId,
        locationNumber: interaction.fields.getTextInputValue("location")
      })
      await renderPlayer(sessionId, session.data.selectedPlayerId)
      return true
    }
    if (customId.startsWith(IDS.playerRemove)) {
      await players.remove({
        discordUserId: interaction.user.id,
        playerId: session.data.selectedPlayerId
      })
      await renderPlayer(sessionId)
      return true
    }
    if (customId.startsWith(IDS.playerGift)) {
      await renderGift(sessionId, session.data.selectedPlayerId)
      return true
    }
    if (customId.startsWith(IDS.giftChange)) {
      await renderGift(sessionId, session.data.selectedPlayerId)
      return true
    }
    if (customId.startsWith(IDS.submitModal)) {
      const result = await gifts.submit({
        discordUserId: interaction.user.id,
        guildId: interaction.guildId,
        code: interaction.fields.getTextInputValue("code")
      })
      const runtime = runtimeProvider()?.status() || {}
      const availability = runtime.verificationEnabled && runtime.verifierConfigured
        ? "Verification is queued."
        : "Verification is currently unavailable; the candidate remains pending."
      await interaction.editReply({
        content: `Gift code ${result.giftCode.code}: ${result.outcome}. ${availability}`,
        components: []
      })
      return true
    }
    if (customId.startsWith(IDS.giftToggle)) {
      const accounts = await gifts.status({
        discordUserId: interaction.user.id,
        playerId: session.data.selectedPlayerId,
        guildId: interaction.guildId
      })
      const current = accounts[0]
      const account = await gifts.setAutomaticRedemption({
        discordUserId: interaction.user.id,
        guildId: interaction.guildId,
        playerId: session.data.selectedPlayerId,
        enabled: !(current.gift_redemption_enabled && current.guild_gift_code_enrolled)
      })
      await community.onAutoRedemptionEnabled(account.engagement_event)
      await renderGift(sessionId, account.player_id)
      return true
    }
    if (customId.startsWith(IDS.giftHistory)) {
      const history = await gifts.history({
        discordUserId: interaction.user.id,
        playerId: session.data.selectedPlayerId
      })
      const lines = history.redemptions.length
        ? history.redemptions.map(row => `${row.code} · ${row.status} · ${row.location_number_snapshot}`)
        : ["No redemption history yet."]
      await interaction.editReply({
        content: `**Redemption History · Player ${history.account.player_id}**\n${lines.join("\n")}`,
        components: [new ActionRowBuilder().addComponents(
          new ButtonBuilder().setCustomId(`${IDS.playerGift}${sessionId}`).setLabel("Back").setStyle(ButtonStyle.Secondary)
        )]
      })
      return true
    }
    if (customId.startsWith(IDS.giftActive)) {
      await renderActiveCodes(sessionId, 0)
      return true
    }
    if (customId.startsWith(IDS.giftActivePrevious)) {
      await renderActiveCodes(sessionId, Math.max(0, Number(session.data.activeCodesPage || 0) - 1))
      return true
    }
    if (customId.startsWith(IDS.giftActiveNext)) {
      await renderActiveCodes(sessionId, Number(session.data.activeCodesPage || 0) + 1)
      return true
    }

    if (customId.startsWith(IDS.adminChannel)) {
      await interaction.editReply({
        content: "**Configure Gift-Code Channel**",
        components: [new ActionRowBuilder().addComponents(
          new ChannelSelectMenuBuilder().setCustomId(`${IDS.adminChannelSelect}${sessionId}`)
            .setPlaceholder("Select a text channel")
            .setChannelTypes(ChannelType.GuildText, ChannelType.GuildAnnouncement)
            .setMinValues(1).setMaxValues(1)
        )]
      })
      return true
    }
    if (customId.startsWith(IDS.adminChannelSelect)) {
      await community.configureChannel(interaction.guildId, interaction.values[0])
      try {
        await renderAdmin(sessionId)
      } catch (error) {
        error.giftCodeHandler = "configure_channel_reload"
        throw error
      }
      return true
    }
    if (customId.startsWith(IDS.adminSourceChannel)) {
      await interaction.editReply({
        content: "**Configure Gift-Code Source Channel**\nMessages in this channel are read for candidates. Verified announcements use the separate gift-code channel.",
        components: [new ActionRowBuilder().addComponents(
          new ChannelSelectMenuBuilder().setCustomId(`${IDS.adminSourceChannelSelect}${sessionId}`)
            .setPlaceholder("Select a mirrored source channel")
            .setChannelTypes(ChannelType.GuildText, ChannelType.GuildAnnouncement)
            .setMinValues(1).setMaxValues(1)
        )]
      })
      return true
    }
    if (customId.startsWith(IDS.adminSourceChannelSelect)) {
      await sourceRepository.configureDiscordChannel({
        guildId: interaction.guildId,
        channelId: interaction.values[0]
      })
      await renderAdmin(sessionId)
      return true
    }
    if (customId.startsWith(IDS.verifyModal)) {
      const submission = await gifts.submit({
        discordUserId: interaction.user.id,
        guildId: interaction.guildId,
        code: interaction.fields.getTextInputValue("code"),
        isAdmin: true
      })
      const result = await runtimeProvider()?.verifyCode(submission.giftCode.code)
      await interaction.editReply({
        content: result?.processed
          ? `TEST verification completed. Result: ${result.result.giftCode.status}.`
          : `TEST verification did not run: ${result?.reason || "runtime unavailable"}.`,
        components: []
      })
      return true
    }
    if (customId.startsWith(IDS.inspectModal)) {
      const code = await gifts.adminCode(interaction.fields.getTextInputValue("code"))
      await interaction.editReply({
        content: formatCodeDiagnostics(code),
        components: []
      })
      return true
    }
    if (customId.startsWith(IDS.adminQueue)) {
      const diagnostics = await gifts.adminStatus()
      await interaction.editReply({
        content: `**Queue Status**\nPending: ${diagnostics.pending_redemptions}\n` +
          `Retrying: ${diagnostics.retry_count}\n` +
          `Oldest pending: ${diagnostics.oldest_pending_at_utc ? `<t:${Math.floor(new Date(diagnostics.oldest_pending_at_utc).getTime() / 1000)}:R>` : "none"}\n` +
          `Next retry: ${diagnostics.next_retry_at_utc ? `<t:${Math.floor(new Date(diagnostics.next_retry_at_utc).getTime() / 1000)}:R>` : "none"}`,
        components: []
      })
      return true
    }
    if (customId.startsWith(IDS.adminStats)) {
      const stats = await community.communityStats(interaction.guildId)
      await interaction.editReply({
        content: formatCommunityStats(stats),
        components: []
      })
      return true
    }
    if (customId.startsWith(IDS.adminRefresh)) {
      await renderAdmin(sessionId)
      return true
    }

    return false
  } catch (error) {
    const expected = error instanceof PlayerValidationError
      || error instanceof PlayerAccountError
      || error instanceof GiftCodeError
      || error instanceof GiftCodeCommunityError
      || error instanceof InteractionSessionError
    const message = expected
      ? error.message
      : "Player and gift-code services are temporarily unavailable."
    if (!(error instanceof InteractionSessionError)) {
      logger.error(JSON.stringify(giftCodePanelFailureDiagnostics(interaction, health, error)))
    }
    if (interaction.deferred || interaction.replied) {
      await interaction.editReply({ content: message, components: [] }).catch(() => {})
    } else {
      await interaction.reply({ content: message, flags: MessageFlags.Ephemeral }).catch(() => {})
    }
    return true
  }
}

module.exports = {
  PREFIX,
  IDS,
  selectedAccount,
  accountMenu,
  playerPanel,
  giftPanel,
  activeCodesPanel,
  registrationModal,
  adminPanel,
  formatCommunityStats,
  giftCodePanelHandler,
  safePanelErrorMessage,
  giftCodePanelFailureDiagnostics,
  isPanelInteraction,
  handleGiftCodePanelInteraction
}
