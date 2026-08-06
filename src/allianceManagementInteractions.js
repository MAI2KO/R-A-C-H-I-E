const crypto = require("crypto")
const {
  ActionRowBuilder,
  ButtonBuilder,
  ButtonStyle,
  MessageFlags,
  ModalBuilder,
  StringSelectMenuBuilder,
  TextInputBuilder,
  TextInputStyle
} = require("discord.js")

const { InteractionSessionError, InteractionSessionStore } = require("./interactionSessions")
const { normalizeAllianceName, SchedulerValidationError } = require("./eventSchedulerService")
const { sessionContext } = require("./eventCreationInteractions")

const allianceSessions = new InteractionSessionStore()
const ALLIANCES_PER_PAGE = 10
const ALLIANCE_IDS = Object.freeze({
  prefix: "am:",
  open: "am:open",
  listPrefix: "am:l:",
  selectPrefix: "am:s:",
  addPrefix: "am:add:",
  addModalPrefix: "am:addm:",
  renamePrefix: "am:ren:",
  renameModalPrefix: "am:renm:",
  deletePrefix: "am:del:",
  deleteConfirmPrefix: "am:delc:"
})

function token() {
  return crypto.randomBytes(6).toString("base64url")
}

function allianceModal(customId, title, value = "") {
  return new ModalBuilder()
    .setCustomId(customId)
    .setTitle(title)
    .addComponents(
      new ActionRowBuilder().addComponents(
        new TextInputBuilder()
          .setCustomId("name")
          .setLabel("Alliance or sub-alliance name")
          .setStyle(TextInputStyle.Short)
          .setRequired(true)
          .setMaxLength(100)
          .setValue(value)
      )
    )
}

function allianceListView(alliances, page, total, sessionId, selectedAllianceId) {
  const totalPages = Math.max(1, Math.ceil(total / ALLIANCES_PER_PAGE))
  const tokenMap = {}
  const lines = alliances.map(alliance => {
    const role = alliance.is_default ? "Main alliance" : "Sub-alliance"
    return `${role}: ${alliance.alliance_name} (${alliance.managed_event_count} active/paused events)`
  })
  const components = []
  if (alliances.length) {
    components.push(new ActionRowBuilder().addComponents(
      new StringSelectMenuBuilder()
        .setCustomId(`${ALLIANCE_IDS.selectPrefix}${sessionId}`)
        .setPlaceholder("Select an alliance to manage")
        .addOptions(alliances.map(alliance => {
          const allianceToken = token()
          tokenMap[allianceToken] = String(alliance.id)
          return {
            label: alliance.alliance_name.slice(0, 100),
            value: allianceToken,
            description: alliance.is_default ? "Main alliance" : "Sub-alliance",
            default: String(alliance.id) === String(selectedAllianceId || "")
          }
        }))
    ))
  }
  components.push(new ActionRowBuilder().addComponents(
    new ButtonBuilder().setCustomId(`${ALLIANCE_IDS.addPrefix}${sessionId}`)
      .setLabel("Add").setStyle(ButtonStyle.Success),
    new ButtonBuilder().setCustomId(`${ALLIANCE_IDS.renamePrefix}${sessionId}`)
      .setLabel("Rename").setStyle(ButtonStyle.Primary).setDisabled(!selectedAllianceId),
    new ButtonBuilder().setCustomId(`${ALLIANCE_IDS.deletePrefix}${sessionId}`)
      .setLabel("Delete").setStyle(ButtonStyle.Danger).setDisabled(!selectedAllianceId),
    new ButtonBuilder().setCustomId(`${ALLIANCE_IDS.listPrefix}${sessionId}:${Math.max(0, page - 1)}`)
      .setLabel("Previous").setStyle(ButtonStyle.Secondary).setDisabled(page === 0),
    new ButtonBuilder().setCustomId(`${ALLIANCE_IDS.listPrefix}${sessionId}:${page + 1}`)
      .setLabel("Next").setStyle(ButtonStyle.Secondary).setDisabled(page >= totalPages - 1)
  ))
  components.push(new ActionRowBuilder().addComponents(
    new ButtonBuilder().setCustomId("es:home").setLabel("Back").setStyle(ButtonStyle.Secondary)
  ))
  return {
    view: {
      content: `Alliances - page ${page + 1} of ${totalPages}\n\n${lines.join("\n") || "No alliances configured."}`,
      components,
      allowedMentions: { parse: [], repliedUser: false }
    },
    tokenMap
  }
}

async function renderAllianceList(
  interaction,
  repository,
  health,
  page = 0,
  sessionId = null
) {
  const context = sessionContext(interaction, health)
  const safePage = Math.max(0, Number(page) || 0)
  const id = sessionId || allianceSessions.create(context, {})
  const session = allianceSessions.get(id, context)
  const result = await repository.listAlliances(interaction.guildId, {
    limit: ALLIANCES_PER_PAGE,
    offset: safePage * ALLIANCES_PER_PAGE
  })
  const built = allianceListView(
    result.alliances,
    safePage,
    result.total,
    id,
    session.data.selectedAllianceId
  )
  allianceSessions.update(id, context, {
    page: safePage,
    tokenMap: built.tokenMap
  })
  if (interaction.deferred || interaction.replied) await interaction.editReply(built.view)
  else await interaction.reply({ ...built.view, flags: MessageFlags.Ephemeral })
}

function duplicateNameError(error) {
  if (error?.code === "23505" && error?.constraint === "event_alliances_name_unique_ci") {
    return new SchedulerValidationError(
      "An alliance with that name already exists in this Discord server and game profile."
    )
  }
  return error
}

async function selectedAlliance(repository, interaction, session) {
  const allianceId = session.data.selectedAllianceId
  if (!allianceId) throw new InteractionSessionError("Select an alliance first.")
  const alliance = await repository.getAlliance(interaction.guildId, allianceId)
  if (!alliance) throw new InteractionSessionError("That alliance is no longer available.")
  return alliance
}

async function handleAllianceManagementInteraction(interaction, { repository, health }) {
  const customId = String(interaction.customId || "")
  if (!customId.startsWith(ALLIANCE_IDS.prefix)) return false
  const context = sessionContext(interaction, health)

  if (interaction.isButton?.() && customId === ALLIANCE_IDS.open) {
    await interaction.deferUpdate()
    await renderAllianceList(interaction, repository, health)
    return true
  }

  if (interaction.isButton?.() && customId.startsWith(ALLIANCE_IDS.listPrefix)) {
    const [sessionId, page] = customId.slice(ALLIANCE_IDS.listPrefix.length).split(":")
    allianceSessions.get(sessionId, context)
    await interaction.deferUpdate()
    await renderAllianceList(interaction, repository, health, Number(page), sessionId)
    return true
  }

  if (interaction.isStringSelectMenu?.() && customId.startsWith(ALLIANCE_IDS.selectPrefix)) {
    const sessionId = customId.slice(ALLIANCE_IDS.selectPrefix.length)
    const session = allianceSessions.get(sessionId, context)
    const allianceId = session.data.tokenMap?.[interaction.values[0]]
    if (!allianceId) throw new InteractionSessionError("That alliance control has expired.")
    allianceSessions.update(sessionId, context, { selectedAllianceId: allianceId })
    await interaction.deferUpdate()
    await renderAllianceList(interaction, repository, health, session.data.page, sessionId)
    return true
  }

  if (interaction.isButton?.() && customId.startsWith(ALLIANCE_IDS.addPrefix)) {
    const sessionId = customId.slice(ALLIANCE_IDS.addPrefix.length)
    allianceSessions.get(sessionId, context)
    await interaction.showModal(allianceModal(
      `${ALLIANCE_IDS.addModalPrefix}${sessionId}`,
      "Add alliance or sub-alliance"
    ))
    return true
  }

  if (interaction.isModalSubmit?.() && customId.startsWith(ALLIANCE_IDS.addModalPrefix)) {
    const sessionId = customId.slice(ALLIANCE_IDS.addModalPrefix.length)
    const session = allianceSessions.get(sessionId, context)
    const allianceName = normalizeAllianceName(interaction.fields.getTextInputValue("name"))
    await interaction.deferReply({ flags: MessageFlags.Ephemeral })
    try {
      const alliance = await repository.createAlliance({
        guildId: interaction.guildId,
        allianceName,
        createdByBotInstance: health.botInstanceName
      })
      allianceSessions.update(sessionId, context, { selectedAllianceId: String(alliance.id) })
    } catch (error) {
      throw duplicateNameError(error)
    }
    await renderAllianceList(interaction, repository, health, session.data.page, sessionId)
    return true
  }

  if (interaction.isButton?.() && customId.startsWith(ALLIANCE_IDS.renamePrefix)) {
    const sessionId = customId.slice(ALLIANCE_IDS.renamePrefix.length)
    const session = allianceSessions.get(sessionId, context)
    const alliance = await selectedAlliance(repository, interaction, session)
    await interaction.showModal(allianceModal(
      `${ALLIANCE_IDS.renameModalPrefix}${sessionId}`,
      "Rename alliance",
      alliance.alliance_name
    ))
    return true
  }

  if (interaction.isModalSubmit?.() && customId.startsWith(ALLIANCE_IDS.renameModalPrefix)) {
    const sessionId = customId.slice(ALLIANCE_IDS.renameModalPrefix.length)
    const session = allianceSessions.get(sessionId, context)
    const allianceName = normalizeAllianceName(interaction.fields.getTextInputValue("name"))
    await interaction.deferReply({ flags: MessageFlags.Ephemeral })
    try {
      const renamed = await repository.renameAlliance({
        guildId: interaction.guildId,
        allianceId: session.data.selectedAllianceId,
        allianceName
      })
      if (!renamed) throw new InteractionSessionError("That alliance is no longer available.")
    } catch (error) {
      throw duplicateNameError(error)
    }
    await renderAllianceList(interaction, repository, health, session.data.page, sessionId)
    return true
  }

  if (interaction.isButton?.() && customId.startsWith(ALLIANCE_IDS.deletePrefix)) {
    const sessionId = customId.slice(ALLIANCE_IDS.deletePrefix.length)
    const session = allianceSessions.get(sessionId, context)
    const alliance = await selectedAlliance(repository, interaction, session)
    await interaction.deferUpdate()
    await interaction.editReply({
      content: `Delete ${alliance.alliance_name}? Alliances with event history cannot be deleted.`,
      components: [new ActionRowBuilder().addComponents(
        new ButtonBuilder().setCustomId(`${ALLIANCE_IDS.deleteConfirmPrefix}${sessionId}`)
          .setLabel("Confirm delete").setStyle(ButtonStyle.Danger),
        new ButtonBuilder().setCustomId(`${ALLIANCE_IDS.listPrefix}${sessionId}:${session.data.page}`)
          .setLabel("Cancel").setStyle(ButtonStyle.Secondary)
      )],
      allowedMentions: { parse: [], repliedUser: false }
    })
    return true
  }

  if (interaction.isButton?.() && customId.startsWith(ALLIANCE_IDS.deleteConfirmPrefix)) {
    const sessionId = customId.slice(ALLIANCE_IDS.deleteConfirmPrefix.length)
    const session = allianceSessions.get(sessionId, context)
    await interaction.deferUpdate()
    const result = await repository.deleteAlliance({
      guildId: interaction.guildId,
      allianceId: session.data.selectedAllianceId
    })
    if (!result.deleted) {
      const message = result.reason === "default"
        ? "The main alliance cannot be deleted. Rename it instead."
        : result.reason === "events"
          ? "That alliance cannot be deleted while event history still belongs to it."
          : "That alliance is no longer available."
      await interaction.editReply({
        content: message,
        components: [new ActionRowBuilder().addComponents(
          new ButtonBuilder().setCustomId(`${ALLIANCE_IDS.listPrefix}${sessionId}:${session.data.page}`)
            .setLabel("Back to alliances").setStyle(ButtonStyle.Secondary)
        )]
      })
      return true
    }
    allianceSessions.update(sessionId, context, { selectedAllianceId: null })
    await renderAllianceList(interaction, repository, health, session.data.page, sessionId)
    return true
  }

  return false
}

module.exports = {
  ALLIANCES_PER_PAGE,
  ALLIANCE_IDS,
  allianceSessions,
  allianceModal,
  allianceListView,
  renderAllianceList,
  handleAllianceManagementInteraction
}
