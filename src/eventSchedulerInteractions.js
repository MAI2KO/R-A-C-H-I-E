const {
  ActionRowBuilder,
  ButtonBuilder,
  ButtonStyle,
  MessageFlags,
  ModalBuilder,
  TextInputBuilder,
  TextInputStyle
} = require("discord.js")

const { getPool } = require("./db")
const { getEventSchedulerHealth } = require("./eventSchedulerHealth")
const { createEventSchedulerRepository } = require("./eventSchedulerRepository")
const {
  SchedulerValidationError,
  normalizeAllianceName,
  resolveSendableChannel
} = require("./eventSchedulerService")

const IDS = Object.freeze({
  prefix: "es:",
  home: "es:home",
  configure: "es:cfg",
  configureModal: "es:cfgm",
  state: "es:state",
  stateModal: "es:stm",
  stateOff: "es:stoff"
})

function isSchedulerInteraction(interaction) {
  return interaction.commandName === "event-scheduler"
    || String(interaction.customId || "").startsWith(IDS.prefix)
}

function buildHomeView(settings, stateLink) {
  const allianceChannel = settings ? `<#${settings.event_channel_id}>` : "Not configured"
  const stateChannel = stateLink
    ? `<#${stateLink.state_event_channel_id}> (${stateLink.sharing_enabled ? "enabled" : "disabled"})`
    : "Not configured"

  return {
    content:
      `Event scheduler\n\n` +
      `Alliance: ${settings?.alliance_name || "Not configured"}\n` +
      `Alliance event channel: ${allianceChannel}\n` +
      `State event channel: ${stateChannel}`,
    components: [
      new ActionRowBuilder().addComponents(
        new ButtonBuilder()
          .setCustomId(IDS.configure)
          .setLabel("Alliance channel")
          .setStyle(ButtonStyle.Primary),
        new ButtonBuilder()
          .setCustomId(IDS.state)
          .setLabel("State sharing")
          .setStyle(ButtonStyle.Secondary),
        new ButtonBuilder()
          .setCustomId(IDS.stateOff)
          .setLabel("Disable state")
          .setStyle(ButtonStyle.Secondary)
          .setDisabled(!stateLink?.sharing_enabled)
      )
    ]
  }
}

function buildAllianceModal(settings) {
  return new ModalBuilder()
    .setCustomId(IDS.configureModal)
    .setTitle("Alliance event channel")
    .addComponents(
      new ActionRowBuilder().addComponents(
        new TextInputBuilder()
          .setCustomId("a")
          .setLabel("Alliance name")
          .setStyle(TextInputStyle.Short)
          .setRequired(true)
          .setMaxLength(100)
          .setValue(settings?.alliance_name || "")
      ),
      new ActionRowBuilder().addComponents(
        new TextInputBuilder()
          .setCustomId("c")
          .setLabel("Alliance event channel ID")
          .setStyle(TextInputStyle.Short)
          .setRequired(true)
          .setMaxLength(22)
          .setValue(settings?.event_channel_id || "")
      )
    )
}

function buildStateModal(stateLink) {
  return new ModalBuilder()
    .setCustomId(IDS.stateModal)
    .setTitle("State event sharing")
    .addComponents(
      new ActionRowBuilder().addComponents(
        new TextInputBuilder()
          .setCustomId("g")
          .setLabel("State Discord guild ID")
          .setStyle(TextInputStyle.Short)
          .setRequired(true)
          .setMaxLength(22)
          .setValue(stateLink?.state_guild_id || "")
      ),
      new ActionRowBuilder().addComponents(
        new TextInputBuilder()
          .setCustomId("c")
          .setLabel("State event channel ID")
          .setStyle(TextInputStyle.Short)
          .setRequired(true)
          .setMaxLength(22)
          .setValue(stateLink?.state_event_channel_id || "")
      )
    )
}

async function loadHome(repository, guildId) {
  const [settings, stateLink] = await Promise.all([
    repository.getGuildSettings(guildId),
    repository.getStateLink(guildId)
  ])
  return buildHomeView(settings, stateLink)
}

async function replyUnavailable(interaction) {
  const payload = {
    content: "The event scheduler is temporarily unavailable. Existing bot features are unaffected.",
    flags: MessageFlags.Ephemeral
  }
  if (interaction.deferred || interaction.replied) {
    await interaction.editReply({ content: payload.content, components: [] })
  } else {
    await interaction.reply(payload)
  }
}

async function handleEventSchedulerInteraction(
  interaction,
  { userCanManageServer, healthProvider = getEventSchedulerHealth } = {}
) {
  if (!isSchedulerInteraction(interaction)) return false

  const health = healthProvider()
  if (!health.available) {
    await replyUnavailable(interaction)
    return true
  }

  if (!(await userCanManageServer(interaction))) {
    await interaction.reply({
      content: "You do not have permission to manage the event scheduler.",
      flags: MessageFlags.Ephemeral
    })
    return true
  }

  const repository = createEventSchedulerRepository(getPool(), health.gameProfile)
  const guildId = interaction.guildId

  try {
    if (interaction.isChatInputCommand?.() && interaction.commandName === "event-scheduler") {
      await interaction.deferReply({ flags: MessageFlags.Ephemeral })
      await interaction.editReply(await loadHome(repository, guildId))
      return true
    }

    if (interaction.isButton?.() && interaction.customId === IDS.home) {
      await interaction.deferUpdate()
      await interaction.editReply(await loadHome(repository, guildId))
      return true
    }

    if (interaction.isButton?.() && interaction.customId === IDS.configure) {
      const settings = await repository.getGuildSettings(guildId)
      await interaction.showModal(buildAllianceModal(settings))
      return true
    }

    if (interaction.isButton?.() && interaction.customId === IDS.state) {
      const settings = await repository.getGuildSettings(guildId)
      if (!settings) {
        await interaction.reply({
          content: "Configure the alliance event channel first.",
          flags: MessageFlags.Ephemeral
        })
        return true
      }
      const stateLink = await repository.getStateLink(guildId)
      await interaction.showModal(buildStateModal(stateLink))
      return true
    }

    if (interaction.isButton?.() && interaction.customId === IDS.stateOff) {
      await interaction.deferUpdate()
      await repository.setStateSharing(guildId, false)
      await interaction.editReply(await loadHome(repository, guildId))
      return true
    }

    if (interaction.isModalSubmit?.() && interaction.customId === IDS.configureModal) {
      await interaction.deferReply({ flags: MessageFlags.Ephemeral })
      const allianceName = normalizeAllianceName(interaction.fields.getTextInputValue("a"))
      const channelId = interaction.fields.getTextInputValue("c")
      const target = await resolveSendableChannel(interaction.client, guildId, channelId)

      await repository.upsertGuildSettings({
        guildId,
        botInstanceName: health.botInstanceName,
        allianceName,
        eventChannelId: target.channelId
      })
      await interaction.editReply(await loadHome(repository, guildId))
      return true
    }

    if (interaction.isModalSubmit?.() && interaction.customId === IDS.stateModal) {
      await interaction.deferReply({ flags: MessageFlags.Ephemeral })
      const stateGuildId = interaction.fields.getTextInputValue("g")
      const stateChannelId = interaction.fields.getTextInputValue("c")
      const target = await resolveSendableChannel(
        interaction.client,
        stateGuildId,
        stateChannelId
      )

      await repository.upsertStateLink({
        allianceGuildId: guildId,
        configuredByBotInstance: health.botInstanceName,
        stateGuildId: target.guildId,
        stateEventChannelId: target.channelId,
        sharingEnabled: true
      })
      await interaction.editReply(await loadHome(repository, guildId))
      return true
    }
  } catch (error) {
    const message = error instanceof SchedulerValidationError
      ? error.message
      : "The event scheduler could not complete that action."

    if (interaction.deferred || interaction.replied) {
      await interaction.editReply({ content: message, components: [] })
    } else {
      await interaction.reply({ content: message, flags: MessageFlags.Ephemeral })
    }
    return true
  }

  return true
}

module.exports = {
  IDS,
  isSchedulerInteraction,
  buildHomeView,
  buildAllianceModal,
  buildStateModal,
  handleEventSchedulerInteraction
}
