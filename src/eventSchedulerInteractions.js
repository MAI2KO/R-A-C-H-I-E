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
  normalizeSnowflake,
  normalizeAllianceName,
  resolveSendableChannel
} = require("./eventSchedulerService")
const { handleEventCreationInteraction } = require("./eventCreationInteractions")
const { handleEventManagementInteraction } = require("./eventManagementInteractions")
const { handleAllianceManagementInteraction } = require("./allianceManagementInteractions")
const { parseUtcTime } = require("./timeParsing")

const IDS = Object.freeze({
  prefix: "es:",
  home: "es:home",
  configure: "es:cfg",
  configureModal: "es:cfgm",
  state: "es:state",
  stateModal: "es:stm",
  stateOff: "es:stoff",
  roundup: "es:round",
  roundupModal: "es:roundm"
})

function isSchedulerInteraction(interaction) {
  return interaction.commandName === "event-scheduler"
    || String(interaction.customId || "").startsWith(IDS.prefix)
    || String(interaction.customId || "").startsWith("ec:")
    || String(interaction.customId || "").startsWith("el:")
    || String(interaction.customId || "").startsWith("ep:")
    || String(interaction.customId || "").startsWith("mg:")
    || String(interaction.customId || "").startsWith("am:")
}

function buildHomeView(settings, stateLink) {
  const allianceChannel = settings ? `<#${settings.event_channel_id}>` : "Not configured"
  const stateChannel = stateLink
    ? `<#${stateLink.state_event_channel_id}> (${stateLink.sharing_enabled ? "enabled" : "disabled"})`
    : "Not configured"
  const roundup = settings?.weekly_roundup_enabled
    ? `<#${settings.weekly_roundup_channel_id}> on weekday ${settings.weekly_roundup_day} at ${String(settings.weekly_roundup_time_utc).slice(0, 5)} UTC`
    : "Disabled"

  return {
    content:
      `Event scheduler\n\n` +
      `Main alliance: ${settings?.alliance_name || "Not configured"}\n` +
      `Alliances: ${settings?.alliance_count ?? 0}\n` +
      `Alliance event channel: ${allianceChannel}\n` +
      `State weekly roundup channel: ${stateChannel}\n` +
      `Weekly roundup: ${roundup}`,
    components: [
      new ActionRowBuilder().addComponents(
        new ButtonBuilder()
          .setCustomId("ec:new")
          .setLabel("Create event")
          .setStyle(ButtonStyle.Success)
          .setDisabled(!settings),
        new ButtonBuilder()
          .setCustomId("el:0")
          .setLabel("View events")
          .setStyle(ButtonStyle.Primary),
        new ButtonBuilder()
          .setCustomId("am:open")
          .setLabel("Alliances")
          .setStyle(ButtonStyle.Primary)
          .setDisabled(!settings)
      ),
      new ActionRowBuilder().addComponents(
        new ButtonBuilder()
          .setCustomId(IDS.configure)
          .setLabel("Alliance channel")
          .setStyle(ButtonStyle.Primary),
        new ButtonBuilder()
          .setCustomId(IDS.state)
          .setLabel("State roundup")
          .setStyle(ButtonStyle.Secondary),
        new ButtonBuilder()
          .setCustomId(IDS.stateOff)
          .setLabel("Disable state roundup")
          .setStyle(ButtonStyle.Secondary)
          .setDisabled(!stateLink?.sharing_enabled)
      ),
      new ActionRowBuilder().addComponents(
        new ButtonBuilder()
          .setCustomId(IDS.roundup)
          .setLabel("Weekly roundup")
          .setStyle(ButtonStyle.Primary)
      )
    ]
  }
}

function buildAllianceModal(settings) {
  const modal = new ModalBuilder()
    .setCustomId(IDS.configureModal)
    .setTitle("Alliance reminder channel")
  if (!settings) {
    modal.addComponents(
      new ActionRowBuilder().addComponents(
        new TextInputBuilder()
          .setCustomId("a")
          .setLabel("Main alliance name")
          .setStyle(TextInputStyle.Short)
          .setRequired(true)
          .setMaxLength(100)
      )
    )
  }
  return modal.addComponents(
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
    .setTitle("State weekly roundup")
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
          .setLabel("State roundup channel ID")
          .setStyle(TextInputStyle.Short)
          .setRequired(true)
          .setMaxLength(22)
          .setValue(stateLink?.state_event_channel_id || "")
      )
    )
}

function buildRoundupModal(settings) {
  return new ModalBuilder()
    .setCustomId(IDS.roundupModal)
    .setTitle("Weekly roundup settings")
    .addComponents(
      new ActionRowBuilder().addComponents(
        new TextInputBuilder().setCustomId("e").setLabel("Enabled? yes or no")
          .setStyle(TextInputStyle.Short).setRequired(true)
          .setValue(settings?.weekly_roundup_enabled ? "yes" : "no")
      ),
      new ActionRowBuilder().addComponents(
        new TextInputBuilder().setCustomId("d").setLabel("Weekday (0 Sunday to 6 Saturday)")
          .setStyle(TextInputStyle.Short).setRequired(true)
          .setValue(String(settings?.weekly_roundup_day ?? 1))
      ),
      new ActionRowBuilder().addComponents(
        new TextInputBuilder().setCustomId("t").setLabel("UTC time")
          .setStyle(TextInputStyle.Short).setRequired(true)
          .setValue(String(settings?.weekly_roundup_time_utc || "09:00").slice(0, 5))
      ),
      new ActionRowBuilder().addComponents(
        new TextInputBuilder().setCustomId("c").setLabel("Alliance roundup channel ID")
          .setStyle(TextInputStyle.Short).setRequired(true).setMaxLength(22)
          .setValue(settings?.weekly_roundup_channel_id || settings?.event_channel_id || "")
      ),
      new ActionRowBuilder().addComponents(
        new TextInputBuilder().setCustomId("x").setLabel("Post when no events? yes or no")
          .setStyle(TextInputStyle.Short).setRequired(true)
          .setValue(settings?.roundup_when_empty ? "yes" : "no")
      )
    )
}

function parseYesNo(value, label) {
  const normalized = String(value || "").trim().toLowerCase()
  if (normalized === "yes") return true
  if (normalized === "no") return false
  throw new SchedulerValidationError(`${label} must be yes or no.`)
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
    if (await handleAllianceManagementInteraction(interaction, { repository, health })) return true

    if (await handleEventManagementInteraction(interaction, { repository, health })) return true

    if (await handleEventCreationInteraction(interaction, {
      repository,
      health,
      loadHome
    })) return true

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

    if (interaction.isButton?.() && interaction.customId === IDS.roundup) {
      const settings = await repository.getGuildSettings(guildId)
      if (!settings) {
        await interaction.reply({
          content: "Configure the alliance event channel first.",
          flags: MessageFlags.Ephemeral
        })
        return true
      }
      await interaction.showModal(buildRoundupModal(settings))
      return true
    }

    if (interaction.isModalSubmit?.() && interaction.customId === IDS.configureModal) {
      await interaction.deferReply({ flags: MessageFlags.Ephemeral })
      const existingSettings = await repository.getGuildSettings(guildId)
      const allianceName = existingSettings?.alliance_name || normalizeAllianceName(
        interaction.fields.getTextInputValue("a")
      )
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

    if (interaction.isModalSubmit?.() && interaction.customId === IDS.roundupModal) {
      await interaction.deferReply({ flags: MessageFlags.Ephemeral })
      const enabled = parseYesNo(interaction.fields.getTextInputValue("e"), "Enabled")
      const weekday = Number(interaction.fields.getTextInputValue("d"))
      if (!Number.isInteger(weekday) || weekday < 0 || weekday > 6) {
        throw new SchedulerValidationError("Weekday must be an integer from 0 to 6.")
      }
      const timeUtc = parseUtcTime(interaction.fields.getTextInputValue("t"))
      const channelInput = interaction.fields.getTextInputValue("c")
      const channelId = enabled
        ? (await resolveSendableChannel(interaction.client, guildId, channelInput)).channelId
        : normalizeSnowflake(channelInput, "Channel ID")
      const postWhenEmpty = parseYesNo(
        interaction.fields.getTextInputValue("x"),
        "Post when empty"
      )
      await repository.configureWeeklyRoundup({
        guildId,
        enabled,
        weekday,
        timeUtc,
        channelId,
        postWhenEmpty
      })
      await interaction.editReply(await loadHome(repository, guildId))
      return true
    }
  } catch (error) {
    const userFacingErrors = new Set([
      "SchedulerValidationError",
      "DateTimeValidationError",
      "EventValidationError",
      "EventImageError",
      "InteractionSessionError",
      "OccurrenceValidationError"
    ])
    const message = error instanceof SchedulerValidationError || userFacingErrors.has(error.name)
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
  buildRoundupModal,
  parseYesNo,
  handleEventSchedulerInteraction
}
