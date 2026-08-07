const {
  ActionRowBuilder,
  ButtonBuilder,
  ButtonStyle,
  ChannelSelectMenuBuilder,
  ChannelType,
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
const { handleEventCreationInteraction } = require("./eventCreationInteractions")
const { handleEventManagementInteraction } = require("./eventManagementInteractions")
const { handleAllianceManagementInteraction } = require("./allianceManagementInteractions")
const { handleEventSchedulerHelpInteraction } = require("./eventSchedulerHelp")
const { parseUtcTime } = require("./timeParsing")
const {
  CODE_TTL_MINUTES,
  generateStateLinkCode,
  hashStateLinkCode
} = require("./stateLinkCodes")

const IDS = Object.freeze({
  prefix: "es:",
  home: "es:home",
  identity: "es:identity",
  identityModal: "es:identitym",
  channels: "es:channels",
  reminderChannel: "es:reminderch",
  roundupChannel: "es:roundch",
  roundupModalPrefix: "es:roundm:",
  roundupOff: "es:roundoff",
  stateDestination: "es:statedest",
  stateDestinationChannel: "es:statedestch",
  stateCode: "es:statecode",
  stateSharing: "es:stateshare",
  stateLink: "es:statelink",
  stateLinkModal: "es:statelinkm",
  stateOff: "es:stateoff"
})

const TEXT_CHANNEL_TYPES = [ChannelType.GuildText, ChannelType.GuildAnnouncement]
const SAFE_MENTIONS = Object.freeze({ parse: [], repliedUser: false })

function isSchedulerInteraction(interaction) {
  return interaction.commandName === "event-scheduler"
    || interaction.commandName === "event-scheduler-help"
    || String(interaction.customId || "").startsWith(IDS.prefix)
    || String(interaction.customId || "").startsWith("ec:")
    || String(interaction.customId || "").startsWith("el:")
    || String(interaction.customId || "").startsWith("ep:")
    || String(interaction.customId || "").startsWith("mg:")
    || String(interaction.customId || "").startsWith("am:")
    || String(interaction.customId || "").startsWith("eh:")
}

function backButton() {
  return new ButtonBuilder()
    .setCustomId(IDS.home)
    .setLabel("Back")
    .setStyle(ButtonStyle.Secondary)
}

function channelSelect(customId, placeholder, defaultChannelId) {
  const menu = new ChannelSelectMenuBuilder()
    .setCustomId(customId)
    .setPlaceholder(placeholder)
    .setChannelTypes(TEXT_CHANNEL_TYPES)
    .setMinValues(1)
    .setMaxValues(1)
  if (defaultChannelId) menu.setDefaultChannels(defaultChannelId)
  return menu
}

function buildHomeView(settings, stateLink, stateDestination) {
  const allianceChannel = settings?.event_channel_id
    ? `<#${settings.event_channel_id}>`
    : "Not configured"
  const state = stateLink
    ? `${stateLink.state_guild_name || "Linked state Discord"} / ${stateLink.state_channel_name || "roundup channel"} (${stateLink.sharing_enabled ? "enabled" : "disabled"})`
    : "Not linked"
  const destination = stateDestination
    ? `<#${stateDestination.state_roundup_channel_id}> (${stateDestination.enabled ? "enabled" : "disabled"})`
    : "Not configured"
  const roundup = settings?.weekly_roundup_enabled
    ? `<#${settings.weekly_roundup_channel_id}> on weekday ${settings.weekly_roundup_day} at ${String(settings.weekly_roundup_time_utc).slice(0, 5)} UTC`
    : "Disabled"

  return {
    content:
      `Event scheduler\n\n` +
      `Main alliance: ${settings?.alliance_name || "Not configured"}\n` +
      `Alliances: ${settings?.alliance_count ?? 0}\n` +
      `Alliance reminder channel: ${allianceChannel}\n` +
      `Alliance weekly roundup: ${roundup}\n` +
      `State sharing: ${state}\n` +
      `This server as state destination: ${destination}`,
    components: [
      new ActionRowBuilder().addComponents(
        new ButtonBuilder().setCustomId("ec:new").setLabel("Create event")
          .setStyle(ButtonStyle.Success).setDisabled(!settings?.event_channel_id),
        new ButtonBuilder().setCustomId("el:0").setLabel("View events")
          .setStyle(ButtonStyle.Primary),
        new ButtonBuilder().setCustomId("am:open").setLabel("Alliances")
          .setStyle(ButtonStyle.Primary).setDisabled(!settings)
      ),
      new ActionRowBuilder().addComponents(
        new ButtonBuilder().setCustomId(IDS.identity).setLabel("Alliance identity")
          .setStyle(ButtonStyle.Primary),
        new ButtonBuilder().setCustomId(IDS.channels).setLabel("Configure channels")
          .setStyle(ButtonStyle.Primary).setDisabled(!settings)
      ),
      new ActionRowBuilder().addComponents(
        new ButtonBuilder().setCustomId(IDS.stateDestination).setLabel("State destination")
          .setStyle(ButtonStyle.Secondary),
        new ButtonBuilder().setCustomId(IDS.stateSharing).setLabel("State sharing")
          .setStyle(ButtonStyle.Secondary).setDisabled(!settings)
      ),
      new ActionRowBuilder().addComponents(
        new ButtonBuilder().setCustomId("eh:home").setLabel("Help")
          .setStyle(ButtonStyle.Secondary)
      )
    ],
    allowedMentions: SAFE_MENTIONS
  }
}

function buildAllianceIdentityModal(settings) {
  return new ModalBuilder()
    .setCustomId(IDS.identityModal)
    .setTitle(settings ? "Rename main alliance" : "Set main alliance")
    .addComponents(new ActionRowBuilder().addComponents(
      new TextInputBuilder()
        .setCustomId("a")
        .setLabel("Main alliance name")
        .setStyle(TextInputStyle.Short)
        .setRequired(true)
        .setMaxLength(100)
        .setValue(settings?.alliance_name || "")
    ))
}

function buildChannelConfigurationView(settings) {
  return {
    content:
      "Configure channels\n\n" +
      "Choose the alliance reminder channel. It must also allow Attach Files for event images.\n\n" +
      "Choose the alliance weekly-roundup channel to enable or update its UTC schedule.",
    components: [
      new ActionRowBuilder().addComponents(channelSelect(
        IDS.reminderChannel,
        "Select alliance reminder channel",
        settings?.event_channel_id
      )),
      new ActionRowBuilder().addComponents(channelSelect(
        IDS.roundupChannel,
        "Select alliance weekly-roundup channel",
        settings?.weekly_roundup_channel_id
      )),
      new ActionRowBuilder().addComponents(
        new ButtonBuilder().setCustomId(IDS.roundupOff).setLabel("Disable weekly roundup")
          .setStyle(ButtonStyle.Secondary).setDisabled(!settings?.weekly_roundup_enabled),
        backButton()
      )
    ],
    allowedMentions: SAFE_MENTIONS
  }
}

function buildRoundupModal(settings, channelId) {
  return new ModalBuilder()
    .setCustomId(`${IDS.roundupModalPrefix}${channelId}`)
    .setTitle("Alliance weekly roundup")
    .addComponents(
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
        new TextInputBuilder().setCustomId("x").setLabel("Post when no events? yes or no")
          .setStyle(TextInputStyle.Short).setRequired(true)
          .setValue(settings?.roundup_when_empty ? "yes" : "no")
      )
    )
}

function buildStateDestinationView(destination, generatedCode = null) {
  const status = destination
    ? `Current state roundup channel: <#${destination.state_roundup_channel_id}>`
    : "Select this server's state weekly-roundup channel."
  const codeText = generatedCode
    ? `\n\nState link code: **${generatedCode}**\nThis one-time ${CODE_TTL_MINUTES}-minute code is for an alliance administrator using the same bot/game profile.`
    : ""
  return {
    content: `State destination\n\n${status}${codeText}`,
    components: [
      new ActionRowBuilder().addComponents(channelSelect(
        IDS.stateDestinationChannel,
        "Select state weekly-roundup channel",
        destination?.state_roundup_channel_id
      )),
      new ActionRowBuilder().addComponents(
        new ButtonBuilder().setCustomId(IDS.stateCode).setLabel("Generate link code")
          .setStyle(ButtonStyle.Primary).setDisabled(!destination?.enabled),
        backButton()
      )
    ],
    allowedMentions: SAFE_MENTIONS
  }
}

function buildStateSharingView(stateLink) {
  const status = stateLink
    ? `${stateLink.state_guild_name || "Linked state Discord"}\nChannel: ${stateLink.state_channel_name || "Configured roundup channel"}\nProfile: ${stateLink.game_profile}\nSharing: ${stateLink.sharing_enabled ? "enabled" : "disabled"}`
    : "No state destination is linked. Obtain a short-lived code from the state Discord."
  return {
    content: `State sharing\n\n${status}`,
    components: [new ActionRowBuilder().addComponents(
      new ButtonBuilder().setCustomId(IDS.stateLink).setLabel("Link with code")
        .setStyle(ButtonStyle.Primary),
      new ButtonBuilder().setCustomId(IDS.stateOff).setLabel("Disable sharing")
        .setStyle(ButtonStyle.Secondary).setDisabled(!stateLink?.sharing_enabled),
      backButton()
    )],
    allowedMentions: SAFE_MENTIONS
  }
}

function buildStateLinkModal() {
  return new ModalBuilder()
    .setCustomId(IDS.stateLinkModal)
    .setTitle("Link state destination")
    .addComponents(new ActionRowBuilder().addComponents(
      new TextInputBuilder()
        .setCustomId("code")
        .setLabel("State link code")
        .setStyle(TextInputStyle.Short)
        .setRequired(true)
        .setMinLength(12)
        .setMaxLength(14)
        .setPlaceholder("ABCD-EFGH-JKLM")
    ))
}

function parseYesNo(value, label) {
  const normalized = String(value || "").trim().toLowerCase()
  if (normalized === "yes") return true
  if (normalized === "no") return false
  throw new SchedulerValidationError(`${label} must be yes or no.`)
}

async function describeStateLink(client, stateLink) {
  if (!stateLink) return null
  try {
    const guild = await client.guilds.fetch(stateLink.state_guild_id)
    const channel = await guild.channels.fetch(stateLink.state_event_channel_id)
    return {
      ...stateLink,
      state_guild_name: guild.name,
      state_channel_name: channel ? `#${channel.name}` : "Configured roundup channel"
    }
  } catch {
    return { ...stateLink }
  }
}

async function loadHome(repository, guildId, client) {
  const [settings, rawStateLink, stateDestination] = await Promise.all([
    repository.getGuildSettings(guildId),
    repository.getStateLink(guildId),
    repository.getStateDestination(guildId)
  ])
  const stateLink = await describeStateLink(client, rawStateLink)
  return buildHomeView(settings, stateLink, stateDestination)
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
  {
    userCanManageServer,
    healthProvider = getEventSchedulerHealth,
    repositoryProvider = health => createEventSchedulerRepository(getPool(), health.gameProfile)
  } = {}
) {
  if (!isSchedulerInteraction(interaction)) return false
  if (await handleEventSchedulerHelpInteraction(interaction)) return true

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

  const repository = repositoryProvider(health)
  const guildId = interaction.guildId

  try {
    if (await handleAllianceManagementInteraction(interaction, { repository, health })) return true
    if (await handleEventManagementInteraction(interaction, { repository, health })) return true
    if (await handleEventCreationInteraction(interaction, {
      repository,
      health,
      loadHome: (repo, id) => loadHome(repo, id, interaction.client)
    })) return true

    if (interaction.isChatInputCommand?.() && interaction.commandName === "event-scheduler") {
      await interaction.deferReply({ flags: MessageFlags.Ephemeral })
      await interaction.editReply(await loadHome(repository, guildId, interaction.client))
      return true
    }

    if (interaction.isButton?.() && interaction.customId === IDS.home) {
      await interaction.deferUpdate()
      await interaction.editReply(await loadHome(repository, guildId, interaction.client))
      return true
    }

    if (interaction.isButton?.() && interaction.customId === IDS.identity) {
      await interaction.showModal(buildAllianceIdentityModal(
        await repository.getGuildSettings(guildId)
      ))
      return true
    }

    if (interaction.isModalSubmit?.() && interaction.customId === IDS.identityModal) {
      await interaction.deferReply({ flags: MessageFlags.Ephemeral })
      const settings = await repository.getGuildSettings(guildId)
      const allianceName = normalizeAllianceName(interaction.fields.getTextInputValue("a"))
      if (settings) {
        await repository.renameAlliance({
          guildId,
          allianceId: settings.default_alliance_id,
          allianceName
        })
      } else {
        await repository.upsertGuildIdentity({
          guildId,
          botInstanceName: health.botInstanceName,
          allianceName
        })
      }
      await interaction.editReply(await loadHome(repository, guildId, interaction.client))
      return true
    }

    if (interaction.isButton?.() && interaction.customId === IDS.channels) {
      await interaction.deferUpdate()
      await interaction.editReply(buildChannelConfigurationView(
        await repository.getGuildSettings(guildId)
      ))
      return true
    }

    if (interaction.isChannelSelectMenu?.() && interaction.customId === IDS.reminderChannel) {
      await interaction.deferUpdate()
      const target = await resolveSendableChannel(interaction.client, guildId, interaction.values[0])
      if (!await repository.setEventChannel({ guildId, eventChannelId: target.channelId })) {
        throw new SchedulerValidationError("Set the main alliance identity first.")
      }
      await interaction.editReply(buildChannelConfigurationView(
        await repository.getGuildSettings(guildId)
      ))
      return true
    }

    if (interaction.isChannelSelectMenu?.() && interaction.customId === IDS.roundupChannel) {
      const settings = await repository.getGuildSettings(guildId)
      const target = await resolveSendableChannel(
        interaction.client,
        guildId,
        interaction.values[0],
        { requireAttachments: false }
      )
      await interaction.showModal(buildRoundupModal(settings, target.channelId))
      return true
    }

    if (interaction.isModalSubmit?.()
      && String(interaction.customId).startsWith(IDS.roundupModalPrefix)) {
      await interaction.deferReply({ flags: MessageFlags.Ephemeral })
      const channelId = String(interaction.customId).slice(IDS.roundupModalPrefix.length)
      await resolveSendableChannel(interaction.client, guildId, channelId, {
        requireAttachments: false
      })
      const weekday = Number(interaction.fields.getTextInputValue("d"))
      if (!Number.isInteger(weekday) || weekday < 0 || weekday > 6) {
        throw new SchedulerValidationError("Weekday must be an integer from 0 to 6.")
      }
      await repository.configureWeeklyRoundup({
        guildId,
        enabled: true,
        weekday,
        timeUtc: parseUtcTime(interaction.fields.getTextInputValue("t")),
        channelId,
        postWhenEmpty: parseYesNo(
          interaction.fields.getTextInputValue("x"),
          "Post when empty"
        )
      })
      await interaction.editReply(await loadHome(repository, guildId, interaction.client))
      return true
    }

    if (interaction.isButton?.() && interaction.customId === IDS.roundupOff) {
      await interaction.deferUpdate()
      const settings = await repository.getGuildSettings(guildId)
      await repository.configureWeeklyRoundup({
        guildId,
        enabled: false,
        weekday: settings.weekly_roundup_day,
        timeUtc: settings.weekly_roundup_time_utc,
        channelId: settings.weekly_roundup_channel_id,
        postWhenEmpty: settings.roundup_when_empty
      })
      await interaction.editReply(buildChannelConfigurationView(
        await repository.getGuildSettings(guildId)
      ))
      return true
    }

    if (interaction.isButton?.() && interaction.customId === IDS.stateDestination) {
      await interaction.deferUpdate()
      await interaction.editReply(buildStateDestinationView(
        await repository.getStateDestination(guildId)
      ))
      return true
    }

    if (interaction.isChannelSelectMenu?.()
      && interaction.customId === IDS.stateDestinationChannel) {
      await interaction.deferUpdate()
      const target = await resolveSendableChannel(
        interaction.client,
        guildId,
        interaction.values[0],
        { requireAttachments: false }
      )
      const destination = await repository.upsertStateDestination({
        stateGuildId: guildId,
        configuredByBotInstance: health.botInstanceName,
        stateRoundupChannelId: target.channelId
      })
      await interaction.editReply(buildStateDestinationView(destination))
      return true
    }

    if (interaction.isButton?.() && interaction.customId === IDS.stateCode) {
      await interaction.deferUpdate()
      const code = generateStateLinkCode()
      const destination = await repository.getStateDestination(guildId)
      const created = await repository.createStateLinkCode({
        stateGuildId: guildId,
        codeHash: hashStateLinkCode(code),
        createdByBotInstance: health.botInstanceName,
        createdByUserId: interaction.user.id,
        expiresAt: new Date(Date.now() + CODE_TTL_MINUTES * 60 * 1000)
      })
      if (!created) throw new SchedulerValidationError("Configure this state destination first.")
      await interaction.editReply(buildStateDestinationView(destination, code))
      return true
    }

    if (interaction.isButton?.() && interaction.customId === IDS.stateSharing) {
      await interaction.deferUpdate()
      const stateLink = await describeStateLink(
        interaction.client,
        await repository.getStateLink(guildId)
      )
      await interaction.editReply(buildStateSharingView(stateLink))
      return true
    }

    if (interaction.isButton?.() && interaction.customId === IDS.stateLink) {
      await interaction.showModal(buildStateLinkModal())
      return true
    }

    if (interaction.isModalSubmit?.() && interaction.customId === IDS.stateLinkModal) {
      await interaction.deferReply({ flags: MessageFlags.Ephemeral })
      const codeHash = hashStateLinkCode(interaction.fields.getTextInputValue("code"))
      if (!codeHash) throw new SchedulerValidationError("Enter a valid state link code.")
      const link = await repository.consumeStateLinkCode({
        allianceGuildId: guildId,
        configuredByBotInstance: health.botInstanceName,
        codeHash
      })
      if (!link) {
        throw new SchedulerValidationError(
          "That state link code is invalid, expired, already used, or belongs to another game profile."
        )
      }
      await interaction.editReply(buildStateSharingView(
        await describeStateLink(interaction.client, link)
      ))
      return true
    }

    if (interaction.isButton?.() && interaction.customId === IDS.stateOff) {
      await interaction.deferUpdate()
      await repository.setStateSharing(guildId, false)
      await interaction.editReply(buildStateSharingView(
        await describeStateLink(interaction.client, await repository.getStateLink(guildId))
      ))
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
  buildAllianceIdentityModal,
  buildChannelConfigurationView,
  buildRoundupModal,
  buildStateDestinationView,
  buildStateSharingView,
  buildStateLinkModal,
  parseYesNo,
  describeStateLink,
  loadHome,
  handleEventSchedulerInteraction
}
