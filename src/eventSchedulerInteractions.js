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

const { getPool } = require("./db")
const { getEventSchedulerHealth } = require("./eventSchedulerHealth")
const { createEventSchedulerRepository } = require("./eventSchedulerRepository")
const {
  SchedulerValidationError,
  normalizeAllianceName,
  resolveSendableChannel
} = require("./eventSchedulerService")
const {
  handleEventCreationModalOpeningInteraction,
  handleEventCreationInteraction
} = require("./eventCreationInteractions")
const {
  handleEventManagementModalOpeningInteraction,
  handleEventManagementInteraction
} = require("./eventManagementInteractions")
const {
  handleAllianceManagementModalOpeningInteraction,
  handleAllianceManagementInteraction
} = require("./allianceManagementInteractions")
const { handleEventSchedulerHelpInteraction } = require("./eventSchedulerHelp")
const {
  acknowledgeSchedulerInteraction,
  isExpectedInteractionResponseError,
  logDiscordApiError,
  logExpectedInteractionResponseError,
  safelyRespondToInteraction
} = require("./interactionResponses")
const { parseUtcTime } = require("./timeParsing")
const { formatWeeklyRoundup } = require("./weeklyRoundupFormatting")
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
  roundupSettings: "es:roundset",
  roundupEnable: "es:roundon",
  roundupDisable: "es:roundoff",
  roundupSchedule: "es:roundsched",
  roundupDayPrefix: "es:roundday:",
  roundupContinuePrefix: "es:roundcontinue:",
  roundupTimeModalPrefix: "es:roundtime:",
  roundupScheduleConfirmPrefix: "es:roundconfirm:",
  roundupChannelView: "es:roundchannel",
  roundupChannel: "es:roundch",
  roundupPreview: "es:roundpreview",
  roundupTest: "es:roundtest",
  roundupTestConfirm: "es:roundtestconfirm",
  stateRoundupToggle: "es:stateroundtoggle",
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
const WEEKDAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
const guildSettingsCache = new Map()

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
          .setStyle(ButtonStyle.Primary).setDisabled(!settings),
        new ButtonBuilder().setCustomId(IDS.roundupSettings).setLabel("Weekly roundup settings")
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
      "Choose the alliance reminder channel. It must also allow Attach Files for event images.",
    components: [
      new ActionRowBuilder().addComponents(channelSelect(
        IDS.reminderChannel,
        "Select alliance reminder channel",
        settings?.event_channel_id
      )),
      new ActionRowBuilder().addComponents(
        backButton()
      )
    ],
    allowedMentions: SAFE_MENTIONS
  }
}

function buildRoundupTimeModal(day, settings, currentSchedule = null) {
  const currentSuffix = currentSchedule
    ? `:${currentSchedule.weekly_roundup_day}:${String(currentSchedule.weekly_roundup_time_utc)
      .slice(0, 5).replace(":", "")}`
    : ""
  return new ModalBuilder()
    .setCustomId(`${IDS.roundupTimeModalPrefix}${day}${currentSuffix}`)
    .setTitle("Weekly roundup schedule")
    .addComponents(
      new ActionRowBuilder().addComponents(
        new TextInputBuilder().setCustomId("t").setLabel("UTC time")
          .setStyle(TextInputStyle.Short).setRequired(true)
          .setValue(String(settings?.weekly_roundup_time_utc || "09:00").slice(0, 5))
      )
    )
}

function roundupScheduleLabel(settings) {
  const day = WEEKDAYS[Number(settings?.weekly_roundup_day ?? 1)] || "Monday"
  const time = String(settings?.weekly_roundup_time_utc || "09:00").slice(0, 5)
  return `${day} at ${time} UTC`
}

function buildRoundupSettingsView(settings) {
  const channel = settings?.weekly_roundup_channel_id
    ? `<#${settings.weekly_roundup_channel_id}>`
    : "Not configured"
  return {
    content:
      `Weekly roundup settings\n\n` +
      `Alliance weekly roundup: ${settings?.weekly_roundup_enabled ? "Enabled" : "Disabled"}\n` +
      `State publishing: ${settings?.state_roundup_enabled ? "Enabled" : "Disabled"}\n` +
      `Schedule: ${roundupScheduleLabel(settings)}\n` +
      `Channel: ${channel}\n\n` +
      "Alliance and state publishing are enabled independently. They use the displayed UTC schedule.",
    components: [
      new ActionRowBuilder().addComponents(
        new ButtonBuilder().setCustomId(IDS.roundupEnable).setLabel("Enable alliance roundup")
          .setStyle(ButtonStyle.Success).setDisabled(Boolean(settings?.weekly_roundup_enabled)),
        new ButtonBuilder().setCustomId(IDS.roundupDisable).setLabel("Disable alliance roundup")
          .setStyle(ButtonStyle.Secondary).setDisabled(!settings?.weekly_roundup_enabled),
        new ButtonBuilder().setCustomId(IDS.stateRoundupToggle)
          .setLabel(settings?.state_roundup_enabled ? "Disable state publishing" : "Enable state publishing")
          .setStyle(ButtonStyle.Secondary)
      ),
      new ActionRowBuilder().addComponents(
        new ButtonBuilder().setCustomId(IDS.roundupSchedule).setLabel("Change schedule")
          .setStyle(ButtonStyle.Primary),
        new ButtonBuilder().setCustomId(IDS.roundupChannelView).setLabel("Change channel")
          .setStyle(ButtonStyle.Primary),
        backButton()
      ),
      new ActionRowBuilder().addComponents(
        new ButtonBuilder().setCustomId(IDS.roundupPreview).setLabel("Preview roundup")
          .setStyle(ButtonStyle.Secondary),
        new ButtonBuilder().setCustomId(IDS.roundupTest).setLabel("Send test roundup")
          .setStyle(ButtonStyle.Secondary).setDisabled(!settings?.weekly_roundup_channel_id)
      )
    ],
    allowedMentions: SAFE_MENTIONS
  }
}

function buildRoundupDayView(settings, selectedDay = Number(settings?.weekly_roundup_day ?? 1)) {
  const currentDay = Number(settings?.weekly_roundup_day ?? 1)
  const currentTime = String(settings?.weekly_roundup_time_utc || "09:00")
    .slice(0, 5)
  const compactTime = currentTime.replace(":", "")
  return {
    content:
      `Change weekly roundup schedule\n\n` +
      `Current schedule: ${roundupScheduleLabel(settings)}\n` +
      `Selected weekday: ${WEEKDAYS[selectedDay]}\n\n` +
      "Select a UTC weekday or keep the current selection, then continue.",
    components: [
      new ActionRowBuilder().addComponents(
        new StringSelectMenuBuilder()
          .setCustomId(`${IDS.roundupDayPrefix}${currentDay}:${compactTime}`)
          .setPlaceholder("Select UTC weekday")
          .addOptions(WEEKDAYS.map((day, index) => ({
            label: day,
            value: String(index),
            default: selectedDay === index
          })))
      ),
      new ActionRowBuilder().addComponents(
        new ButtonBuilder()
          .setCustomId(`${IDS.roundupContinuePrefix}${selectedDay}:${currentDay}:${compactTime}`)
          .setLabel("Continue")
          .setStyle(ButtonStyle.Primary),
        new ButtonBuilder().setCustomId(IDS.roundupSettings).setLabel("Back")
          .setStyle(ButtonStyle.Secondary)
      )
    ]
  }
}

function buildRoundupTestConfirmation(settings) {
  return {
    content:
      `Send test weekly roundup\n\n` +
      `Destination: <#${settings.weekly_roundup_channel_id}>\n` +
      `Configured schedule: ${roundupScheduleLabel(settings)}\n\n` +
      "This sends a clearly marked test and does not affect scheduled-roundup history.",
    components: [new ActionRowBuilder().addComponents(
      new ButtonBuilder().setCustomId(IDS.roundupTestConfirm).setLabel("Confirm test send")
        .setStyle(ButtonStyle.Danger),
      new ButtonBuilder().setCustomId(IDS.roundupSettings).setLabel("Cancel")
        .setStyle(ButtonStyle.Secondary)
    )]
  }
}

function roundupPreviewResponse(payload, messages) {
  if (!messages.length) {
    return {
      content: "Roundup preview\n\nNo eligible events are in the current configured weekly window.",
      embeds: [],
      components: [new ActionRowBuilder().addComponents(
        new ButtonBuilder().setCustomId(IDS.roundupSettings).setLabel("Back")
          .setStyle(ButtonStyle.Secondary)
      )]
    }
  }
  return {
    content: "Roundup preview - nothing has been sent.",
    embeds: messages.map(message => message.embeds?.[0]).filter(Boolean).slice(0, 10),
    components: [new ActionRowBuilder().addComponents(
      new ButtonBuilder().setCustomId(IDS.roundupSettings).setLabel("Back")
        .setStyle(ButtonStyle.Secondary)
    )],
    allowedMentions: SAFE_MENTIONS
  }
}

function buildRoundupChannelView(settings) {
  return {
    content: "Change alliance weekly-roundup channel\n\nSelect a text channel. Roundups do not require Attach Files.",
    components: [
      new ActionRowBuilder().addComponents(channelSelect(
        IDS.roundupChannel,
        "Select alliance weekly-roundup channel",
        settings?.weekly_roundup_channel_id
      )),
      new ActionRowBuilder().addComponents(
        new ButtonBuilder().setCustomId(IDS.roundupSettings).setLabel("Back")
          .setStyle(ButtonStyle.Secondary)
      )
    ]
  }
}

function buildRoundupSchedulePreview(day, timeUtc, currentSchedule = null) {
  const normalized = String(timeUtc).replace(":", "")
  const current = currentSchedule
    ? `Current schedule: ${roundupScheduleLabel(currentSchedule)}\n`
    : ""
  return {
    content:
      `Weekly roundup schedule preview\n\n` +
      current +
      `Proposed schedule: ${WEEKDAYS[day]} at ${timeUtc} UTC\n\n` +
      "This takes effect from the next future scheduled roundup and will not replay an earlier roundup.",
    components: [new ActionRowBuilder().addComponents(
      new ButtonBuilder().setCustomId(`${IDS.roundupScheduleConfirmPrefix}${day}:${normalized}`)
        .setLabel("Save schedule").setStyle(ButtonStyle.Success),
      new ButtonBuilder().setCustomId(IDS.roundupSettings).setLabel("Cancel")
        .setStyle(ButtonStyle.Secondary)
    )]
  }
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
  if (settings) guildSettingsCache.set(String(guildId), settings)
  else guildSettingsCache.delete(String(guildId))
  const stateLink = await describeStateLink(client, rawStateLink)
  return buildHomeView(settings, stateLink, stateDestination)
}

async function replyUnavailable(interaction) {
  const payload = {
    content: "The event scheduler is temporarily unavailable. Existing bot features are unaffected.",
    flags: MessageFlags.Ephemeral
  }
  await safelyRespondToInteraction(interaction, payload)
}

async function handleSchedulerModalOpeningInteraction(interaction, { health }) {
  if (interaction.isButton?.() && interaction.customId === IDS.identity) {
    await interaction.showModal(buildAllianceIdentityModal(
      guildSettingsCache.get(String(interaction.guildId)) || null
    ))
    return true
  }
  if (interaction.isButton?.()
    && String(interaction.customId).startsWith(IDS.roundupContinuePrefix)) {
    const [dayValue, currentDayValue, compactTime = "0900"] = String(interaction.customId)
      .slice(IDS.roundupContinuePrefix.length)
      .split(":")
    const day = Number(dayValue)
    const currentDay = Number(currentDayValue)
    if (![day, currentDay].every(value => Number.isInteger(value) && value >= 0 && value <= 6)) {
      throw new SchedulerValidationError("Select a valid UTC weekday.")
    }
    await interaction.showModal(buildRoundupTimeModal(day, {
      weekly_roundup_time_utc: parseUtcTime(compactTime)
    }, {
      weekly_roundup_day: currentDay,
      weekly_roundup_time_utc: parseUtcTime(compactTime)
    }))
    return true
  }
  if (interaction.isButton?.() && interaction.customId === IDS.stateLink) {
    await interaction.showModal(buildStateLinkModal())
    return true
  }
  if (await handleAllianceManagementModalOpeningInteraction(interaction, { health })) return true
  if (await handleEventManagementModalOpeningInteraction(interaction, { health })) return true
  return handleEventCreationModalOpeningInteraction(interaction, { health })
}

async function handleEventSchedulerInteraction(
  interaction,
  {
    userCanManageServer,
    healthProvider = getEventSchedulerHealth,
    repositoryProvider = health => createEventSchedulerRepository(getPool(), health.gameProfile),
    roundupNow = () => new Date(),
    roundupFormatter = formatWeeklyRoundup,
    roundupTargetResolver = resolveSendableChannel,
    logger = console
  } = {}
) {
  if (!isSchedulerInteraction(interaction)) return false

  try {
    if (await handleEventSchedulerHelpInteraction(interaction)) return true

    const health = healthProvider()
    if (await handleSchedulerModalOpeningInteraction(interaction, { health })) return true

    await acknowledgeSchedulerInteraction(interaction)
    if (!health.available) {
      await replyUnavailable(interaction)
      return true
    }
    if (!(await userCanManageServer(interaction))) {
      await safelyRespondToInteraction(interaction, {
        content: "You do not have permission to manage the event scheduler.",
        flags: MessageFlags.Ephemeral
      }, { logger })
      return true
    }

    const repository = repositoryProvider(health)
    const guildId = interaction.guildId

    if (await handleAllianceManagementInteraction(interaction, { repository, health })) return true
    if (await handleEventManagementInteraction(interaction, { repository, health })) return true
    if (await handleEventCreationInteraction(interaction, {
      repository,
      health,
      loadHome: (repo, id) => loadHome(repo, id, interaction.client)
    })) return true

    if (interaction.isChatInputCommand?.() && interaction.commandName === "event-scheduler") {
      await acknowledgeSchedulerInteraction(interaction)
      await interaction.editReply(await loadHome(repository, guildId, interaction.client))
      return true
    }

    if (interaction.isButton?.() && interaction.customId === IDS.home) {
      await acknowledgeSchedulerInteraction(interaction)
      await interaction.editReply(await loadHome(repository, guildId, interaction.client))
      return true
    }

    if (interaction.isButton?.() && interaction.customId === IDS.identity) {
      return false
    }

    if (interaction.isModalSubmit?.() && interaction.customId === IDS.identityModal) {
      await acknowledgeSchedulerInteraction(interaction)
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
      await acknowledgeSchedulerInteraction(interaction)
      await interaction.editReply(buildChannelConfigurationView(
        await repository.getGuildSettings(guildId)
      ))
      return true
    }

    if (interaction.isChannelSelectMenu?.() && interaction.customId === IDS.reminderChannel) {
      await acknowledgeSchedulerInteraction(interaction)
      const target = await resolveSendableChannel(interaction.client, guildId, interaction.values[0])
      if (!await repository.setEventChannel({ guildId, eventChannelId: target.channelId })) {
        throw new SchedulerValidationError("Set the main alliance identity first.")
      }
      await interaction.editReply(buildChannelConfigurationView(
        await repository.getGuildSettings(guildId)
      ))
      return true
    }

    if (interaction.isButton?.() && interaction.customId === IDS.roundupSettings) {
      await acknowledgeSchedulerInteraction(interaction)
      await interaction.editReply(buildRoundupSettingsView(
        await repository.getGuildSettings(guildId)
      ))
      return true
    }

    if (interaction.isButton?.() && interaction.customId === IDS.roundupChannelView) {
      await acknowledgeSchedulerInteraction(interaction)
      await interaction.editReply(buildRoundupChannelView(
        await repository.getGuildSettings(guildId)
      ))
      return true
    }

    if (interaction.isChannelSelectMenu?.() && interaction.customId === IDS.roundupChannel) {
      await acknowledgeSchedulerInteraction(interaction)
      const settings = await repository.getGuildSettings(guildId)
      const target = await resolveSendableChannel(
        interaction.client,
        guildId,
        interaction.values[0],
        { requireAttachments: false }
      )
      await repository.configureWeeklyRoundup({
        guildId,
        enabled: settings.weekly_roundup_enabled,
        weekday: settings.weekly_roundup_day,
        timeUtc: settings.weekly_roundup_time_utc,
        channelId: target.channelId,
        postWhenEmpty: settings.roundup_when_empty,
        stateEnabled: settings.state_roundup_enabled
      })
      await interaction.editReply(buildRoundupSettingsView(
        await repository.getGuildSettings(guildId)
      ))
      return true
    }

    if (interaction.isButton?.()
      && [IDS.roundupEnable, IDS.roundupDisable, IDS.stateRoundupToggle]
        .includes(interaction.customId)) {
      await acknowledgeSchedulerInteraction(interaction)
      const settings = await repository.getGuildSettings(guildId)
      const enableAlliance = interaction.customId === IDS.roundupEnable
        ? true
        : interaction.customId === IDS.roundupDisable
          ? false
          : settings.weekly_roundup_enabled
      if (enableAlliance && !settings.weekly_roundup_channel_id) {
        throw new SchedulerValidationError("Choose the alliance roundup channel before enabling it.")
      }
      await repository.configureWeeklyRoundup({
        guildId,
        enabled: enableAlliance,
        weekday: settings.weekly_roundup_day,
        timeUtc: settings.weekly_roundup_time_utc,
        channelId: settings.weekly_roundup_channel_id,
        postWhenEmpty: settings.roundup_when_empty,
        stateEnabled: interaction.customId === IDS.stateRoundupToggle
          ? !settings.state_roundup_enabled
          : settings.state_roundup_enabled
      })
      await interaction.editReply(buildRoundupSettingsView(
        await repository.getGuildSettings(guildId)
      ))
      return true
    }

    if (interaction.isButton?.() && interaction.customId === IDS.roundupSchedule) {
      await acknowledgeSchedulerInteraction(interaction)
      await interaction.editReply(buildRoundupDayView(
        await repository.getGuildSettings(guildId)
      ))
      return true
    }

    if (interaction.isStringSelectMenu?.()
      && String(interaction.customId).startsWith(IDS.roundupDayPrefix)) {
      const [currentDayValue, compactTime = "0900"] = String(interaction.customId)
        .slice(IDS.roundupDayPrefix.length)
        .split(":")
      const currentDay = Number(currentDayValue)
      const selectedDay = Number(interaction.values?.[0])
      if (![currentDay, selectedDay].every(day => Number.isInteger(day) && day >= 0 && day <= 6)) {
        throw new SchedulerValidationError("Select a valid UTC weekday.")
      }
      await interaction.editReply(buildRoundupDayView({
        weekly_roundup_day: currentDay,
        weekly_roundup_time_utc: parseUtcTime(compactTime)
      }, selectedDay))
      return true
    }

    if (interaction.isModalSubmit?.()
      && String(interaction.customId).startsWith(IDS.roundupTimeModalPrefix)) {
      await acknowledgeSchedulerInteraction(interaction)
      const [dayValue, currentDayValue, compactTime] = String(interaction.customId)
        .slice(IDS.roundupTimeModalPrefix.length)
        .split(":")
      const day = Number(dayValue)
      if (!Number.isInteger(day) || day < 0 || day > 6) {
        throw new SchedulerValidationError("Select a valid UTC weekday.")
      }
      const timeUtc = parseUtcTime(interaction.fields.getTextInputValue("t"))
      const currentDay = Number(currentDayValue)
      const currentSchedule = Number.isInteger(currentDay) && compactTime
        ? {
            weekly_roundup_day: currentDay,
            weekly_roundup_time_utc: parseUtcTime(compactTime)
          }
        : null
      await interaction.editReply(buildRoundupSchedulePreview(day, timeUtc, currentSchedule))
      return true
    }

    if (interaction.isButton?.() && interaction.customId === IDS.roundupPreview) {
      const payload = await repository.getWeeklyRoundupPreview(guildId, { now: roundupNow() })
      if (!payload) throw new SchedulerValidationError("Configure the event scheduler first.")
      const messages = roundupFormatter(payload)
      await interaction.editReply(roundupPreviewResponse(payload, messages))
      return true
    }

    if (interaction.isButton?.() && interaction.customId === IDS.roundupTest) {
      const settings = await repository.getGuildSettings(guildId)
      if (!settings?.weekly_roundup_channel_id) {
        throw new SchedulerValidationError("Choose the alliance roundup channel first.")
      }
      await interaction.editReply(buildRoundupTestConfirmation(settings))
      return true
    }

    if (interaction.isButton?.() && interaction.customId === IDS.roundupTestConfirm) {
      const settings = await repository.getGuildSettings(guildId)
      if (!settings?.weekly_roundup_channel_id) {
        throw new SchedulerValidationError("Choose the alliance roundup channel first.")
      }
      const payload = await repository.getWeeklyRoundupPreview(guildId, { now: roundupNow() })
      if (!payload) throw new SchedulerValidationError("Configure the event scheduler first.")
      const messages = roundupFormatter({
        ...payload,
        claim: { ...payload.claim, isTest: true }
      })
      if (!messages.length) {
        throw new SchedulerValidationError(
          "No eligible events are in the current configured weekly window."
        )
      }
      const target = await roundupTargetResolver(
        interaction.client,
        guildId,
        settings.weekly_roundup_channel_id,
        { requireAttachments: false }
      )
      for (const message of messages) await target.channel.send(message)
      await interaction.editReply({
        content:
          `Sent ${messages.length} test roundup message${messages.length === 1 ? "" : "s"} ` +
          `to <#${settings.weekly_roundup_channel_id}>. Scheduled history was not changed.`,
        embeds: [],
        components: [new ActionRowBuilder().addComponents(
          new ButtonBuilder().setCustomId(IDS.roundupSettings).setLabel("Back")
            .setStyle(ButtonStyle.Secondary)
        )]
      })
      return true
    }

    if (interaction.isButton?.()
      && String(interaction.customId).startsWith(IDS.roundupScheduleConfirmPrefix)) {
      await acknowledgeSchedulerInteraction(interaction)
      const [dayValue, compactTime] = String(interaction.customId)
        .slice(IDS.roundupScheduleConfirmPrefix.length)
        .split(":")
      const day = Number(dayValue)
      if (!Number.isInteger(day) || day < 0 || day > 6) {
        throw new SchedulerValidationError("Select a valid UTC weekday.")
      }
      const timeUtc = parseUtcTime(compactTime)
      const settings = await repository.getGuildSettings(guildId)
      await repository.configureWeeklyRoundup({
        guildId,
        enabled: settings.weekly_roundup_enabled,
        weekday: day,
        timeUtc,
        channelId: settings.weekly_roundup_channel_id,
        postWhenEmpty: settings.roundup_when_empty,
        stateEnabled: settings.state_roundup_enabled
      })
      await interaction.editReply(buildRoundupSettingsView(
        await repository.getGuildSettings(guildId)
      ))
      return true
    }

    if (interaction.isButton?.() && interaction.customId === IDS.stateDestination) {
      await acknowledgeSchedulerInteraction(interaction)
      await interaction.editReply(buildStateDestinationView(
        await repository.getStateDestination(guildId)
      ))
      return true
    }

    if (interaction.isChannelSelectMenu?.()
      && interaction.customId === IDS.stateDestinationChannel) {
      await acknowledgeSchedulerInteraction(interaction)
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
      await acknowledgeSchedulerInteraction(interaction)
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
      await acknowledgeSchedulerInteraction(interaction)
      const stateLink = await describeStateLink(
        interaction.client,
        await repository.getStateLink(guildId)
      )
      await interaction.editReply(buildStateSharingView(stateLink))
      return true
    }

    if (interaction.isButton?.() && interaction.customId === IDS.stateLink) {
      return false
    }

    if (interaction.isModalSubmit?.() && interaction.customId === IDS.stateLinkModal) {
      await acknowledgeSchedulerInteraction(interaction)
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
      await acknowledgeSchedulerInteraction(interaction)
      await repository.setStateSharing(guildId, false)
      await interaction.editReply(buildStateSharingView(
        await describeStateLink(interaction.client, await repository.getStateLink(guildId))
      ))
      return true
    }
  } catch (error) {
    if (isExpectedInteractionResponseError(error)) {
      logExpectedInteractionResponseError(error, interaction, logger)
      return true
    }
    const userFacingErrors = new Set([
      "SchedulerValidationError",
      "DateTimeValidationError",
      "EventValidationError",
      "EventImageError",
      "InteractionSessionError",
      "OccurrenceValidationError"
    ])
    const isUserFacing = error instanceof SchedulerValidationError || userFacingErrors.has(error.name)
    if (!isUserFacing && !logDiscordApiError(error, interaction, logger)) {
      logger.error("[Event scheduler] Interaction failed:", error)
    }
    const message = isUserFacing
      ? error.message
      : "The event scheduler could not complete that action."
    try {
      await safelyRespondToInteraction(interaction, {
        content: message,
        flags: MessageFlags.Ephemeral,
        components: []
      }, { logger })
    } catch (responseError) {
      if (isExpectedInteractionResponseError(responseError)) {
        logExpectedInteractionResponseError(responseError, interaction, logger)
      } else {
        throw responseError
      }
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
  buildRoundupTimeModal,
  roundupScheduleLabel,
  buildRoundupSettingsView,
  buildRoundupDayView,
  buildRoundupChannelView,
  buildRoundupSchedulePreview,
  buildRoundupTestConfirmation,
  roundupPreviewResponse,
  buildStateDestinationView,
  buildStateSharingView,
  buildStateLinkModal,
  parseYesNo,
  describeStateLink,
  loadHome,
  handleSchedulerModalOpeningInteraction,
  handleEventSchedulerInteraction
}
