const crypto = require("crypto")
const {
  ActionRowBuilder,
  ButtonBuilder,
  ButtonStyle,
  FileUploadBuilder,
  LabelBuilder,
  MessageFlags,
  ModalBuilder,
  StringSelectMenuBuilder,
  TextInputBuilder,
  TextInputStyle
} = require("discord.js")

const { parseIsoDate, parseTimeOrGroups } = require("./timeParsing")
const {
  EventValidationError,
  normalizeCustomMessage,
  validateEventDraft
} = require("./eventValidation")
const { downloadEventImage } = require("./eventImage")
const { InteractionSessionStore } = require("./interactionSessions")
const {
  EVENTS_PER_PAGE,
  formatEventPreview,
  formatEventListPage,
  formatUpcomingOccurrencePreview
} = require("./eventSchedulerFormatting")
const { getNextOccurrences } = require("./occurrenceCalculation")

const creationSessions = new InteractionSessionStore()
const OCCURRENCES_PER_PREVIEW = 5
const ALLIANCES_PER_PAGE = 25

const CREATION_IDS = Object.freeze({
  newEvent: "ec:new",
  corePrefix: "ec:m:",
  allianceSelectPrefix: "ec:as:",
  alliancePagePrefix: "ec:ap:",
  allianceChangePrefix: "ec:ac:",
  recurrencePrefix: "ec:r:",
  advancePrefix: "ec:a:",
  startPrefix: "ec:s:",
  messagesPrefix: "ec:cm:",
  messagesModalPrefix: "ec:cmm:",
  optionsPrefix: "ec:n:",
  alliancePrefix: "ec:pa:",
  statePrefix: "ec:ps:",
  roundupPrefix: "ec:pr:",
  previewPrefix: "ec:pv:",
  createPrefix: "ec:ok:",
  editPrefix: "ec:e:",
  cancelPrefix: "ec:x:",
  listPrefix: "el:",
  occurrencePreviewPrefix: "ep:"
})

function sessionContext(interaction, health) {
  return {
    userId: interaction.user.id,
    guildId: interaction.guildId,
    gameProfile: health.gameProfile
  }
}

function idSuffix(customId, prefix) {
  return String(customId).slice(prefix.length)
}

function textInput(
  customId,
  label,
  value,
  { paragraph = false, maximum = 100, required = true } = {}
) {
  const input = new TextInputBuilder()
    .setCustomId(customId)
    .setLabel(label)
    .setStyle(paragraph ? TextInputStyle.Paragraph : TextInputStyle.Short)
    .setRequired(required)
    .setMaxLength(maximum)
  if (value) input.setValue(value)
  return new ActionRowBuilder().addComponents(input)
}

function coreTimeValue(data) {
  if (data.groups?.length) {
    return data.groups.map(group => `${group.groupName} = ${group.eventTimeUtc}`).join("\n")
  }
  return data.eventTimeUtc || ""
}

function buildCoreModal(sessionId, data = {}) {
  const modal = new ModalBuilder()
    .setCustomId(`${CREATION_IDS.corePrefix}${sessionId}`)
    .setTitle(data.mode === "edit" ? "Edit scheduled event" : "Create scheduled event")
    .addComponents(
      textInput("e", "Event name", data.eventName),
      textInput("d", "First date (YYYY-MM-DD)", data.firstOccurrenceDate),
      textInput(
        "t",
        "UTC time or Group = UTC time per line",
        coreTimeValue(data),
        { paragraph: true, maximum: 1000 }
      )
    )

  modal.addLabelComponents(
    new LabelBuilder()
      .setLabel("Event image (optional)")
      .setDescription("PNG, JPEG, GIF or WebP; maximum 8 MB")
      .setFileUploadComponent(
        new FileUploadBuilder()
          .setCustomId("img")
          .setRequired(data.imageAction === "replace")
          .setMinValues(data.imageAction === "replace" ? 1 : 0)
          .setMaxValues(1)
      )
  )
  return modal
}

function buildMessagesModal(sessionId, data = {}) {
  return new ModalBuilder()
    .setCustomId(`${CREATION_IDS.messagesModalPrefix}${sessionId}`)
    .setTitle("Custom reminder messages")
    .addComponents(
      textInput("advance", "Advance message (optional)", data.advanceReminderMessage, {
        paragraph: true,
        maximum: 500,
        required: false
      }),
      textInput("final", "Final message (optional)", data.finalReminderMessage, {
        paragraph: true,
        maximum: 500,
        required: false
      })
    )
}

function opaqueToken() {
  return crypto.randomBytes(6).toString("base64url")
}

function buildAllianceSelectionView(sessionId, alliances, page, total, selectedAllianceId) {
  const totalPages = Math.max(1, Math.ceil(total / ALLIANCES_PER_PAGE))
  const tokenMap = {}
  const options = alliances.map(alliance => {
    const token = opaqueToken()
    tokenMap[token] = String(alliance.id)
    return {
      label: alliance.alliance_name.slice(0, 100),
      value: token,
      description: alliance.is_default ? "Main alliance" : "Sub-alliance",
      default: String(alliance.id) === String(selectedAllianceId || "")
    }
  })
  return {
    view: {
      content: `Select alliance or sub-alliance - page ${page + 1} of ${totalPages}`,
      components: [
        new ActionRowBuilder().addComponents(
          new StringSelectMenuBuilder()
            .setCustomId(`${CREATION_IDS.allianceSelectPrefix}${sessionId}`)
            .setPlaceholder("Alliance or sub-alliance")
            .addOptions(options)
        ),
        new ActionRowBuilder().addComponents(
          new ButtonBuilder()
            .setCustomId(`${CREATION_IDS.alliancePagePrefix}${sessionId}:${Math.max(0, page - 1)}`)
            .setLabel("Previous")
            .setStyle(ButtonStyle.Secondary)
            .setDisabled(page === 0),
          new ButtonBuilder()
            .setCustomId(`${CREATION_IDS.alliancePagePrefix}${sessionId}:${page + 1}`)
            .setLabel("Next")
            .setStyle(ButtonStyle.Secondary)
            .setDisabled(page >= totalPages - 1),
          new ButtonBuilder()
            .setCustomId(`${CREATION_IDS.cancelPrefix}${sessionId}`)
            .setLabel("Cancel")
            .setStyle(ButtonStyle.Secondary)
        )
      ]
    },
    tokenMap
  }
}

async function renderAllianceSelection(
  interaction,
  repository,
  sessionStore,
  sessionId,
  context,
  page = 0
) {
  const safePage = Math.max(0, Number(page) || 0)
  const result = await repository.listAlliances(interaction.guildId, {
    limit: ALLIANCES_PER_PAGE,
    offset: safePage * ALLIANCES_PER_PAGE
  })
  if (!result.alliances.length) {
    throw new EventValidationError("Add an alliance before creating an event.")
  }
  const session = sessionStore.get(sessionId, context)
  const built = buildAllianceSelectionView(
    sessionId,
    result.alliances,
    safePage,
    result.total,
    session.data.allianceId
  )
  sessionStore.update(sessionId, context, {
    allianceSelectionPage: safePage,
    allianceTokenMap: built.tokenMap
  })
  if (interaction.deferred || interaction.replied) await interaction.editReply(built.view)
  else await interaction.reply({ ...built.view, flags: MessageFlags.Ephemeral })
}

function selectOption(label, value, selected) {
  return { label, value, default: String(selected) === value }
}

function buildTimingView(sessionId, data) {
  return {
    content:
      `Event options\n\n` +
      `Recurrence: every ${data.recurrenceDays} days\n` +
      `Advance reminder: ${data.advanceReminderMinutes ?? "none"}\n` +
      `Advance custom message: ${data.advanceReminderMessage ? "Yes" : "No"}\n` +
      `Final announcement (1 minute before): ${data.reminderAtStart ? "Yes" : "No"}\n` +
      `Final custom message: ${data.finalReminderMessage ? "Yes" : "No"}`,
    components: [
      new ActionRowBuilder().addComponents(
        new StringSelectMenuBuilder()
          .setCustomId(`${CREATION_IDS.recurrencePrefix}${sessionId}`)
          .setPlaceholder("Recurrence")
          .addOptions(
            selectOption("Every 3 days", "3", data.recurrenceDays),
            selectOption("Every week", "7", data.recurrenceDays),
            selectOption("Every 2 weeks", "14", data.recurrenceDays),
            selectOption("Every 4 weeks", "28", data.recurrenceDays)
          )
      ),
      new ActionRowBuilder().addComponents(
        new StringSelectMenuBuilder()
          .setCustomId(`${CREATION_IDS.advancePrefix}${sessionId}`)
          .setPlaceholder("Advance reminder")
          .addOptions(
            selectOption("No advance reminder", "none", data.advanceReminderMinutes ?? "none"),
            selectOption("5 minutes before", "5", data.advanceReminderMinutes),
            selectOption("10 minutes before", "10", data.advanceReminderMinutes),
            selectOption("15 minutes before", "15", data.advanceReminderMinutes),
            selectOption("20 minutes before", "20", data.advanceReminderMinutes),
            selectOption("30 minutes before", "30", data.advanceReminderMinutes)
          )
      ),
      new ActionRowBuilder().addComponents(
        new ButtonBuilder()
          .setCustomId(`${CREATION_IDS.startPrefix}${sessionId}`)
          .setLabel(data.reminderAtStart ? "Final: On" : "Final: Off")
          .setStyle(data.reminderAtStart ? ButtonStyle.Success : ButtonStyle.Secondary),
        new ButtonBuilder()
          .setCustomId(`${CREATION_IDS.messagesPrefix}${sessionId}`)
          .setLabel("Messages")
          .setStyle(ButtonStyle.Secondary),
        new ButtonBuilder()
          .setCustomId(`${CREATION_IDS.optionsPrefix}${sessionId}`)
          .setLabel("Publishing options")
          .setStyle(ButtonStyle.Primary),
        new ButtonBuilder()
          .setCustomId(`${CREATION_IDS.cancelPrefix}${sessionId}`)
          .setLabel("Cancel")
          .setStyle(ButtonStyle.Secondary)
      )
    ]
  }
}

function buildPublishingView(sessionId, data) {
  return {
    content:
      `Publishing options\n\n` +
      `Alliance reminders: ${data.publishToAlliance ? "Yes" : "No"}\n` +
      `Alliance weekly roundup: ${data.includeInWeeklyRoundup ? "Yes" : "No"}\n` +
      `State-roundup eligibility: ${data.publishToState ? "Yes" : "No"}\n` +
      `State weekly roundup: ${data.includeInWeeklyRoundup && data.publishToState ? "Yes" : "No"}`,
    components: [
      new ActionRowBuilder().addComponents(
        new ButtonBuilder()
          .setCustomId(`${CREATION_IDS.alliancePrefix}${sessionId}`)
          .setLabel(data.publishToAlliance ? "Alliance: On" : "Alliance: Off")
          .setStyle(data.publishToAlliance ? ButtonStyle.Success : ButtonStyle.Secondary),
        new ButtonBuilder()
          .setCustomId(`${CREATION_IDS.statePrefix}${sessionId}`)
          .setLabel(data.publishToState ? "State eligible: On" : "State eligible: Off")
          .setStyle(data.publishToState ? ButtonStyle.Success : ButtonStyle.Secondary),
        new ButtonBuilder()
          .setCustomId(`${CREATION_IDS.roundupPrefix}${sessionId}`)
          .setLabel(data.includeInWeeklyRoundup ? "Roundup: On" : "Roundup: Off")
          .setStyle(data.includeInWeeklyRoundup ? ButtonStyle.Success : ButtonStyle.Secondary)
      ),
      new ActionRowBuilder().addComponents(
        new ButtonBuilder()
          .setCustomId(`${CREATION_IDS.previewPrefix}${sessionId}`)
          .setLabel("Preview")
          .setStyle(ButtonStyle.Primary),
        new ButtonBuilder()
          .setCustomId(`${CREATION_IDS.cancelPrefix}${sessionId}`)
          .setLabel("Cancel")
          .setStyle(ButtonStyle.Secondary)
      )
    ]
  }
}

function buildPreviewView(sessionId, data) {
  return {
    content: formatEventPreview(data),
    allowedMentions: { parse: [], repliedUser: false },
    components: [
      new ActionRowBuilder().addComponents(
        new ButtonBuilder()
          .setCustomId(`${CREATION_IDS.createPrefix}${sessionId}`)
          .setLabel(data.mode === "edit" ? "Save changes" : "Create")
          .setStyle(ButtonStyle.Success),
        new ButtonBuilder()
          .setCustomId(`${CREATION_IDS.editPrefix}${sessionId}`)
          .setLabel("Edit details")
          .setStyle(ButtonStyle.Secondary),
        new ButtonBuilder()
          .setCustomId(`${CREATION_IDS.allianceChangePrefix}${sessionId}`)
          .setLabel("Change alliance")
          .setStyle(ButtonStyle.Secondary),
        new ButtonBuilder()
          .setCustomId(`${CREATION_IDS.cancelPrefix}${sessionId}`)
          .setLabel("Cancel")
          .setStyle(ButtonStyle.Secondary)
      )
    ]
  }
}

function buildListView(events, page, total) {
  const totalPages = Math.max(1, Math.ceil(total / EVENTS_PER_PAGE))
  const components = []
  if (events.length > 0) {
    components.push(
      new ActionRowBuilder().addComponents(
        ...events.map((event, index) =>
          new ButtonBuilder()
            .setCustomId(`${CREATION_IDS.occurrencePreviewPrefix}${event.id}:${page}`)
            .setLabel(`Preview ${index + 1}`)
            .setStyle(ButtonStyle.Primary)
        )
      )
    )
  }
  components.push(
    new ActionRowBuilder().addComponents(
      new ButtonBuilder()
        .setCustomId(`${CREATION_IDS.listPrefix}${Math.max(0, page - 1)}`)
        .setLabel("Previous")
        .setStyle(ButtonStyle.Secondary)
        .setDisabled(page === 0),
      new ButtonBuilder()
        .setCustomId(`${CREATION_IDS.listPrefix}${page + 1}`)
        .setLabel("Next")
        .setStyle(ButtonStyle.Secondary)
        .setDisabled(page >= totalPages - 1),
      new ButtonBuilder()
        .setCustomId("es:home")
        .setLabel("Back")
        .setStyle(ButtonStyle.Secondary)
    )
  )
  return {
    content: formatEventListPage(events, page, total),
    components
  }
}

function buildOccurrencePreviewView(event, occurrences, page) {
  return {
    content: formatUpcomingOccurrencePreview(event, occurrences),
    allowedMentions: { parse: [], repliedUser: false },
    components: [
      new ActionRowBuilder().addComponents(
        new ButtonBuilder()
          .setCustomId(`${CREATION_IDS.listPrefix}${page}`)
          .setLabel("Back to events")
          .setStyle(ButtonStyle.Secondary)
      )
    ]
  }
}

async function stateLinkIsEnabled(repository, guildId) {
  const link = await repository.getStateLink(guildId)
  return Boolean(link?.sharing_enabled)
}

async function validateSessionEvent(repository, guildId, draft) {
  const alliance = await repository.getAlliance(guildId, draft.allianceId)
  if (!alliance) throw new EventValidationError("Select a valid alliance or sub-alliance.")
  return validateEventDraft({
    ...draft,
    allianceId: String(alliance.id),
    allianceName: alliance.alliance_name
  }, {
    stateLinkEnabled: await stateLinkIsEnabled(repository, guildId)
  })
}

async function showList(interaction, repository, page) {
  const safePage = Math.max(0, Number.isInteger(page) ? page : 0)
  const result = await repository.listEvents(interaction.guildId, {
    limit: EVENTS_PER_PAGE,
    offset: safePage * EVENTS_PER_PAGE
  })
  const view = buildListView(result.events, safePage, result.total)
  if (interaction.deferred || interaction.replied) await interaction.editReply(view)
  else await interaction.reply({ ...view, flags: MessageFlags.Ephemeral })
}

async function handleEventCreationInteraction(
  interaction,
  { repository, health, loadHome, sessionStore = creationSessions }
) {
  const customId = String(interaction.customId || "")
  const relevant = customId.startsWith("ec:")
    || customId.startsWith("el:")
    || customId.startsWith("ep:")
  if (!relevant) return false

  const context = sessionContext(interaction, health)

  if (interaction.isButton?.() && customId === CREATION_IDS.newEvent) {
    const settings = await repository.getGuildSettings(interaction.guildId)
    if (!settings) {
      await interaction.reply({
        content: "Configure the alliance event channel before creating events.",
        flags: MessageFlags.Ephemeral
      })
      return true
    }
    const sessionId = sessionStore.create(context, {
      allianceId: null,
      allianceName: null,
      recurrenceDays: 7,
      advanceReminderMinutes: null,
      advanceReminderMessage: null,
      reminderAtStart: false,
      finalReminderMessage: null,
      publishToAlliance: true,
      publishToState: false,
      includeInWeeklyRoundup: false,
      groups: []
    })
    await renderAllianceSelection(
      interaction,
      repository,
      sessionStore,
      sessionId,
      context,
      0
    )
    return true
  }

  if (interaction.isButton?.() && customId.startsWith(CREATION_IDS.listPrefix)) {
    await interaction.deferUpdate()
    const page = Number(idSuffix(customId, CREATION_IDS.listPrefix))
    await showList(interaction, repository, page)
    return true
  }

  if (
    interaction.isButton?.()
    && customId.startsWith(CREATION_IDS.occurrencePreviewPrefix)
  ) {
    await interaction.deferUpdate()
    const [eventId, pageValue] = idSuffix(
      customId,
      CREATION_IDS.occurrencePreviewPrefix
    ).split(":")
    const page = Math.max(0, Number(pageValue) || 0)
    const event = await repository.getEvent(interaction.guildId, eventId)
    if (!event) {
      await interaction.editReply({
        content: "That event is no longer available for preview.",
        components: []
      })
      return true
    }
    const occurrences = getNextOccurrences(event, new Date(), OCCURRENCES_PER_PREVIEW)
    await interaction.editReply(buildOccurrencePreviewView(event, occurrences, page))
    return true
  }

  if (interaction.isButton?.() && customId.startsWith(CREATION_IDS.alliancePagePrefix)) {
    const [sessionId, pageValue] = idSuffix(
      customId,
      CREATION_IDS.alliancePagePrefix
    ).split(":")
    sessionStore.get(sessionId, context)
    await interaction.deferUpdate()
    await renderAllianceSelection(
      interaction,
      repository,
      sessionStore,
      sessionId,
      context,
      Number(pageValue)
    )
    return true
  }

  if (
    interaction.isStringSelectMenu?.()
    && customId.startsWith(CREATION_IDS.allianceSelectPrefix)
  ) {
    const sessionId = idSuffix(customId, CREATION_IDS.allianceSelectPrefix)
    const session = sessionStore.get(sessionId, context)
    const allianceId = session.data.allianceTokenMap?.[interaction.values[0]]
    if (!allianceId) throw new EventValidationError("That alliance selection has expired.")
    const alliance = await repository.getAlliance(interaction.guildId, allianceId)
    if (!alliance) throw new EventValidationError("That alliance is no longer available.")
    const updated = sessionStore.update(sessionId, context, {
      allianceId: String(alliance.id),
      allianceName: alliance.alliance_name,
      allianceTokenMap: null
    })
    await interaction.showModal(buildCoreModal(sessionId, updated.data))
    return true
  }

  if (
    interaction.isModalSubmit?.()
    && customId.startsWith(CREATION_IDS.messagesModalPrefix)
  ) {
    const sessionId = idSuffix(customId, CREATION_IDS.messagesModalPrefix)
    sessionStore.get(sessionId, context)
    await interaction.deferReply({ flags: MessageFlags.Ephemeral })
    const session = sessionStore.update(sessionId, context, {
      advanceReminderMessage: normalizeCustomMessage(
        interaction.fields.getTextInputValue("advance"),
        "Advance reminder message"
      ),
      finalReminderMessage: normalizeCustomMessage(
        interaction.fields.getTextInputValue("final"),
        "Final announcement message"
      )
    })
    await interaction.editReply(buildTimingView(sessionId, session.data))
    return true
  }

  if (interaction.isModalSubmit?.() && customId.startsWith(CREATION_IDS.corePrefix)) {
    const sessionId = idSuffix(customId, CREATION_IDS.corePrefix)
    sessionStore.get(sessionId, context)
    await interaction.deferReply({ flags: MessageFlags.Ephemeral })

    const parsedDate = parseIsoDate(interaction.fields.getTextInputValue("d"))
    const timeDetails = parseTimeOrGroups(interaction.fields.getTextInputValue("t"))
    let image
    const uploads = interaction.fields.getUploadedFiles("img")
    const attachment = uploads?.first?.()
    if (attachment) image = await downloadEventImage(attachment)

    const previous = sessionStore.get(sessionId, context).data
    if (previous.mode === "edit" && previous.imageAction === "replace" && !image) {
      throw new EventValidationError("Select one replacement image.")
    }
    const nextImage = previous.mode === "edit"
      ? (previous.imageAction === "replace" ? image : previous.imageAction === "remove" ? null : previous.image)
      : (image || previous.image || null)
    const session = sessionStore.update(sessionId, context, {
      eventName: interaction.fields.getTextInputValue("e").trim(),
      firstOccurrenceDate: parsedDate.value,
      firstDateIsPast: parsedDate.isPast,
      eventTimeUtc: timeDetails.eventTimeUtc,
      groups: timeDetails.groups,
      grouped: timeDetails.groups.length > 0,
      image: nextImage
    })
    await interaction.editReply(buildTimingView(sessionId, session.data))
    return true
  }

  const prefixes = Object.values(CREATION_IDS)
    .filter(value => value.endsWith(":"))
  const prefix = prefixes.find(value => customId.startsWith(value))
  if (!prefix) return false
  const sessionId = idSuffix(customId, prefix)
  const session = sessionStore.get(sessionId, context)

  if (interaction.isStringSelectMenu?.() && prefix === CREATION_IDS.recurrencePrefix) {
    sessionStore.update(sessionId, context, { recurrenceDays: Number(interaction.values[0]) })
    await interaction.deferUpdate()
    await interaction.editReply(buildTimingView(sessionId, session.data))
    return true
  }
  if (interaction.isStringSelectMenu?.() && prefix === CREATION_IDS.advancePrefix) {
    const value = interaction.values[0]
    sessionStore.update(sessionId, context, {
      advanceReminderMinutes: value === "none" ? null : Number(value)
    })
    await interaction.deferUpdate()
    await interaction.editReply(buildTimingView(sessionId, session.data))
    return true
  }

  if (interaction.isButton?.() && prefix === CREATION_IDS.startPrefix) {
    sessionStore.update(sessionId, context, { reminderAtStart: !session.data.reminderAtStart })
    await interaction.deferUpdate()
    await interaction.editReply(buildTimingView(sessionId, session.data))
    return true
  }
  if (interaction.isButton?.() && prefix === CREATION_IDS.messagesPrefix) {
    await interaction.showModal(buildMessagesModal(sessionId, session.data))
    return true
  }
  if (interaction.isButton?.() && prefix === CREATION_IDS.optionsPrefix) {
    await interaction.deferUpdate()
    await interaction.editReply(buildPublishingView(sessionId, session.data))
    return true
  }
  if (interaction.isButton?.() && prefix === CREATION_IDS.alliancePrefix) {
    sessionStore.update(sessionId, context, { publishToAlliance: !session.data.publishToAlliance })
    await interaction.deferUpdate()
    await interaction.editReply(buildPublishingView(sessionId, session.data))
    return true
  }
  if (interaction.isButton?.() && prefix === CREATION_IDS.statePrefix) {
    if (!session.data.publishToState && !(await stateLinkIsEnabled(repository, interaction.guildId))) {
      await interaction.reply({
        content: "Enable a valid state roundup link before including this event in state roundups.",
        flags: MessageFlags.Ephemeral
      })
      return true
    }
    sessionStore.update(sessionId, context, { publishToState: !session.data.publishToState })
    await interaction.deferUpdate()
    await interaction.editReply(buildPublishingView(sessionId, session.data))
    return true
  }
  if (interaction.isButton?.() && prefix === CREATION_IDS.roundupPrefix) {
    sessionStore.update(sessionId, context, {
      includeInWeeklyRoundup: !session.data.includeInWeeklyRoundup
    })
    await interaction.deferUpdate()
    await interaction.editReply(buildPublishingView(sessionId, session.data))
    return true
  }
  if (interaction.isButton?.() && prefix === CREATION_IDS.previewPrefix) {
    const event = await validateSessionEvent(repository, interaction.guildId, session.data)
    sessionStore.update(sessionId, context, event)
    await interaction.deferUpdate()
    await interaction.editReply(buildPreviewView(sessionId, event))
    return true
  }
  if (interaction.isButton?.() && prefix === CREATION_IDS.editPrefix) {
    await interaction.showModal(buildCoreModal(sessionId, session.data))
    return true
  }
  if (interaction.isButton?.() && prefix === CREATION_IDS.allianceChangePrefix) {
    await interaction.deferUpdate()
    await renderAllianceSelection(
      interaction,
      repository,
      sessionStore,
      sessionId,
      context,
      0
    )
    return true
  }
  if (interaction.isButton?.() && prefix === CREATION_IDS.cancelPrefix) {
    sessionStore.cancel(sessionId, context)
    await interaction.deferUpdate()
    await interaction.editReply(await loadHome(repository, interaction.guildId))
    return true
  }
  if (interaction.isButton?.() && prefix === CREATION_IDS.createPrefix) {
    await interaction.deferUpdate()
    if (session.data.creationInProgress) {
      await interaction.editReply({
        content: "This event is already being created.",
        components: []
      })
      return true
    }
    sessionStore.update(sessionId, context, { creationInProgress: true })
    try {
      const event = await validateSessionEvent(
        repository,
        interaction.guildId,
        session.data
      )
      if (session.data.mode === "edit") {
        const updated = await repository.updateEvent({
          guildId: interaction.guildId,
          eventId: session.data.eventId,
          event,
          imageAction: session.data.imageAction
        })
        if (!updated) throw new Error("That event is no longer available.")
      } else {
        await repository.createEvent({
          guildId: interaction.guildId,
          createdByUserId: interaction.user.id,
          createdByBotInstance: health.botInstanceName,
          event
        })
      }
      sessionStore.complete(sessionId, context)
      await interaction.editReply({
        content: `${session.data.mode === "edit" ? "Updated" : "Created"} ${event.eventName}. No public message was posted.`,
        components: []
      })
    } catch (error) {
      try {
        sessionStore.update(sessionId, context, { creationInProgress: false })
      } catch {
        // The original database or validation error is more useful than an expired session.
      }
      throw error
    }
    return true
  }

  return false
}

module.exports = {
  CREATION_IDS,
  ALLIANCES_PER_PAGE,
  creationSessions,
  sessionContext,
  buildCoreModal,
  buildMessagesModal,
  buildAllianceSelectionView,
  renderAllianceSelection,
  buildTimingView,
  buildPublishingView,
  buildPreviewView,
  buildListView,
  buildOccurrencePreviewView,
  handleEventCreationInteraction
}
