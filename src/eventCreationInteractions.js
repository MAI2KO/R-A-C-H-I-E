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

const { parseIsoDate, parseUtcTime } = require("./timeParsing")
const { acknowledgeSchedulerInteraction } = require("./interactionResponses")
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
  timingChoicePrefix: "ec:tc:",
  singleTimePrefix: "ec:ts:",
  singleTimeModalPrefix: "ec:tsm:",
  groupsPrefix: "ec:tg:",
  groupSelectPrefix: "ec:gs:",
  groupAddPrefix: "ec:ga:",
  groupEditPrefix: "ec:ge:",
  groupRemovePrefix: "ec:gr:",
  groupContinuePrefix: "ec:gc:",
  groupModalPrefix: "ec:gm:",
  imageManagePrefix: "ec:im:",
  imageKeepPrefix: "ec:ik:",
  imageReplacePrefix: "ec:ir:",
  imageRemovePrefix: "ec:id:",
  imageUploadPrefix: "ec:iu:",
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

function buildCoreModal(sessionId, data = {}) {
  return new ModalBuilder()
    .setCustomId(`${CREATION_IDS.corePrefix}${sessionId}`)
    .setTitle(data.mode === "edit" ? "Edit scheduled event" : "Create scheduled event")
    .addComponents(
      textInput("e", "Event name", data.eventName),
      textInput("d", "First date (YYYY-MM-DD)", data.firstOccurrenceDate)
    )
}

function buildTimingChoiceView(sessionId, data = {}) {
  const current = data.groups?.length
    ? `${data.groups.length} configured group${data.groups.length === 1 ? "" : "s"}`
    : data.eventTimeUtc
      ? `${data.eventTimeUtc} UTC`
      : "Not configured"
  return {
    content: `Event timing\n\nCurrent timing: ${current}\n\nChoose one event time or manage separate named groups.`,
    components: [new ActionRowBuilder().addComponents(
      new ButtonBuilder().setCustomId(`${CREATION_IDS.singleTimePrefix}${sessionId}`)
        .setLabel("Single time").setStyle(ButtonStyle.Primary),
      new ButtonBuilder().setCustomId(`${CREATION_IDS.groupsPrefix}${sessionId}`)
        .setLabel("Multiple groups").setStyle(ButtonStyle.Primary),
      new ButtonBuilder().setCustomId(`${CREATION_IDS.cancelPrefix}${sessionId}`)
        .setLabel("Cancel").setStyle(ButtonStyle.Secondary)
    )]
  }
}

function buildSingleTimeModal(sessionId, data = {}) {
  return new ModalBuilder()
    .setCustomId(`${CREATION_IDS.singleTimeModalPrefix}${sessionId}`)
    .setTitle("Single event time")
    .addComponents(textInput("t", "Event time (UTC)", data.eventTimeUtc, { maximum: 20 }))
}

function normalizeGroupInput(name, time, groups, editingIndex = null) {
  const groupName = String(name || "").trim()
  if (!groupName || groupName.length > 100) {
    throw new EventValidationError("Group name must be 1 to 100 characters.")
  }
  const duplicate = groups.some((group, index) =>
    index !== editingIndex && group.groupName.toLowerCase() === groupName.toLowerCase()
  )
  if (duplicate) throw new EventValidationError(`Duplicate group name: ${groupName}.`)
  return { groupName, eventTimeUtc: parseUtcTime(time) }
}

function upsertGroup(groups, group, editingIndex = null) {
  const updated = [...groups]
  if (editingIndex === null) {
    if (updated.length >= 20) throw new EventValidationError("An event can have at most 20 groups.")
    updated.push(group)
  } else {
    if (!Number.isInteger(editingIndex) || !updated[editingIndex]) {
      throw new EventValidationError("Select a group to edit.")
    }
    updated[editingIndex] = group
  }
  return updated.map((item, index) => ({ ...item, sortOrder: index }))
}

function removeGroup(groups, index) {
  if (!Number.isInteger(index) || !groups[index]) {
    throw new EventValidationError("Select a group to remove.")
  }
  return groups
    .filter((_, groupIndex) => groupIndex !== index)
    .map((group, groupIndex) => ({ ...group, sortOrder: groupIndex }))
}

function buildGroupModal(sessionId, mode, group = {}) {
  return new ModalBuilder()
    .setCustomId(`${CREATION_IDS.groupModalPrefix}${sessionId}:${mode}`)
    .setTitle(mode === "edit" ? "Edit event group" : "Add event group")
    .addComponents(
      textInput("name", "Group name", group.groupName, { maximum: 100 }),
      textInput("time", "Event time (UTC)", group.eventTimeUtc, { maximum: 20 })
    )
}

function buildGroupManagerView(sessionId, data = {}) {
  const groups = data.groups || []
  const selected = Number.isInteger(data.selectedGroupIndex) ? data.selectedGroupIndex : null
  const lines = groups.length
    ? groups.map((group, index) => `${index + 1}. ${group.groupName} - ${group.eventTimeUtc} UTC`).join("\n")
    : "No groups configured."
  const components = []
  if (groups.length) {
    components.push(new ActionRowBuilder().addComponents(
      new StringSelectMenuBuilder()
        .setCustomId(`${CREATION_IDS.groupSelectPrefix}${sessionId}`)
        .setPlaceholder("Select a group to edit or remove")
        .addOptions(groups.slice(0, 25).map((group, index) => ({
          label: group.groupName.slice(0, 100),
          description: `${group.eventTimeUtc} UTC`,
          value: String(index),
          default: selected === index
        })))
    ))
  }
  components.push(new ActionRowBuilder().addComponents(
    new ButtonBuilder().setCustomId(`${CREATION_IDS.groupAddPrefix}${sessionId}`)
      .setLabel("Add group").setStyle(ButtonStyle.Primary).setDisabled(groups.length >= 20),
    new ButtonBuilder().setCustomId(`${CREATION_IDS.groupEditPrefix}${sessionId}`)
      .setLabel("Edit group").setStyle(ButtonStyle.Secondary).setDisabled(selected === null),
    new ButtonBuilder().setCustomId(`${CREATION_IDS.groupRemovePrefix}${sessionId}`)
      .setLabel("Remove group").setStyle(ButtonStyle.Danger).setDisabled(selected === null),
    new ButtonBuilder().setCustomId(`${CREATION_IDS.groupContinuePrefix}${sessionId}`)
      .setLabel("Continue").setStyle(ButtonStyle.Success).setDisabled(groups.length === 0)
  ))
  return { content: `Groups\n\n${lines}`, components }
}

function buildImageChoiceView(sessionId, data = {}) {
  return {
    content: `Manage image\n\n${data.image ? "An event image is currently stored." : "No event image is stored."}`,
    components: [new ActionRowBuilder().addComponents(
      new ButtonBuilder().setCustomId(`${CREATION_IDS.imageKeepPrefix}${sessionId}`)
        .setLabel("Keep current image").setStyle(ButtonStyle.Primary),
      new ButtonBuilder().setCustomId(`${CREATION_IDS.imageReplacePrefix}${sessionId}`)
        .setLabel("Replace image").setStyle(ButtonStyle.Secondary),
      new ButtonBuilder().setCustomId(`${CREATION_IDS.imageRemovePrefix}${sessionId}`)
        .setLabel("Remove image").setStyle(ButtonStyle.Danger).setDisabled(!data.image)
    )]
  }
}

function buildImageUploadModal(sessionId) {
  return new ModalBuilder()
    .setCustomId(`${CREATION_IDS.imageUploadPrefix}${sessionId}`)
    .setTitle("Replace event image")
    .addLabelComponents(
      new LabelBuilder()
        .setLabel("Event image")
        .setDescription("PNG, JPEG, GIF or WebP; maximum 8 MB")
        .setFileUploadComponent(
          new FileUploadBuilder()
            .setCustomId("img")
            .setRequired(true)
            .setMinValues(1)
            .setMaxValues(1)
        )
      )
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
  const tokenNameMap = {}
  const options = alliances.map(alliance => {
    const token = opaqueToken()
    tokenMap[token] = String(alliance.id)
    tokenNameMap[token] = alliance.alliance_name
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
    tokenMap,
    tokenNameMap
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
    allianceTokenMap: built.tokenMap,
    allianceTokenNameMap: built.tokenNameMap
  })
  if (interaction.deferred || interaction.replied) await interaction.editReply(built.view)
  else await interaction.reply({ ...built.view, flags: MessageFlags.Ephemeral })
}

async function handleEventCreationModalOpeningInteraction(
  interaction,
  { health, sessionStore = creationSessions }
) {
  const customId = String(interaction.customId || "")
  if (!(customId.startsWith("ec:") || customId.startsWith("el:")
    || customId.startsWith("ep:"))) return false
  const context = sessionContext(interaction, health)

  if (interaction.isStringSelectMenu?.()
    && customId.startsWith(CREATION_IDS.allianceSelectPrefix)) {
    const sessionId = idSuffix(customId, CREATION_IDS.allianceSelectPrefix)
    const session = sessionStore.get(sessionId, context)
    if (session.data.allianceChangeOnly) return false
    const selectedToken = interaction.values?.[0]
    const allianceId = session.data.allianceTokenMap?.[selectedToken]
    const allianceName = session.data.allianceTokenNameMap?.[selectedToken]
    if (!allianceId || !allianceName) {
      throw new EventValidationError("That alliance selection has expired.")
    }
    const updated = sessionStore.update(sessionId, context, {
      allianceId,
      allianceName,
      allianceTokenMap: null,
      allianceTokenNameMap: null
    })
    await interaction.showModal(buildCoreModal(sessionId, updated.data))
    return true
  }

  const prefixes = Object.values(CREATION_IDS).filter(value => value.endsWith(":"))
  const prefix = prefixes.find(value => customId.startsWith(value))
  const modalPrefixes = new Set([
    CREATION_IDS.singleTimePrefix,
    CREATION_IDS.groupAddPrefix,
    CREATION_IDS.groupEditPrefix,
    CREATION_IDS.messagesPrefix,
    CREATION_IDS.imageReplacePrefix,
    CREATION_IDS.editPrefix
  ])
  if (!prefix || !modalPrefixes.has(prefix) || !interaction.isButton?.()) return false
  const sessionId = idSuffix(customId, prefix)
  const session = sessionStore.get(sessionId, context)

  if (prefix === CREATION_IDS.singleTimePrefix) {
    await interaction.showModal(buildSingleTimeModal(sessionId, session.data))
    return true
  }
  if (prefix === CREATION_IDS.groupAddPrefix) {
    await interaction.showModal(buildGroupModal(sessionId, "add"))
    return true
  }
  if (prefix === CREATION_IDS.groupEditPrefix) {
    const index = session.data.selectedGroupIndex
    if (!Number.isInteger(index) || !session.data.groups?.[index]) {
      throw new EventValidationError("Select a group to edit.")
    }
    await interaction.showModal(buildGroupModal(sessionId, "edit", session.data.groups[index]))
    return true
  }
  if (prefix === CREATION_IDS.messagesPrefix) {
    await interaction.showModal(buildMessagesModal(sessionId, session.data))
    return true
  }
  if (prefix === CREATION_IDS.imageReplacePrefix) {
    await interaction.showModal(buildImageUploadModal(sessionId))
    return true
  }
  if (prefix === CREATION_IDS.editPrefix) {
    await interaction.showModal(buildCoreModal(sessionId, session.data))
    return true
  }
  return false
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
      `Final custom message: ${data.finalReminderMessage ? "Yes" : "No"}\n\n` +
      "Advance sends one alliance reminder at the selected offset. " +
      "Final says About to start one minute before the event.",
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
          .setCustomId(`${CREATION_IDS.imageManagePrefix}${sessionId}`)
          .setLabel("Manage image")
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
      `State weekly roundup: ${data.includeInWeeklyRoundup && data.publishToState ? "Yes" : "No"}\n\n` +
      "State eligibility includes this event in the combined state weekly roundup. " +
      "It never sends an individual state reminder.",
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
    if (!settings?.event_channel_id) {
      await interaction.editReply({
        content: "Configure the alliance event channel before creating events.",
        components: []
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
    await acknowledgeSchedulerInteraction(interaction)
    const page = Number(idSuffix(customId, CREATION_IDS.listPrefix))
    await showList(interaction, repository, page)
    return true
  }

  if (
    interaction.isButton?.()
    && customId.startsWith(CREATION_IDS.occurrencePreviewPrefix)
  ) {
    await acknowledgeSchedulerInteraction(interaction)
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
    await acknowledgeSchedulerInteraction(interaction)
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
    const event = await validateSessionEvent(repository, interaction.guildId, updated.data)
    sessionStore.update(sessionId, context, { ...event, allianceChangeOnly: false })
    await acknowledgeSchedulerInteraction(interaction)
    await interaction.editReply(buildPreviewView(sessionId, event))
    return true
  }

  if (interaction.isModalSubmit?.()
    && customId.startsWith(CREATION_IDS.singleTimeModalPrefix)) {
    const sessionId = idSuffix(customId, CREATION_IDS.singleTimeModalPrefix)
    sessionStore.get(sessionId, context)
    await acknowledgeSchedulerInteraction(interaction)
    const session = sessionStore.update(sessionId, context, {
      eventTimeUtc: parseUtcTime(interaction.fields.getTextInputValue("t")),
      groups: [],
      grouped: false,
      selectedGroupIndex: null
    })
    await interaction.editReply(buildTimingView(sessionId, session.data))
    return true
  }

  if (interaction.isModalSubmit?.() && customId.startsWith(CREATION_IDS.groupModalPrefix)) {
    const suffix = idSuffix(customId, CREATION_IDS.groupModalPrefix)
    const separator = suffix.lastIndexOf(":")
    const sessionId = suffix.slice(0, separator)
    const mode = suffix.slice(separator + 1)
    if (!sessionId || !["add", "edit"].includes(mode)) {
      throw new EventValidationError("That group action is invalid.")
    }
    const current = sessionStore.get(sessionId, context).data
    const editingIndex = mode === "edit" ? current.selectedGroupIndex : null
    if (mode === "edit" && !Number.isInteger(editingIndex)) {
      throw new EventValidationError("Select a group to edit.")
    }
    const group = normalizeGroupInput(
      interaction.fields.getTextInputValue("name"),
      interaction.fields.getTextInputValue("time"),
      current.groups || [],
      editingIndex
    )
    const groups = upsertGroup(current.groups || [], group, editingIndex)
    await acknowledgeSchedulerInteraction(interaction)
    const session = sessionStore.update(sessionId, context, {
      groups,
      eventTimeUtc: null,
      grouped: true,
      selectedGroupIndex: mode === "edit" ? editingIndex : groups.length - 1
    })
    await interaction.editReply(buildGroupManagerView(sessionId, session.data))
    return true
  }

  if (interaction.isModalSubmit?.() && customId.startsWith(CREATION_IDS.imageUploadPrefix)) {
    const sessionId = idSuffix(customId, CREATION_IDS.imageUploadPrefix)
    const current = sessionStore.get(sessionId, context).data
    const attachment = interaction.fields.getUploadedFiles("img")?.first?.()
    if (!attachment) throw new EventValidationError("Select one replacement image.")
    await acknowledgeSchedulerInteraction(interaction)
    const image = await downloadEventImage(attachment)
    const session = sessionStore.update(sessionId, context, { image, imageAction: "replace" })
    await interaction.editReply(current.imageReturnView === "preview"
      ? buildPreviewView(sessionId, session.data)
      : buildTimingView(sessionId, session.data))
    return true
  }

  if (
    interaction.isModalSubmit?.()
    && customId.startsWith(CREATION_IDS.messagesModalPrefix)
  ) {
    const sessionId = idSuffix(customId, CREATION_IDS.messagesModalPrefix)
    sessionStore.get(sessionId, context)
    await acknowledgeSchedulerInteraction(interaction)
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
    await acknowledgeSchedulerInteraction(interaction)

    const parsedDate = parseIsoDate(interaction.fields.getTextInputValue("d"))
    const session = sessionStore.update(sessionId, context, {
      eventName: interaction.fields.getTextInputValue("e").trim(),
      firstOccurrenceDate: parsedDate.value,
      firstDateIsPast: parsedDate.isPast
    })
    await interaction.editReply(buildTimingChoiceView(sessionId, session.data))
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
    await acknowledgeSchedulerInteraction(interaction)
    await interaction.editReply(buildTimingView(sessionId, session.data))
    return true
  }
  if (interaction.isStringSelectMenu?.() && prefix === CREATION_IDS.groupSelectPrefix) {
    sessionStore.update(sessionId, context, { selectedGroupIndex: Number(interaction.values[0]) })
    await acknowledgeSchedulerInteraction(interaction)
    await interaction.editReply(buildGroupManagerView(sessionId, session.data))
    return true
  }
  if (interaction.isStringSelectMenu?.() && prefix === CREATION_IDS.advancePrefix) {
    const value = interaction.values[0]
    sessionStore.update(sessionId, context, {
      advanceReminderMinutes: value === "none" ? null : Number(value)
    })
    await acknowledgeSchedulerInteraction(interaction)
    await interaction.editReply(buildTimingView(sessionId, session.data))
    return true
  }

  if (interaction.isButton?.() && prefix === CREATION_IDS.startPrefix) {
    sessionStore.update(sessionId, context, { reminderAtStart: !session.data.reminderAtStart })
    await acknowledgeSchedulerInteraction(interaction)
    await interaction.editReply(buildTimingView(sessionId, session.data))
    return true
  }
  if (interaction.isButton?.() && prefix === CREATION_IDS.singleTimePrefix) {
    return false
  }
  if (interaction.isButton?.() && prefix === CREATION_IDS.groupsPrefix) {
    sessionStore.update(sessionId, context, {
      groups: session.data.groups || [],
      eventTimeUtc: null,
      grouped: true,
      selectedGroupIndex: null
    })
    await acknowledgeSchedulerInteraction(interaction)
    await interaction.editReply(buildGroupManagerView(sessionId, session.data))
    return true
  }
  if (interaction.isButton?.() && prefix === CREATION_IDS.groupAddPrefix) {
    return false
  }
  if (interaction.isButton?.() && prefix === CREATION_IDS.groupEditPrefix) {
    const index = session.data.selectedGroupIndex
    if (!Number.isInteger(index) || !session.data.groups?.[index]) {
      throw new EventValidationError("Select a group to edit.")
    }
    return false
  }
  if (interaction.isButton?.() && prefix === CREATION_IDS.groupRemovePrefix) {
    const index = session.data.selectedGroupIndex
    if (!Number.isInteger(index) || !session.data.groups?.[index]) {
      throw new EventValidationError("Select a group to remove.")
    }
    const groups = removeGroup(session.data.groups, index)
    sessionStore.update(sessionId, context, { groups, selectedGroupIndex: null })
    await acknowledgeSchedulerInteraction(interaction)
    await interaction.editReply(buildGroupManagerView(sessionId, session.data))
    return true
  }
  if (interaction.isButton?.() && prefix === CREATION_IDS.groupContinuePrefix) {
    if (!session.data.groups?.length) {
      throw new EventValidationError("A grouped event requires at least one group.")
    }
    await acknowledgeSchedulerInteraction(interaction)
    await interaction.editReply(buildTimingView(sessionId, session.data))
    return true
  }
  if (interaction.isButton?.() && prefix === CREATION_IDS.messagesPrefix) {
    return false
  }
  if (interaction.isButton?.() && prefix === CREATION_IDS.imageManagePrefix) {
    sessionStore.update(sessionId, context, { imageReturnView: "timing" })
    await acknowledgeSchedulerInteraction(interaction)
    await interaction.editReply(buildImageChoiceView(sessionId, session.data))
    return true
  }
  if (interaction.isButton?.() && prefix === CREATION_IDS.imageKeepPrefix) {
    sessionStore.update(sessionId, context, { imageAction: "retain" })
    await acknowledgeSchedulerInteraction(interaction)
    await interaction.editReply(session.data.imageReturnView === "preview"
      ? buildPreviewView(sessionId, session.data)
      : buildTimingView(sessionId, session.data))
    return true
  }
  if (interaction.isButton?.() && prefix === CREATION_IDS.imageReplacePrefix) {
    return false
  }
  if (interaction.isButton?.() && prefix === CREATION_IDS.imageRemovePrefix) {
    sessionStore.update(sessionId, context, { image: null, imageAction: "remove" })
    await acknowledgeSchedulerInteraction(interaction)
    await interaction.editReply(session.data.imageReturnView === "preview"
      ? buildPreviewView(sessionId, session.data)
      : buildTimingView(sessionId, session.data))
    return true
  }
  if (interaction.isButton?.() && prefix === CREATION_IDS.optionsPrefix) {
    await acknowledgeSchedulerInteraction(interaction)
    await interaction.editReply(buildPublishingView(sessionId, session.data))
    return true
  }
  if (interaction.isButton?.() && prefix === CREATION_IDS.alliancePrefix) {
    sessionStore.update(sessionId, context, { publishToAlliance: !session.data.publishToAlliance })
    await acknowledgeSchedulerInteraction(interaction)
    await interaction.editReply(buildPublishingView(sessionId, session.data))
    return true
  }
  if (interaction.isButton?.() && prefix === CREATION_IDS.statePrefix) {
    if (!session.data.publishToState && !(await stateLinkIsEnabled(repository, interaction.guildId))) {
      await interaction.editReply({
        content: "Enable a valid state roundup link before including this event in state roundups.",
        components: []
      })
      return true
    }
    sessionStore.update(sessionId, context, { publishToState: !session.data.publishToState })
    await acknowledgeSchedulerInteraction(interaction)
    await interaction.editReply(buildPublishingView(sessionId, session.data))
    return true
  }
  if (interaction.isButton?.() && prefix === CREATION_IDS.roundupPrefix) {
    sessionStore.update(sessionId, context, {
      includeInWeeklyRoundup: !session.data.includeInWeeklyRoundup
    })
    await acknowledgeSchedulerInteraction(interaction)
    await interaction.editReply(buildPublishingView(sessionId, session.data))
    return true
  }
  if (interaction.isButton?.() && prefix === CREATION_IDS.previewPrefix) {
    const event = await validateSessionEvent(repository, interaction.guildId, session.data)
    sessionStore.update(sessionId, context, event)
    await acknowledgeSchedulerInteraction(interaction)
    await interaction.editReply(buildPreviewView(sessionId, event))
    return true
  }
  if (interaction.isButton?.() && prefix === CREATION_IDS.editPrefix) {
    return false
  }
  if (interaction.isButton?.() && prefix === CREATION_IDS.allianceChangePrefix) {
    sessionStore.update(sessionId, context, { allianceChangeOnly: true })
    await acknowledgeSchedulerInteraction(interaction)
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
    await acknowledgeSchedulerInteraction(interaction)
    await interaction.editReply(await loadHome(repository, interaction.guildId))
    return true
  }
  if (interaction.isButton?.() && prefix === CREATION_IDS.createPrefix) {
    await acknowledgeSchedulerInteraction(interaction)
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
  buildTimingChoiceView,
  buildSingleTimeModal,
  normalizeGroupInput,
  upsertGroup,
  removeGroup,
  buildGroupModal,
  buildGroupManagerView,
  buildImageChoiceView,
  buildImageUploadModal,
  buildMessagesModal,
  buildAllianceSelectionView,
  renderAllianceSelection,
  buildTimingView,
  buildPublishingView,
  buildPreviewView,
  buildListView,
  buildOccurrencePreviewView,
  handleEventCreationModalOpeningInteraction,
  handleEventCreationInteraction
}
