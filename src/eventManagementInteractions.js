const crypto = require("crypto")
const {
  ActionRowBuilder,
  ButtonBuilder,
  ButtonStyle,
  MessageFlags,
  StringSelectMenuBuilder
} = require("discord.js")

const { InteractionSessionError, InteractionSessionStore } = require("./interactionSessions")
const { acknowledgeSchedulerInteraction } = require("./interactionResponses")
const {
  creationSessions,
  sessionContext,
  renderAllianceSelection,
  buildCoreModal,
  buildImageChoiceView
} = require("./eventCreationInteractions")
const {
  EVENTS_PER_PAGE,
  formatEventEntry,
  formatUpcomingOccurrencePreview
} = require("./eventSchedulerFormatting")
const { getNextOccurrences } = require("./occurrenceCalculation")

const managementSessions = new InteractionSessionStore()
const MANAGEMENT_PREFIX = "mg:"
const ALLIANCES_PER_FILTER = 24
const FILTER_IDS = Object.freeze({
  selectPrefix: `${MANAGEMENT_PREFIX}f:`,
  changePrefix: `${MANAGEMENT_PREFIX}fc:`
})

function token() {
  return crypto.randomBytes(6).toString("base64url")
}

function uniqueToken(tokenMap) {
  let value
  do value = token()
  while (Object.hasOwn(tokenMap, value))
  return value
}

function eventDraft(event) {
  return {
    mode: "edit",
    eventId: String(event.id),
    allianceId: String(event.alliance_id),
    allianceName: event.alliance_name,
    eventName: event.event_name,
    firstOccurrenceDate: String(event.first_occurrence_date).slice(0, 10),
    firstDateIsPast: false,
    eventTimeUtc: event.event_time_utc ? String(event.event_time_utc).slice(0, 5) : null,
    groups: (event.groups || []).map(group => ({
      groupName: group.group_name,
      eventTimeUtc: String(group.event_time_utc).slice(0, 5),
      sortOrder: group.sort_order
    })),
    grouped: Boolean(event.groups?.length),
    recurrenceDays: event.recurrence_days,
    advanceReminderMinutes: event.advance_reminder_minutes,
    advanceReminderMessage: event.advance_reminder_message,
    reminderAtStart: event.reminder_at_start,
    finalReminderMessage: event.final_reminder_message,
    publishToAlliance: event.publish_to_alliance,
    publishToState: event.publish_to_state,
    includeInWeeklyRoundup: event.include_in_weekly_roundup,
    image: event.image_filename ? {
      originalFilename: event.image_filename,
      byteSize: event.image_byte_size,
      contentType: event.image_content_type
    } : null,
    imageAction: "retain"
  }
}

function actionOptions(event) {
  return [
    { label: "Preview", value: "preview" },
    { label: "Edit details", value: "edit" },
    { label: "Change alliance", value: "alliance" },
    { label: "Manage image", value: "image" },
    {
      label: event.status === "paused" ? "Resume" : "Pause",
      value: event.status === "paused" ? "resume" : "pause"
    },
    { label: "Delete", value: "delete" }
  ]
}

function allianceFilterView(alliances, sessionId) {
  const allianceMap = {}
  const options = [{ label: "All alliances", value: "all", description: "Every alliance in this server" }]
  for (const alliance of alliances) {
    const allianceToken = uniqueToken(allianceMap)
    allianceMap[allianceToken] = {
      id: String(alliance.id),
      name: alliance.alliance_name,
      isDefault: alliance.is_default === true
    }
    options.push({
      label: alliance.alliance_name.slice(0, 100),
      value: allianceToken,
      description: alliance.is_default ? "Main alliance" : "Sub-alliance"
    })
  }
  return {
    view: {
      content: "View events\n\nSelect an alliance or choose All alliances.",
      components: [
        new ActionRowBuilder().addComponents(
          new StringSelectMenuBuilder()
            .setCustomId(`${FILTER_IDS.selectPrefix}${sessionId}`)
            .setPlaceholder("Select alliance")
            .addOptions(options)
        ),
        new ActionRowBuilder().addComponents(
          new ButtonBuilder().setCustomId("es:home").setLabel("Back")
            .setStyle(ButtonStyle.Secondary)
        )
      ]
    },
    allianceMap
  }
}

async function renderAllianceFilter(interaction, repository, health, sessionId = null) {
  const context = sessionContext(interaction, health)
  const id = sessionId || managementSessions.create(context, {})
  managementSessions.get(id, context)
  const result = await repository.listAlliances(interaction.guildId, {
    limit: ALLIANCES_PER_FILTER,
    offset: 0
  })
  const built = allianceFilterView(result.alliances, id)
  managementSessions.update(id, context, { allianceMap: built.allianceMap })
  await interaction.editReply(built.view)
}

function listView(
  events,
  page,
  total,
  sessionId,
  eventMap,
  eventDraftMap = {},
  { filterLabel = "All alliances", allAlliances = true } = {}
) {
  const totalPages = Math.max(1, Math.ceil(total / EVENTS_PER_PAGE))
  const components = events.map((event, index) => {
    const eventToken = uniqueToken(eventMap)
    eventMap[eventToken] = String(event.id)
    eventDraftMap[eventToken] = eventDraft(event)
    return new ActionRowBuilder().addComponents(
      new StringSelectMenuBuilder()
        .setCustomId(`${MANAGEMENT_PREFIX}a:${sessionId}:${eventToken}`)
        .setPlaceholder(`${index + 1}. ${allAlliances ? `${event.alliance_name} — ` : ""}${event.event_name}`.slice(0, 100))
        .addOptions(actionOptions(event))
    )
  })
  const navigation = [
    new ButtonBuilder().setCustomId(`${FILTER_IDS.changePrefix}${sessionId}`)
      .setLabel("Change alliance").setStyle(ButtonStyle.Secondary)
  ]
  if (events.length) {
    navigation.push(
      new ButtonBuilder()
        .setCustomId(`${MANAGEMENT_PREFIX}l:${sessionId}:${Math.max(0, page - 1)}`)
        .setLabel("Previous").setStyle(ButtonStyle.Secondary).setDisabled(page === 0),
      new ButtonBuilder()
        .setCustomId(`${MANAGEMENT_PREFIX}l:${sessionId}:${page + 1}`)
        .setLabel("Next").setStyle(ButtonStyle.Secondary).setDisabled(page >= totalPages - 1)
    )
  }
  navigation.push(
    new ButtonBuilder().setCustomId("es:home").setLabel("Back").setStyle(ButtonStyle.Secondary)
  )
  components.push(new ActionRowBuilder().addComponents(...navigation))
  const content = events.length
    ? `Scheduled events — ${filterLabel}\nPage ${page + 1} of ${totalPages}\n\n` +
      events.map(formatEventEntry).join("\n\n")
    : `No scheduled events for ${filterLabel}.`
  return { content: content.slice(0, 1950), components }
}

async function renderList(interaction, repository, health, page, sessionId = null) {
  const context = sessionContext(interaction, health)
  const safePage = Math.max(0, Number(page) || 0)
  const id = sessionId || managementSessions.create(context, {})
  const session = managementSessions.get(id, context)
  const result = await repository.listEvents(interaction.guildId, {
    limit: EVENTS_PER_PAGE,
    offset: safePage * EVENTS_PER_PAGE,
    allianceId: session.data.selectedAllianceId || null
  })
  const eventMap = {}
  const eventDraftMap = {}
  managementSessions.update(id, context, { page: safePage, eventMap, eventDraftMap })
  const view = listView(result.events, safePage, result.total, id, eventMap, eventDraftMap, {
    filterLabel: session.data.selectedAllianceName || "All alliances",
    allAlliances: !session.data.selectedAllianceId
  })
  if (interaction.deferred || interaction.replied) await interaction.editReply(view)
  else await interaction.reply({ ...view, flags: MessageFlags.Ephemeral })
}

async function handleEventManagementModalOpeningInteraction(interaction, { health }) {
  const customId = String(interaction.customId || "")
  if (!(interaction.isStringSelectMenu?.()
    && customId.startsWith(`${MANAGEMENT_PREFIX}a:`))) return false
  const [, , sessionId, eventToken] = customId.split(":")
  const action = String(interaction.values?.[0] || "")
  if (action !== "edit") return false
  const context = sessionContext(interaction, health)
  const session = managementSessions.get(sessionId, context)
  const eventId = session.data.eventMap?.[eventToken]
  const draft = session.data.eventDraftMap?.[eventToken]
  if (!eventId || !draft || String(draft.eventId) !== String(eventId) || !draft.allianceId) {
    throw new InteractionSessionError("That event control is no longer valid.")
  }
  const editSessionId = creationSessions.create(context, draft)
  await interaction.showModal(buildCoreModal(editSessionId, draft))
  return true
}

function confirmationView(sessionId, action, eventToken, event, page = 0) {
  const verb = action === "delete" ? "Delete" : action === "pause" ? "Pause" : "Resume"
  const effect = action === "delete"
    ? "Delete softly removes future reminders and roundups while preserving history."
    : action === "pause"
      ? "Pause stops future reminders and roundups while preserving the recurrence schedule."
      : "Resume restores future eligible reminders from the original recurrence schedule."
  return {
    content: `${verb} ${event.event_name}?\n\n${effect}\n\nThis change requires confirmation.`,
    components: [new ActionRowBuilder().addComponents(
      new ButtonBuilder()
        .setCustomId(`${MANAGEMENT_PREFIX}c:${sessionId}:${action}:${eventToken}`)
        .setLabel(`Confirm ${verb.toLowerCase()}`)
        .setStyle(action === "delete" ? ButtonStyle.Danger : ButtonStyle.Primary),
      new ButtonBuilder()
        .setCustomId(`${MANAGEMENT_PREFIX}l:${sessionId}:${page}`)
        .setLabel("Cancel").setStyle(ButtonStyle.Secondary)
    )]
  }
}

async function handleEventManagementInteraction(
  interaction,
  { repository, health }
) {
  const customId = String(interaction.customId || "")
  if (!(customId.startsWith("el:") || customId.startsWith(MANAGEMENT_PREFIX))) return false
  const context = sessionContext(interaction, health)

  if (interaction.isButton?.() && customId.startsWith("el:")) {
    await acknowledgeSchedulerInteraction(interaction)
    await renderAllianceFilter(interaction, repository, health)
    return true
  }

  if (interaction.isButton?.() && customId.startsWith(FILTER_IDS.changePrefix)) {
    const sessionId = customId.slice(FILTER_IDS.changePrefix.length)
    managementSessions.get(sessionId, context)
    await acknowledgeSchedulerInteraction(interaction)
    await renderAllianceFilter(interaction, repository, health, sessionId)
    return true
  }

  if (interaction.isStringSelectMenu?.() && customId.startsWith(FILTER_IDS.selectPrefix)) {
    const sessionId = customId.slice(FILTER_IDS.selectPrefix.length)
    const session = managementSessions.get(sessionId, context)
    const selectedToken = String(interaction.values?.[0] || "")
    let selectedAllianceId = null
    let selectedAllianceName = "All alliances"
    if (selectedToken !== "all") {
      const selected = session.data.allianceMap?.[selectedToken]
      if (!selected?.id) throw new InteractionSessionError("That alliance control has expired.")
      const alliance = await repository.getAlliance(interaction.guildId, selected.id)
      if (!alliance) throw new InteractionSessionError("That alliance is no longer available.")
      selectedAllianceId = String(alliance.id)
      selectedAllianceName = alliance.alliance_name
    }
    managementSessions.update(sessionId, context, {
      selectedAllianceId,
      selectedAllianceName,
      page: 0
    })
    await acknowledgeSchedulerInteraction(interaction)
    await renderList(interaction, repository, health, 0, sessionId)
    return true
  }

  if (interaction.isButton?.() && customId.startsWith(`${MANAGEMENT_PREFIX}l:`)) {
    const [, , sessionId, page] = customId.split(":")
    managementSessions.get(sessionId, context)
    await acknowledgeSchedulerInteraction(interaction)
    await renderList(interaction, repository, health, Number(page), sessionId)
    return true
  }

  if (interaction.isStringSelectMenu?.() && customId.startsWith(`${MANAGEMENT_PREFIX}a:`)) {
    const [, , sessionId, eventToken] = customId.split(":")
    const session = managementSessions.get(sessionId, context)
    const action = String(interaction.values[0])
    if (!["preview", "edit", "alliance", "image", "pause", "resume", "delete"].includes(action)) {
      throw new InteractionSessionError("That event action is invalid.")
    }
    const eventId = session.data.eventMap[eventToken]
    if (!eventId) throw new Error("That event control is no longer valid.")
    const event = await repository.getEvent(interaction.guildId, eventId)
    if (!event) throw new Error("That event is no longer available.")

    if (action === "preview") {
      await acknowledgeSchedulerInteraction(interaction)
      const occurrences = getNextOccurrences(event, new Date(), 5)
      await interaction.editReply({
        content: formatUpcomingOccurrencePreview(event, occurrences),
        allowedMentions: { parse: [], repliedUser: false },
        components: [new ActionRowBuilder().addComponents(
          new ButtonBuilder().setCustomId(`${MANAGEMENT_PREFIX}l:${sessionId}:${session.data.page}`)
            .setLabel("Back to events").setStyle(ButtonStyle.Secondary)
        )]
      })
      return true
    }
    if (action === "edit") {
      return false
    }
    if (action === "alliance") {
      const draft = { ...eventDraft(event), allianceChangeOnly: true }
      const editSessionId = creationSessions.create(context, draft)
      await acknowledgeSchedulerInteraction(interaction)
      await renderAllianceSelection(
        interaction,
        repository,
        creationSessions,
        editSessionId,
        context,
        0
      )
      return true
    }
    if (action === "image") {
      const draft = { ...eventDraft(event), imageReturnView: "preview" }
      const editSessionId = creationSessions.create(context, draft)
      await acknowledgeSchedulerInteraction(interaction)
      await interaction.editReply(buildImageChoiceView(editSessionId, draft))
      return true
    }
    if (["pause", "resume", "delete"].includes(action)) {
      await acknowledgeSchedulerInteraction(interaction)
      await interaction.editReply(confirmationView(
        sessionId,
        action,
        eventToken,
        event,
        session.data.page
      ))
      return true
    }
  }

  if (interaction.isButton?.() && customId.startsWith(`${MANAGEMENT_PREFIX}c:`)) {
    const [, , sessionId, action, eventToken] = customId.split(":")
    if (!["pause", "resume", "delete"].includes(action)) {
      throw new InteractionSessionError("That event action is invalid.")
    }
    const session = managementSessions.get(sessionId, context)
    const eventId = session.data.eventMap[eventToken]
    if (!eventId) throw new Error("That event control is no longer valid.")
    const status = action === "pause" ? "paused" : action === "resume" ? "active" : "deleted"
    await acknowledgeSchedulerInteraction(interaction)
    const changed = await repository.setEventStatus({
      guildId: interaction.guildId,
      eventId,
      status
    })
    if (!changed) throw new Error("That event is no longer available.")
    managementSessions.complete(sessionId, context)
    await interaction.editReply({
      content: `${action === "delete" ? "Deleted" : action === "pause" ? "Paused" : "Resumed"} ${changed.event_name}.`,
      components: []
    })
    return true
  }
  return false
}

module.exports = {
  MANAGEMENT_PREFIX,
  FILTER_IDS,
  managementSessions,
  eventDraft,
  actionOptions,
  allianceFilterView,
  renderAllianceFilter,
  listView,
  imageChoiceView: buildImageChoiceView,
  confirmationView,
  handleEventManagementModalOpeningInteraction,
  handleEventManagementInteraction
}
