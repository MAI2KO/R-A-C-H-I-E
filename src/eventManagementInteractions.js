const crypto = require("crypto")
const {
  ActionRowBuilder,
  ButtonBuilder,
  ButtonStyle,
  MessageFlags,
  StringSelectMenuBuilder
} = require("discord.js")

const { InteractionSessionError, InteractionSessionStore } = require("./interactionSessions")
const {
  creationSessions,
  sessionContext,
  renderAllianceSelection,
  buildCoreModal,
  buildImageChoiceView
} = require("./eventCreationInteractions")
const {
  EVENTS_PER_PAGE,
  formatEventListPage,
  formatUpcomingOccurrencePreview
} = require("./eventSchedulerFormatting")
const { getNextOccurrences } = require("./occurrenceCalculation")

const managementSessions = new InteractionSessionStore()
const MANAGEMENT_PREFIX = "mg:"

function token() {
  return crypto.randomBytes(6).toString("base64url")
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

function actionOptions(event, eventToken) {
  return [
    { label: "Preview", value: `preview:${eventToken}` },
    { label: "Edit details", value: `edit:${eventToken}` },
    { label: "Change alliance", value: `alliance:${eventToken}` },
    { label: "Manage image", value: `image:${eventToken}` },
    {
      label: event.status === "paused" ? "Resume" : "Pause",
      value: `${event.status === "paused" ? "resume" : "pause"}:${eventToken}`
    },
    { label: "Delete", value: `delete:${eventToken}` }
  ]
}

function listView(events, page, total, sessionId, eventMap) {
  const totalPages = Math.max(1, Math.ceil(total / EVENTS_PER_PAGE))
  const components = events.map((event, index) => {
    const eventToken = token()
    eventMap[eventToken] = String(event.id)
    return new ActionRowBuilder().addComponents(
      new StringSelectMenuBuilder()
        .setCustomId(`${MANAGEMENT_PREFIX}a:${sessionId}`)
        .setPlaceholder(`${index + 1}. ${event.event_name}`.slice(0, 100))
        .addOptions(actionOptions(event, eventToken))
    )
  })
  components.push(new ActionRowBuilder().addComponents(
    new ButtonBuilder()
      .setCustomId(`${MANAGEMENT_PREFIX}l:${sessionId}:${Math.max(0, page - 1)}`)
      .setLabel("Previous").setStyle(ButtonStyle.Secondary).setDisabled(page === 0),
    new ButtonBuilder()
      .setCustomId(`${MANAGEMENT_PREFIX}l:${sessionId}:${page + 1}`)
      .setLabel("Next").setStyle(ButtonStyle.Secondary).setDisabled(page >= totalPages - 1),
    new ButtonBuilder().setCustomId("es:home").setLabel("Back").setStyle(ButtonStyle.Secondary)
  ))
  return { content: formatEventListPage(events, page, total), components }
}

async function renderList(interaction, repository, health, page, sessionId = null) {
  const context = sessionContext(interaction, health)
  const safePage = Math.max(0, Number(page) || 0)
  const result = await repository.listEvents(interaction.guildId, {
    limit: EVENTS_PER_PAGE,
    offset: safePage * EVENTS_PER_PAGE
  })
  const id = sessionId || managementSessions.create(context, {})
  const eventMap = {}
  managementSessions.update(id, context, { page: safePage, eventMap })
  const view = listView(result.events, safePage, result.total, id, eventMap)
  if (interaction.deferred || interaction.replied) await interaction.editReply(view)
  else await interaction.reply({ ...view, flags: MessageFlags.Ephemeral })
}

function confirmationView(sessionId, action, eventToken, event) {
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
        .setCustomId(`${MANAGEMENT_PREFIX}l:${sessionId}:0`)
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
    await interaction.deferUpdate()
    await renderList(interaction, repository, health, Number(customId.slice(3)))
    return true
  }

  if (interaction.isButton?.() && customId.startsWith(`${MANAGEMENT_PREFIX}l:`)) {
    const [, , sessionId, page] = customId.split(":")
    managementSessions.get(sessionId, context)
    await interaction.deferUpdate()
    await renderList(interaction, repository, health, Number(page), sessionId)
    return true
  }

  if (interaction.isStringSelectMenu?.() && customId.startsWith(`${MANAGEMENT_PREFIX}a:`)) {
    const sessionId = customId.split(":")[2]
    const session = managementSessions.get(sessionId, context)
    const [action, eventToken] = String(interaction.values[0]).split(":")
    if (!["preview", "edit", "alliance", "image", "pause", "resume", "delete"].includes(action)) {
      throw new InteractionSessionError("That event action is invalid.")
    }
    const eventId = session.data.eventMap[eventToken]
    if (!eventId) throw new Error("That event control is no longer valid.")
    const event = await repository.getEvent(interaction.guildId, eventId)
    if (!event) throw new Error("That event is no longer available.")

    if (action === "preview") {
      await interaction.deferUpdate()
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
      const draft = eventDraft(event)
      const editSessionId = creationSessions.create(context, draft)
      await interaction.showModal(buildCoreModal(editSessionId, draft))
      return true
    }
    if (action === "alliance") {
      const draft = { ...eventDraft(event), allianceChangeOnly: true }
      const editSessionId = creationSessions.create(context, draft)
      await interaction.deferUpdate()
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
      await interaction.deferUpdate()
      await interaction.editReply(buildImageChoiceView(editSessionId, draft))
      return true
    }
    if (["pause", "resume", "delete"].includes(action)) {
      await interaction.deferUpdate()
      await interaction.editReply(confirmationView(sessionId, action, eventToken, event))
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
    await interaction.deferUpdate()
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
  managementSessions,
  eventDraft,
  actionOptions,
  listView,
  imageChoiceView: buildImageChoiceView,
  confirmationView,
  handleEventManagementInteraction
}
