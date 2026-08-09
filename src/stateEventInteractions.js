const crypto = require("crypto")
const {
  ActionRowBuilder,
  ButtonBuilder,
  ButtonStyle,
  FileUploadBuilder,
  LabelBuilder,
  ModalBuilder,
  StringSelectMenuBuilder,
  TextInputBuilder,
  TextInputStyle
} = require("discord.js")

const { prepareStoredEventImage } = require("./discordEventDelivery")
const { downloadEventImage } = require("./eventImage")
const { formatStateEventDelivery } = require("./stateEventDeliveryFormatting")
const { getStateEventOccurrencesInRange } = require("./stateEventOccurrenceCalculation")
const {
  normalizeRecurrenceDays,
  normalizeStateEventName,
  normalizeStatePhase,
  validateStateEventDraft
} = require("./stateEventValidation")
const { InteractionSessionError, InteractionSessionStore } = require("./interactionSessions")
const { parseIsoDate } = require("./timeParsing")

const stateEventSessions = new InteractionSessionStore()
const EVENTS_PER_PAGE = 10
const CRITICAL_ALERT_WARNING = "Exact-time alerts may cover the game screen. Consider disabling them for critical phases such as battle or red-zone start."

const STATE_EVENT_IDS = Object.freeze({
  prefix: "se:",
  home: "se:home",
  newEvent: "se:new",
  basicModalPrefix: "se:bm:",
  recurrencePrefix: "se:r:",
  phasesPrefix: "se:ph:",
  phaseSelectPrefix: "se:ps:",
  phaseAddPrefix: "se:pa:",
  phaseEditPrefix: "se:pe:",
  phaseRemovePrefix: "se:pd:",
  phaseModalPrefix: "se:pm:",
  phasePrePrefix: "se:pp:",
  phaseExactPrefix: "se:px:",
  phaseMessagesPrefix: "se:pg:",
  phaseMessagesModalPrefix: "se:pgm:",
  phaseMediaPrefix: "se:mi:",
  phaseMediaReplacePrefix: "se:mr:",
  phaseMediaUploadPrefix: "se:mu:",
  phaseMediaRemovePrefix: "se:md:",
  phaseMediaBackPrefix: "se:mb:",
  phaseSavePrefix: "se:sv:",
  phaseCancelPrefix: "se:pc:",
  reviewPrefix: "se:rv:",
  confirmPrefix: "se:ok:",
  cancelPrefix: "se:cx:",
  listPrefix: "se:li:",
  eventSelectPrefix: "se:es:",
  managePrefix: "se:mg:",
  editDetailsPrefix: "se:ed:",
  managePhasesPrefix: "se:mp:",
  previewPrefix: "se:pv:",
  statusPrefix: "se:st:",
  statusConfirmPrefix: "se:sc:",
  testPrefix: "se:tt:",
  testPhasePrefix: "se:tp:",
  testKindPrefix: "se:tk:"
})

const RECURRENCES = Object.freeze([
  ["One-time", "none"],
  ["Every 2 days", "2"],
  ["Every 3 days", "3"],
  ["Every 1 week", "7"],
  ["Every 2 weeks", "14"],
  ["Every 3 weeks", "21"],
  ["Every 4 weeks", "28"],
  ["Every 5 weeks", "35"],
  ["Every 6 weeks", "42"]
])
const PRE_ALERTS = Object.freeze([
  ["None", "none"],
  ["5 minutes", "5"],
  ["10 minutes", "10"],
  ["15 minutes", "15"],
  ["20 minutes", "20"],
  ["30 minutes", "30"]
])

function sessionContext(interaction, health) {
  return {
    userId: interaction.user.id,
    guildId: interaction.guildId,
    gameProfile: health.gameProfile
  }
}

function suffix(customId, prefix) {
  return String(customId).slice(prefix.length)
}

function opaqueToken() {
  return crypto.randomBytes(6).toString("base64url")
}

function input(customId, label, value, { paragraph = false, required = true, maximum = 100 } = {}) {
  const field = new TextInputBuilder()
    .setCustomId(customId)
    .setLabel(label)
    .setStyle(paragraph ? TextInputStyle.Paragraph : TextInputStyle.Short)
    .setRequired(required)
    .setMaxLength(maximum)
  if (value) field.setValue(String(value))
  return new ActionRowBuilder().addComponents(field)
}

function recurrenceLabel(value) {
  const normalized = value === null ? "none" : String(value)
  return RECURRENCES.find(([, candidate]) => candidate === normalized)?.[0] || "Unknown"
}

function phaseDraft(phase) {
  return {
    id: phase.id || null,
    phaseName: phase.phaseName ?? phase.phase_name,
    phaseTimeUtc: String(phase.phaseTimeUtc ?? phase.phase_time_utc ?? "").slice(0, 5),
    preAlertMinutes: phase.preAlertMinutes ?? phase.pre_alert_minutes ?? null,
    preAlertMessage: phase.preAlertMessage ?? phase.pre_alert_message ?? null,
    announceExact: phase.announceExact ?? phase.announce_exact ?? true,
    exactMessage: phase.exactMessage ?? phase.exact_message ?? null,
    preAlertMedia: phase.preAlertMedia ?? phase.pre_alert_media ?? null,
    exactMedia: phase.exactMedia ?? phase.exact_media ?? null,
    sortOrder: Number(phase.sortOrder ?? phase.sort_order ?? 0)
  }
}

function eventDraft(event) {
  return {
    mode: "edit",
    eventId: String(event.id),
    eventName: event.event_name,
    firstOccurrenceDate: event.first_occurrence_date,
    recurrenceDays: event.recurrence_days,
    stateNumber: event.state_number,
    status: event.status,
    phases: (event.phases || []).map(phaseDraft)
  }
}

function sortPhases(phases) {
  return [...phases]
    .sort((left, right) => String(left.phaseTimeUtc).localeCompare(String(right.phaseTimeUtc))
      || Number(left.sortOrder || 0) - Number(right.sortOrder || 0)
      || String(left.phaseName).localeCompare(String(right.phaseName)))
    .map((phase, index) => ({ ...phase, sortOrder: index }))
}

function phaseLines(phases, { localDate = null } = {}) {
  if (!phases?.length) return "No phases configured."
  return sortPhases(phases).map(phase => {
    let local = ""
    if (localDate) {
      const instant = new Date(`${localDate}T${phase.phaseTimeUtc}:00.000Z`)
      local = `\nLocal: <t:${Math.floor(instant.getTime() / 1000)}:t>`
    }
    return `${phase.phaseName} - ${phase.phaseTimeUtc} UTC${local}\n` +
      `Pre-alert: ${phase.preAlertMinutes === null ? "None" : `${phase.preAlertMinutes} min`}\n` +
      `Exact alert: ${phase.announceExact ? "On" : "Off"}`
  }).join("\n\n")
}

function buildStateEventHome(destination) {
  return {
    content: `State events\n\nState ${destination.state_number || "number not set"}\n` +
      "Canonical state events publish to this state Discord and its linked alliance Discords.",
    components: [
      new ActionRowBuilder().addComponents(
        new ButtonBuilder().setCustomId(STATE_EVENT_IDS.newEvent).setLabel("Create state event")
          .setStyle(ButtonStyle.Success).setDisabled(!destination.state_number),
        new ButtonBuilder().setCustomId(`${STATE_EVENT_IDS.listPrefix}0`).setLabel("View state events")
          .setStyle(ButtonStyle.Primary)
      ),
      new ActionRowBuilder().addComponents(
        new ButtonBuilder().setCustomId("eh:state-events").setLabel("Help")
          .setStyle(ButtonStyle.Secondary),
        new ButtonBuilder().setCustomId("es:home").setLabel("Back")
          .setStyle(ButtonStyle.Secondary)
      )
    ]
  }
}

function buildBasicModal(sessionId, data = {}) {
  return new ModalBuilder()
    .setCustomId(`${STATE_EVENT_IDS.basicModalPrefix}${sessionId}`)
    .setTitle(data.mode === "edit" ? "Edit state event" : "Create state event")
    .addComponents(
      input("name", "Event name", data.eventName),
      input("date", "First occurrence (YYYY-MM-DD)", data.firstOccurrenceDate)
    )
}

function buildRecurrenceView(sessionId, data) {
  const selected = data.recurrenceDays === null ? "none" : String(data.recurrenceDays ?? "")
  return {
    content: `State event recurrence\n\n${data.eventName}\nFirst occurrence: ${data.firstOccurrenceDate}\n` +
      `Recurrence: ${selected ? recurrenceLabel(data.recurrenceDays) : "Select one"}`,
    components: [
      new ActionRowBuilder().addComponents(
        new StringSelectMenuBuilder().setCustomId(`${STATE_EVENT_IDS.recurrencePrefix}${sessionId}`)
          .setPlaceholder("Choose recurrence").addOptions(RECURRENCES.map(([label, value]) => ({
            label, value, default: selected === value
          })))
      ),
      new ActionRowBuilder().addComponents(
        new ButtonBuilder().setCustomId(`${STATE_EVENT_IDS.cancelPrefix}${sessionId}`)
          .setLabel("Cancel").setStyle(ButtonStyle.Secondary)
      )
    ]
  }
}

function buildPhaseManagerView(sessionId, data) {
  const phases = sortPhases(data.phases || [])
  const selected = Number.isInteger(data.selectedPhaseIndex) ? data.selectedPhaseIndex : null
  const components = []
  if (phases.length) {
    components.push(new ActionRowBuilder().addComponents(
      new StringSelectMenuBuilder().setCustomId(`${STATE_EVENT_IDS.phaseSelectPrefix}${sessionId}`)
        .setPlaceholder("Select a phase to edit or remove")
        .addOptions(phases.slice(0, 25).map((phase, index) => ({
          label: phase.phaseName.slice(0, 100),
          description: `${phase.phaseTimeUtc} UTC`,
          value: String(index),
          default: selected === index
        })))
    ))
  }
  components.push(new ActionRowBuilder().addComponents(
    new ButtonBuilder().setCustomId(`${STATE_EVENT_IDS.phaseAddPrefix}${sessionId}`)
      .setLabel("Add phase").setStyle(ButtonStyle.Primary).setDisabled(phases.length >= 25),
    new ButtonBuilder().setCustomId(`${STATE_EVENT_IDS.phaseEditPrefix}${sessionId}`)
      .setLabel("Edit phase").setStyle(ButtonStyle.Secondary).setDisabled(selected === null),
    new ButtonBuilder().setCustomId(`${STATE_EVENT_IDS.phaseRemovePrefix}${sessionId}`)
      .setLabel("Remove phase").setStyle(ButtonStyle.Danger)
      .setDisabled(selected === null || phases.length <= 1),
    new ButtonBuilder().setCustomId(`${STATE_EVENT_IDS.reviewPrefix}${sessionId}`)
      .setLabel("Continue / Review").setStyle(ButtonStyle.Success).setDisabled(phases.length === 0)
  ))
  components.push(new ActionRowBuilder().addComponents(
    new ButtonBuilder().setCustomId(`${STATE_EVENT_IDS.cancelPrefix}${sessionId}`)
      .setLabel("Cancel").setStyle(ButtonStyle.Secondary)
  ))
  return {
    content: `Manage phases\n\n${phaseLines(phases)}\n\n${CRITICAL_ALERT_WARNING}`,
    components
  }
}

function buildPhaseModal(sessionId, mode, phase = {}) {
  return new ModalBuilder()
    .setCustomId(`${STATE_EVENT_IDS.phaseModalPrefix}${sessionId}:${mode}`)
    .setTitle(mode === "edit" ? "Edit phase" : "Add phase")
    .addComponents(
      input("name", "Phase name", phase.phaseName),
      input("time", "UTC time (e.g. 10am or 18:30)", phase.phaseTimeUtc, { maximum: 20 })
    )
}

function mediaLabel(media) {
  return media ? `${media.originalFilename} (${media.contentType})` : "None"
}

function buildPhaseConfigurationView(sessionId, data) {
  const phase = data.workingPhase
  return {
    content: `${phase.phaseName}\n${phase.phaseTimeUtc} UTC\nLocal: ` +
      `<t:${Math.floor(new Date(`${data.firstOccurrenceDate}T${phase.phaseTimeUtc}:00.000Z`).getTime() / 1000)}:t>\n\n` +
      `Pre-alert: ${phase.preAlertMinutes === null ? "None" : `${phase.preAlertMinutes} minutes`}\n` +
      `Exact alert: ${phase.announceExact ? "On" : "Off"}\n` +
      `Pre-alert media: ${mediaLabel(phase.preAlertMedia)}\n` +
      `Exact-time media: ${mediaLabel(phase.exactMedia)}\n\n${CRITICAL_ALERT_WARNING}`,
    components: [
      new ActionRowBuilder().addComponents(
        new StringSelectMenuBuilder().setCustomId(`${STATE_EVENT_IDS.phasePrePrefix}${sessionId}`)
          .setPlaceholder("Pre-alert")
          .addOptions(PRE_ALERTS.map(([label, value]) => ({
            label, value,
            default: (phase.preAlertMinutes === null ? "none" : String(phase.preAlertMinutes)) === value
          })))
      ),
      new ActionRowBuilder().addComponents(
        new ButtonBuilder().setCustomId(`${STATE_EVENT_IDS.phaseExactPrefix}${sessionId}`)
          .setLabel(`Exact alert: ${phase.announceExact ? "On" : "Off"}`)
          .setStyle(phase.announceExact ? ButtonStyle.Success : ButtonStyle.Secondary),
        new ButtonBuilder().setCustomId(`${STATE_EVENT_IDS.phaseMessagesPrefix}${sessionId}`)
          .setLabel("Messages").setStyle(ButtonStyle.Secondary)
      ),
      new ActionRowBuilder().addComponents(
        new ButtonBuilder().setCustomId(`${STATE_EVENT_IDS.phaseMediaPrefix}${sessionId}:pre_alert`)
          .setLabel("Pre-alert media").setStyle(ButtonStyle.Secondary),
        new ButtonBuilder().setCustomId(`${STATE_EVENT_IDS.phaseMediaPrefix}${sessionId}:exact`)
          .setLabel("Exact-time media").setStyle(ButtonStyle.Secondary)
      ),
      new ActionRowBuilder().addComponents(
        new ButtonBuilder().setCustomId(`${STATE_EVENT_IDS.phaseSavePrefix}${sessionId}`)
          .setLabel("Save phase").setStyle(ButtonStyle.Success),
        new ButtonBuilder().setCustomId(`${STATE_EVENT_IDS.phaseCancelPrefix}${sessionId}`)
          .setLabel("Cancel phase").setStyle(ButtonStyle.Secondary)
      )
    ]
  }
}

function buildMessagesModal(sessionId, phase) {
  return new ModalBuilder()
    .setCustomId(`${STATE_EVENT_IDS.phaseMessagesModalPrefix}${sessionId}`)
    .setTitle("Phase messages")
    .addComponents(
      input("pre", "Pre-alert message (optional)", phase.preAlertMessage,
        { paragraph: true, required: false, maximum: 500 }),
      input("exact", "Exact-time message (optional)", phase.exactMessage,
        { paragraph: true, required: false, maximum: 500 })
    )
}

function buildMediaView(sessionId, data, kind) {
  const key = kind === "pre_alert" ? "preAlertMedia" : "exactMedia"
  return {
    content: `${kind === "pre_alert" ? "Pre-alert" : "Exact-time"} media\n\nCurrent: ` +
      `${mediaLabel(data.workingPhase[key])}\nPNG, JPEG, GIF or WebP; maximum 8 MB.`,
    components: [new ActionRowBuilder().addComponents(
      new ButtonBuilder().setCustomId(`${STATE_EVENT_IDS.phaseMediaReplacePrefix}${sessionId}:${kind}`)
        .setLabel(data.workingPhase[key] ? "Replace" : "Upload").setStyle(ButtonStyle.Primary),
      new ButtonBuilder().setCustomId(`${STATE_EVENT_IDS.phaseMediaRemovePrefix}${sessionId}:${kind}`)
        .setLabel("Remove").setStyle(ButtonStyle.Danger).setDisabled(!data.workingPhase[key]),
      new ButtonBuilder().setCustomId(`${STATE_EVENT_IDS.phaseMediaBackPrefix}${sessionId}`)
        .setLabel("Back").setStyle(ButtonStyle.Secondary)
    )]
  }
}

function buildMediaUploadModal(sessionId, kind) {
  return new ModalBuilder()
    .setCustomId(`${STATE_EVENT_IDS.phaseMediaUploadPrefix}${sessionId}:${kind}`)
    .setTitle(kind === "pre_alert" ? "Pre-alert media" : "Exact-time media")
    .addLabelComponents(new LabelBuilder().setLabel("Phase media")
      .setDescription("PNG, JPEG, GIF or WebP; maximum 8 MB")
      .setFileUploadComponent(new FileUploadBuilder().setCustomId("img").setRequired(true)
        .setMinValues(1).setMaxValues(1)))
}

function buildReviewView(sessionId, data) {
  const draft = validateStateEventDraft(data)
  return {
    content: `State ${data.stateNumber}\n${draft.eventName}\nFirst occurrence: ${draft.firstOccurrenceDate}\n` +
      `Recurrence: ${recurrenceLabel(draft.recurrenceDays)}\n\nPhases\n\n${phaseLines(draft.phases)}`,
    components: [new ActionRowBuilder().addComponents(
      new ButtonBuilder().setCustomId(`${STATE_EVENT_IDS.confirmPrefix}${sessionId}`)
        .setLabel(data.mode === "edit" ? "Apply changes" : "Confirm")
        .setStyle(ButtonStyle.Success),
      new ButtonBuilder().setCustomId(`${STATE_EVENT_IDS.phasesPrefix}${sessionId}`)
        .setLabel("Back / Edit").setStyle(ButtonStyle.Secondary),
      new ButtonBuilder().setCustomId(`${STATE_EVENT_IDS.cancelPrefix}${sessionId}`)
        .setLabel("Cancel").setStyle(ButtonStyle.Secondary)
    )]
  }
}

function buildEventListView(events, page, total, sessionId) {
  const pages = Math.max(1, Math.ceil(total / EVENTS_PER_PAGE))
  const tokenMap = {}
  const options = events.map(event => {
    const token = opaqueToken()
    tokenMap[token] = String(event.id)
    return {
      label: event.event_name.slice(0, 100),
      description: `${event.status === "paused" ? "Paused" : "Active"} - ${recurrenceLabel(event.recurrence_days)}`,
      value: token
    }
  })
  const components = []
  if (options.length) {
    components.push(new ActionRowBuilder().addComponents(
      new StringSelectMenuBuilder().setCustomId(`${STATE_EVENT_IDS.eventSelectPrefix}${sessionId}`)
        .setPlaceholder("Select a state event").addOptions(options)
    ))
  }
  components.push(new ActionRowBuilder().addComponents(
    new ButtonBuilder().setCustomId(`${STATE_EVENT_IDS.listPrefix}${Math.max(0, page - 1)}`)
      .setLabel("Previous").setStyle(ButtonStyle.Secondary).setDisabled(page === 0),
    new ButtonBuilder().setCustomId(`${STATE_EVENT_IDS.listPrefix}${page + 1}`)
      .setLabel("Next").setStyle(ButtonStyle.Secondary).setDisabled(page >= pages - 1),
    new ButtonBuilder().setCustomId(STATE_EVENT_IDS.home).setLabel("Back")
      .setStyle(ButtonStyle.Secondary)
  ))
  return {
    view: {
      content: `State events - page ${page + 1} of ${pages}\n\n` +
        (events.length ? events.map(event =>
          `${event.status === "paused" ? "[PAUSED]" : "[ACTIVE]"} ${event.event_name} - ${event.phases.length} phase${event.phases.length === 1 ? "" : "s"}`
        ).join("\n") : "No active or paused state events."),
      components
    },
    tokenMap
  }
}

function buildManageView(sessionId, data) {
  const event = data.selectedEvent
  const paused = event.status === "paused"
  return {
    content: `Manage state event\n\nState ${event.state_number}\n${event.event_name}\n` +
      `${paused ? "Paused" : "Active"} - ${recurrenceLabel(event.recurrence_days)}\n` +
      `${event.phases.length} phase${event.phases.length === 1 ? "" : "s"}`,
    components: [
      new ActionRowBuilder().addComponents(
        new ButtonBuilder().setCustomId(`${STATE_EVENT_IDS.previewPrefix}${sessionId}`)
          .setLabel("Preview next occurrence").setStyle(ButtonStyle.Primary),
        new ButtonBuilder().setCustomId(`${STATE_EVENT_IDS.editDetailsPrefix}${sessionId}`)
          .setLabel("Edit event details").setStyle(ButtonStyle.Secondary),
        new ButtonBuilder().setCustomId(`${STATE_EVENT_IDS.managePhasesPrefix}${sessionId}`)
          .setLabel("Manage phases").setStyle(ButtonStyle.Secondary)
      ),
      new ActionRowBuilder().addComponents(
        new ButtonBuilder().setCustomId(`${STATE_EVENT_IDS.statusPrefix}${sessionId}:${paused ? "resume" : "pause"}`)
          .setLabel(paused ? "Resume" : "Pause").setStyle(ButtonStyle.Secondary),
        new ButtonBuilder().setCustomId(`${STATE_EVENT_IDS.statusPrefix}${sessionId}:delete`)
          .setLabel("Delete").setStyle(ButtonStyle.Danger),
        new ButtonBuilder().setCustomId(`${STATE_EVENT_IDS.testPrefix}${sessionId}`)
          .setLabel("Test announcement").setStyle(ButtonStyle.Secondary),
        new ButtonBuilder().setCustomId(`${STATE_EVENT_IDS.listPrefix}${data.page || 0}`)
          .setLabel("Back").setStyle(ButtonStyle.Secondary)
      )
    ]
  }
}

function nextOccurrencePhases(event, now = new Date()) {
  const anchor = new Date(`${event.first_occurrence_date}T00:00:00.000Z`)
  const intervalDays = event.recurrence_days || 1
  const rangeStart = new Date(Math.min(now.getTime(), anchor.getTime()))
  const rangeEnd = new Date(Math.max(now.getTime() + (intervalDays + 2) * 86400000,
    anchor.getTime() + 2 * 86400000))
  const future = getStateEventOccurrencesInRange(event, rangeStart, rangeEnd)
    .filter(item => item.occurrenceAt >= now)
  if (!future.length) return []
  const index = future[0].occurrenceIndex
  return future.filter(item => item.occurrenceIndex === index)
}

function buildOccurrencePreviewView(sessionId, event, now = new Date()) {
  const occurrences = nextOccurrencePhases(event, now)
  if (!occurrences.length) {
    return {
      content: `State ${event.state_number} - ${event.event_name}\n\nNo future occurrence remains.`,
      components: [new ActionRowBuilder().addComponents(
        new ButtonBuilder().setCustomId(`${STATE_EVENT_IDS.managePrefix}${sessionId}`)
          .setLabel("Back").setStyle(ButtonStyle.Secondary)
      )]
    }
  }
  const date = occurrences[0].occurrenceAt.toISOString().slice(0, 10)
  const byId = new Map(event.phases.map(phase => [String(phase.id), phaseDraft(phase)]))
  return {
    content: `State ${event.state_number} - ${event.event_name}\nNext occurrence: ${date}\n\n` +
      occurrences.map(item => {
        const phase = byId.get(String(item.phaseId))
        return `${item.phaseName}\n${item.occurrenceAt.toISOString().slice(11, 16)} UTC\n` +
          `Local: <t:${Math.floor(item.occurrenceAt.getTime() / 1000)}:t>\n` +
          `Pre-alert: ${phase.preAlertMinutes === null ? "None" : `${phase.preAlertMinutes} min`}\n` +
          `Exact alert: ${phase.announceExact ? "On" : "Off"}`
      }).join("\n\n"),
    components: [new ActionRowBuilder().addComponents(
      new ButtonBuilder().setCustomId(`${STATE_EVENT_IDS.managePrefix}${sessionId}`)
        .setLabel("Back").setStyle(ButtonStyle.Secondary)
    )]
  }
}

function buildStatusConfirmation(sessionId, action, event) {
  const detail = action === "delete"
    ? "Future alerts and roundup entries will stop. Existing sent history is retained."
    : action === "pause"
      ? "Future alerts and roundup entries will stop until resumed. The original anchor and history are retained."
      : "Future scheduling resumes from the original anchor. Expired alerts will not be replayed."
  return {
    content: `${action[0].toUpperCase()}${action.slice(1)} ${event.event_name}?\n\n${detail}`,
    components: [new ActionRowBuilder().addComponents(
      new ButtonBuilder().setCustomId(`${STATE_EVENT_IDS.statusConfirmPrefix}${sessionId}:${action}`)
        .setLabel(`Confirm ${action}`).setStyle(ButtonStyle.Danger),
      new ButtonBuilder().setCustomId(`${STATE_EVENT_IDS.managePrefix}${sessionId}`)
        .setLabel("Cancel").setStyle(ButtonStyle.Secondary)
    )]
  }
}

function buildTestView(sessionId, data) {
  const phases = data.selectedEvent.phases.map(phaseDraft)
  const selected = Number.isInteger(data.testPhaseIndex) ? data.testPhaseIndex : null
  const components = [new ActionRowBuilder().addComponents(
    new StringSelectMenuBuilder().setCustomId(`${STATE_EVENT_IDS.testPhasePrefix}${sessionId}`)
      .setPlaceholder("Choose phase").addOptions(phases.map((phase, index) => ({
        label: phase.phaseName.slice(0, 100), value: String(index),
        description: `${phase.phaseTimeUtc} UTC`, default: selected === index
      })))
  )]
  if (selected !== null) {
    const phase = phases[selected]
    components.push(new ActionRowBuilder().addComponents(
      new ButtonBuilder().setCustomId(`${STATE_EVENT_IDS.testKindPrefix}${sessionId}:pre_alert`)
        .setLabel("Send pre-alert test").setStyle(ButtonStyle.Primary)
        .setDisabled(phase.preAlertMinutes === null),
      new ButtonBuilder().setCustomId(`${STATE_EVENT_IDS.testKindPrefix}${sessionId}:exact`)
        .setLabel("Send exact-time test").setStyle(ButtonStyle.Primary)
        .setDisabled(!phase.announceExact),
      new ButtonBuilder().setCustomId(`${STATE_EVENT_IDS.managePrefix}${sessionId}`)
        .setLabel("Back").setStyle(ButtonStyle.Secondary)
    ))
  }
  return {
    content: "Test state-event announcement\n\nChoose one phase and version. The test goes only to this state destination and does not change production history.",
    components
  }
}

async function requireDestination(schedulerRepository, guildId) {
  const destination = await schedulerRepository.getStateDestination(guildId)
  if (!destination?.enabled) {
    throw new InteractionSessionError("State events are available only in an enabled state destination for this game profile.")
  }
  return destination
}

async function renderList(interaction, stateRepository, context, page = 0, sessionStore = stateEventSessions) {
  const safePage = Math.max(0, Number(page) || 0)
  const result = await stateRepository.listStateEvents(interaction.guildId, {
    limit: EVENTS_PER_PAGE,
    offset: safePage * EVENTS_PER_PAGE
  })
  const sessionId = sessionStore.create(context, { page: safePage })
  const built = buildEventListView(result.events, safePage, result.total, sessionId)
  sessionStore.update(sessionId, context, { eventTokenMap: built.tokenMap })
  await interaction.editReply(built.view)
}

async function handleStateEventModalOpeningInteraction(
  interaction,
  { health, sessionStore = stateEventSessions }
) {
  const customId = String(interaction.customId || "")
  if (!customId.startsWith(STATE_EVENT_IDS.prefix)) return false
  const context = sessionContext(interaction, health)

  if (interaction.isButton?.() && customId === STATE_EVENT_IDS.newEvent) {
    const sessionId = sessionStore.create(context, {
      mode: "create", recurrenceDays: undefined, phases: [], selectedPhaseIndex: null
    })
    await interaction.showModal(buildBasicModal(sessionId))
    return true
  }
  if (interaction.isButton?.() && customId.startsWith(STATE_EVENT_IDS.editDetailsPrefix)) {
    const sessionId = suffix(customId, STATE_EVENT_IDS.editDetailsPrefix)
    const session = sessionStore.get(sessionId, context)
    Object.assign(session.data, eventDraft(session.data.selectedEvent), {
      page: session.data.page,
      selectedEvent: session.data.selectedEvent
    })
    await interaction.showModal(buildBasicModal(sessionId, session.data))
    return true
  }
  for (const [prefix, mode] of [
    [STATE_EVENT_IDS.phaseAddPrefix, "add"],
    [STATE_EVENT_IDS.phaseEditPrefix, "edit"]
  ]) {
    if (interaction.isButton?.() && customId.startsWith(prefix)) {
      const sessionId = suffix(customId, prefix)
      const session = sessionStore.get(sessionId, context)
      const phase = mode === "edit" ? session.data.phases?.[session.data.selectedPhaseIndex] : {}
      if (mode === "edit" && !phase) throw new InteractionSessionError("Select a phase to edit.")
      await interaction.showModal(buildPhaseModal(sessionId, mode, phase))
      return true
    }
  }
  if (interaction.isButton?.() && customId.startsWith(STATE_EVENT_IDS.phaseMessagesPrefix)) {
    const sessionId = suffix(customId, STATE_EVENT_IDS.phaseMessagesPrefix)
    const session = sessionStore.get(sessionId, context)
    if (!session.data.workingPhase) throw new InteractionSessionError("That phase editor expired.")
    await interaction.showModal(buildMessagesModal(sessionId, session.data.workingPhase))
    return true
  }
  if (interaction.isButton?.() && customId.startsWith(STATE_EVENT_IDS.phaseMediaReplacePrefix)) {
    const [sessionId, kind] = suffix(customId, STATE_EVENT_IDS.phaseMediaReplacePrefix).split(":")
    sessionStore.get(sessionId, context)
    if (!["pre_alert", "exact"].includes(kind)) throw new InteractionSessionError("Invalid media type.")
    await interaction.showModal(buildMediaUploadModal(sessionId, kind))
    return true
  }
  return false
}

async function handleStateEventInteraction(interaction, {
  schedulerRepository,
  stateRepository,
  health,
  now = () => new Date(),
  targetResolver,
  formatter = formatStateEventDelivery,
  sessionStore = stateEventSessions
}) {
  const customId = String(interaction.customId || "")
  if (!customId.startsWith(STATE_EVENT_IDS.prefix)) return false
  const context = sessionContext(interaction, health)
  const destination = await requireDestination(schedulerRepository, interaction.guildId)

  if (customId === STATE_EVENT_IDS.home) {
    await interaction.editReply(buildStateEventHome(destination))
    return true
  }

  if (interaction.isModalSubmit?.() && customId.startsWith(STATE_EVENT_IDS.basicModalPrefix)) {
    const sessionId = suffix(customId, STATE_EVENT_IDS.basicModalPrefix)
    sessionStore.get(sessionId, context)
    const parsedDate = parseIsoDate(interaction.fields.getTextInputValue("date"))
    const session = sessionStore.update(sessionId, context, {
      eventName: normalizeStateEventName(interaction.fields.getTextInputValue("name")),
      firstOccurrenceDate: parsedDate.value,
      stateNumber: destination.state_number
    })
    await interaction.editReply(buildRecurrenceView(sessionId, session.data))
    return true
  }

  if (interaction.isStringSelectMenu?.() && customId.startsWith(STATE_EVENT_IDS.recurrencePrefix)) {
    const sessionId = suffix(customId, STATE_EVENT_IDS.recurrencePrefix)
    const session = sessionStore.update(sessionId, context, {
      recurrenceDays: normalizeRecurrenceDays(interaction.values[0])
    })
    await interaction.editReply(session.data.mode === "edit"
      ? buildReviewView(sessionId, session.data)
      : buildPhaseManagerView(sessionId, session.data))
    return true
  }

  if (interaction.isButton?.() && customId.startsWith(STATE_EVENT_IDS.phasesPrefix)) {
    const sessionId = suffix(customId, STATE_EVENT_IDS.phasesPrefix)
    const session = sessionStore.get(sessionId, context)
    await interaction.editReply(buildPhaseManagerView(sessionId, session.data))
    return true
  }

  if (interaction.isStringSelectMenu?.() && customId.startsWith(STATE_EVENT_IDS.phaseSelectPrefix)) {
    const sessionId = suffix(customId, STATE_EVENT_IDS.phaseSelectPrefix)
    const session = sessionStore.get(sessionId, context)
    const index = Number(interaction.values[0])
    if (!session.data.phases?.[index]) throw new InteractionSessionError("That phase is unavailable.")
    sessionStore.update(sessionId, context, { selectedPhaseIndex: index })
    await interaction.editReply(buildPhaseManagerView(sessionId, session.data))
    return true
  }

  if (interaction.isModalSubmit?.() && customId.startsWith(STATE_EVENT_IDS.phaseModalPrefix)) {
    const details = suffix(customId, STATE_EVENT_IDS.phaseModalPrefix)
    const split = details.lastIndexOf(":")
    const sessionId = details.slice(0, split)
    const mode = details.slice(split + 1)
    const session = sessionStore.get(sessionId, context)
    const current = mode === "edit" ? session.data.phases?.[session.data.selectedPhaseIndex] : {}
    const workingPhase = normalizeStatePhase({
      ...current,
      phaseName: interaction.fields.getTextInputValue("name"),
      phaseTimeUtc: interaction.fields.getTextInputValue("time"),
      announceExact: current?.announceExact ?? true
    })
    sessionStore.update(sessionId, context, { workingPhase, workingMode: mode })
    await interaction.editReply(buildPhaseConfigurationView(sessionId, session.data))
    return true
  }

  if (interaction.isStringSelectMenu?.() && customId.startsWith(STATE_EVENT_IDS.phasePrePrefix)) {
    const sessionId = suffix(customId, STATE_EVENT_IDS.phasePrePrefix)
    const session = sessionStore.get(sessionId, context)
    session.data.workingPhase.preAlertMinutes = interaction.values[0] === "none"
      ? null : Number(interaction.values[0])
    await interaction.editReply(buildPhaseConfigurationView(sessionId, session.data))
    return true
  }

  if (interaction.isButton?.() && customId.startsWith(STATE_EVENT_IDS.phaseExactPrefix)) {
    const sessionId = suffix(customId, STATE_EVENT_IDS.phaseExactPrefix)
    const session = sessionStore.get(sessionId, context)
    session.data.workingPhase.announceExact = !session.data.workingPhase.announceExact
    await interaction.editReply(buildPhaseConfigurationView(sessionId, session.data))
    return true
  }

  if (interaction.isModalSubmit?.() && customId.startsWith(STATE_EVENT_IDS.phaseMessagesModalPrefix)) {
    const sessionId = suffix(customId, STATE_EVENT_IDS.phaseMessagesModalPrefix)
    const session = sessionStore.get(sessionId, context)
    const phase = normalizeStatePhase({
      ...session.data.workingPhase,
      preAlertMessage: interaction.fields.getTextInputValue("pre"),
      exactMessage: interaction.fields.getTextInputValue("exact")
    })
    sessionStore.update(sessionId, context, { workingPhase: phase })
    await interaction.editReply(buildPhaseConfigurationView(sessionId, session.data))
    return true
  }

  if (interaction.isButton?.() && customId.startsWith(STATE_EVENT_IDS.phaseMediaPrefix)) {
    const [sessionId, kind] = suffix(customId, STATE_EVENT_IDS.phaseMediaPrefix).split(":")
    const session = sessionStore.get(sessionId, context)
    await interaction.editReply(buildMediaView(sessionId, session.data, kind))
    return true
  }

  if (interaction.isModalSubmit?.() && customId.startsWith(STATE_EVENT_IDS.phaseMediaUploadPrefix)) {
    const [sessionId, kind] = suffix(customId, STATE_EVENT_IDS.phaseMediaUploadPrefix).split(":")
    const session = sessionStore.get(sessionId, context)
    const attachment = interaction.fields.getUploadedFiles("img")?.first?.()
    if (!attachment) throw new InteractionSessionError("Choose one phase image.")
    const media = await downloadEventImage(attachment)
    session.data.workingPhase[kind === "pre_alert" ? "preAlertMedia" : "exactMedia"] = media
    await interaction.editReply(buildMediaView(sessionId, session.data, kind))
    return true
  }

  if (interaction.isButton?.() && customId.startsWith(STATE_EVENT_IDS.phaseMediaRemovePrefix)) {
    const [sessionId, kind] = suffix(customId, STATE_EVENT_IDS.phaseMediaRemovePrefix).split(":")
    const session = sessionStore.get(sessionId, context)
    session.data.workingPhase[kind === "pre_alert" ? "preAlertMedia" : "exactMedia"] = null
    await interaction.editReply(buildMediaView(sessionId, session.data, kind))
    return true
  }

  if (interaction.isButton?.() && customId.startsWith(STATE_EVENT_IDS.phaseMediaBackPrefix)) {
    const sessionId = suffix(customId, STATE_EVENT_IDS.phaseMediaBackPrefix)
    const session = sessionStore.get(sessionId, context)
    await interaction.editReply(buildPhaseConfigurationView(sessionId, session.data))
    return true
  }

  if (interaction.isButton?.() && customId.startsWith(STATE_EVENT_IDS.phaseSavePrefix)) {
    const sessionId = suffix(customId, STATE_EVENT_IDS.phaseSavePrefix)
    const session = sessionStore.get(sessionId, context)
    const phase = normalizeStatePhase(session.data.workingPhase)
    const phases = [...(session.data.phases || [])]
    const index = session.data.workingMode === "edit" ? session.data.selectedPhaseIndex : phases.length
    const duplicate = phases.some((item, itemIndex) => itemIndex !== index
      && item.phaseName.toLowerCase() === phase.phaseName.toLowerCase())
    if (duplicate) throw new InteractionSessionError(`Duplicate phase name: ${phase.phaseName}.`)
    phases[index] = phase
    sessionStore.update(sessionId, context, {
      phases: sortPhases(phases), selectedPhaseIndex: index,
      workingPhase: null, workingMode: null
    })
    await interaction.editReply(buildPhaseManagerView(sessionId, session.data))
    return true
  }

  if (interaction.isButton?.() && customId.startsWith(STATE_EVENT_IDS.phaseCancelPrefix)) {
    const sessionId = suffix(customId, STATE_EVENT_IDS.phaseCancelPrefix)
    const session = sessionStore.update(sessionId, context, { workingPhase: null, workingMode: null })
    await interaction.editReply(buildPhaseManagerView(sessionId, session.data))
    return true
  }

  if (interaction.isButton?.() && customId.startsWith(STATE_EVENT_IDS.phaseRemovePrefix)) {
    const sessionId = suffix(customId, STATE_EVENT_IDS.phaseRemovePrefix)
    const session = sessionStore.get(sessionId, context)
    if (session.data.phases.length <= 1) {
      throw new InteractionSessionError("A state event must retain at least one phase.")
    }
    const phases = session.data.phases.filter((_, index) => index !== session.data.selectedPhaseIndex)
    sessionStore.update(sessionId, context, { phases: sortPhases(phases), selectedPhaseIndex: null })
    await interaction.editReply(buildPhaseManagerView(sessionId, session.data))
    return true
  }

  if (interaction.isButton?.() && customId.startsWith(STATE_EVENT_IDS.reviewPrefix)) {
    const sessionId = suffix(customId, STATE_EVENT_IDS.reviewPrefix)
    const session = sessionStore.get(sessionId, context)
    await interaction.editReply(buildReviewView(sessionId, session.data))
    return true
  }

  if (interaction.isButton?.() && customId.startsWith(STATE_EVENT_IDS.confirmPrefix)) {
    const sessionId = suffix(customId, STATE_EVENT_IDS.confirmPrefix)
    const session = sessionStore.get(sessionId, context)
    await requireDestination(schedulerRepository, interaction.guildId)
    const draft = validateStateEventDraft(session.data)
    const saved = session.data.mode === "edit"
      ? await stateRepository.updateStateEvent({
        stateGuildId: interaction.guildId, eventId: session.data.eventId, event: draft
      })
      : await stateRepository.createStateEvent({
        stateGuildId: interaction.guildId,
        createdByUserId: interaction.user.id,
        createdByBotInstance: health.botInstanceName,
        event: draft
      })
    if (!saved) throw new InteractionSessionError("That state event is no longer available.")
    sessionStore.complete(sessionId, context)
    await interaction.editReply({
      content: `${session.data.mode === "edit" ? "Updated" : "Created"} ${draft.eventName}. No public message was posted.`,
      components: [new ActionRowBuilder().addComponents(
        new ButtonBuilder().setCustomId(`${STATE_EVENT_IDS.listPrefix}0`)
          .setLabel("View state events").setStyle(ButtonStyle.Secondary)
      )]
    })
    return true
  }

  if (interaction.isButton?.() && customId.startsWith(STATE_EVENT_IDS.cancelPrefix)) {
    const sessionId = suffix(customId, STATE_EVENT_IDS.cancelPrefix)
    sessionStore.cancel(sessionId, context)
    await interaction.editReply({ content: "State-event changes cancelled. Nothing was saved.", components: [] })
    return true
  }

  if (interaction.isButton?.() && customId.startsWith(STATE_EVENT_IDS.listPrefix)) {
    await renderList(interaction, stateRepository, context,
      Number(suffix(customId, STATE_EVENT_IDS.listPrefix)), sessionStore)
    return true
  }

  if (interaction.isStringSelectMenu?.() && customId.startsWith(STATE_EVENT_IDS.eventSelectPrefix)) {
    const sessionId = suffix(customId, STATE_EVENT_IDS.eventSelectPrefix)
    const session = sessionStore.get(sessionId, context)
    const eventId = session.data.eventTokenMap?.[interaction.values[0]]
    if (!eventId) throw new InteractionSessionError("That state-event control expired.")
    const event = await stateRepository.getStateEvent(interaction.guildId, eventId)
    if (!event) throw new InteractionSessionError("That state event is no longer available.")
    sessionStore.update(sessionId, context, { selectedEvent: event })
    await interaction.editReply(buildManageView(sessionId, session.data))
    return true
  }

  if (interaction.isButton?.() && customId.startsWith(STATE_EVENT_IDS.managePrefix)) {
    const sessionId = suffix(customId, STATE_EVENT_IDS.managePrefix)
    const session = sessionStore.get(sessionId, context)
    const event = await stateRepository.getStateEvent(interaction.guildId, session.data.selectedEvent.id)
    if (!event) throw new InteractionSessionError("That state event is no longer available.")
    sessionStore.update(sessionId, context, { selectedEvent: event })
    await interaction.editReply(buildManageView(sessionId, session.data))
    return true
  }

  if (interaction.isButton?.() && customId.startsWith(STATE_EVENT_IDS.managePhasesPrefix)) {
    const sessionId = suffix(customId, STATE_EVENT_IDS.managePhasesPrefix)
    const session = sessionStore.get(sessionId, context)
    Object.assign(session.data, eventDraft(session.data.selectedEvent), {
      page: session.data.page, selectedEvent: session.data.selectedEvent
    })
    await interaction.editReply(buildPhaseManagerView(sessionId, session.data))
    return true
  }

  if (interaction.isButton?.() && customId.startsWith(STATE_EVENT_IDS.previewPrefix)) {
    const sessionId = suffix(customId, STATE_EVENT_IDS.previewPrefix)
    const session = sessionStore.get(sessionId, context)
    await interaction.editReply(buildOccurrencePreviewView(sessionId, session.data.selectedEvent, now()))
    return true
  }

  if (interaction.isButton?.() && customId.startsWith(STATE_EVENT_IDS.statusPrefix)) {
    const [sessionId, action] = suffix(customId, STATE_EVENT_IDS.statusPrefix).split(":")
    const session = sessionStore.get(sessionId, context)
    if (!["pause", "resume", "delete"].includes(action)) throw new InteractionSessionError("Invalid action.")
    await interaction.editReply(buildStatusConfirmation(sessionId, action, session.data.selectedEvent))
    return true
  }

  if (interaction.isButton?.() && customId.startsWith(STATE_EVENT_IDS.statusConfirmPrefix)) {
    const [sessionId, action] = suffix(customId, STATE_EVENT_IDS.statusConfirmPrefix).split(":")
    const session = sessionStore.get(sessionId, context)
    await requireDestination(schedulerRepository, interaction.guildId)
    const status = action === "pause" ? "paused" : action === "resume" ? "active" : "deleted"
    const changed = await stateRepository.setStateEventStatus({
      stateGuildId: interaction.guildId,
      eventId: session.data.selectedEvent.id,
      status
    })
    if (!changed) throw new InteractionSessionError("That state event is no longer available.")
    sessionStore.complete(sessionId, context)
    await interaction.editReply({
      content: `${action === "delete" ? "Deleted" : action === "pause" ? "Paused" : "Resumed"} ${changed.event_name}.`,
      components: [new ActionRowBuilder().addComponents(
        new ButtonBuilder().setCustomId(`${STATE_EVENT_IDS.listPrefix}0`)
          .setLabel("Back to state events").setStyle(ButtonStyle.Secondary)
      )]
    })
    return true
  }

  if (interaction.isButton?.() && customId.startsWith(STATE_EVENT_IDS.testPrefix)) {
    const sessionId = suffix(customId, STATE_EVENT_IDS.testPrefix)
    const session = sessionStore.update(sessionId, context, { testPhaseIndex: null })
    await interaction.editReply(buildTestView(sessionId, session.data))
    return true
  }

  if (interaction.isStringSelectMenu?.() && customId.startsWith(STATE_EVENT_IDS.testPhasePrefix)) {
    const sessionId = suffix(customId, STATE_EVENT_IDS.testPhasePrefix)
    const session = sessionStore.update(sessionId, context, { testPhaseIndex: Number(interaction.values[0]) })
    await interaction.editReply(buildTestView(sessionId, session.data))
    return true
  }

  if (interaction.isButton?.() && customId.startsWith(STATE_EVENT_IDS.testKindPrefix)) {
    const [sessionId, kind] = suffix(customId, STATE_EVENT_IDS.testKindPrefix).split(":")
    const session = sessionStore.get(sessionId, context)
    await requireDestination(schedulerRepository, interaction.guildId)
    const event = session.data.selectedEvent
    const phase = phaseDraft(event.phases[session.data.testPhaseIndex])
    if (!phase || !["pre_alert", "exact"].includes(kind)) {
      throw new InteractionSessionError("Choose a valid test announcement.")
    }
    if (kind === "pre_alert" && phase.preAlertMinutes === null) {
      throw new InteractionSessionError("That phase has no pre-alert to test.")
    }
    if (kind === "exact" && !phase.announceExact) {
      throw new InteractionSessionError("That phase has no exact-time announcement to test.")
    }
    const occurrenceAt = new Date(`${event.first_occurrence_date}T${phase.phaseTimeUtc}:00.000Z`)
    const media = kind === "pre_alert" ? phase.preAlertMedia : phase.exactMedia
    const prepared = media ? prepareStoredEventImage(media) : null
    const target = await targetResolver(interaction.client, interaction.guildId,
      destination.state_roundup_channel_id, { requireAttachments: Boolean(prepared) })
    const message = formatter({
      claim: { deliveryKind: kind, occurrenceAt },
      stateEvent: { stateNumber: destination.state_number, eventName: event.event_name },
      phase: {
        name: phase.phaseName,
        preAlertMinutes: phase.preAlertMinutes,
        preAlertMessage: phase.preAlertMessage,
        announceExact: phase.announceExact,
        exactMessage: phase.exactMessage
      }
    }, { imageFilename: prepared?.filename || null, test: true })
    if (prepared) message.files = [prepared.file]
    await target.channel.send(message)
    await interaction.editReply({
      content: `Sent one TEST announcement to <#${destination.state_roundup_channel_id}>. Production scheduling and history were not changed.`,
      components: [new ActionRowBuilder().addComponents(
        new ButtonBuilder().setCustomId(`${STATE_EVENT_IDS.managePrefix}${sessionId}`)
          .setLabel("Back").setStyle(ButtonStyle.Secondary)
      )]
    })
    return true
  }

  return customId === STATE_EVENT_IDS.newEvent
}

module.exports = {
  STATE_EVENT_IDS,
  RECURRENCES,
  PRE_ALERTS,
  CRITICAL_ALERT_WARNING,
  stateEventSessions,
  recurrenceLabel,
  phaseDraft,
  eventDraft,
  sortPhases,
  buildStateEventHome,
  buildBasicModal,
  buildRecurrenceView,
  buildPhaseManagerView,
  buildPhaseModal,
  buildPhaseConfigurationView,
  buildMessagesModal,
  buildMediaView,
  buildMediaUploadModal,
  buildReviewView,
  buildEventListView,
  buildManageView,
  nextOccurrencePhases,
  buildOccurrencePreviewView,
  buildStatusConfirmation,
  buildTestView,
  handleStateEventModalOpeningInteraction,
  handleStateEventInteraction
}
