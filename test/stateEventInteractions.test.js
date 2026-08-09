const test = require("node:test")
const assert = require("node:assert/strict")

const { buildHomeView, handleEventSchedulerInteraction } = require("../src/eventSchedulerInteractions")
const { validateAttachmentMetadata } = require("../src/eventImage")
const { InteractionSessionStore } = require("../src/interactionSessions")
const {
  STATE_EVENT_IDS,
  RECURRENCES,
  PRE_ALERTS,
  CRITICAL_ALERT_WARNING,
  buildEventListView,
  buildManageView,
  buildMediaView,
  buildOccurrencePreviewView,
  buildPhaseConfigurationView,
  buildPhaseManagerView,
  buildReviewView,
  buildStateEventHome,
  buildTestView,
  eventDraft,
  handleStateEventInteraction,
  handleStateEventModalOpeningInteraction,
  nextOccurrencePhases,
  phaseDraft,
  sortPhases
} = require("../src/stateEventInteractions")

const health = { available: true, gameProfile: "wos", botInstanceName: "rachie-wos" }
const context = { userId: "user-1", guildId: "state-guild", gameProfile: "wos" }
const destination = {
  state_guild_id: "state-guild",
  state_roundup_channel_id: "state-channel",
  state_number: "689",
  enabled: true
}

function storedEvent(overrides = {}) {
  return {
    id: "11111111-1111-4111-8111-111111111111",
    state_guild_id: "state-guild",
    state_number: "689",
    event_name: "SvS",
    first_occurrence_date: "2026-08-10",
    recurrence_days: 28,
    status: "active",
    phases: [{
      id: "22222222-2222-4222-8222-222222222222",
      phase_name: "Borders open",
      phase_time_utc: "10:00:00",
      pre_alert_minutes: 30,
      pre_alert_message: "Prepare",
      announce_exact: true,
      exact_message: "Open",
      pre_alert_media: null,
      exact_media: null,
      sort_order: 0
    }],
    ...overrides
  }
}

function interaction(customId, type = "button", overrides = {}) {
  return {
    customId,
    commandName: null,
    guildId: "state-guild",
    user: { id: "user-1" },
    values: [],
    fields: { getTextInputValue() { return "" } },
    client: {},
    isChatInputCommand: () => false,
    isButton: () => type === "button",
    isModalSubmit: () => type === "modal",
    isStringSelectMenu: () => type === "select",
    isChannelSelectMenu: () => false,
    async showModal(modal) { this.modal = modal },
    async deferUpdate() { this.deferred = true },
    async deferReply() { this.deferred = true },
    async editReply(payload) { this.edited = payload },
    ...overrides
  }
}

function options({ schedulerRepository = null, stateRepository = null, sessionStore = null } = {}) {
  return {
    schedulerRepository: schedulerRepository || {
      async getStateDestination(guildId) {
        return guildId === destination.state_guild_id ? destination : null
      }
    },
    stateRepository: stateRepository || {},
    health,
    now: () => new Date("2026-08-09T12:00:00Z"),
    targetResolver: async () => ({ channel: { async send() { return { id: "test-message" } } } }),
    sessionStore: sessionStore || new InteractionSessionStore()
  }
}

function componentIds(view) {
  return (view.components || []).flatMap(row => row.toJSON().components)
    .map(component => component.custom_id).filter(Boolean)
}

test("state-event home is exposed only for an enabled state destination", () => {
  const stateHome = buildHomeView(null, null, destination)
  const ordinaryHome = buildHomeView({ alliance_name: "YOU" }, { sharing_enabled: true }, null)
  assert.ok(componentIds(stateHome).includes(STATE_EVENT_IDS.home))
  assert.ok(!componentIds(ordinaryHome).includes(STATE_EVENT_IDS.home))
  assert.deepEqual(componentIds(buildStateEventHome(destination)).slice(0, 2), [
    STATE_EVENT_IDS.newEvent, `${STATE_EVENT_IDS.listPrefix}0`
  ])
})

test("guided creation supports one or many phases and persists only after confirmation", async () => {
  const sessionStore = new InteractionSessionStore()
  let created = null
  const stateRepository = {
    async createStateEvent(input) { created = input; return { id: "created" } }
  }
  const handlerOptions = options({ stateRepository, sessionStore })

  const opener = interaction(STATE_EVENT_IDS.newEvent)
  assert.equal(await handleStateEventModalOpeningInteraction(opener, { health, sessionStore }), true)
  assert.equal(opener.deferred, undefined)
  const sessionId = opener.modal.toJSON().custom_id.slice(STATE_EVENT_IDS.basicModalPrefix.length)

  const basic = interaction(`${STATE_EVENT_IDS.basicModalPrefix}${sessionId}`, "modal", {
    fields: { getTextInputValue(id) { return id === "name" ? "SvS" : "2026-08-10" } }
  })
  await handleStateEventInteraction(basic, handlerOptions)
  assert.match(basic.edited.content, /State event recurrence/)

  const recurrence = interaction(`${STATE_EVENT_IDS.recurrencePrefix}${sessionId}`, "select", {
    values: ["2"]
  })
  await handleStateEventInteraction(recurrence, handlerOptions)
  assert.match(recurrence.edited.content, /No phases configured/)
  assert.equal(componentIds(recurrence.edited)
    .find(id => id.startsWith(STATE_EVENT_IDS.reviewPrefix)), `${STATE_EVENT_IDS.reviewPrefix}${sessionId}`)

  const add = interaction(`${STATE_EVENT_IDS.phaseAddPrefix}${sessionId}`)
  await handleStateEventModalOpeningInteraction(add, { health, sessionStore })
  assert.equal(add.deferred, undefined)
  const addPhase = interaction(`${STATE_EVENT_IDS.phaseModalPrefix}${sessionId}:add`, "modal", {
    fields: { getTextInputValue(id) { return id === "name" ? "Borders open" : "10am" } }
  })
  await handleStateEventInteraction(addPhase, handlerOptions)
  assert.match(addPhase.edited.content, /10:00 UTC/)
  assert.match(addPhase.edited.content, /Local: <t:/)
  assert.match(addPhase.edited.content, new RegExp(CRITICAL_ALERT_WARNING.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")))

  await handleStateEventInteraction(interaction(
    `${STATE_EVENT_IDS.phasePrePrefix}${sessionId}`, "select", { values: ["30"] }
  ), handlerOptions)
  await handleStateEventInteraction(interaction(
    `${STATE_EVENT_IDS.phaseExactPrefix}${sessionId}`
  ), handlerOptions)
  const save = interaction(`${STATE_EVENT_IDS.phaseSavePrefix}${sessionId}`)
  await handleStateEventInteraction(save, handlerOptions)
  assert.match(save.edited.content, /Borders open - 10:00 UTC/)
  assert.match(save.edited.content, /Pre-alert: 30 min/)
  assert.match(save.edited.content, /Exact alert: Off/)

  const review = interaction(`${STATE_EVENT_IDS.reviewPrefix}${sessionId}`)
  await handleStateEventInteraction(review, handlerOptions)
  assert.match(review.edited.content, /Every 2 days/)
  assert.equal(created, null)

  const confirm = interaction(`${STATE_EVENT_IDS.confirmPrefix}${sessionId}`)
  await handleStateEventInteraction(confirm, handlerOptions)
  assert.equal(created.event.recurrenceDays, 2)
  assert.equal(created.event.phases.length, 1)
  assert.equal(created.event.phases[0].announceExact, false)
})

test("phase controls parse shared free-form times, order phases and reject duplicates", () => {
  const phases = sortPhases([
    phaseDraft({ phase_name: "Battle", phase_time_utc: "18:30", announce_exact: false }),
    phaseDraft({ phase_name: "Borders", phase_time_utc: "1000", announce_exact: true })
  ].map(phase => ({ ...phase, phaseTimeUtc: require("../src/timeParsing").parseUtcTime(phase.phaseTimeUtc) })))
  assert.deepEqual(phases.map(phase => phase.phaseTimeUtc), ["10:00", "18:30"])
  assert.deepEqual(PRE_ALERTS.map(([, value]) => value), ["none", "5", "10", "15", "20", "30"])
  assert.deepEqual(RECURRENCES.map(([, value]) => value), [
    "none", "2", "3", "7", "14", "21", "28", "35", "42"
  ])
  assert.throws(() => require("../src/stateEventValidation").validateStateEventDraft({
    eventName: "SvS", firstOccurrenceDate: "2026-08-10", recurrenceDays: 42,
    phases: [
      { phaseName: "Battle", phaseTimeUtc: "6:30pm" },
      { phaseName: "battle", phaseTimeUtc: "1830" }
    ]
  }), /Duplicate phase name/)
})

test("phase editor retains, replaces and removes media independently including GIF", () => {
  const gif = {
    originalFilename: "phase.gif", contentType: "image/gif", byteSize: 6,
    imageData: Buffer.from("GIF89a")
  }
  assert.equal(validateAttachmentMetadata({
    name: gif.originalFilename, contentType: gif.contentType, size: gif.byteSize,
    url: "https://cdn.discordapp.com/attachments/1/2/phase.gif"
  }).contentType, "image/gif")
  const draft = phaseDraft({
    phase_name: "Battle", phase_time_utc: "12:00",
    pre_alert_media: gif, exact_media: null
  })
  assert.equal(draft.preAlertMedia, gif)
  assert.equal(draft.exactMedia, null)
  const data = { workingPhase: draft }
  assert.match(buildMediaView("opaque", data, "pre_alert").content, /phase.gif/)
  data.workingPhase.exactMedia = gif
  data.workingPhase.preAlertMedia = null
  assert.match(buildMediaView("opaque", data, "pre_alert").content, /Current: None/)
  assert.match(buildMediaView("opaque", data, "exact").content, /image\/gif/)
})

test("list, management, review, preview and test views use opaque unique controls", () => {
  const event = storedEvent({
    phases: [
      storedEvent().phases[0],
      { ...storedEvent().phases[0], id: "33333333-3333-4333-8333-333333333333",
        phase_name: "Battle", phase_time_utc: "11:00:00", announce_exact: false }
    ]
  })
  const list = buildEventListView([event], 0, 1, "session")
  assert.ok(!JSON.stringify(list.view).includes(event.id))
  assert.equal(Object.values(list.tokenMap)[0], event.id)
  const views = [
    list.view,
    buildManageView("session", { selectedEvent: event, page: 0 }),
    buildReviewView("session", { ...eventDraft(event), stateNumber: "689" }),
    buildOccurrencePreviewView("session", event, new Date("2026-08-09T00:00:00Z")),
    buildTestView("session", { selectedEvent: event, testPhaseIndex: 0 }),
    buildPhaseManagerView("session", eventDraft(event)),
    buildPhaseConfigurationView("session", {
      ...eventDraft(event), workingPhase: phaseDraft(event.phases[0])
    })
  ]
  for (const view of views) {
    const ids = componentIds(view)
    assert.equal(new Set(ids).size, ids.length)
  }
  assert.equal(nextOccurrencePhases(event, new Date("2026-08-09T00:00:00Z")).length, 2)
  assert.match(views[3].content, /Local: <t:/)
})

test("cancel, preview and test announcement do not mutate production state", async () => {
  const sessionStore = new InteractionSessionStore()
  let createCount = 0
  let mutationCount = 0
  let sent = null
  const stateRepository = {
    async createStateEvent() { createCount += 1 },
    async getStateEvent() { return storedEvent() },
    async listStateEvents() { return { events: [storedEvent()], total: 1 } },
    async updateStateEvent() { mutationCount += 1 },
    async setStateEventStatus() { mutationCount += 1 }
  }
  const handlerOptions = options({ stateRepository, sessionStore })
  handlerOptions.targetResolver = async () => ({ channel: { async send(message) { sent = message } } })

  const cancelSession = sessionStore.create(context, { mode: "create" })
  await handleStateEventInteraction(interaction(
    `${STATE_EVENT_IDS.cancelPrefix}${cancelSession}`
  ), handlerOptions)
  assert.equal(createCount, 0)

  const manageSession = sessionStore.create(context, {
    selectedEvent: storedEvent(), page: 0, testPhaseIndex: 0
  })
  await handleStateEventInteraction(interaction(
    `${STATE_EVENT_IDS.previewPrefix}${manageSession}`
  ), handlerOptions)
  await handleStateEventInteraction(interaction(
    `${STATE_EVENT_IDS.testKindPrefix}${manageSession}:exact`
  ), handlerOptions)
  assert.equal(mutationCount, 0)
  assert.match(sent.embeds[0].toJSON().title, /^TEST - /)
})

test("pause, resume and soft delete are confirmed and call the status API", async () => {
  for (const [action, expected] of [["pause", "paused"], ["resume", "active"], ["delete", "deleted"]]) {
    const sessionStore = new InteractionSessionStore()
    let status = null
    const stateRepository = {
      async setStateEventStatus(input) { status = input.status; return { event_name: "SvS" } }
    }
    const sessionId = sessionStore.create(context, { selectedEvent: storedEvent() })
    const confirm = interaction(`${STATE_EVENT_IDS.statusConfirmPrefix}${sessionId}:${action}`)
    await handleStateEventInteraction(confirm, options({ stateRepository, sessionStore }))
    assert.equal(status, expected)
  }
})

test("sessions isolate user, guild and profile and ordinary guilds cannot enter state events", async () => {
  const sessionStore = new InteractionSessionStore()
  const id = sessionStore.create(context, {})
  assert.throws(() => sessionStore.get(id, { ...context, userId: "user-2" }), /another user/)
  assert.throws(() => sessionStore.get(id, { ...context, guildId: "alliance-guild" }), /another Discord/)
  assert.throws(() => sessionStore.get(id, { ...context, gameProfile: "kingshot" }), /another game profile/)

  await assert.rejects(handleStateEventInteraction(
    interaction(STATE_EVENT_IDS.home, "button", { guildId: "alliance-guild" }),
    options({ schedulerRepository: { async getStateDestination() { return null } } })
  ), /only in an enabled state destination/)
})

test("scheduler acknowledges state buttons before destination and repository work", async () => {
  const order = []
  const schedulerRepository = {
    async getStateDestination() { order.push("destination"); return destination }
  }
  const stateRepository = {
    async listStateEvents() { order.push("list"); return { events: [], total: 0 } }
  }
  const target = interaction(`${STATE_EVENT_IDS.listPrefix}0`, "button", {
    async deferUpdate() { order.push("defer"); this.deferred = true }
  })
  await handleEventSchedulerInteraction(target, {
    async userCanManageServer() { order.push("auth"); return true },
    healthProvider: () => health,
    repositoryProvider: () => schedulerRepository,
    stateRepositoryProvider: () => stateRepository
  })
  assert.deepEqual(order, ["defer", "auth", "destination", "list"])
})
