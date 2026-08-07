const test = require("node:test")
const assert = require("node:assert/strict")

const {
  allianceListView
} = require("../src/allianceManagementInteractions")
const {
  CREATION_IDS,
  buildAllianceSelectionView,
  buildGroupManagerView,
  buildImageChoiceView,
  buildListView,
  buildOccurrencePreviewView,
  buildPreviewView,
  buildPublishingView,
  buildTimingChoiceView,
  buildTimingView,
  creationSessions
} = require("../src/eventCreationInteractions")
const {
  allianceFilterView,
  confirmationView,
  eventDraft,
  handleEventManagementModalOpeningInteraction,
  listView,
  managementSessions
} = require("../src/eventManagementInteractions")
const {
  buildChannelConfigurationView,
  buildHomeView,
  buildRoundupChannelView,
  buildRoundupDayView,
  buildRoundupSchedulePreview,
  buildRoundupSettingsView,
  buildStateDestinationView,
  buildStateSharingView,
  handleEventSchedulerInteraction
} = require("../src/eventSchedulerInteractions")
const { buildSchedulerHelpView } = require("../src/eventSchedulerHelp")
const { InteractionSessionError } = require("../src/interactionSessions")

const context = Object.freeze({
  userId: "user-1",
  guildId: "guild-1",
  gameProfile: "wos"
})
const health = Object.freeze({
  available: true,
  gameProfile: "wos",
  botInstanceName: "rachie-wos"
})

function scheduledEvent(overrides = {}) {
  return {
    id: "1",
    guild_id: context.guildId,
    game_profile: context.gameProfile,
    alliance_id: "alliance-1",
    alliance_name: "HnC",
    is_default_alliance: true,
    event_name: "Bear Trap",
    first_occurrence_date: "2028-02-28",
    event_time_utc: null,
    recurrence_days: 7,
    advance_reminder_minutes: 30,
    advance_reminder_message: null,
    reminder_at_start: true,
    final_reminder_message: null,
    publish_to_alliance: true,
    publish_to_state: false,
    include_in_weekly_roundup: true,
    image_filename: null,
    image_content_type: null,
    image_byte_size: null,
    status: "active",
    groups: [
      { group_name: "Best Bear", event_time_utc: "18:00:00", sort_order: 0 },
      { group_name: "Smelly Bear", event_time_utc: "19:00:00", sort_order: 1 }
    ],
    ...overrides
  }
}

const liveFixture = Object.freeze([
  scheduledEvent(),
  scheduledEvent({
    id: "2",
    alliance_id: "alliance-2",
    alliance_name: "HwC",
    is_default_alliance: false,
    groups: [
      { group_name: "Gay Bear", event_time_utc: "18:30:00", sort_order: 0 },
      { group_name: "Fat Bear", event_time_utc: "19:30:00", sort_order: 1 }
    ]
  }),
  scheduledEvent({
    id: "3",
    alliance_id: "alliance-2",
    alliance_name: "HwC",
    is_default_alliance: false,
    event_name: "SvS",
    event_time_utc: "20:00:00",
    groups: []
  })
])

function componentData(view) {
  return (view.components || []).flatMap(row => {
    const data = typeof row.toJSON === "function" ? row.toJSON() : row
    return data.components || []
  })
}

function assertUniqueComponentIds(view, label) {
  const ids = componentData(view)
    .map(component => component.custom_id)
    .filter(Boolean)
  assert.equal(new Set(ids).size, ids.length, `${label} contains duplicate custom IDs`)
  for (const id of ids) {
    assert.ok(id.length <= 100, `${label} custom ID exceeds Discord's 100-character limit`)
  }
  return ids
}

function managementView(events = liveFixture, page = 0, total = events.length, sessionId = "session") {
  const eventMap = {}
  const eventDraftMap = {}
  const view = listView(events, page, total, sessionId, eventMap, eventDraftMap)
  return { view, eventMap, eventDraftMap }
}

function componentInteraction(customId, values = []) {
  const isSelect = customId.startsWith("mg:a:") || customId.startsWith("mg:f:")
  return {
    commandName: null,
    customId,
    guildId: context.guildId,
    user: { id: context.userId },
    values,
    deferred: false,
    replied: false,
    isChatInputCommand: () => false,
    isButton: () => !isSelect,
    isStringSelectMenu: () => isSelect,
    isChannelSelectMenu: () => false,
    isModalSubmit: () => false,
    async deferUpdate() { this.deferred = true },
    async editReply(payload) { this.edited = payload },
    async reply(payload) { this.replied = true; this.replyPayload = payload },
    async showModal(modal) { this.modal = modal }
  }
}

function allianceRepository(events = liveFixture) {
  const alliances = [
    { id: "alliance-1", alliance_name: "HnC", is_default: true },
    { id: "alliance-2", alliance_name: "HwC", is_default: false },
    { id: "alliance-3", alliance_name: "Empty", is_default: false }
  ]
  const calls = []
  return {
    calls,
    async listAlliances() { return { alliances, total: alliances.length } },
    async getAlliance(guildId, allianceId) {
      if (guildId !== context.guildId) return null
      return alliances.find(alliance => alliance.id === String(allianceId)) || null
    },
    async listEvents(guildId, { limit, offset, allianceId }) {
      calls.push({ guildId, limit, offset, allianceId })
      const filtered = allianceId
        ? events.filter(event => event.alliance_id === String(allianceId))
        : events
      return { events: filtered.slice(offset, offset + limit), total: filtered.length }
    },
    async getEvent(guildId, eventId) {
      if (guildId !== context.guildId) return null
      return events.find(event => event.id === String(eventId)) || null
    }
  }
}

function schedulerOptions(repository, overrides = {}) {
  return {
    healthProvider: () => health,
    userCanManageServer: async () => true,
    repositoryProvider: () => repository,
    logger: { error() {}, warn() {} },
    ...overrides
  }
}

async function openEventFilter(repository, options = schedulerOptions(repository)) {
  const interaction = componentInteraction("el:0")
  await handleEventSchedulerInteraction(interaction, options)
  const select = componentData(interaction.edited)
    .find(component => component.custom_id?.startsWith("mg:f:"))
  return { interaction, select, options }
}

test("the live duplicate-name event list has unique component IDs", () => {
  const { view, eventMap } = managementView()
  const ids = assertUniqueComponentIds(view, "event management list")
  assert.equal(ids.filter(id => id.startsWith("mg:a:")).length, 3)
  assert.equal(new Set(Object.values(eventMap)).size, 3)
  assert.equal((view.content.match(/\*\*Bear Trap\*\*/g) || []).length, 2)
  assert.match(view.content, /Alliance: HnC/)
  assert.match(view.content, /Alliance: HwC/)
  assert.match(view.content, /\*\*SvS\*\*/)
})

test("each duplicate-named Bear Trap opens its own database-backed draft", async () => {
  const sessionId = managementSessions.create(context, {})
  const built = managementView(liveFixture, 0, liveFixture.length, sessionId)
  managementSessions.update(sessionId, context, {
    page: 0,
    eventMap: built.eventMap,
    eventDraftMap: built.eventDraftMap
  })
  const selectorIds = assertUniqueComponentIds(built.view, "duplicate event list")
    .filter(id => id.startsWith("mg:a:"))

  for (const [index, expected] of [
    [0, { eventId: "1", allianceId: "alliance-1", group: "Best Bear" }],
    [1, { eventId: "2", allianceId: "alliance-2", group: "Gay Bear" }]
  ]) {
    const interaction = componentInteraction(selectorIds[index], ["edit"])
    assert.equal(await handleEventManagementModalOpeningInteraction(interaction, { health }), true)
    assert.equal(interaction.deferred, false)
    const modalId = interaction.modal.toJSON().custom_id
    const editSessionId = modalId.slice(CREATION_IDS.corePrefix.length)
    const draft = creationSessions.get(editSessionId, context).data
    assert.equal(draft.eventId, expected.eventId)
    assert.equal(draft.allianceId, expected.allianceId)
    assert.equal(draft.eventName, "Bear Trap")
    assert.equal(draft.groups[0].groupName, expected.group)
  }
})

test("event list pagination remains unique and navigates both pages", async () => {
  const fourth = scheduledEvent({ id: "4", event_name: "Foundry", event_time_utc: "21:00:00", groups: [] })
  const events = [...liveFixture, fourth]
  const calls = []
  const repository = {
    async listAlliances() {
      return {
        alliances: [
          { id: "alliance-1", alliance_name: "HnC", is_default: true },
          { id: "alliance-2", alliance_name: "HwC", is_default: false }
        ],
        total: 2
      }
    },
    async listEvents(guildId, { limit, offset, allianceId }) {
      calls.push({ guildId, limit, offset, allianceId })
      return { events: events.slice(offset, offset + limit), total: events.length }
    }
  }
  const options = {
    healthProvider: () => health,
    userCanManageServer: async () => true,
    repositoryProvider: () => repository,
    logger: { error() {}, warn() {} }
  }

  const pageOne = componentInteraction("el:0")
  await handleEventSchedulerInteraction(pageOne, options)
  assert.match(pageOne.edited.content, /select an alliance/i)
  const filter = componentData(pageOne.edited).find(component => component.custom_id?.startsWith("mg:f:"))

  const allAlliances = componentInteraction(filter.custom_id, ["all"])
  await handleEventSchedulerInteraction(allAlliances, options)
  assert.match(allAlliances.edited.content, /page 1 of 2/i)
  assert.match(componentData(allAlliances.edited)[0].placeholder, /HnC — Bear Trap/)
  assertUniqueComponentIds(allAlliances.edited, "event list page 1")
  const next = componentData(allAlliances.edited).find(component => component.label === "Next")
  assert.ok(next.custom_id.startsWith("mg:l:"))

  const pageTwo = componentInteraction(next.custom_id)
  await handleEventSchedulerInteraction(pageTwo, options)
  assert.match(pageTwo.edited.content, /page 2 of 2/i)
  assert.match(pageTwo.edited.content, /Foundry/)
  assertUniqueComponentIds(pageTwo.edited, "event list page 2")
  const previous = componentData(pageTwo.edited).find(component => component.label === "Previous")

  const pageOneAgain = componentInteraction(previous.custom_id)
  await handleEventSchedulerInteraction(pageOneAgain, options)
  assert.match(pageOneAgain.edited.content, /page 1 of 2/i)
  assert.equal(calls.at(-1).offset, 0)
  assert.equal(calls.at(-1).allianceId, null)
})

test("view events requires an alliance choice and filters duplicate event names", async () => {
  const repository = allianceRepository()
  const opened = await openEventFilter(repository)
  assert.match(opened.interaction.edited.content, /Select an alliance/)
  assert.doesNotMatch(opened.interaction.edited.content, /Bear Trap/)
  const labels = opened.select.options.map(option => option.label)
  assert.deepEqual(labels, ["All alliances", "HnC", "HwC", "Empty"])
  assert.equal(opened.select.options[1].description, "Main alliance")
  assert.equal(opened.select.options[2].description, "Sub-alliance")
  assert.ok(opened.select.options.every(option => !/^\d+$/.test(option.value)))

  for (const allianceName of ["HnC", "HwC"]) {
    const fresh = await openEventFilter(repository)
    const option = fresh.select.options.find(item => item.label === allianceName)
    const selected = componentInteraction(fresh.select.custom_id, [option.value])
    await handleEventSchedulerInteraction(selected, fresh.options)
    assert.match(selected.edited.content, new RegExp(`Scheduled events — ${allianceName}`))
    const selectors = componentData(selected.edited)
      .filter(component => component.custom_id?.startsWith("mg:a:"))
    assert.ok(selectors.length > 0)
    assert.ok(selectors.every(component => !component.placeholder.includes("—")))
    const expectedAllianceId = allianceName === "HnC" ? "alliance-1" : "alliance-2"
    assert.equal(repository.calls.at(-1).allianceId, expectedAllianceId)
    if (allianceName === "HnC") assert.doesNotMatch(selected.edited.content, /Alliance: HwC/)
    if (allianceName === "HwC") assert.doesNotMatch(selected.edited.content, /Alliance: HnC/)
  }

  const all = await openEventFilter(repository)
  const selectedAll = componentInteraction(all.select.custom_id, ["all"])
  await handleEventSchedulerInteraction(selectedAll, all.options)
  const allSelectors = componentData(selectedAll.edited)
    .filter(component => component.custom_id?.startsWith("mg:a:"))
  assert.match(allSelectors[0].placeholder, /HnC — Bear Trap/)
  assert.match(allSelectors[1].placeholder, /HwC — Bear Trap/)
})

test("filtered pagination and event preview return preserve the alliance filter", async () => {
  const hncEvents = Array.from({ length: 4 }, (_, index) => scheduledEvent({
    id: String(index + 10),
    event_name: `HnC Event ${index + 1}`,
    groups: [],
    event_time_utc: `${10 + index}:00:00`
  }))
  const repository = allianceRepository([...hncEvents, liveFixture[1]])
  const opened = await openEventFilter(repository)
  const hnc = opened.select.options.find(option => option.label === "HnC")
  const selected = componentInteraction(opened.select.custom_id, [hnc.value])
  await handleEventSchedulerInteraction(selected, opened.options)
  assert.match(selected.edited.content, /Page 1 of 2/)

  const next = componentData(selected.edited).find(component => component.label === "Next")
  const pageTwo = componentInteraction(next.custom_id)
  await handleEventSchedulerInteraction(pageTwo, opened.options)
  assert.match(pageTwo.edited.content, /Page 2 of 2/)
  assert.equal(repository.calls.at(-1).allianceId, "alliance-1")

  const selector = componentData(pageTwo.edited)
    .find(component => component.custom_id?.startsWith("mg:a:"))
  const preview = componentInteraction(selector.custom_id, ["preview"])
  await handleEventSchedulerInteraction(preview, opened.options)
  const back = componentData(preview.edited).find(component => component.label === "Back to events")
  const returned = componentInteraction(back.custom_id)
  await handleEventSchedulerInteraction(returned, opened.options)
  assert.match(returned.edited.content, /Scheduled events — HnC/)
  assert.match(returned.edited.content, /Page 2 of 2/)
  assert.equal(repository.calls.at(-1).allianceId, "alliance-1")
})

test("an empty alliance has a clear state without pagination", async () => {
  const repository = allianceRepository()
  const opened = await openEventFilter(repository)
  const empty = opened.select.options.find(option => option.label === "Empty")
  const selected = componentInteraction(opened.select.custom_id, [empty.value])
  await handleEventSchedulerInteraction(selected, opened.options)
  assert.equal(selected.edited.content, "No scheduled events for Empty.")
  const controls = componentData(selected.edited)
  assert.deepEqual(controls.map(component => component.label), ["Change alliance", "Back"])
})

test("alliance filters reject cross-guild and cross-profile session use", async () => {
  const repository = allianceRepository()
  const opened = await openEventFilter(repository)
  const hnc = opened.select.options.find(option => option.label === "HnC")

  const wrongGuild = componentInteraction(opened.select.custom_id, [hnc.value])
  wrongGuild.guildId = "guild-2"
  await handleEventSchedulerInteraction(wrongGuild, opened.options)
  assert.match(wrongGuild.edited.content, /another Discord server/)
  assert.equal(repository.calls.length, 0)

  const wrongProfile = componentInteraction(opened.select.custom_id, [hnc.value])
  await handleEventSchedulerInteraction(wrongProfile, schedulerOptions(repository, {
    healthProvider: () => ({ ...health, gameProfile: "kingshot" })
  }))
  assert.match(wrongProfile.edited.content, /another game profile/)
  assert.equal(repository.calls.length, 0)
})

test("event controls reject forged tokens and cross-guild or cross-profile sessions", async () => {
  const sessionId = managementSessions.create(context, {})
  const built = managementView([liveFixture[0]], 0, 1, sessionId)
  managementSessions.update(sessionId, context, {
    page: 0,
    eventMap: built.eventMap,
    eventDraftMap: built.eventDraftMap
  })
  const selectorId = componentData(built.view)[0].custom_id

  const forged = componentInteraction(`mg:a:${sessionId}:not-valid`, ["edit"])
  await assert.rejects(
    handleEventManagementModalOpeningInteraction(forged, { health }),
    InteractionSessionError
  )

  const wrongGuild = componentInteraction(selectorId, ["edit"])
  wrongGuild.guildId = "guild-2"
  await assert.rejects(
    handleEventManagementModalOpeningInteraction(wrongGuild, { health }),
    /another Discord server/
  )

  const wrongProfile = componentInteraction(selectorId, ["edit"])
  await assert.rejects(
    handleEventManagementModalOpeningInteraction(wrongProfile, {
      health: { ...health, gameProfile: "kingshot" }
    }),
    /another game profile/
  )
})

test("users without scheduler management permission cannot load event data", async () => {
  let repositoryCreated = false
  const interaction = componentInteraction("el:0")
  await handleEventSchedulerInteraction(interaction, {
    healthProvider: () => health,
    userCanManageServer: async () => false,
    repositoryProvider() { repositoryCreated = true; throw new Error("must not access repository") },
    logger: { error() {}, warn() {} }
  })
  assert.equal(interaction.deferred, true)
  assert.equal(repositoryCreated, false)
  assert.match(interaction.edited.content, /do not have permission/)
})

test("all scheduler message builders emit unique, concise custom IDs", () => {
  const draft = eventDraft(liveFixture[0])
  const settings = {
    alliance_name: "HnC",
    alliance_count: 2,
    event_channel_id: "channel-1",
    weekly_roundup_channel_id: "channel-2",
    weekly_roundup_enabled: true,
    state_roundup_enabled: true,
    weekly_roundup_day: 1,
    weekly_roundup_time_utc: "09:00:00"
  }
  const alliances = [
    { id: "alliance-1", alliance_name: "HnC", is_default: true, managed_event_count: 1 },
    { id: "alliance-2", alliance_name: "HwC", is_default: false, managed_event_count: 2 }
  ]
  const views = [
    ["home", buildHomeView(settings, null, null)],
    ["channels", buildChannelConfigurationView(settings)],
    ["roundup settings", buildRoundupSettingsView(settings)],
    ["roundup day", buildRoundupDayView(settings)],
    ["roundup channel", buildRoundupChannelView(settings)],
    ["roundup preview", buildRoundupSchedulePreview(1, "09:00")],
    ["state destination", buildStateDestinationView({ state_roundup_channel_id: "channel-3", enabled: true })],
    ["state sharing", buildStateSharingView({ game_profile: "wos", sharing_enabled: true })],
    ["help", buildSchedulerHelpView("managing-events")],
    ["alliance management", allianceListView(alliances, 0, 2, "session", null).view],
    ["alliance selection", buildAllianceSelectionView("session", alliances, 0, 2, null).view],
    ["timing choice", buildTimingChoiceView("session", draft)],
    ["group management", buildGroupManagerView("session", { groups: draft.groups, selectedGroupIndex: 0 })],
    ["image management", buildImageChoiceView("session", draft)],
    ["event timing", buildTimingView("session", draft)],
    ["publishing", buildPublishingView("session", draft)],
    ["event preview", buildPreviewView("session", draft)],
    ["event listing", buildListView(liveFixture, 0, liveFixture.length)],
    ["occurrence preview", buildOccurrencePreviewView(liveFixture[0], [], 0)],
    ["event management", managementView().view],
    ["event alliance filter", allianceFilterView(alliances, "session").view],
    ["event confirmation", confirmationView("session", "delete", "event-token", liveFixture[0])]
  ]

  for (const [label, view] of views) assertUniqueComponentIds(view, label)
})
