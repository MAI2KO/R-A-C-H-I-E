const test = require("node:test")
const assert = require("node:assert/strict")

const {
  IDS,
  buildRoundupDayView,
  handleEventSchedulerInteraction
} = require("../src/eventSchedulerInteractions")
const { generateMissingRoundupClaims } = require("../src/weeklyRoundupGeneration")

const guildId = "7777777777777701"
const channelId = "7777777777777702"
const health = Object.freeze({
  available: true,
  gameProfile: "wos",
  botInstanceName: "rachie-wos"
})
const settings = Object.freeze({
  alliance_name: "HnC",
  weekly_roundup_enabled: true,
  state_roundup_enabled: true,
  weekly_roundup_day: 5,
  weekly_roundup_time_utc: "15:20:00",
  weekly_roundup_channel_id: channelId,
  roundup_when_empty: false
})

function roundupPayload() {
  return {
    claim: {
      gameProfile: "wos",
      targetKind: "alliance",
      targetGuildId: guildId,
      targetChannelId: channelId,
      targetIsCurrent: true,
      weekStart: new Date("2028-03-03T00:00:00Z"),
      weekEnd: new Date("2028-03-10T00:00:00Z"),
      postWhenEmpty: false
    },
    allianceName: "HnC",
    occurrences: [{
      eventId: "event-1",
      sourceGuildId: guildId,
      allianceId: "alliance-1",
      allianceName: "HnC",
      isMainAlliance: true,
      eventName: "Bear Trap",
      occurrenceAt: new Date("2028-03-03T18:00:00Z")
    }]
  }
}

function stateRoundupPayload() {
  return {
    ...roundupPayload(),
    claim: {
      ...roundupPayload().claim,
      targetKind: "state",
      targetGuildId: "state-guild",
      targetChannelId: "state-channel"
    },
    occurrences: [
      ...roundupPayload().occurrences,
      {
        eventId: "event-2",
        sourceGuildId: "guild-2",
        allianceId: "alliance-2",
        allianceName: "HwC",
        isMainAlliance: true,
        eventName: "Foundry",
        occurrenceAt: new Date("2028-03-04T20:00:00Z")
      }
    ]
  }
}

function interaction(type, customId, overrides = {}) {
  return {
    commandName: null,
    customId,
    guildId,
    user: { id: "user-1" },
    client: {},
    values: [],
    deferred: false,
    replied: false,
    isChatInputCommand: () => false,
    isButton: () => type === "button",
    isStringSelectMenu: () => type === "select",
    isChannelSelectMenu: () => false,
    isModalSubmit: () => type === "modal",
    async deferUpdate() { this.deferred = true },
    async deferReply() { this.deferred = true },
    async editReply(payload) { this.edited = payload },
    async reply(payload) { this.replied = true; this.replyPayload = payload },
    async showModal(modal) { this.modal = modal },
    ...overrides
  }
}

function options(repository, overrides = {}) {
  return {
    healthProvider: () => health,
    userCanManageServer: async () => true,
    repositoryProvider: () => repository,
    roundupNow: () => new Date("2028-03-03T16:00:00Z"),
    logger: { error() {}, warn() {} },
    ...overrides
  }
}

function components(view) {
  return view.components.flatMap(row => row.toJSON().components)
}

test("roundup preview is ephemeral and does not mutate scheduled history", async () => {
  const history = { status: "sent", sentMessageId: "production-message" }
  const repository = {
    async getWeeklyRoundupPreview() { return roundupPayload() }
  }
  const preview = interaction("button", IDS.roundupAlliancePreview)
  await handleEventSchedulerInteraction(preview, options(repository))
  assert.equal(preview.deferred, true)
  assert.match(preview.edited.content, /nothing has been sent/i)
  assert.match(preview.edited.embeds[0].toJSON().title, /weekly roundup$/i)
  assert.deepEqual(history, { status: "sent", sentMessageId: "production-message" })
})

test("confirmed test sends are repeatable and never consume production history", async () => {
  const sent = []
  const targets = []
  const history = { status: "sent", sentMessageId: "production-message" }
  const repository = {
    async getGuildSettings() { return settings },
    async getWeeklyRoundupPreview() { return roundupPayload() }
  }
  const resolver = async (client, targetGuildId, targetChannelId) => {
    targets.push({ targetGuildId, targetChannelId })
    return {
    channel: { async send(message) { sent.push(message); return { id: `test-${sent.length}` } } }
    }
  }

  const request = interaction("button", IDS.roundupAllianceTest)
  await handleEventSchedulerInteraction(request, options(repository, {
    roundupTargetResolver: resolver
  }))
  assert.equal(sent.length, 0)
  assert.match(request.edited.content, /does not affect scheduled-roundup history/i)
  assert.ok(components(request.edited).some(component => component.label === "Confirm test send"))

  for (let count = 1; count <= 2; count += 1) {
    const confirm = interaction("button", IDS.roundupAllianceTestConfirm)
    await handleEventSchedulerInteraction(confirm, options(repository, {
      roundupTargetResolver: resolver
    }))
    assert.equal(sent.length, count)
    assert.match(sent.at(-1).embeds[0].toJSON().title, /— TEST/)
    assert.match(confirm.edited.content, /Scheduled history was not changed/)
  }
  assert.deepEqual(targets, [
    { targetGuildId: guildId, targetChannelId: channelId },
    { targetGuildId: guildId, targetChannelId: channelId }
  ])
  assert.deepEqual(history, { status: "sent", sentMessageId: "production-message" })
})

test("state preview and repeated test sends use the combined linked destination", async () => {
  const sent = []
  const targets = []
  const history = { status: "sent", sentMessageId: "production-state-message" }
  const repository = {
    async getGuildSettings() { return settings },
    async getStateWeeklyRoundupPreview() { return stateRoundupPayload() }
  }
  const resolver = async (client, targetGuildId, targetChannelId) => {
    targets.push({ targetGuildId, targetChannelId })
    return {
      channel: { async send(message) { sent.push(message); return { id: `state-${sent.length}` } } }
    }
  }

  const preview = interaction("button", IDS.roundupStatePreview)
  await handleEventSchedulerInteraction(preview, options(repository))
  assert.match(preview.edited.content, /State roundup preview.*nothing has been sent/is)
  const previewJson = preview.edited.embeds[0].toJSON()
  assert.match(previewJson.title, /State weekly roundup/)
  assert.match(previewJson.description, /\*\*HnC\*\*/)
  assert.match(previewJson.description, /\*\*HwC\*\*/)

  const request = interaction("button", IDS.roundupStateTest)
  await handleEventSchedulerInteraction(request, options(repository))
  assert.match(request.edited.content, /Destination: <#state-channel>/)
  assert.equal(sent.length, 0)

  for (let count = 1; count <= 2; count += 1) {
    const confirm = interaction("button", IDS.roundupStateTestConfirm)
    await handleEventSchedulerInteraction(confirm, options(repository, {
      roundupTargetResolver: resolver
    }))
    assert.equal(sent.length, count)
    const sentJson = sent.at(-1).embeds[0].toJSON()
    assert.match(sentJson.title, /State weekly roundup — TEST/)
    assert.match(sentJson.description, /\*\*HnC\*\*/)
    assert.match(sentJson.description, /\*\*HwC\*\*/)
  }
  assert.deepEqual(targets, [
    { targetGuildId: "state-guild", targetChannelId: "state-channel" },
    { targetGuildId: "state-guild", targetChannelId: "state-channel" }
  ])
  assert.deepEqual(history, { status: "sent", sentMessageId: "production-state-message" })
})

test("state roundup testing explains missing destination configuration", async () => {
  const repository = {
    async getGuildSettings() { return settings },
    async getStateWeeklyRoundupPreview() { return null }
  }
  for (const customId of [IDS.roundupStatePreview, IDS.roundupStateTest]) {
    const request = interaction("button", customId)
    await handleEventSchedulerInteraction(request, options(repository))
    assert.match(request.edited.content, /valid enabled state destination and link/i)
  }
})

test("scheduled roundup replay protection remains independent from test sends", async () => {
  const inserted = new Set()
  const repository = {
    async listRoundupConfigurations() {
      return [{
        source_guild_id: guildId,
        game_profile: "wos",
        weekly_roundup_day: 5,
        weekly_roundup_time_utc: "15:20:00",
        weekly_roundup_channel_id: channelId,
        weekly_roundup_enabled: true,
        state_roundup_enabled: false,
        weekly_roundup_not_before: "2028-01-01T00:00:00Z",
        roundup_when_empty: false
      }]
    },
    async insertMissingClaims(claims) {
      let count = 0
      for (const claim of claims) {
        const key = `${claim.weekStartDate}:${claim.targetKind}:${claim.targetGuildId}`
        if (!inserted.has(key)) { inserted.add(key); count += 1 }
      }
      return count
    }
  }
  const input = {
    repository,
    gameProfile: "wos",
    now: new Date("2028-03-03T15:20:00Z"),
    graceMinutes: 60
  }
  assert.equal(await generateMissingRoundupClaims(input), 1)
  assert.equal(await generateMissingRoundupClaims(input), 0)
})

test("the current weekday can continue, a new weekday can be drafted, and Back does not save", async () => {
  let configured = null
  const repository = {
    async getGuildSettings() { return settings },
    async configureWeeklyRoundup(input) { configured = input; return input }
  }
  const initial = buildRoundupDayView(settings)
  assert.match(initial.content, /Current schedule: Friday at 15:20 UTC/)
  assert.match(initial.content, /Selected weekday: Friday/)
  const initialComponents = components(initial)
  const fridayContinue = initialComponents.find(component => component.label === "Continue")

  const retain = interaction("button", fridayContinue.custom_id)
  await handleEventSchedulerInteraction(retain, options(repository))
  assert.equal(retain.deferred, false)
  assert.equal(retain.modal.toJSON().custom_id, `${IDS.roundupTimeModalPrefix}5:5:1520`)
  assert.equal(retain.modal.toJSON().components[0].components[0].value, "15:20")
  assert.equal(configured, null)

  const weekdayMenu = initialComponents.find(component => component.custom_id?.startsWith(IDS.roundupDayPrefix))
  const change = interaction("select", weekdayMenu.custom_id, { values: ["6"] })
  await handleEventSchedulerInteraction(change, options(repository))
  assert.match(change.edited.content, /Current schedule: Friday at 15:20 UTC/)
  assert.match(change.edited.content, /Selected weekday: Saturday/)
  const saturdayContinue = components(change.edited).find(component => component.label === "Continue")
  assert.match(saturdayContinue.custom_id, /:6:5:1520$/)
  assert.equal(configured, null)

  const back = interaction("button", IDS.roundupSettings)
  await handleEventSchedulerInteraction(back, options(repository))
  assert.match(back.edited.content, /Schedule: Friday at 15:20 UTC/)
  assert.equal(configured, null)

  const saturday = interaction("button", saturdayContinue.custom_id)
  await handleEventSchedulerInteraction(saturday, options(repository))
  const modal = interaction("modal", saturday.modal.toJSON().custom_id, {
    fields: { getTextInputValue: () => "16:45" }
  })
  await handleEventSchedulerInteraction(modal, options(repository))
  assert.equal(configured, null)
  assert.match(modal.edited.content, /Current schedule: Friday at 15:20 UTC/)
  assert.match(modal.edited.content, /Proposed schedule: Saturday at 16:45 UTC/)
  const save = components(modal.edited).find(component => component.label === "Save schedule")
  const confirm = interaction("button", save.custom_id)
  await handleEventSchedulerInteraction(confirm, options(repository))
  assert.equal(configured.weekday, 6)
  assert.equal(configured.timeUtc, "16:45")
})
