const test = require("node:test")
const assert = require("node:assert/strict")
const { ChannelType, PermissionFlagsBits } = require("discord.js")

const { buildDeliveryClaims } = require("../src/eventDeliveryGeneration")
const {
  PermanentDeliveryError
} = require("../src/eventDeliveryWorker")
const {
  createDiscordEventDeliveryHandler,
  resolveDeliveryTarget
} = require("../src/discordEventDelivery")
const { createDiscordWeeklyRoundupDelivery } = require("../src/discordWeeklyRoundup")

const WINDOW_START = new Date("2026-08-06T11:00:00Z")
const WINDOW_END = new Date("2026-08-06T13:00:00Z")

function event(overrides = {}) {
  return {
    id: "71",
    guild_id: "alliance-guild",
    game_profile: "wos",
    alliance_name: "North",
    event_name: "Bear Hunt",
    first_occurrence_date: "2026-08-06",
    event_time_utc: "12:10:00",
    recurrence_days: 3,
    advance_reminder_minutes: 30,
    reminder_at_start: true,
    publish_to_alliance: true,
    publish_to_state: true,
    status: "active",
    event_channel_id: "alliance-channel",
    sharing_enabled: true,
    state_guild_id: "state-guild",
    state_event_channel_id: "state-channel",
    groups: [],
    ...overrides
  }
}

function generationOptions(gameProfile = "wos") {
  return { gameProfile, windowStart: WINDOW_START, windowEnd: WINDOW_END }
}

test("individual generation is alliance-only even when state publishing is enabled", () => {
  const claims = buildDeliveryClaims([event()], generationOptions())
  assert.deepEqual(claims.map(claim => `${claim.targetKind}:${claim.deliveryKind}`), [
    "alliance:advance_reminder",
    "alliance:final_reminder"
  ])
  assert.equal(claims[0].deliverAt.toISOString(), "2026-08-06T11:40:00.000Z")
  assert.equal(claims[1].deliverAt.toISOString(), "2026-08-06T12:09:00.000Z")
  assert.equal(claims.some(claim => claim.targetKind === "state"), false)
  assert.equal(claims.some(claim => claim.deliverAt.getTime() === claim.occurrenceAt.getTime()), false)
})

test("state-roundup-only and grouped events create no individual state claims", () => {
  assert.deepEqual(buildDeliveryClaims([event({ publish_to_alliance: false })], generationOptions()), [])
  const grouped = buildDeliveryClaims([event({
    event_time_utc: null,
    groups: [
      { group_id: "1", group_name: "Alpha", event_time_utc: "12:10", sort_order: 0 },
      { group_id: "2", group_name: "Beta", event_time_utc: "12:20", sort_order: 1 }
    ]
  })], generationOptions())
  assert.deepEqual(grouped.map(claim => `${claim.groupId}:${claim.deliveryKind}`), [
    "1:advance_reminder",
    "1:final_reminder",
    "2:advance_reminder",
    "2:final_reminder"
  ])
  assert.ok(grouped.every(claim => claim.targetKind === "alliance"))
})

function discordFixture({ send = async () => ({ id: "message" }) } = {}) {
  const permissions = new Set([
    PermissionFlagsBits.ViewChannel,
    PermissionFlagsBits.SendMessages,
    PermissionFlagsBits.EmbedLinks,
    PermissionFlagsBits.AttachFiles
  ])
  const member = { id: "bot" }
  const channels = new Map()
  function addChannel(id, guildId) {
    const channel = {
      id,
      guildId,
      type: ChannelType.GuildText,
      isTextBased: () => true,
      isSendable: () => true,
      permissionsFor: () => ({ has: permission => permissions.has(permission) }),
      send
    }
    channels.set(`${guildId}:${id}`, channel)
    return channel
  }
  const allianceChannel = addChannel("alliance-channel", "alliance-guild")
  const stateChannel = addChannel("state-channel", "state-guild")
  function guild(id) {
    return {
      id,
      members: { me: member, fetchMe: async () => member },
      channels: {
        cache: new Map([...channels.values()].filter(channel => channel.guildId === id)
          .map(channel => [channel.id, channel])),
        fetch: async channelId => channels.get(`${id}:${channelId}`) || null
      }
    }
  }
  const guilds = new Map([
    ["alliance-guild", guild("alliance-guild")],
    ["state-guild", guild("state-guild")]
  ])
  return {
    client: {
      isReady: () => true,
      guilds: { cache: guilds, fetch: async id => guilds.get(id) || null }
    },
    allianceChannel,
    stateChannel
  }
}

function stateIndividualPayload() {
  const imageData = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
  return {
    claim: {
      gameProfile: "wos",
      deliveryKind: "advance_reminder",
      targetKind: "state",
      targetGuildId: "state-guild",
      targetChannelId: "state-channel",
      targetIsCurrent: true,
      occurrenceAt: new Date("2026-08-10T18:30:00Z"),
      deliverAt: new Date("2026-08-10T18:00:00Z")
    },
    event: { guildId: "alliance-guild", eventName: "Bear Hunt", recurrenceDays: 7 },
    alliance: { guildId: "alliance-guild", name: "North" },
    group: null,
    image: {
      originalFilename: "event.png",
      contentType: "image/png",
      byteSize: imageData.length,
      imageData
    }
  }
}

test("individual state delivery is permanently rejected before sending or attaching images", async () => {
  let sends = 0
  const fixture = discordFixture({ send: async () => { sends += 1; return { id: "unexpected" } } })
  const handler = createDiscordEventDeliveryHandler({ client: fixture.client, gameProfile: "wos" })
  await assert.rejects(handler(stateIndividualPayload()), error =>
    error instanceof PermanentDeliveryError && /state reminders are disabled/i.test(error.message)
  )
  assert.equal(sends, 0)
})

test("state target validation remains available only for weekly roundups", async () => {
  const sent = []
  const fixture = discordFixture({
    send: async message => { sent.push(message); return { id: `roundup-${sent.length}` } }
  })
  const target = await resolveDeliveryTarget(fixture.client, stateIndividualPayload(), {
    hasImage: false
  })
  assert.equal(target, fixture.stateChannel)

  const deliver = createDiscordWeeklyRoundupDelivery({ client: fixture.client, gameProfile: "wos" })
  const prepared = await deliver({
    claim: {
      gameProfile: "wos",
      targetKind: "state",
      targetGuildId: "state-guild",
      targetChannelId: "state-channel",
      targetIsCurrent: true,
      weekStart: new Date("2026-08-10T00:00:00Z"),
      weekEnd: new Date("2026-08-17T00:00:00Z"),
      postWhenEmpty: true
    },
    allianceName: "North",
    occurrences: []
  })
  assert.equal(prepared.messages.length, 1)
  assert.equal(prepared.messages[0].files, undefined)
  assert.equal(await prepared.sendPart(0), "roundup-1")
  assert.equal(sent.length, 1)
})

test("profile mismatch still blocks individual delivery", async () => {
  const fixture = discordFixture()
  const handler = createDiscordEventDeliveryHandler({ client: fixture.client, gameProfile: "kingshot" })
  await assert.rejects(handler({
    ...stateIndividualPayload(),
    claim: {
      ...stateIndividualPayload().claim,
      targetKind: "alliance",
      targetGuildId: "alliance-guild",
      targetChannelId: "alliance-channel"
    }
  }), PermanentDeliveryError)
})
