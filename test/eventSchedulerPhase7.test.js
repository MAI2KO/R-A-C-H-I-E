const test = require("node:test")
const assert = require("node:assert/strict")
const {
  ChannelType,
  PermissionFlagsBits
} = require("discord.js")

const { buildDeliveryClaims } = require("../src/eventDeliveryGeneration")
const {
  EMBED_TITLE_LIMIT,
  EMBED_DESCRIPTION_LIMIT,
  EMBED_FIELD_VALUE_LIMIT,
  formatStateEventDelivery
} = require("../src/eventDeliveryFormatting")
const {
  PermanentDeliveryError,
  RetryableDeliveryError,
  createEventDeliveryWorker
} = require("../src/eventDeliveryWorker")
const {
  resolveDeliveryTarget,
  createDiscordEventDeliveryHandler
} = require("../src/discordEventDelivery")

const WINDOW_START = new Date("2026-08-06T11:00:00Z")
const WINDOW_END = new Date("2026-08-06T13:00:00Z")

function stateEvent(overrides = {}) {
  return {
    id: "71",
    guild_id: "alliance-guild",
    game_profile: "wos",
    alliance_name: "North",
    event_name: "Bear Hunt",
    first_occurrence_date: "2026-08-06",
    event_time_utc: "12:10:00",
    recurrence_days: 3,
    advance_reminder_minutes: 10,
    reminder_at_start: true,
    publish_to_alliance: false,
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

test("eligible state events generate advance and start claims", () => {
  const claims = buildDeliveryClaims([stateEvent()], generationOptions())
  assert.deepEqual(claims.map(claim => claim.deliveryKind), [
    "advance_reminder",
    "event_start"
  ])
  assert.ok(claims.every(claim => claim.targetKind === "state"))
  assert.ok(claims.every(claim => claim.targetGuildId === "state-guild"))
  assert.ok(claims.every(claim => claim.targetChannelId === "state-channel"))
})

test("one occurrence can generate independent alliance and state claims", () => {
  const claims = buildDeliveryClaims([stateEvent({ publish_to_alliance: true })], generationOptions())
  assert.deepEqual(claims.map(claim => `${claim.targetKind}:${claim.deliveryKind}`), [
    "alliance:advance_reminder",
    "alliance:event_start",
    "state:advance_reminder",
    "state:event_start"
  ])
})

test("state generation requires publishing, an enabled link and complete target", () => {
  const excluded = [
    stateEvent({ publish_to_state: false }),
    stateEvent({ sharing_enabled: false }),
    stateEvent({ sharing_enabled: null, state_guild_id: null, state_event_channel_id: null }),
    stateEvent({ state_guild_id: null }),
    stateEvent({ state_event_channel_id: null }),
    stateEvent({ status: "paused" }),
    stateEvent({ status: "deleted" })
  ]
  for (const definition of excluded) {
    assert.deepEqual(buildDeliveryClaims([definition], generationOptions()), [])
  }
})

test("grouped state events generate each reminder kind per group", () => {
  const claims = buildDeliveryClaims([stateEvent({
    event_time_utc: null,
    groups: [
      { group_id: "1", group_name: "Alpha", event_time_utc: "12:10:00", sort_order: 0 },
      { group_id: "2", group_name: "Beta", event_time_utc: "12:20:00", sort_order: 1 }
    ]
  })], generationOptions())
  assert.deepEqual(claims.map(claim => `${claim.groupId}:${claim.deliveryKind}`), [
    "1:advance_reminder",
    "1:event_start",
    "2:advance_reminder",
    "2:event_start"
  ])
})

test("state generation remains isolated in both profile directions", () => {
  const wos = stateEvent()
  const kingshot = stateEvent({ id: "72", game_profile: "kingshot" })
  assert.ok(buildDeliveryClaims([wos, kingshot], generationOptions("wos"))
    .every(claim => claim.gameProfile === "wos"))
  assert.ok(buildDeliveryClaims([wos, kingshot], generationOptions("kingshot"))
    .every(claim => claim.gameProfile === "kingshot"))
})

function statePayload(overrides = {}) {
  const base = {
    claim: {
      id: "state-claim",
      gameProfile: "wos",
      attemptCount: 1,
      deliveryKind: "advance_reminder",
      targetKind: "state",
      targetGuildId: "state-guild",
      targetChannelId: "state-channel",
      targetIsCurrent: true,
      occurrenceAt: new Date("2026-08-10T18:30:00Z"),
      deliverAt: new Date("2026-08-10T18:20:00Z")
    },
    event: {
      id: "71",
      guildId: "alliance-guild",
      eventName: "Bear Hunt",
      recurrenceDays: 7
    },
    alliance: { name: "North", guildId: "alliance-guild" },
    group: null,
    image: null
  }
  return {
    ...base,
    ...overrides,
    claim: { ...base.claim, ...overrides.claim },
    event: { ...base.event, ...overrides.event },
    alliance: { ...base.alliance, ...overrides.alliance }
  }
}

function stateDiscordFixture({
  permissions = [
    PermissionFlagsBits.ViewChannel,
    PermissionFlagsBits.SendMessages,
    PermissionFlagsBits.EmbedLinks,
    PermissionFlagsBits.AttachFiles
  ],
  channelOverrides = {},
  send = async () => ({ id: "state-message" })
} = {}) {
  const allowed = new Set(permissions)
  const member = { id: "bot" }
  const channel = {
    id: "state-channel",
    guildId: "state-guild",
    type: ChannelType.GuildText,
    isTextBased: () => true,
    isSendable: () => true,
    permissionsFor: () => ({ has: permission => allowed.has(permission) }),
    send,
    ...channelOverrides
  }
  const guild = {
    id: "state-guild",
    members: { me: member, fetchMe: async () => member },
    channels: {
      cache: new Map([["state-channel", channel]]),
      fetch: async id => id === "state-channel" ? channel : null
    }
  }
  const client = {
    isReady: () => true,
    guilds: {
      cache: new Map([["state-guild", guild]]),
      fetch: async id => id === "state-guild" ? guild : null
    }
  }
  return { client, guild, channel }
}

test("state reminder format makes alliance prominent and remains mention safe", () => {
  const message = formatStateEventDelivery(statePayload({
    alliance: { name: "@everyone North" },
    group: { id: "1", name: "Alpha", eventTimeUtc: "18:30", sortOrder: 0 }
  }))
  const embed = message.embeds[0].toJSON()
  assert.match(embed.title, /North - Bear Hunt/)
  assert.match(embed.description, /State event reminder/)
  assert.match(embed.description, /Alliance:/)
  assert.match(embed.fields.find(field => field.name === "Group").value, /Alpha/)
  assert.match(embed.fields.find(field => field.name === "When").value, /2026-08-10 18:30 UTC/)
  assert.match(embed.fields.find(field => field.name === "When").value, /<t:1786386600:F>/)
  assert.match(embed.fields.find(field => field.name === "Recurrence").value, /Every week/)
  assert.doesNotMatch(JSON.stringify(embed), /@everyone/)
  assert.deepEqual(message.allowedMentions, { parse: [], repliedUser: false })
})

test("state event-start formatting supports ungrouped events and limits", () => {
  const message = formatStateEventDelivery(statePayload({
    claim: {
      deliveryKind: "event_start",
      deliverAt: new Date("2026-08-10T18:30:00Z")
    },
    alliance: { name: "A".repeat(5000) },
    event: { eventName: "E".repeat(5000), recurrenceDays: 28 }
  }))
  const embed = message.embeds[0].toJSON()
  assert.match(embed.description, /starting now/i)
  assert.equal(embed.fields.find(field => field.name === "Status").value, "Starting now")
  assert.equal(embed.fields.find(field => field.name === "Recurrence").value, "Every 4 weeks")
  assert.ok(embed.title.length <= EMBED_TITLE_LIMIT)
  assert.ok(embed.description.length <= EMBED_DESCRIPTION_LIMIT)
  assert.ok(embed.fields.every(field => field.value.length <= EMBED_FIELD_VALUE_LIMIT))
  assert.equal(embed.fields.some(field => field.name === "Group"), false)
})

test("state target validation succeeds for an accessible current target", async () => {
  const fixture = stateDiscordFixture()
  assert.equal(
    await resolveDeliveryTarget(fixture.client, statePayload(), { hasImage: false }),
    fixture.channel
  )
  const result = await createDiscordEventDeliveryHandler({
    client: fixture.client,
    gameProfile: "wos"
  })(statePayload())
  assert.deepEqual(result, { sentMessageId: "state-message" })
})

test("state delivery reuses persisted images without filesystem writes", async () => {
  let sentMessage
  const fixture = stateDiscordFixture({
    send: async message => {
      sentMessage = message
      return { id: "state-image-message" }
    }
  })
  const imageData = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
  const result = await createDiscordEventDeliveryHandler({
    client: fixture.client,
    gameProfile: "wos"
  })(statePayload({
    image: {
      originalFilename: "event.png",
      contentType: "image/png",
      byteSize: imageData.length,
      imageData
    }
  }))

  assert.deepEqual(result, { sentMessageId: "state-image-message" })
  assert.equal(sentMessage.files[0].name, "event-image.png")
  assert.deepEqual(sentMessage.files[0].attachment, imageData)
  assert.equal(sentMessage.embeds[0].toJSON().image.url, "attachment://event-image.png")
})

test("state target validation rejects stale, missing and unsupported targets", async () => {
  const stale = stateDiscordFixture()
  await assert.rejects(
    resolveDeliveryTarget(stale.client, statePayload({
      claim: { targetIsCurrent: false }
    }), { hasImage: false }),
    /sharing disabled or target changed/i
  )

  const missingGuild = stateDiscordFixture()
  missingGuild.client.guilds.cache.clear()
  missingGuild.client.guilds.fetch = async () => null
  await assert.rejects(
    resolveDeliveryTarget(missingGuild.client, statePayload(), { hasImage: false }),
    PermanentDeliveryError
  )

  const missingChannel = stateDiscordFixture()
  missingChannel.guild.channels.cache.clear()
  missingChannel.guild.channels.fetch = async () => null
  await assert.rejects(
    resolveDeliveryTarget(missingChannel.client, statePayload(), { hasImage: false }),
    PermanentDeliveryError
  )
  const wrongGuild = stateDiscordFixture({ channelOverrides: { guildId: "other" } })
  await assert.rejects(
    resolveDeliveryTarget(wrongGuild.client, statePayload(), { hasImage: false }),
    /another guild/
  )
  const unsupported = stateDiscordFixture({ channelOverrides: { type: ChannelType.GuildForum } })
  await assert.rejects(
    resolveDeliveryTarget(unsupported.client, statePayload(), { hasImage: false }),
    PermanentDeliveryError
  )
})

test("state target permissions and temporary fetch errors are classified correctly", async () => {
  const required = [
    PermissionFlagsBits.ViewChannel,
    PermissionFlagsBits.SendMessages,
    PermissionFlagsBits.EmbedLinks,
    PermissionFlagsBits.AttachFiles
  ]
  for (const missing of required.slice(0, 3)) {
    const fixture = stateDiscordFixture({
      permissions: required.filter(permission => permission !== missing)
    })
    await assert.rejects(
      resolveDeliveryTarget(fixture.client, statePayload(), { hasImage: false }),
      PermanentDeliveryError
    )
  }
  const noFiles = stateDiscordFixture({
    permissions: required.filter(permission => permission !== PermissionFlagsBits.AttachFiles)
  })
  await assert.rejects(
    resolveDeliveryTarget(noFiles.client, statePayload(), { hasImage: true }),
    /Attach Files/
  )

  const temporary = stateDiscordFixture()
  temporary.client.guilds.cache.clear()
  temporary.client.guilds.fetch = async () => {
    throw Object.assign(new Error("network"), { code: "ECONNRESET" })
  }
  const handler = createDiscordEventDeliveryHandler({
    client: temporary.client,
    gameProfile: "wos"
  })
  await assert.rejects(handler(statePayload()), RetryableDeliveryError)
})

function independentRepository(claims, payloads) {
  const outcomes = []
  return {
    gameProfile: "wos",
    outcomes,
    async listActiveEventDefinitions() { return [] },
    async insertMissingDeliveryClaims() { return 0 },
    async claimDueDeliveries() { return claims.splice(0) },
    async getClaimPayload({ claimId }) { return payloads.get(claimId) },
    async markClaimSent(input) { outcomes.push({ status: "sent", ...input }); return true },
    async markClaimFailed(input) { outcomes.push({ status: "retry", ...input }); return true },
    async markClaimPermanentlyFailed(input) {
      outcomes.push({ status: "permanent", ...input })
      return true
    }
  }
}

async function runIndependentDeliveries(resultByTarget) {
  const claims = [
    { id: "alliance", attempt_count: 1 },
    { id: "state", attempt_count: 1 }
  ]
  const payloads = new Map([
    ["alliance", { ...statePayload(), claim: {
      ...statePayload().claim,
      id: "alliance",
      targetKind: "alliance",
      targetGuildId: "alliance-guild",
      targetChannelId: "alliance-channel",
      targetIsCurrent: true
    } }],
    ["state", statePayload({ claim: { id: "state" } })]
  ])
  const repository = independentRepository(claims, payloads)
  const worker = createEventDeliveryWorker({
    env: { EVENT_SCHEDULER_ENABLED: "true" },
    health: { available: true, gameProfile: "wos", botInstanceName: "rachie-wos" },
    repository,
    gameProfile: "wos",
    botInstanceName: "rachie-wos",
    deliveryHandler: async delivery => {
      const result = resultByTarget[delivery.claim.targetKind]
      if (result instanceof Error) throw result
      return { sentMessageId: result }
    },
    logger: { error() {} },
    workerId: "phase7-worker",
    config: {
      lookaheadMinutes: 1440,
      graceMinutes: 60,
      pollIntervalMs: 5000,
      batchSize: 10,
      claimLeaseSeconds: 60,
      handlerTimeoutMs: 1000
    }
  })
  await worker.tick()
  return repository.outcomes
}

test("alliance and state outcomes remain independent in both directions", async () => {
  const allianceSent = await runIndependentDeliveries({
    alliance: "alliance-message",
    state: new RetryableDeliveryError("temporary")
  })
  assert.ok(allianceSent.some(item => item.claimId === "alliance" && item.status === "sent"))
  assert.ok(allianceSent.some(item => item.claimId === "state" && item.status === "retry"))

  const stateSent = await runIndependentDeliveries({
    alliance: new RetryableDeliveryError("temporary"),
    state: "state-message"
  })
  assert.ok(stateSent.some(item => item.claimId === "alliance" && item.status === "retry"))
  assert.ok(stateSent.some(item => item.claimId === "state" && item.status === "sent"))

  const statePermanent = await runIndependentDeliveries({
    alliance: "alliance-message",
    state: new PermanentDeliveryError("deleted channel")
  })
  assert.ok(statePermanent.some(item => item.claimId === "alliance" && item.status === "sent"))
  assert.ok(statePermanent.some(item => item.claimId === "state" && item.status === "permanent"))

  const alliancePermanent = await runIndependentDeliveries({
    alliance: new PermanentDeliveryError("deleted channel"),
    state: "state-message"
  })
  assert.ok(alliancePermanent.some(item => item.claimId === "alliance" && item.status === "permanent"))
  assert.ok(alliancePermanent.some(item => item.claimId === "state" && item.status === "sent"))
})

test("successful alliance and state deliveries retain separate message IDs", async () => {
  const outcomes = await runIndependentDeliveries({
    alliance: "alliance-message",
    state: "state-message"
  })
  assert.deepEqual(
    outcomes.filter(item => item.status === "sent")
      .map(item => `${item.claimId}:${item.sentMessageId}`).sort(),
    ["alliance:alliance-message", "state:state-message"]
  )
})
