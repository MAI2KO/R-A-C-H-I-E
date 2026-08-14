const test = require("node:test")
const assert = require("node:assert/strict")
const fs = require("fs")
const path = require("path")
const { EventEmitter } = require("events")
const {
  ChannelType,
  PermissionFlagsBits
} = require("discord.js")

const {
  EMBED_TITLE_LIMIT,
  EMBED_DESCRIPTION_LIMIT,
  formatAllianceEventDelivery
} = require("../src/eventDeliveryFormatting")
const {
  PermanentDeliveryError,
  RetryableDeliveryError,
  createEventDeliveryWorker
} = require("../src/eventDeliveryWorker")
const {
  normalizeDiscordDeliveryError,
  prepareStoredEventImage,
  resolveAllianceTarget,
  createDiscordEventDeliveryHandler
} = require("../src/discordEventDelivery")
const {
  createAllianceEventDeliveryRuntime
} = require("../src/allianceEventDeliveryRuntime")
const {
  createEventDeliveryRepository
} = require("../src/eventDeliveryRepository")

const OCCURRENCE = new Date("2026-08-10T18:30:00Z")

function payload(overrides = {}) {
  const base = {
    claim: {
      id: "1",
      gameProfile: "wos",
      attemptCount: 1,
      deliveryKind: "advance_reminder",
      targetKind: "alliance",
      targetGuildId: "guild-1",
      targetChannelId: "channel-1",
      occurrenceAt: new Date(OCCURRENCE),
      deliverAt: new Date("2026-08-10T18:20:00Z")
    },
    event: {
      id: "41",
      guildId: "guild-1",
      eventName: "Bear Hunt",
      recurrenceDays: 7
    },
    alliance: { name: "North", guildId: "guild-1" },
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

function embedJson(message) {
  return message.embeds[0].toJSON()
}

test("advance reminder embeds cover 10, 30 and grouped events", () => {
  const ten = formatAllianceEventDelivery(payload())
  const tenEmbed = embedJson(ten)
  assert.match(tenEmbed.title, /Event reminder: Bear Hunt/)
  assert.equal(
    tenEmbed.description,
    "**North**\n**Bear Hunt**\n\nStarts in 10 minutes\n\n" +
      "Start time\n18:30 UTC\nLocal time: <t:1786386600:t>"
  )
  assert.equal(tenEmbed.fields, undefined)
  assert.doesNotMatch(JSON.stringify(tenEmbed), /Alliance:|Group:|Status:|When|recurrence|2026-08-10/i)

  const thirty = embedJson(formatAllianceEventDelivery(payload({
    claim: { deliverAt: new Date("2026-08-10T18:00:00Z") }
  })))
  assert.match(thirty.description, /Starts in 30 minutes/)

  const grouped = embedJson(formatAllianceEventDelivery(payload({
    group: { id: "2", name: "Alpha", eventTimeUtc: "18:30", sortOrder: 0 }
  })))
  assert.match(grouped.description, /^\*\*North\*\*\n\*\*Bear Hunt\*\*\nAlpha/)
})

test("final reminder embeds say about to start one minute before", () => {
  const final = embedJson(formatAllianceEventDelivery(payload({
    claim: { deliveryKind: "final_reminder", deliverAt: new Date("2026-08-10T18:29:00Z") },
    event: { recurrenceDays: 3 }
  })))
  assert.match(final.title, /About to start: Bear Hunt/)
  assert.equal(
    final.description,
    "**North**\n**Bear Hunt**\n\nAbout to start\nStarts in approximately 1 minute\n\n" +
      "Start time\n18:30 UTC\nLocal time: <t:1786386600:t>"
  )
  assert.equal(final.fields, undefined)
  assert.doesNotMatch(JSON.stringify(final), /Starting now|Has started|Event start/i)

  const grouped = embedJson(formatAllianceEventDelivery(payload({
    claim: { deliveryKind: "final_reminder", deliverAt: new Date("2026-08-10T18:29:00Z") },
    group: { id: "2", name: "Beta", eventTimeUtc: "18:30", sortOrder: 0 }
  })))
  assert.match(grouped.description, /Beta/)
  assert.throws(() => formatAllianceEventDelivery(payload({
    claim: { deliveryKind: "event_start", deliverAt: new Date(OCCURRENCE) }
  })), PermanentDeliveryError)
})

test("formatting rejects malformed reminder intervals", () => {
  assert.throws(() => formatAllianceEventDelivery(payload({
    claim: { deliverAt: new Date("2026-08-10T18:19:31Z") }
  })), PermanentDeliveryError)
})

test("formatting neutralizes mentions and remains within Discord limits", () => {
  const message = formatAllianceEventDelivery(payload({
    event: { eventName: `@everyone <@123> ${"x".repeat(5000)}` },
    alliance: { name: `@here ${"y".repeat(5000)}` },
    group: { name: `<@&456> ${"z".repeat(5000)}` }
  }))
  const embed = embedJson(message)
  const serialized = JSON.stringify(embed)
  assert.doesNotMatch(serialized, /@everyone|@here|<@123>|<@&456>/)
  assert.deepEqual(message.allowedMentions, { parse: [], repliedUser: false })
  assert.ok(embed.title.length <= EMBED_TITLE_LIMIT)
  assert.ok(embed.description.length <= EMBED_DESCRIPTION_LIMIT)
  assert.equal(embed.fields, undefined)
})

const imageFixtures = Object.freeze({
  "image/png": Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
  "image/jpeg": Buffer.from([0xff, 0xd8, 0xff]),
  "image/gif": Buffer.from("GIF89a", "ascii"),
  "image/webp": Buffer.from("RIFFxxxxWEBP", "ascii")
})

test("stored PNG, JPEG, GIF and WebP images receive safe generated filenames", () => {
  const expected = ["event-image.png", "event-image.jpg", "event-image.gif", "event-image.webp"]
  const prepared = Object.entries(imageFixtures).map(([contentType, data]) =>
    prepareStoredEventImage({
      originalFilename: "@everyone/../../unsafe.exe",
      contentType,
      byteSize: data.length,
      imageData: data
    })
  )
  assert.deepEqual(prepared.map(item => item.filename), expected)
  assert.ok(prepared.every(item => Buffer.isBuffer(item.file.attachment)))
  assert.equal(prepareStoredEventImage(null), null)
})

test("corrupt or unsupported stored images fail permanently", () => {
  assert.throws(() => prepareStoredEventImage({
    contentType: "image/png",
    byteSize: 9,
    imageData: imageFixtures["image/png"]
  }), PermanentDeliveryError)
  assert.throws(() => prepareStoredEventImage({
    contentType: "text/plain",
    byteSize: 3,
    imageData: Buffer.from("bad")
  }), PermanentDeliveryError)
  assert.throws(() => prepareStoredEventImage({
    contentType: "image/png",
    byteSize: 3,
    imageData: Buffer.from("bad")
  }), PermanentDeliveryError)
  const deliverySource = fs.readFileSync(
    path.join(__dirname, "..", "src", "discordEventDelivery.js"),
    "utf8"
  )
  assert.doesNotMatch(deliverySource, /writeFile|createWriteStream|tmpdir/)
})

function discordFixture({
  permissions = [
    PermissionFlagsBits.ViewChannel,
    PermissionFlagsBits.SendMessages,
    PermissionFlagsBits.EmbedLinks,
    PermissionFlagsBits.AttachFiles
  ],
  channelOverrides = {},
  guildOverrides = {},
  send = async () => ({ id: "discord-message-1" }),
  ready = true
} = {}) {
  const permissionSet = new Set(permissions)
  const botMember = { id: "bot-member" }
  const channel = {
    id: "channel-1",
    guildId: "guild-1",
    type: ChannelType.GuildText,
    isTextBased: () => true,
    isSendable: () => true,
    permissionsFor: member => member === botMember
      ? { has: permission => permissionSet.has(permission) }
      : null,
    send,
    ...channelOverrides
  }
  const guild = {
    id: "guild-1",
    members: { me: botMember, fetchMe: async () => botMember },
    channels: {
      cache: new Map([["channel-1", channel]]),
      fetch: async id => id === "channel-1" ? channel : null
    },
    ...guildOverrides
  }
  const client = {
    isReady: () => ready,
    guilds: {
      cache: new Map([["guild-1", guild]]),
      fetch: async id => id === "guild-1" ? guild : null
    }
  }
  return { client, guild, channel, botMember }
}

test("alliance target validation accepts guild text and announcement channels", async () => {
  for (const type of [ChannelType.GuildText, ChannelType.GuildAnnouncement]) {
    const fixture = discordFixture({ channelOverrides: { type } })
    assert.equal(
      await resolveAllianceTarget(fixture.client, payload(), { hasImage: false }),
      fixture.channel
    )
  }
})

test("target validation rejects missing or mismatched guilds and channels", async () => {
  const missingGuild = discordFixture()
  missingGuild.client.guilds.cache.clear()
  missingGuild.client.guilds.fetch = async () => null
  await assert.rejects(
    resolveAllianceTarget(missingGuild.client, payload(), { hasImage: false }),
    PermanentDeliveryError
  )

  const missingChannel = discordFixture()
  missingChannel.guild.channels.cache.clear()
  missingChannel.guild.channels.fetch = async () => null
  await assert.rejects(
    resolveAllianceTarget(missingChannel.client, payload(), { hasImage: false }),
    PermanentDeliveryError
  )

  const wrongGuild = discordFixture({ channelOverrides: { guildId: "guild-2" } })
  await assert.rejects(
    resolveAllianceTarget(wrongGuild.client, payload(), { hasImage: false }),
    /another guild/
  )
  const unsupported = discordFixture({ channelOverrides: { type: ChannelType.GuildForum } })
  await assert.rejects(
    resolveAllianceTarget(unsupported.client, payload(), { hasImage: false }),
    /not a sendable/
  )
})

test("target validation enforces each required permission", async () => {
  const basePermissions = [
    PermissionFlagsBits.ViewChannel,
    PermissionFlagsBits.SendMessages,
    PermissionFlagsBits.EmbedLinks,
    PermissionFlagsBits.AttachFiles
  ]
  for (const missing of basePermissions.slice(0, 3)) {
    const fixture = discordFixture({
      permissions: basePermissions.filter(permission => permission !== missing)
    })
    await assert.rejects(
      resolveAllianceTarget(fixture.client, payload(), { hasImage: false }),
      PermanentDeliveryError
    )
  }
  const noAttachment = discordFixture({
    permissions: basePermissions.filter(permission => permission !== PermissionFlagsBits.AttachFiles)
  })
  assert.equal(
    await resolveAllianceTarget(noAttachment.client, payload(), { hasImage: false }),
    noAttachment.channel
  )
  await assert.rejects(
    resolveAllianceTarget(noAttachment.client, payload(), { hasImage: true }),
    /Attach Files/
  )
})

test("Discord error classification separates transient and permanent failures safely", () => {
  for (const error of [
    Object.assign(new Error("socket"), { code: "ECONNRESET" }),
    Object.assign(new Error("server"), { status: 503 }),
    Object.assign(new Error("rate"), { status: 429 })
  ]) {
    assert.ok(normalizeDiscordDeliveryError(error) instanceof RetryableDeliveryError)
  }
  for (const error of [
    Object.assign(new Error("unknown channel"), { code: 10003, status: 404 }),
    Object.assign(new Error("missing permission"), { code: 50013, status: 403 }),
    Object.assign(new Error("invalid body"), { code: 50035, status: 400 })
  ]) {
    assert.ok(normalizeDiscordDeliveryError(error) instanceof PermanentDeliveryError)
  }
  const sensitive = normalizeDiscordDeliveryError(new PermanentDeliveryError(
    `BOT_TOKEN=sensitive-value https://cdn.discordapp.com/attachment ${"x".repeat(600)}`
  ))
  assert.doesNotMatch(sensitive.message, /sensitive-value|discordapp/)
  assert.ok(sensitive.message.length <= 500)
})

test("Discord handler sends one mention-safe alliance embed and returns its message ID", async () => {
  let sentOptions
  const fixture = discordFixture({
    send: async options => {
      sentOptions = options
      return { id: "discord-message-42" }
    }
  })
  const imageData = imageFixtures["image/png"]
  const handler = createDiscordEventDeliveryHandler({ client: fixture.client, gameProfile: "wos" })
  const result = await handler(payload({
    claim: { deliverAt: new Date("2026-08-10T18:00:00Z") },
    image: {
      originalFilename: "unsafe.png",
      contentType: "image/png",
      byteSize: imageData.length,
      imageData
    }
  }))
  assert.deepEqual(result, { sentMessageId: "discord-message-42" })
  assert.deepEqual(sentOptions.allowedMentions, { parse: [], repliedUser: false })
  assert.equal(sentOptions.files[0].name, "event-image.png")
  assert.equal(sentOptions.embeds[0].toJSON().image.url, "attachment://event-image.png")
})

test("stored images are attached to advance reminders and omitted from final announcements", async () => {
  const sent = []
  const fixture = discordFixture({
    permissions: [
      PermissionFlagsBits.ViewChannel,
      PermissionFlagsBits.SendMessages,
      PermissionFlagsBits.EmbedLinks,
      PermissionFlagsBits.AttachFiles
    ],
    send: async options => {
      sent.push(options)
      return { id: `text-${sent.length}` }
    }
  })
  const imageData = imageFixtures["image/png"]
  const image = {
    originalFilename: "event.png",
    contentType: "image/png",
    byteSize: imageData.length,
    imageData
  }
  const handler = createDiscordEventDeliveryHandler({ client: fixture.client, gameProfile: "wos" })
  await handler(payload({ image }))
  await handler(payload({
    claim: {
      deliveryKind: "final_reminder",
      deliverAt: new Date("2026-08-10T18:29:00Z")
    },
    image
  }))
  assert.equal(sent.length, 2)
  assert.equal(sent[0].files.length, 1)
  assert.equal(sent[0].embeds[0].toJSON().image.url, "attachment://event-image.png")
  assert.equal(sent[1].files, undefined)
  assert.equal(sent[1].embeds[0].toJSON().image, undefined)
})

test("Discord handler sends text-only embeds without requiring attachments", async () => {
  let sentOptions
  const fixture = discordFixture({
    permissions: [
      PermissionFlagsBits.ViewChannel,
      PermissionFlagsBits.SendMessages,
      PermissionFlagsBits.EmbedLinks
    ],
    send: async options => {
      sentOptions = options
      return { id: "text-only" }
    }
  })
  const result = await createDiscordEventDeliveryHandler({
    client: fixture.client,
    gameProfile: "wos"
  })(payload())
  assert.equal(result.sentMessageId, "text-only")
  assert.equal(sentOptions.files, undefined)
})

test("Discord handlers reject state and cross-profile claims permanently", async () => {
  const fixture = discordFixture()
  const wos = createDiscordEventDeliveryHandler({ client: fixture.client, gameProfile: "wos" })
  const kingshot = createDiscordEventDeliveryHandler({
    client: fixture.client,
    gameProfile: "kingshot"
  })
  await assert.rejects(wos(payload({ claim: { targetKind: "state" } })), PermanentDeliveryError)
  await assert.rejects(wos(payload({ claim: { gameProfile: "kingshot" } })), PermanentDeliveryError)
  await assert.rejects(kingshot(payload()), PermanentDeliveryError)
  assert.equal(
    (await kingshot(payload({ claim: { gameProfile: "kingshot" } }))).sentMessageId,
    "discord-message-1"
  )
})

test("alliance repository scope excludes state claims before handler delivery", async () => {
  const calls = []
  const client = {
    async query(text, values) {
      calls.push({ text, values })
      return { rows: [], rowCount: 0 }
    },
    release() {}
  }
  const repository = createEventDeliveryRepository({
    query: client.query.bind(client),
    async connect() { return client }
  }, "wos", { targetKind: "alliance" })
  await repository.claimDueDeliveries({
    now: new Date(),
    batchSize: 10,
    leaseSeconds: 60,
    botInstanceName: "rachie-wos",
    workerId: "worker"
  })
  const claimQuery = calls.find(call => call.text.includes("FOR UPDATE SKIP LOCKED"))
  assert.match(claimQuery.text, /target_kind = \$8/)
  assert.equal(claimQuery.values[7], "alliance")
})

function workerRepository(claims) {
  const calls = []
  return {
    gameProfile: "wos",
    calls,
    async listActiveEventDefinitions() { return [] },
    async insertMissingDeliveryClaims() { return 0 },
    async claimDueDeliveries() { return claims.splice(0) },
    async getClaimPayload() { return payload() },
    async markClaimSent(input) { calls.push({ method: "sent", input }); return true },
    async markClaimFailed(input) { calls.push({ method: "failed", input }); return true },
    async markClaimPermanentlyFailed(input) {
      calls.push({ method: "permanent", input })
      return true
    }
  }
}

function workerConfig() {
  return {
    lookaheadMinutes: 1440,
    graceMinutes: 60,
    pollIntervalMs: 5000,
    batchSize: 10,
    claimLeaseSeconds: 60,
    handlerTimeoutMs: 1000
  }
}

test("real handler integrates with worker success, retry and permanent paths", async () => {
  const scenarios = [
    {
      send: async () => ({ id: "worker-message" }),
      expected: "sent"
    },
    {
      send: async () => { throw Object.assign(new Error("server"), { status: 503 }) },
      expected: "failed"
    },
    {
      send: async () => { throw Object.assign(new Error("forbidden"), { code: 50013 }) },
      expected: "permanent"
    }
  ]
  for (const scenario of scenarios) {
    const fixture = discordFixture({ send: scenario.send })
    const repository = workerRepository([{ id: "1", attempt_count: 1 }])
    const worker = createEventDeliveryWorker({
      env: { EVENT_SCHEDULER_ENABLED: "true" },
      health: { available: true, gameProfile: "wos", botInstanceName: "rachie-wos" },
      repository,
      gameProfile: "wos",
      botInstanceName: "rachie-wos",
      deliveryHandler: createDiscordEventDeliveryHandler({
        client: fixture.client,
        gameProfile: "wos"
      }),
      logger: { error() {} },
      workerId: "phase6-test-worker",
      config: workerConfig()
    })
    await worker.tick()
    assert.ok(repository.calls.some(call => call.method === scenario.expected))
  }
})

function runtimeFixture(overrides = {}) {
  let ready = overrides.ready ?? true
  const calls = []
  const worker = {
    start() { calls.push("worker.start"); return { started: true, workerId: "worker" } },
    async stop() { calls.push("worker.stop"); return { drained: true } }
  }
  const client = {
    isReady: () => ready,
    destroy() { calls.push("client.destroy") }
  }
  const runtime = createAllianceEventDeliveryRuntime({
    client,
    initializationPromise: Promise.resolve(overrides.health || {
      available: true,
      gameProfile: "wos",
      botInstanceName: "rachie-wos"
    }),
    env: { EVENT_SCHEDULER_ENABLED: "true" },
    logger: { log() {}, error() {} },
    getPoolFn() { calls.push("getPool"); return {} },
    createRepositoryFn(_pool, _profile, options) {
      calls.push(`createRepository:${options.targetKind}`)
      return { gameProfile: "wos" }
    },
    createHandlerFn() { calls.push("createHandler"); return async () => ({}) },
    createRoundupRepositoryFn() {
      calls.push("createRoundupRepository")
      return { gameProfile: "wos" }
    },
    createRoundupDeliveryFn() { calls.push("createRoundupDelivery"); return async () => ({}) },
    createRoundupProcessorFn() { calls.push("createRoundupProcessor"); return { tick: async () => 0 } },
    createWorkerFn() { calls.push("createWorker"); return worker },
    async shutdownFn({ worker: activeWorker }) {
      calls.push("shutdown")
      if (activeWorker) await activeWorker.stop()
      return { workerDrained: true }
    }
  })
  return { runtime, client, calls, setReady(value) { ready = value } }
}

test("runtime waits for client readiness and starts exactly once", async () => {
  const fixture = runtimeFixture({ ready: false })
  assert.deepEqual(await fixture.runtime.start(), {
    started: false,
    reason: "Discord client not ready"
  })
  assert.deepEqual(fixture.calls, [])
  fixture.setReady(true)
  assert.equal((await fixture.runtime.start()).started, true)
  assert.equal((await fixture.runtime.start()).started, true)
  assert.ok(fixture.calls.includes("createRepository:alliance"))
  assert.equal(fixture.calls.filter(call => call === "worker.start").length, 1)
})

test("runtime does not start when scheduler database health is unavailable", async () => {
  const fixture = runtimeFixture({ health: { available: false, reason: "database failure" } })
  assert.deepEqual(await fixture.runtime.start(), {
    started: false,
    reason: "database failure"
  })
  assert.ok(!fixture.calls.includes("createWorker"))
})

test("runtime shutdown is composable and signal handlers install once", async () => {
  const fixture = runtimeFixture()
  await fixture.runtime.start()
  const processRef = new EventEmitter()
  assert.equal(fixture.runtime.installShutdownHandlers(processRef), true)
  assert.equal(fixture.runtime.installShutdownHandlers(processRef), false)
  processRef.emit("SIGTERM")
  await new Promise(resolve => setImmediate(resolve))
  assert.ok(fixture.calls.includes("worker.stop"))
  assert.ok(fixture.calls.includes("client.destroy"))
})

test("shutdown during initialization prevents a later worker start", async () => {
  let resolveHealth
  const initializationPromise = new Promise(resolve => { resolveHealth = resolve })
  const calls = []
  const runtime = createAllianceEventDeliveryRuntime({
    client: { isReady: () => true },
    initializationPromise,
    env: { EVENT_SCHEDULER_ENABLED: "true" },
    logger: { log() {}, error() {} },
    getPoolFn() { calls.push("getPool"); return {} },
    createRepositoryFn() { return { gameProfile: "wos" } },
    createHandlerFn() { return async () => ({}) },
    createWorkerFn() { calls.push("createWorker"); return { start: () => ({ started: true }) } },
    async shutdownFn() { calls.push("shutdown"); return { workerDrained: true } }
  })
  const starting = runtime.start()
  await runtime.stop()
  resolveHealth({ available: true, gameProfile: "wos", botInstanceName: "rachie-wos" })
  assert.deepEqual(await starting, { started: false, reason: "stopped" })
  assert.ok(!calls.includes("getPool"))
  assert.ok(!calls.includes("createWorker"))
})

test("index startup wiring preserves registration, login and Apps Script contracts", () => {
  const source = fs.readFileSync(path.join(__dirname, "..", "index.js"), "utf8")
  const registration = source.indexOf("await registerCommands()")
  const login = source.indexOf("await client.login(process.env.BOT_TOKEN)")
  assert.ok(registration > 0 && login > registration)
  assert.match(source, /client\.once\("clientReady"/)
  assert.match(source, /allianceEventDeliveryRuntime\?\.start\(\)/)
  assert.match(source, /getEventSchedulerHelpCommandData/)
  assert.match(source, /commands\.push\(eventSchedulerHelpCommand\)/)
  assert.match(source, /createAppsScriptTransport\(\)/)
  assert.match(source, /createBookingAppsScriptClient/)
  assert.match(source, /createStateAppsScriptClient/)
  assert.match(source, /createConfigAppsScriptClient/)
  assert.match(source, /createBanterAppsScriptClient/)
  assert.match(source, /announcement_channel_id/)
  assert.doesNotMatch(source, /await initializeEventSchedulerSubsystem\(\)/)
})
