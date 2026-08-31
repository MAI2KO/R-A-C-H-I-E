const test = require("node:test")
const assert = require("node:assert/strict")
const {
  ChannelType,
  MessageFlags,
  PermissionFlagsBits
} = require("discord.js")

const {
  BOT_SETUP_IDS,
  PROFILE_NAMES,
  CHANNELS,
  permissionOverwrites,
  giftCard,
  ministerCard,
  setupStatus,
  createBotSetupService
} = require("../src/botSetupService")
const {
  buildBotSetupCommand,
  handleBotSetupInteraction,
  handlePersistentOnboardingInteraction,
  setupPublicError
} = require("../src/botSetupInteractions")
const { profileTerminology } = require("../src/giftCodes/terminology")
const community = { communityNumber: "9999", allianceAbbreviation: "HWC" }

function memoryRepository(initialDestinations = { gift: null, event: null },
  schedulerSummary = { configured: false, scheduledEvents: 0, stateLinked: false }) {
  let stored = null
  const destinations = []
  return {
    async get() { return stored ? { ...stored } : null },
    async save(_guildId, values) { stored = { ...values }; return stored },
    async getDestinations() { return initialDestinations },
    async getSchedulerSummary() { return schedulerSummary },
    async reconcileDestinations(input) {
      destinations.push(input)
    },
    state: () => stored,
    destinations
  }
}

function discordSetupFixture({ manageChannels = true } = {}) {
  let nextId = 100000000000000000n
  const channels = new Map()
  const created = []
  const unrelated = {
    id: "999999999999999999",
    name: "unrelated-chat",
    type: ChannelType.GuildText,
    parentId: null
  }
  channels.set(unrelated.id, unrelated)

  function makeChannel(input) {
    const id = String(nextId++)
    const messages = new Map()
    const sent = []
    const edits = []
    const overwriteSets = []
    const channel = {
      id,
      name: input.name,
      type: input.type,
      parentId: input.parent || null,
      initialOverwrites: input.permissionOverwrites || [],
      permissionOverwrites: {
        async set(values) { overwriteSets.push(values) }
      },
      messages: {
        async fetch(messageId) { return messages.get(messageId) || null }
      },
      async send(payload) {
        const message = {
          id: String(nextId++),
          payload,
          async edit(nextPayload) { this.payload = nextPayload; edits.push(nextPayload); return this }
        }
        messages.set(message.id, message)
        sent.push(payload)
        return message
      },
      messagesMap: messages,
      sent,
      edits,
      overwriteSets
    }
    channels.set(id, channel)
    created.push(channel)
    return channel
  }

  const guild = {
    id: "777777777777777777",
    name: "HoboswithCandy",
    roles: { everyone: { id: "777777777777777777" } },
    members: {
      me: {
        id: "888888888888888888",
        permissions: { has: permission => permission !== PermissionFlagsBits.ManageChannels || manageChannels }
      }
    },
    channels: {
      async fetch(id) { return id === undefined ? channels : channels.get(id) || null },
      async create(input) { return makeChannel(input) }
    }
  }
  return {
    client: { guilds: { async fetch() { return guild } } },
    guild,
    channels,
    created,
    unrelated
  }
}

test("bot setup creates one managed category, five channels and persistent cards", async () => {
  for (const profile of ["wos", "kingshot"]) {
    const discord = discordSetupFixture()
    const repository = memoryRepository()
    const service = createBotSetupService({
      repository,
      client: discord.client,
      gameProfile: profile, botInstanceName: `test-${profile}`,
      logger: { log() {} }
    })
    const first = await service.reconcile(discord.guild.id, community)
    assert.match(first.content, new RegExp(PROFILE_NAMES[profile].botName.replaceAll(".", "\\.")))
    assert.equal(discord.created.filter(channel => channel.type === ChannelType.GuildCategory).length, 1)
    assert.equal(discord.created.filter(channel => channel.type === ChannelType.GuildText).length, 5)
    assert.deepEqual(
      discord.created.filter(channel => channel.type === ChannelType.GuildText).map(channel => channel.name),
      CHANNELS.map(channel => channel.name)
    )
    assert.equal(repository.destinations.length, 1)
    assert.equal(discord.channels.has(discord.unrelated.id), true)

    const before = discord.created.length
    await service.reconcile(discord.guild.id, community)
    assert.equal(discord.created.length, before, "repeated setup created duplicate channels")
    const managed = repository.state()
    assert.equal(discord.channels.get(managed.gift_auto_redeem_channel_id).sent.length, 1)
    assert.equal(discord.channels.get(managed.gift_auto_redeem_channel_id).edits.length, 1)
  }
})

test("setup permissions keep parents read-only and allow public threads only in announcement channels", () => {
  const discord = discordSetupFixture()
  for (const allowThreads of [false, true]) {
    const overwrites = permissionOverwrites(discord.guild, allowThreads)
    const member = overwrites[0]
    const bot = overwrites[1]
    assert.ok(member.allow.includes(PermissionFlagsBits.ViewChannel))
    assert.ok(member.deny.includes(PermissionFlagsBits.SendMessages))
    assert.equal(member.allow.includes(PermissionFlagsBits.CreatePublicThreads), allowThreads)
    assert.equal(member.allow.includes(PermissionFlagsBits.SendMessagesInThreads), allowThreads)
    assert.equal(member.deny.includes(PermissionFlagsBits.CreatePublicThreads), !allowThreads)
    for (const permission of [
      PermissionFlagsBits.ViewChannel,
      PermissionFlagsBits.SendMessages,
      PermissionFlagsBits.EmbedLinks,
      PermissionFlagsBits.AttachFiles,
      PermissionFlagsBits.ReadMessageHistory
    ]) assert.ok(bot.allow.includes(permission))
  }
})

test("missing Manage Channels stops setup before partial creation with clear guidance", async () => {
  const discord = discordSetupFixture({ manageChannels: false })
  const service = createBotSetupService({
    repository: memoryRepository(),
    client: discord.client,
    gameProfile: "wos"
  })
  await assert.rejects(
    service.reconcile(discord.guild.id, community),
    error => error.code === "MANAGE_CHANNELS_REQUIRED"
      && /Grant it to the existing bot role/.test(error.message)
  )
  assert.equal(discord.created.length, 0)
})

test("deleted setup messages are recreated while stored live messages are edited", async () => {
  const discord = discordSetupFixture()
  const repository = memoryRepository()
  const service = createBotSetupService({
    repository,
    client: discord.client,
    gameProfile: "wos",
    logger: { log() {} }
  })
  await service.reconcile(discord.guild.id, community)
  const first = repository.state()
  const giftChannel = discord.channels.get(first.gift_auto_redeem_channel_id)
  giftChannel.messagesMap.delete(first.gift_auto_redeem_message_id)
  await service.reconcile(discord.guild.id, community)
  assert.equal(giftChannel.sent.length, 2)
  assert.notEqual(repository.state().gift_auto_redeem_message_id, first.gift_auto_redeem_message_id)
  assert.equal(discord.channels.get(first.minister_sign_up_channel_id).edits.length, 1)
})

test("partial setup checkpoints make a restart safe without duplicate resources", async () => {
  const discord = discordSetupFixture()
  const repository = memoryRepository()
  const service = createBotSetupService({
    repository,
    client: discord.client,
    gameProfile: "wos",
    logger: { log() {} }
  })
  const create = discord.guild.channels.create
  let calls = 0
  discord.guild.channels.create = async input => {
    calls += 1
    if (calls === 3) throw new Error("temporary Discord failure")
    return create(input)
  }
  await assert.rejects(service.reconcile(discord.guild.id, community), /temporary Discord failure/)
  assert.equal(discord.created.length, 2)

  discord.guild.channels.create = create
  await service.reconcile(discord.guild.id, community)
  assert.equal(discord.created.filter(channel => channel.type === ChannelType.GuildCategory).length, 1)
  assert.equal(discord.created.filter(channel => channel.type === ChannelType.GuildText).length, 5)
})

test("setup preview is read-only and a deleted expected channel alone is recreated", async () => {
  const discord = discordSetupFixture()
  const repository = memoryRepository()
  const service = createBotSetupService({
    repository, client: discord.client, gameProfile: "wos", logger: { log() {} }
  })
  const preview = await service.preview(discord.guild.id, community)
  assert.equal(discord.created.length, 0)
  assert.equal(repository.state(), null)
  assert.match(preview.content, /Missing:[\s\S]*#gift-code-announcements/)

  await service.reconcile(discord.guild.id, community)
  const managed = repository.state()
  discord.channels.delete(managed.event_announcements_channel_id)
  const before = discord.created.length
  const result = await service.reconcile(discord.guild.id, community)
  assert.equal(discord.created.length, before + 1)
  assert.deepEqual(result.created, ["#event-announcements"])
  assert.ok(result.reused.includes("#event-scheduler"))
  assert.notEqual(repository.state().event_announcements_channel_id,
    managed.event_announcements_channel_id)
})

test("setup preserves valid custom announcement destinations", async () => {
  const discord = discordSetupFixture()
  const gift = await discord.guild.channels.create({ name: "custom-gifts", type: ChannelType.GuildText })
  const events = await discord.guild.channels.create({ name: "custom-events", type: ChannelType.GuildText })
  const roundups = await discord.guild.channels.create({ name: "custom-roundups", type: ChannelType.GuildText })
  const repository = memoryRepository({
    gift: { gift_code_channel_id: gift.id },
    event: { event_channel_id: events.id, weekly_roundup_channel_id: roundups.id }
  }, { configured: true, scheduledEvents: 4, stateLinked: true })
  const service = createBotSetupService({
    repository, client: discord.client, gameProfile: "kingshot",
    botInstanceName: "test-kingshot", logger: { log() {} }
  })
  const result = await service.reconcile(discord.guild.id, community,
    "native community created and linked; bookings closed", { created: true })
  assert.deepEqual(result.destinations, {
    giftDestination: gift.id,
    eventDestination: events.id,
    roundupDestination: roundups.id
  })
  assert.deepEqual(repository.destinations[0], {
    guildId: discord.guild.id, giftChannelId: gift.id, eventChannelId: events.id,
    roundupChannelId: roundups.id, botInstanceName: "test-kingshot",
    allianceAbbreviation: "HWC"
  })
  assert.ok(result.created.includes("Native booking community"))
  assert.ok(result.reused.includes("Existing event scheduler configuration"))
  assert.ok(result.reused.includes("4 scheduled events"))
  assert.ok(result.reused.includes("State/Kingdom roundup linkage"))
  assert.match(result.content, /Native booking community/)
  assert.match(result.content, /4 scheduled events/)
})

test("transient Discord fetch errors do not create replacement channels", async () => {
  const discord = discordSetupFixture()
  const repository = memoryRepository()
  const service = createBotSetupService({
    repository,
    client: discord.client,
    gameProfile: "wos",
    logger: { log() {} }
  })
  await service.reconcile(discord.guild.id, community)
  const count = discord.created.length
  discord.guild.channels.fetch = async () => {
    throw Object.assign(new Error("temporary fetch failure"), { code: "ECONNRESET" })
  }
  await assert.rejects(service.reconcile(discord.guild.id, community), /temporary fetch failure/)
  assert.equal(discord.created.length, count)
})

test("persistent cards expose only canonical player-facing registration controls", async () => {
  for (const profile of ["wos", "kingshot"]) {
    const card = giftCard(profileTerminology(profile))
    assert.equal(card.components[0].components.length, 1)
    assert.equal(card.components[0].components[0].data.custom_id, BOT_SETUP_IDS.registerCharacter)
    assert.match(card.content, new RegExp(profile === "wos" ? "State" : "Kingdom"))
    assert.doesNotMatch(card.content, /diagnostic|queue|source|verifier/i)
  }

  const minister = ministerCard(profileTerminology("wos"))
  assert.equal(minister.components.length, 0)
  assert.match(minister.content, /authenticated community website/)

  const interaction = {
    customId: BOT_SETUP_IDS.registerMinisters,
    isButton: () => true,
    async reply(value) { this.replied = value }
  }
  assert.equal(await handlePersistentOnboardingInteraction(interaction), true)
  assert.match(interaction.replied.content, /moved to the authenticated community website/)
  assert.equal(interaction.replied.flags, MessageFlags.Ephemeral)
})

test("bot setup command remains admin-scoped and reports service errors privately", async () => {
  assert.equal(buildBotSetupCommand().toJSON().name, "setup")
  const interaction = {
    commandName: "setup",
    guildId: "777777777777777777",
    guild: { ownerId: "111111111111111111" },
    user: { id: "111111111111111111" },
    client: {},
    isChatInputCommand: () => true,
    async showModal(value) { this.modal = value }
  }
  assert.equal(await handleBotSetupInteraction(interaction, {
    userCanManageServer: async () => true,
    healthProvider: () => ({ available: true, gameProfile: "wos" }),
    bookingApi: { communitySetup: async () => ({ status: "ready" }) }
  }), true)
  assert.equal(interaction.modal.data.custom_id, "botsetup:community")
  assert.equal(interaction.modal.components[0].components[0].data.label, "State number")
  assert.equal(interaction.modal.components[1].components[0].data.label, "Alliance abbreviation")
})

test("canonical setup status reports reconciliation without pulling specialised configuration in", () => {
  for (const profile of ["wos", "kingshot"]) {
    const status = setupStatus(profile, {
      community: { communityNumber: "9999", guildName: "HoboswithCandy", allianceAbbreviation: "HWC" },
      created: ["#gift-code-announcements"], reused: ["#event-scheduler"], bookingStatus: "linked"
    })
    assert.match(status, profile === "wos" ? /State: 9999/ : /Kingdom: 9999/)
    assert.match(status, /`\/event-scheduler`/)
    assert.doesNotMatch(status, /bot-admin|banter/i)
  }
})

test("native setup conflicts and unsupported Kingshot defaults have controlled guidance", () => {
  assert.match(setupPublicError({ code: "community_claim_conflict" }, "wos"),
    /State.*already linked.*Platform approval/)
  assert.match(setupPublicError({ code: "guild_conflict" }, "kingshot"),
    /Kingdom.*already linked.*no mapping was changed/)
  assert.match(setupPublicError({ code: "kingshot_defaults_unavailable" }, "kingshot"),
    /not yet configured.*No setup was changed/)
  assert.equal(setupPublicError({ code: "network_error" }, "wos"), null)
})
