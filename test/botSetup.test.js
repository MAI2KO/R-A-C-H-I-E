const test = require("node:test")
const assert = require("node:assert/strict")
const {
  ChannelType,
  PermissionFlagsBits
} = require("discord.js")

const {
  BOT_SETUP_IDS,
  PROFILE_NAMES,
  CHANNELS,
  permissionOverwrites,
  giftCard,
  createBotSetupService
} = require("../src/botSetupService")
const {
  buildBotSetupCommand,
  handleBotSetupInteraction,
  handlePersistentOnboardingInteraction
} = require("../src/botSetupInteractions")
const { profileTerminology } = require("../src/giftCodes/terminology")

function memoryRepository() {
  let stored = null
  const destinations = []
  return {
    async get() { return stored ? { ...stored } : null },
    async save(_guildId, values) { stored = { ...values }; return stored },
    async reconcileDestinations(_guildId, giftChannelId, eventChannelId) {
      destinations.push({ giftChannelId, eventChannelId })
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
    name: "gift-code-announcements",
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
    roles: { everyone: { id: "777777777777777777" } },
    members: {
      me: {
        id: "888888888888888888",
        permissions: { has: permission => permission !== PermissionFlagsBits.ManageChannels || manageChannels }
      }
    },
    channels: {
      async fetch(id) { return channels.get(id) || null },
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
      gameProfile: profile,
      logger: { log() {} }
    })
    const first = await service.reconcile(discord.guild.id)
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
    await service.reconcile(discord.guild.id)
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
    service.reconcile(discord.guild.id),
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
  await service.reconcile(discord.guild.id)
  const first = repository.state()
  const giftChannel = discord.channels.get(first.gift_auto_redeem_channel_id)
  giftChannel.messagesMap.delete(first.gift_auto_redeem_message_id)
  await service.reconcile(discord.guild.id)
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
  await assert.rejects(service.reconcile(discord.guild.id), /temporary Discord failure/)
  assert.equal(discord.created.length, 2)

  discord.guild.channels.create = create
  await service.reconcile(discord.guild.id)
  assert.equal(discord.created.filter(channel => channel.type === ChannelType.GuildCategory).length, 1)
  assert.equal(discord.created.filter(channel => channel.type === ChannelType.GuildText).length, 5)
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
  await service.reconcile(discord.guild.id)
  const count = discord.created.length
  discord.guild.channels.fetch = async () => {
    throw Object.assign(new Error("temporary fetch failure"), { code: "ECONNRESET" })
  }
  await assert.rejects(service.reconcile(discord.guild.id), /temporary fetch failure/)
  assert.equal(discord.created.length, count)
})

test("persistent cards expose only canonical player-facing registration buttons", async () => {
  for (const profile of ["wos", "kingshot"]) {
    const card = giftCard(profileTerminology(profile))
    assert.equal(card.components[0].components.length, 1)
    assert.equal(card.components[0].components[0].data.custom_id, BOT_SETUP_IDS.registerCharacter)
    assert.match(card.content, new RegExp(profile === "wos" ? "State" : "Kingdom"))
    assert.doesNotMatch(card.content, /diagnostic|queue|source|verifier/i)
  }

  const modal = { id: "canonical-minister-modal" }
  const interaction = {
    customId: BOT_SETUP_IDS.registerMinisters,
    isButton: () => true,
    async showModal(value) { this.modal = value }
  }
  assert.equal(await handlePersistentOnboardingInteraction(interaction, {
    ministerModalBuilder: () => modal
  }), true)
  assert.equal(interaction.modal, modal)
})

test("bot setup command remains admin-scoped and reports service errors privately", async () => {
  assert.equal(buildBotSetupCommand().toJSON().name, "bot-setup")
  const interaction = {
    commandName: "bot-setup",
    guildId: "777777777777777777",
    client: {},
    isChatInputCommand: () => true,
    async deferReply(value) { this.deferred = value },
    async editReply(value) { this.edited = value }
  }
  assert.equal(await handleBotSetupInteraction(interaction, {
    userCanManageServer: async () => true,
    healthProvider: () => ({ available: true, gameProfile: "wos" }),
    poolProvider: () => ({}),
    repositoryFactory: () => ({}),
    serviceFactory: () => ({ async reconcile() { return { content: "Configured" } } })
  }), true)
  assert.equal(interaction.edited, "Configured")
})
