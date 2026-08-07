const test = require("node:test")
const assert = require("node:assert/strict")
const { ChannelType, PermissionFlagsBits } = require("discord.js")

const {
  getEventSchedulerCommandData,
  getEventSchedulerHelpCommandData
} = require("../src/eventSchedulerCommands")
const {
  createEventSchedulerRepository
} = require("../src/eventSchedulerRepository")
const {
  SchedulerValidationError,
  normalizeAllianceName,
  resolveSendableChannel
} = require("../src/eventSchedulerService")

test("scheduler command is registered only when exactly enabled", () => {
  assert.equal(getEventSchedulerCommandData({}), null)
  assert.equal(getEventSchedulerHelpCommandData({}), null)
  assert.equal(getEventSchedulerCommandData({ EVENT_SCHEDULER_ENABLED: "TRUE" }), null)
  assert.equal(getEventSchedulerHelpCommandData({ EVENT_SCHEDULER_ENABLED: "TRUE" }), null)
  assert.equal(
    getEventSchedulerCommandData({ EVENT_SCHEDULER_ENABLED: "true" }).name,
    "event-scheduler"
  )
  assert.equal(
    getEventSchedulerHelpCommandData({ EVENT_SCHEDULER_ENABLED: "true" }).name,
    "event-scheduler-help"
  )
})

test("repository reads guild settings through the current profile", async () => {
  const calls = []
  const pool = {
    async query(text, values) {
      calls.push({ text, values })
      return { rows: [] }
    }
  }

  await createEventSchedulerRepository(pool, "wos").getGuildSettings("guild-1")
  await createEventSchedulerRepository(pool, "kingshot").getGuildSettings("guild-1")

  assert.deepEqual(calls[0].values, ["guild-1", "wos"])
  assert.deepEqual(calls[1].values, ["guild-1", "kingshot"])
  assert.match(calls[0].text, /game_profile = \$2/)
})

test("same guild can upsert separate profile settings", async () => {
  const calls = []
  const pool = {
    async query(text, values) {
      calls.push({ text, values })
      return { rows: [{ guild_id: values[0], game_profile: values[1] }] }
    }
  }
  const input = {
    guildId: "123",
    allianceName: "Alliance",
    eventChannelId: "456",
    botInstanceName: "instance"
  }

  await createEventSchedulerRepository(pool, "wos").upsertGuildSettings(input)
  await createEventSchedulerRepository(pool, "kingshot").upsertGuildSettings(input)

  assert.equal(calls[0].values[1], "wos")
  assert.equal(calls[1].values[1], "kingshot")
  assert.match(calls[0].text, /ON CONFLICT \(guild_id, game_profile\)/)
})

test("alliance names are normalized and bounded", () => {
  assert.equal(normalizeAllianceName("  North  "), "North")
  assert.throws(() => normalizeAllianceName(""), SchedulerValidationError)
  assert.throws(() => normalizeAllianceName("x".repeat(101)), SchedulerValidationError)
})

test("Discord target validation requires an accessible sendable channel", async () => {
  const permissions = { has: () => true }
  const channel = {
    guildId: "1234567890123456",
    type: ChannelType.GuildText,
    isTextBased: () => true,
    isSendable: () => true,
    permissionsFor: () => permissions
  }
  const guild = {
    channels: { fetch: async () => channel },
    members: { me: { id: "bot" } }
  }
  const client = { guilds: { fetch: async () => guild } }

  const result = await resolveSendableChannel(
    client,
    "1234567890123456",
    "2345678901234567"
  )
  assert.equal(result.channel, channel)

  channel.isTextBased = () => false
  await assert.rejects(
    resolveSendableChannel(client, "1234567890123456", "2345678901234567"),
    SchedulerValidationError
  )

  channel.isTextBased = () => true
  channel.type = ChannelType.GuildForum
  await assert.rejects(
    resolveSendableChannel(client, "1234567890123456", "2345678901234567"),
    SchedulerValidationError
  )

  channel.type = ChannelType.GuildText
  permissions.has = permission => permission !== PermissionFlagsBits.EmbedLinks
  await assert.rejects(
    resolveSendableChannel(client, "1234567890123456", "2345678901234567"),
    /Embed Links/
  )

  permissions.has = permission => permission !== PermissionFlagsBits.AttachFiles
  await assert.rejects(
    resolveSendableChannel(client, "1234567890123456", "2345678901234567"),
    /Attach Files/
  )
  await assert.doesNotReject(
    resolveSendableChannel(
      client,
      "1234567890123456",
      "2345678901234567",
      { requireAttachments: false }
    )
  )
})
