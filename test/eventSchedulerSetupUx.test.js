const test = require("node:test")
const assert = require("node:assert/strict")
const { ChannelType, ComponentType, MessageFlags, PermissionFlagsBits } = require("discord.js")

const {
  IDS,
  buildAllianceIdentityModal,
  buildChannelConfigurationView,
  buildHomeView,
  buildStateDestinationView,
  buildStateLinkModal,
  handleEventSchedulerInteraction
} = require("../src/eventSchedulerInteractions")
const {
  SchedulerValidationError,
  resolveSendableChannel
} = require("../src/eventSchedulerService")
const {
  generateStateLinkCode,
  hashStateLinkCode,
  normalizeStateLinkCode
} = require("../src/stateLinkCodes")

const guildId = "1234567890123456"
const channelId = "2345678901234567"

function discordTarget({
  targetGuildId = guildId,
  type = ChannelType.GuildText,
  deniedPermission = null
} = {}) {
  const permissions = { has: permission => permission !== deniedPermission }
  const channel = {
    id: channelId,
    name: "events",
    guildId: targetGuildId,
    type,
    isTextBased: () => true,
    isSendable: () => true,
    permissionsFor: () => permissions
  }
  const guild = {
    id: guildId,
    name: "Alliance server",
    channels: { fetch: async () => channel },
    members: { me: { id: "bot" } }
  }
  return { channel, guild, client: { guilds: { fetch: async () => guild } } }
}

function interactionForChannel(customId, client) {
  return {
    commandName: null,
    customId,
    guildId,
    values: [channelId],
    client,
    user: { id: "3456789012345678" },
    isChatInputCommand: () => false,
    isButton: () => false,
    isModalSubmit: () => false,
    isChannelSelectMenu: () => true,
    async deferUpdate() { this.deferred = true },
    async editReply(payload) { this.edited = payload },
    async showModal(modal) { this.modal = modal }
  }
}

function handlerOptions(repository) {
  return {
    async userCanManageServer() { return true },
    healthProvider() {
      return {
        available: true,
        gameProfile: "wos",
        botInstanceName: "rachie-wos"
      }
    },
    repositoryProvider() { return repository }
  }
}

test("alliance reminder and roundup channels use native text-channel selectors", () => {
  const view = buildChannelConfigurationView({
    event_channel_id: channelId,
    weekly_roundup_channel_id: "3456789012345678",
    weekly_roundup_enabled: true
  })
  const components = view.components.map(row => row.toJSON())
  assert.equal(components[0].components[0].type, ComponentType.ChannelSelect)
  assert.equal(components[0].components[0].custom_id, IDS.reminderChannel)
  assert.deepEqual(components[0].components[0].channel_types, [
    ChannelType.GuildText,
    ChannelType.GuildAnnouncement
  ])
  assert.equal(components[1].components[0].type, ComponentType.ChannelSelect)
  assert.equal(components[1].components[0].custom_id, IDS.roundupChannel)
  assert.equal(components[0].components[0].default_values[0].id, channelId)
})

test("normal setup modals contain names, settings or link codes but no Discord IDs", () => {
  const serialized = JSON.stringify([
    buildAllianceIdentityModal(null).toJSON(),
    buildStateLinkModal().toJSON()
  ])
  assert.doesNotMatch(serialized, /guild id|server id|channel id/i)
  assert.doesNotMatch(serialized, /event database id|alliance database id/i)
})

test("channel validation rejects incompatible, cross-guild and missing permissions", async () => {
  await assert.rejects(
    resolveSendableChannel(
      discordTarget({ type: ChannelType.GuildForum }).client,
      guildId,
      channelId
    ),
    SchedulerValidationError
  )
  await assert.rejects(
    resolveSendableChannel(
      discordTarget({ targetGuildId: "9999999999999999" }).client,
      guildId,
      channelId
    ),
    /text channel in that guild/i
  )
  for (const [permission, label] of [
    [PermissionFlagsBits.ViewChannel, "View Channel"],
    [PermissionFlagsBits.SendMessages, "Send Messages"],
    [PermissionFlagsBits.EmbedLinks, "Embed Links"],
    [PermissionFlagsBits.AttachFiles, "Attach Files"]
  ]) {
    await assert.rejects(
      resolveSendableChannel(
        discordTarget({ deniedPermission: permission }).client,
        guildId,
        channelId
      ),
      new RegExp(label)
    )
  }
  await assert.doesNotReject(resolveSendableChannel(
    discordTarget({ deniedPermission: PermissionFlagsBits.AttachFiles }).client,
    guildId,
    channelId,
    { requireAttachments: false }
  ))
})

test("successful native reminder selection stores the selected channel ID", async () => {
  let saved
  const settings = {
    alliance_name: "YOU",
    event_channel_id: channelId,
    weekly_roundup_enabled: false
  }
  const repository = {
    async setEventChannel(input) { saved = input; return settings },
    async getGuildSettings() { return settings }
  }
  const interaction = interactionForChannel(IDS.reminderChannel, discordTarget().client)
  assert.equal(await handleEventSchedulerInteraction(interaction, handlerOptions(repository)), true)
  assert.deepEqual(saved, { guildId, eventChannelId: channelId })
  assert.equal(interaction.deferred, true)
})

test("main-alliance identity editing uses reconciled alliance rename", async () => {
  let renamed
  const settings = {
    default_alliance_id: "42",
    alliance_name: "YOU",
    alliance_count: 1,
    event_channel_id: channelId,
    weekly_roundup_enabled: false
  }
  const repository = {
    async getGuildSettings() { return settings },
    async renameAlliance(input) { renamed = input },
    async getStateLink() { return null },
    async getStateDestination() { return null }
  }
  const interaction = {
    commandName: null,
    customId: IDS.identityModal,
    guildId,
    client: discordTarget().client,
    fields: { getTextInputValue: () => "  YOU Prime  " },
    isChatInputCommand: () => false,
    isButton: () => false,
    isModalSubmit: () => true,
    isChannelSelectMenu: () => false,
    async deferReply() { this.deferred = true },
    async editReply(payload) { this.edited = payload }
  }
  await handleEventSchedulerInteraction(interaction, handlerOptions(repository))
  assert.deepEqual(renamed, {
    guildId,
    allianceId: "42",
    allianceName: "YOU Prime"
  })
})

test("alliance roundup native selection validates without requiring Attach Files", async () => {
  const settings = { weekly_roundup_day: 1, weekly_roundup_time_utc: "09:00:00" }
  const repository = { async getGuildSettings() { return settings } }
  const target = discordTarget({ deniedPermission: PermissionFlagsBits.AttachFiles })
  const interaction = interactionForChannel(IDS.roundupChannel, target.client)
  await handleEventSchedulerInteraction(interaction, handlerOptions(repository))
  assert.equal(interaction.modal.toJSON().custom_id, `${IDS.roundupModalPrefix}${channelId}`)
})

test("alliance roundup modal stores the natively selected channel ID", async () => {
  let saved
  const repository = {
    async configureWeeklyRoundup(input) { saved = input; return input },
    async getGuildSettings() {
      return {
        alliance_name: "YOU",
        alliance_count: 1,
        event_channel_id: channelId,
        weekly_roundup_enabled: true,
        weekly_roundup_channel_id: channelId,
        weekly_roundup_day: 1,
        weekly_roundup_time_utc: "09:00:00",
        roundup_when_empty: false
      }
    },
    async getStateLink() { return null },
    async getStateDestination() { return null }
  }
  const interaction = {
    commandName: null,
    customId: `${IDS.roundupModalPrefix}${channelId}`,
    guildId,
    client: discordTarget({ deniedPermission: PermissionFlagsBits.AttachFiles }).client,
    fields: {
      getTextInputValue(id) { return { d: "1", t: "09:00", x: "no" }[id] }
    },
    isChatInputCommand: () => false,
    isButton: () => false,
    isModalSubmit: () => true,
    isChannelSelectMenu: () => false,
    async deferReply() { this.deferred = true },
    async editReply(payload) { this.edited = payload }
  }
  await handleEventSchedulerInteraction(interaction, handlerOptions(repository))
  assert.equal(saved.channelId, channelId)
  assert.equal(saved.enabled, true)
})

test("state destination selection uses the interaction guild automatically", async () => {
  let saved
  const destination = {
    state_guild_id: guildId,
    state_roundup_channel_id: channelId,
    enabled: true
  }
  const repository = {
    async upsertStateDestination(input) { saved = input; return destination }
  }
  const target = discordTarget({ deniedPermission: PermissionFlagsBits.AttachFiles })
  const interaction = interactionForChannel(IDS.stateDestinationChannel, target.client)
  await handleEventSchedulerInteraction(interaction, handlerOptions(repository))
  assert.deepEqual(saved, {
    stateGuildId: guildId,
    configuredByBotInstance: "rachie-wos",
    stateRoundupChannelId: channelId
  })
  assert.match(interaction.edited.content, /State destination/)
})

test("existing stored channel IDs remain display and selector defaults", () => {
  const home = buildHomeView({
    alliance_name: "YOU",
    alliance_count: 1,
    event_channel_id: channelId,
    weekly_roundup_enabled: false
  }, {
    game_profile: "wos",
    state_guild_name: "State 123",
    state_channel_name: "#weekly-events",
    sharing_enabled: true
  }, null)
  assert.match(home.content, new RegExp(`<#${channelId}>`))
  assert.match(home.content, /State 123 \/ #weekly-events/)
  assert.doesNotMatch(home.content, /state_guild_id|state_event_channel_id/)

  const destinationView = buildStateDestinationView({
    state_roundup_channel_id: channelId,
    enabled: true
  })
  assert.equal(
    destinationView.components[0].toJSON().components[0].default_values[0].id,
    channelId
  )
})

test("state link codes are random, normalized, hashed and reject malformed input", () => {
  const first = generateStateLinkCode()
  const second = generateStateLinkCode()
  assert.match(first, /^[2-9A-HJ-NP-Z]{4}(?:-[2-9A-HJ-NP-Z]{4}){2}$/)
  assert.notEqual(first, second)
  assert.equal(normalizeStateLinkCode(first.toLowerCase()), first.replaceAll("-", ""))
  assert.match(hashStateLinkCode(first), /^[0-9a-f]{64}$/)
  assert.equal(hashStateLinkCode("not-a-code"), null)
})

test("scheduler responses remain private for native setup entry", async () => {
  let deferred
  const repository = {
    async getGuildSettings() { return null },
    async getStateLink() { return null },
    async getStateDestination() { return null }
  }
  const interaction = {
    commandName: "event-scheduler",
    customId: null,
    guildId,
    client: discordTarget().client,
    isChatInputCommand: () => true,
    isButton: () => false,
    isModalSubmit: () => false,
    isChannelSelectMenu: () => false,
    async deferReply(payload) { deferred = payload; this.deferred = true },
    async editReply() {}
  }
  await handleEventSchedulerInteraction(interaction, handlerOptions(repository))
  assert.equal(deferred.flags, MessageFlags.Ephemeral)
})
