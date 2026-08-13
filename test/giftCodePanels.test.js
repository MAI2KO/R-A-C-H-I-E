const test = require("node:test")
const assert = require("node:assert/strict")
const fs = require("node:fs")
const path = require("node:path")
const { ChannelType, MessageFlags } = require("discord.js")

const { InteractionSessionStore } = require("../src/interactionSessions")
const { createBanterConfigLookup } = require("../src/banterConfig")
const { profileTerminology } = require("../src/giftCodes/terminology")
const { PlayerAccountError } = require("../src/giftCodes/playerService")
const {
  IDS,
  selectedAccount,
  accountMenu,
  playerPanel,
  releaseConfirmationPanel,
  operatorReleaseConfirmation,
  giftPanel,
  activeCodesPanel,
  registrationModal,
  completeCharacterRegistration,
  adminPanel,
  formatCommunityStats,
  giftCodePanelFailureDiagnostics,
  handleGiftCodePanelInteraction
} = require("../src/giftCodes/discord/panelInteractions")
const {
  buildPlayerRegisterCommand,
  buildPlayerAdminCommand,
  buildGiftCodesCommand,
  buildGiftCodeAddCommand,
  buildGiftCodesAdminCommand,
  getGiftCommandData
} = require("../src/giftCodes/discord/commands")
const {
  codeProgressMessage,
  joinMessage,
  staleDiscordReference,
  createGiftCodeCommunityService
} = require("../src/giftCodes/communityService")

const wos = profileTerminology("wos")
const kingshot = profileTerminology("kingshot")
const accounts = [
  {
    id: "a1",
    player_id: "111",
    state_or_kingdom_number: "689",
    is_primary: true,
    is_active: true,
    gift_redemption_enabled: true,
    guild_gift_code_enrolled: true,
    last_redemption_status: "success"
  },
  {
    id: "a2",
    player_id: "222",
    state_or_kingdom_number: "700",
    is_primary: false,
    is_active: true,
    gift_redemption_enabled: false,
    guild_gift_code_enrolled: false,
    last_redemption_status: null
  }
]

test("visible command surface is consolidated while legacy booking register remains", () => {
  assert.equal(buildPlayerRegisterCommand("wos").toJSON().name, "player-register")
  assert.equal(buildPlayerAdminCommand("wos").toJSON().name, "player-admin")
  assert.equal(buildGiftCodesCommand("wos").toJSON().name, "gift-codes")
  assert.equal(buildGiftCodeAddCommand("wos").toJSON().name, "gift-code-add")
  assert.equal(buildGiftCodesAdminCommand("wos").toJSON().name, "gift-codes-admin")
  const registered = [
    buildPlayerRegisterCommand("wos").toJSON(),
    ...getGiftCommandData({ PLAYER_GIFT_CODES_ENABLED: "true", GAME_PROFILE: "wos" })
  ].map(command => command.name)
  assert.deepEqual(registered, [
    "player-register", "player-admin", "gift-code-add", "gift-codes", "gift-codes-admin"
  ])
  assert.ok(!registered.includes("player"))
  assert.ok(!registered.includes("gift"))
  assert.ok(!registered.includes("gift-admin"))
  const indexSource = fs.readFileSync(path.join(__dirname, "..", "index.js"), "utf8")
  assert.match(indexSource, /\.setName\("register"\)/)
  assert.match(indexSource, /register_player_for_server/)
})

test("player panel supports empty registration and multiple-account selection", () => {
  const empty = playerPanel({ sessionId: "s", accounts: [], selected: null, terms: wos })
  assert.match(empty.content, /Register your game account/)
  assert.equal(empty.components[0].components[0].data.label, "Register Character")

  const selected = selectedAccount(accounts, "222")
  const panel = playerPanel({ sessionId: "s", accounts, selected, terms: wos })
  assert.match(panel.content, /State: 700/)
  assert.match(panel.content, /Primary: No/)
  const menu = panel.components[0].components[0].toJSON()
  assert.equal(menu.options.length, 2)
  assert.equal(menu.options[1].default, true)
  assert.deepEqual(panel.components[1].components.map(button => button.data.label), [
    "Add Character", "Change State", "Enable Auto-Redeem", "Remove Character", "Release Character"
  ])
})

test("inactive historical accounts are excluded from player selectors and empty states", () => {
  const inactive = {
    ...accounts[0],
    player_id: "93986200",
    is_active: false,
    is_primary: false,
    gift_redemption_enabled: false
  }
  assert.equal(selectedAccount([inactive], inactive.player_id), null)
  const empty = playerPanel({
    sessionId: "s", accounts: [inactive], selected: selectedAccount([inactive]), terms: wos
  })
  assert.match(empty.content, /Register your game account/)
  assert.equal(empty.components[0].components[0].data.label, "Register Character")
  assert.doesNotMatch(empty.content, /93986200|Active: No/)

  const menu = accountMenu("s", [inactive, accounts[1]], accounts[1])
  assert.equal(menu, null, "one active account should not render a historical selector")
})

test("registration modal and player panels use State or Kingdom correctly", () => {
  const wosModal = registrationModal("s", wos).toJSON()
  const kingshotModal = registrationModal("s", kingshot).toJSON()
  assert.equal(wosModal.components[1].components[0].label, "State")
  assert.equal(kingshotModal.components[1].components[0].label, "Kingdom")
  assert.match(playerPanel({ sessionId: "s", accounts, selected: accounts[0], terms: kingshot }).content, /Kingdom: 689/)
})

test("release views distinguish destructive self release and private operator recovery", () => {
  const release = releaseConfirmationPanel("session", "12345")
  assert.match(release.content, /permanently disconnects/)
  assert.match(release.content, /Use Remove Character instead/)
  assert.deepEqual(release.components[0].components.map(button => button.data.label), [
    "CONFIRM RELEASE", "Cancel"
  ])

  const operator = operatorReleaseConfirmation("session", {
    player_id: "12345",
    discord_user_id: "707866087248756736",
    state_or_kingdom_number: "689",
    is_active: true,
    gift_redemption_enabled: true,
    guild_enrolment_count: 2
  }, wos)
  assert.match(operator.content, /Current Discord owner: 707866087248756736/)
  assert.match(operator.content, /State: 689/)
  assert.deepEqual(operator.allowedMentions, { parse: [], repliedUser: false })
})

test("gift panel focuses on codes and routes character management to the canonical panel", () => {
  const panel = giftPanel({
    sessionId: "s",
    accounts,
    selected: accounts[0],
    terms: wos,
    maximumEnabled: 2
  })
  assert.match(panel.content, /Characters covered: 1 \/ 2/)
  assert.match(panel.content, /Recent result: success/)
  assert.deepEqual(panel.components.at(-1).components.map(button => button.data.label), [
    "Submit Gift Code", "Redemption History", "Active Codes", "Manage Characters"
  ])
  const noAccount = giftPanel({
    sessionId: "s", accounts: [], selected: null, terms: wos, maximumEnabled: 2
  })
  assert.deepEqual(noAccount.components[0].components.map(button => button.data.label), [
    "Submit Gift Code", "Active Codes", "Manage Characters"
  ])
})

test("active-code panel lists only supplied active codes and paginates cleanly", () => {
  const first = activeCodesPanel({
    sessionId: "s",
    visibility: {
      codes: [{ code: "NEWEST" }, { code: "OLDER" }],
      activeCount: 17,
      expiredCount: 27,
      page: 0,
      pageSize: 15
    }
  })
  assert.match(first.content, /NEWEST\nOLDER/)
  assert.match(first.content, /Active codes: 17/)
  assert.match(first.content, /Expired codes recorded: 27/)
  assert.match(first.content, /Page 1 of 2/)
  assert.equal(first.components[0].components[0].data.disabled, true)
  assert.equal(first.components[0].components[2].data.disabled, false)

  const empty = activeCodesPanel({
    sessionId: "s",
    visibility: { codes: [], activeCount: 0, expiredCount: 4, page: 0, pageSize: 15 }
  })
  assert.match(empty.content, /No active gift codes/)
  assert.doesNotMatch(empty.content, /INVALID|REVIEW|RESTRICTED/)
})

test("admin panel reports worker and configuration health without private records", () => {
  const panel = adminPanel({
    sessionId: "s",
    terms: wos,
    runtime: { verificationEnabled: true, redemptionEnabled: false, verifierConfigured: true },
    diagnostics: {
      pending_candidates: 2,
      active_codes: 3,
      expired_codes: 4,
      invalid_codes: 5,
      restricted_review_codes: 6,
      pending_redemptions: 4,
      retry_count: 1
    },
    configuration: {
      settings: { gift_code_channel_id: "123", contributor_role_id: "456" },
      channelAvailable: true,
      roleAvailable: false,
      roleStatus: "error"
    }
  })
  assert.match(panel.content, /Gift-code channel: <#123>/)
  assert.match(panel.content, /Contributor reward role: Unable to create\/assign/)
  assert.match(panel.content, /Expired codes: 4/)
  assert.match(panel.content, /Invalid codes: 5/)
  assert.match(panel.content, /Restricted\/review: 6/)
  assert.ok(!panel.content.includes("Player ID"))
  assert.deepEqual(panel.components.flatMap(row => row.components).map(button => button.data.label), [
    "Configure Channel", "Verify Code", "Inspect Code",
    "Configure Source Channel", "Queue Status", "Community Stats", "Refresh"
  ])
})

test("admin panel exposes bounded profile source health", () => {
  const panel = adminPanel({
    sessionId: "s",
    terms: kingshot,
    runtime: {
      verificationEnabled: false,
      redemptionEnabled: false,
      verifierConfigured: false,
      sourcePollingEnabled: true
    },
    diagnostics: {
      pending_candidates: 0,
      active_codes: 0,
      expired_codes: 0,
      invalid_codes: 0,
      restricted_review_codes: 0,
      pending_redemptions: 0,
      retry_count: 0
    },
    configuration: { settings: null, channelAvailable: false, roleAvailable: false },
    sourceStatus: {
      channels: [{ enabled: true }, { enabled: true }],
      sources: [
        {
          source_type: "discord_mirror",
          last_observation_at_utc: "2026-08-10T00:00:00Z",
          last_candidate_at_utc: "2026-08-09T00:00:00Z"
        },
        {
          source_type: "discord_mirror",
          last_observation_at_utc: "2026-08-11T00:00:00Z",
          last_candidate_at_utc: "2026-08-10T00:00:00Z"
        },
        {
          source_type: "public_catalogue",
          last_poll_at_utc: "2026-08-12T00:00:00Z",
          last_successful_poll_at_utc: "2026-08-11T23:45:00Z",
          observations_count: 7,
          candidates_count: 2,
          last_error: "ETIMEDOUT"
        }
      ]
    }
  })
  assert.match(panel.content, /Official Discord mirror: Enabled \(2\)/)
  assert.match(panel.content, /Last mirror candidate: <t:/)
  assert.match(panel.content, /Public catalogue: Enabled/)
  assert.doesNotMatch(panel.content, /WOSRewards|KingshotRewards/)
  assert.match(panel.content, /Codes observed: 7/)
  assert.match(panel.content, /New candidates: 2/)
  assert.match(panel.content, /Last source error: ETIMEDOUT/)
  assert.doesNotMatch(panel.content, /<html|raw body/i)
})

test("admin panel uses effective runtime state instead of persisted source enablement", () => {
  const panel = adminPanel({
    sessionId: "s",
    terms: wos,
    runtime: { sourcePollingEnabled: false },
    diagnostics: {
      pending_candidates: 0,
      active_codes: 0,
      expired_codes: 0,
      invalid_codes: 0,
      restricted_review_codes: 0,
      pending_redemptions: 0,
      retry_count: 0
    },
    configuration: { settings: null },
    sourceStatus: {
      channels: [],
      sources: [{ source_type: "public_catalogue", enabled: true }]
    }
  })
  assert.match(panel.content, /Public catalogue: Disabled/)
})

test("community statistics use engagement wording without pricing or guarantees", () => {
  const output = formatCommunityStats({
    auto_redeem_players: 8,
    enabled_accounts: 10,
    successful_redemptions: 42,
    verified_codes: 3,
    already_redeemed: 7,
    successful_this_month: 12,
    latest_verified_code: "ABC123",
    unique_contributors: 2
  })
  assert.match(output, /Players using Auto-Redeem: 8/)
  assert.match(output, /Characters covered: 10/)
  assert.match(output, /Successful redemptions: 42/)
  assert.match(output, /Already claimed: 7/)
  assert.match(output, /Verified gift codes: 3/)
  assert.doesNotMatch(output, /\b(?:free|paid|premium|subscription|protected)\b/i)
})

function discordFixture({
  roleEditable = true,
  channelAvailable = true,
  manageRoles = true,
  persistedRole = true,
  fetchError = null,
  editError = null,
  sendError = null
} = {}) {
  const sent = []
  const edits = []
  const assigned = []
  const created = []
  const privateMessages = []
  const messagesByNonce = new Map()
  const fetchedMessages = []
  const roles = new Map()
  const role = {
    id: "456",
    name: "🍭",
    editable: roleEditable,
    managed: false,
    mentionable: false,
    permissions: { bitfield: 0n }
  }
  if (persistedRole) roles.set(role.id, role)
  const member = {
    roles: {
      cache: { has: () => assigned.length > 0 },
      async add(value) { assigned.push(value.id) }
    }
  }
  const message = {
    id: "789",
    async edit(payload) {
      if (editError) throw editError
      edits.push(payload)
      return this
    }
  }
  const channel = {
    id: "123",
    type: ChannelType.GuildText,
    isSendable: () => channelAvailable,
    permissionsFor: () => ({ has: () => channelAvailable }),
    async send(payload) {
      if (sendError) throw sendError
      sent.push(payload)
      return message
    },
    messages: {
      async fetch(messageId) {
        fetchedMessages.push(messageId)
        if (fetchError) throw fetchError
        return message
      }
    }
  }
  const guild = {
    members: {
      me: { permissions: { has: () => manageRoles } },
      async fetch() { return member }
    },
    channels: { async fetch() { return channelAvailable ? channel : null } },
    roles: {
      async fetch(roleId) { return roleId ? roles.get(roleId) || null : roles },
      async create(input) {
        created.push(input)
        const newRole = { ...role, id: `created-${created.length}` }
        roles.set(newRole.id, newRole)
        return newRole
      }
    }
  }
  return {
    client: {
      guilds: { async fetch() { return guild } },
      users: {
        async fetch(userId) {
          return {
            async send(payload) {
              const existing = messagesByNonce.get(payload.nonce)
              if (payload.enforceNonce && existing) return existing
              const sentMessage = { ...payload, userId }
              messagesByNonce.set(payload.nonce, sentMessage)
              privateMessages.push(sentMessage)
              return sentMessage
            }
          }
        }
      }
    },
    sent,
    edits,
    assigned,
    created,
    privateMessages,
    fetchedMessages,
    roles,
    channel,
    role
  }
}

function statusCardRepository({
  accountStats = {},
  communityStats = {},
  status = "pending",
  channelId = null,
  messageId = null,
  managedChannelId = null,
  configuredChannelId = "123"
} = {}) {
  const event = {
    id: "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa",
    event_type: "auto_redeem_join",
    guild_id: "777",
    discord_user_id: "999",
    player_account_id: "account-id",
    status,
    channel_id: channelId,
    message_id: messageId
  }
  const state = {
    event,
    accountStats: {
      enabledCount: 1,
      successfulRedemptions: 0,
      alreadyRedeemed: 0,
      locationNumber: "689",
      ...accountStats
    },
    communityStats: {
      auto_redeem_players: 1,
      enabled_accounts: 1,
      successful_redemptions: 0,
      already_redeemed: 0,
      ...communityStats
    },
    completions: [],
    failures: [],
    clears: []
  }
  return {
    state,
    async claimEvent() {
      if (event.status !== "pending") return null
      event.status = "claimed"
      return { ...event }
    },
    async claimStatusCard() {
      if (!["pending", "completed", "failed"].includes(event.status)) return null
      event.status = "claimed"
      return { ...event }
    },
    async getEventPayload() {
      return {
        ...event,
        gift_code_channel_id: configuredChannelId,
        managed_gift_channel_id: managedChannelId,
        state_or_kingdom_number: "688"
      }
    },
    async accountOwnerStats() { return { ...state.accountStats } },
    async communityStats() { return { ...state.communityStats } },
    async completeEvent(_id, _worker, values) {
      event.status = "completed"
      event.channel_id = values.channelId || event.channel_id
      event.message_id = values.messageId || event.message_id
      event.metadata = { ...(event.metadata || {}), ...(values.metadata || {}) }
      state.completions.push(values)
      return { ...event }
    },
    async clearStatusCardReference(_id, _worker, errorCode) {
      event.status = "completed"
      event.channel_id = null
      event.message_id = null
      state.clears.push(errorCode)
      return { ...event }
    },
    async failEvent(_id, _worker, errorCode) {
      if (event.status === "claimed") event.status = "failed"
      state.failures.push(errorCode)
    },
    async statusCardTargetsForAccount(playerAccountId) {
      assert.equal(playerAccountId, "account-id")
      return [{ guild_id: event.guild_id, discord_user_id: event.discord_user_id }]
    },
    async claimProgressRefresh() { return null }
  }
}

function statusCardMigrationDiscord({
  managedSendError = null,
  oldFetchError = null,
  oldDeleteError = null,
  oldChannelError = null,
  allowLegacySend = false
} = {}) {
  const sent = []
  const edits = []
  const deletes = []
  const nonces = new Map()
  let managedMessage = null
  const oldMessage = {
    id: "legacy-message",
    async edit(payload) { edits.push(payload); return this },
    async delete() {
      if (oldDeleteError) throw oldDeleteError
      deletes.push(this.id)
    }
  }
  const legacy = {
    id: "legacy-channel",
    type: ChannelType.GuildText,
    isSendable: () => true,
    permissionsFor: () => ({ has: () => true }),
    messages: {
      async fetch() {
        if (oldFetchError) throw oldFetchError
        return oldMessage
      }
    },
    async send(payload) {
      if (!allowLegacySend) throw new Error("legacy channel should not receive a migration post")
      const message = {
        id: "legacy-new-message",
        async edit(value) { edits.push(value); return this },
        async delete() { deletes.push(this.id) }
      }
      sent.push({ ...payload, destination: "legacy" })
      return message
    }
  }
  const managed = {
    id: "managed-channel",
    type: ChannelType.GuildText,
    isSendable: () => true,
    permissionsFor: () => ({ has: () => true }),
    messages: { async fetch() { return managedMessage } },
    async send(payload) {
      if (managedSendError) throw managedSendError
      if (payload.enforceNonce && nonces.has(payload.nonce)) return nonces.get(payload.nonce)
      const message = {
        id: "managed-message",
        async edit(value) { edits.push(value); return this },
        async delete() { deletes.push(this.id) }
      }
      managedMessage = message
      nonces.set(payload.nonce, message)
      sent.push({ ...payload, destination: "managed" })
      return message
    }
  }
  const guild = {
    members: { me: {} },
    channels: {
      async fetch(channelId) {
        if (channelId === managed.id) return managed
        if (channelId === legacy.id) {
          if (oldChannelError) throw oldChannelError
          return legacy
        }
        return null
      }
    }
  }
  return {
    client: { guilds: { async fetch() { return guild } } },
    sent,
    edits,
    deletes,
    oldMessage,
    managed
  }
}

test("community messages expose aggregates and location but never full Player ID", () => {
  const payload = {
    code: "ABC123",
    discord_user_id: "999",
    state_or_kingdom_number: "689",
    metadata: { queuedCount: 63 }
  }
  const progress = {
    successful: 57,
    already_redeemed: 4,
    account_issues: 2,
    restricted: 0,
    remaining: 3,
    total: 66
  }
  const codeMessage = codeProgressMessage(payload, progress, wos, { initial: true })
  assert.match(codeMessage, /Accounts queued: 63/)
  assert.match(codeMessage, /Successful redemptions: 57/)
  assert.ok(!codeMessage.includes("282021376"))
  const joined = joinMessage(payload, {
    registered_users: 47,
    auto_redeem_players: 40,
    enabled_accounts: 63,
    successful_redemptions: 418,
    already_redeemed: 27
  }, { enabledCount: 2, successfulRedemptions: 14, alreadyRedeemed: 3 }, wos, 2)
  assert.match(joined, /State: 689/)
  assert.match(joined, /Gift Code Auto-Redeem Activated/)
  assert.match(joined, /Characters covered: 2 \/ 2/)
  assert.match(joined, /Players using Auto-Redeem: 40/)
  assert.match(joined, /Successful redemptions: 418/)
  assert.match(joined, /Already claimed: 27/)
  assert.match(joined, /Their successful redemptions: 14/)
  assert.match(joined, /Their already claimed: 3/)
  assert.ok(!joined.includes("Player ID"))
  assert.doesNotMatch(`${codeMessage}\n${joined}`, /\b(?:free|paid|premium|subscription)\b/i)
})

test("activation stores one status card and later changes edit the same message", async () => {
  const discord = discordFixture()
  const repository = statusCardRepository()
  const service = createGiftCodeCommunityService({
    repository,
    client: discord.client,
    gameProfile: "wos",
    maximumEnabledAccounts: 2,
    logger: { warn() {} }
  })

  await service.onAutoRedemptionEnabled(repository.state.event)
  assert.equal(discord.sent.length, 1)
  assert.equal(repository.state.event.channel_id, "123")
  assert.equal(repository.state.event.message_id, "789")

  repository.state.accountStats.enabledCount = 2
  repository.state.accountStats.successfulRedemptions = 4
  repository.state.accountStats.alreadyRedeemed = 3
  repository.state.communityStats.enabled_accounts = 7
  repository.state.communityStats.successful_redemptions = 12
  repository.state.communityStats.already_redeemed = 8
  assert.equal(await service.onAutoRedemptionEnabled(null, {
    guildId: "777",
    discordUserId: "999"
  }), true)
  assert.equal(discord.sent.length, 1, "refresh posted a duplicate activation card")
  assert.equal(discord.edits.length, 1)
  assert.match(discord.edits[0].content, /Characters covered: 2 \/ 2/)
  assert.match(discord.edits[0].content, /Their successful redemptions: 4/)
  assert.match(discord.edits[0].content, /Their already claimed: 3/)
  assert.match(discord.edits[0].content, /Characters covered: 7/)
  assert.match(discord.edits[0].content, /Successful redemptions: 12/)

  repository.state.accountStats.enabledCount = 0
  repository.state.communityStats.enabled_accounts = 5
  assert.equal(await service.refreshStatusCard("777", "999"), true)
  assert.match(discord.edits[1].content, /Characters covered: 0 \/ 2/)
  assert.match(discord.edits[1].content, /Characters covered: 5/)
})

test("persisted status card survives service restart and keeps profile wording isolated", async () => {
  for (const [gameProfile, expected, unexpected] of [
    ["wos", "State: 689", "Kingdom:"],
    ["kingshot", "Kingdom: 689", "State:"]
  ]) {
    const discord = discordFixture()
    const repository = statusCardRepository({ status: "completed", channelId: "123", messageId: "789" })
    const restarted = createGiftCodeCommunityService({
      repository,
      client: discord.client,
      gameProfile,
      logger: { warn() {} }
    })
    assert.equal(await restarted.refreshStatusCard("777", "999"), true)
    assert.deepEqual(discord.fetchedMessages, ["789"])
    assert.equal(discord.sent.length, 0)
    assert.match(discord.edits[0].content, new RegExp(expected))
    assert.doesNotMatch(discord.edits[0].content, new RegExp(unexpected))
  }
})

test("legacy status cards migrate once to the managed announcement channel on refresh", async () => {
  const discord = statusCardMigrationDiscord()
  const repository = statusCardRepository({
    status: "completed",
    channelId: "legacy-channel",
    messageId: "legacy-message",
    managedChannelId: "managed-channel",
    configuredChannelId: "legacy-channel",
    accountStats: { successfulRedemptions: 4, alreadyRedeemed: 2 },
    communityStats: { successful_redemptions: 12, already_redeemed: 7 }
  })
  const service = createGiftCodeCommunityService({
    repository,
    client: discord.client,
    gameProfile: "wos",
    logger: { warn() {} }
  })

  const results = await Promise.all([
    service.refreshStatusCard("777", "999"),
    service.refreshStatusCard("777", "999")
  ])
  assert.deepEqual(results.sort(), [false, true])
  assert.equal(discord.sent.length, 1)
  assert.match(discord.sent[0].content, /Their successful redemptions: 4/)
  assert.match(discord.sent[0].content, /Already claimed: 7/)
  assert.equal(repository.state.event.channel_id, "managed-channel")
  assert.equal(repository.state.event.message_id, "managed-message")
  assert.equal(repository.state.event.metadata.statusCardGeneration, 1)
  assert.deepEqual(discord.deletes, ["legacy-message"])

  assert.equal(await service.refreshStatusCard("777", "999"), true)
  assert.equal(discord.sent.length, 1, "future refresh should edit the canonical managed card")
  assert.equal(discord.edits.length, 1)
})

test("legacy cleanup failures and stale old references never undo managed migration", async () => {
  for (const fixtureOptions of [
    { oldDeleteError: Object.assign(new Error("forbidden"), { code: 50013 }) },
    { oldFetchError: Object.assign(new Error("unknown message"), { code: 10008 }) },
    { oldChannelError: Object.assign(new Error("unknown channel"), { code: 10003 }) }
  ]) {
    const discord = statusCardMigrationDiscord(fixtureOptions)
    const repository = statusCardRepository({
      status: "completed",
      channelId: "legacy-channel",
      messageId: "legacy-message",
      managedChannelId: "managed-channel",
      configuredChannelId: "legacy-channel"
    })
    const warnings = []
    const service = createGiftCodeCommunityService({
      repository,
      client: discord.client,
      gameProfile: "wos",
      logger: { warn(message) { warnings.push(message) } }
    })
    assert.equal(await service.refreshStatusCard("777", "999"), true)
    assert.equal(repository.state.event.channel_id, "managed-channel")
    assert.equal(repository.state.event.message_id, "managed-message")
    assert.equal(repository.state.failures.length, 0)
    assert.deepEqual(warnings, [], "expected stale cleanup should remain quiet")
  }
})

test("managed destination failure preserves and refreshes the legacy canonical card", async () => {
  const discord = statusCardMigrationDiscord({ managedSendError: new Error("temporary send failure") })
  const repository = statusCardRepository({
    status: "completed",
    channelId: "legacy-channel",
    messageId: "legacy-message",
    managedChannelId: "managed-channel",
    configuredChannelId: "legacy-channel"
  })
  const service = createGiftCodeCommunityService({
    repository,
    client: discord.client,
    gameProfile: "kingshot",
    logger: { warn() {} }
  })

  assert.equal(await service.refreshStatusCard("777", "999"), true)
  assert.equal(discord.sent.length, 0)
  assert.equal(discord.edits.length, 1)
  assert.equal(repository.state.event.channel_id, "legacy-channel")
  assert.equal(repository.state.event.message_id, "legacy-message")
  assert.match(discord.edits[0].content, /Kingdom: 689/)
})

test("new status cards use the managed destination and no managed setup keeps legacy behavior", async () => {
  const managedDiscord = statusCardMigrationDiscord()
  const managedRepository = statusCardRepository({
    managedChannelId: "managed-channel",
    configuredChannelId: "legacy-channel"
  })
  const managedService = createGiftCodeCommunityService({
    repository: managedRepository,
    client: managedDiscord.client,
    gameProfile: "wos",
    logger: { warn() {} }
  })
  await managedService.onAutoRedemptionEnabled(managedRepository.state.event)
  assert.equal(managedRepository.state.event.channel_id, "managed-channel")
  assert.equal(managedDiscord.sent.length, 1)

  const legacyDiscord = discordFixture()
  const legacyRepository = statusCardRepository()
  const legacyService = createGiftCodeCommunityService({
    repository: legacyRepository,
    client: legacyDiscord.client,
    gameProfile: "wos",
    logger: { warn() {} }
  })
  await legacyService.onAutoRedemptionEnabled(legacyRepository.state.event)
  assert.equal(legacyRepository.state.event.channel_id, "123")
  assert.equal(legacyDiscord.sent.length, 1)

  const fallbackDiscord = statusCardMigrationDiscord({
    managedSendError: new Error("managed unavailable"),
    allowLegacySend: true
  })
  const fallbackRepository = statusCardRepository({
    managedChannelId: "managed-channel",
    configuredChannelId: "legacy-channel"
  })
  const fallbackService = createGiftCodeCommunityService({
    repository: fallbackRepository,
    client: fallbackDiscord.client,
    gameProfile: "wos",
    logger: { warn() {} }
  })
  await fallbackService.onAutoRedemptionEnabled(fallbackRepository.state.event)
  assert.equal(fallbackRepository.state.event.channel_id, "legacy-channel")
  assert.equal(fallbackDiscord.sent[0].destination, "legacy")
})

test("successful and already-claimed redemptions refresh personal status counters", async () => {
  const discord = discordFixture()
  const repository = statusCardRepository({ status: "completed", channelId: "123", messageId: "789" })
  const service = createGiftCodeCommunityService({
    repository,
    client: discord.client,
    gameProfile: "wos",
    logger: { warn() {} }
  })

  repository.state.accountStats.successfulRedemptions = 1
  assert.equal(await service.onRedemptionUpdated("code-a", "account-id", "success"), false)
  repository.state.accountStats.alreadyRedeemed = 1
  assert.equal(await service.onRedemptionUpdated("code-b", "account-id", "already_redeemed"), false)
  assert.equal(discord.edits.length, 2)
  assert.match(discord.edits[0].content, /Their successful redemptions: 1/)
  assert.match(discord.edits[1].content, /Their already claimed: 1/)
})

test("deleted status card is replaced once while inaccessible references are cleared", async () => {
  const unknownMessage = Object.assign(new Error("Unknown Message"), { code: 10008 })
  const deleted = discordFixture({ fetchError: unknownMessage })
  const deletedRepository = statusCardRepository({
    status: "completed", channelId: "123", messageId: "old-message"
  })
  const deletedService = createGiftCodeCommunityService({
    repository: deletedRepository,
    client: deleted.client,
    gameProfile: "wos",
    logger: { warn() {} }
  })
  assert.equal(await deletedService.refreshStatusCard("777", "999"), true)
  assert.equal(deleted.sent.length, 1)
  assert.equal(deletedRepository.state.event.message_id, "789")

  const unavailable = discordFixture({ channelAvailable: false })
  const unavailableRepository = statusCardRepository({
    status: "completed", channelId: "123", messageId: "old-message"
  })
  const unavailableService = createGiftCodeCommunityService({
    repository: unavailableRepository,
    client: unavailable.client,
    gameProfile: "wos",
    logger: { warn() {} }
  })
  assert.equal(await unavailableService.refreshStatusCard("777", "999"), false)
  assert.equal(unavailableRepository.state.event.message_id, null)
  assert.equal(unavailableRepository.state.clears.length, 1)
  assert.equal(staleDiscordReference(unknownMessage), true)
})

test("status card edit failure is contained without posting a duplicate", async () => {
  const discord = discordFixture({ editError: new Error("temporary Discord failure") })
  const repository = statusCardRepository({ status: "completed", channelId: "123", messageId: "789" })
  const warnings = []
  const service = createGiftCodeCommunityService({
    repository,
    client: discord.client,
    gameProfile: "wos",
    logger: { warn(message) { warnings.push(message) } }
  })
  assert.equal(await service.onRedemptionUpdated(
    "code-id", "account-id", "success"
  ), false)
  assert.equal(discord.sent.length, 0)
  assert.equal(repository.state.failures.length, 1)
  assert.equal(warnings.length, 1)
})

test("new-code announcement and contributor confirmation are idempotent while duplicates stay silent", async () => {
  const discord = discordFixture()
  const completed = []
  const events = [
    {
      id: "11111111-1111-4111-8111-111111111111",
      event_type: "code_progress"
    },
    {
      id: "22222222-2222-4222-8222-222222222222",
      event_type: "contributor_role"
    }
  ]
  let prepared = false
  const repository = {
    async prepareCodeEngagement() {
      if (prepared) return []
      prepared = true
      return events
    },
    async claimEvent(eventId) { return events.find(event => event.id === eventId) },
    async getEventPayload(eventId) {
      const event = events.find(value => value.id === eventId)
      return {
        ...event,
        guild_id: "777",
        gift_code_id: "code-id",
        gift_code_channel_id: "123",
        contributor_role_id: "456",
        discord_user_id: "999",
        code: "ABC123",
        metadata: { queuedCount: 2 }
      }
    },
    async codeProgress() {
      return { successful: 0, already_redeemed: 0, account_issues: 0, restricted: 0, remaining: 2, total: 2 }
    },
    async getSettings() { return { contributor_role_id: "456" } },
    async completeEvent(id) { completed.push(id); return {} },
    async failEvent() { throw new Error("must not fail") }
  }
  const service = createGiftCodeCommunityService({
    repository,
    client: discord.client,
    gameProfile: "wos",
    logger: { warn() {} }
  })
  assert.equal(await service.onCodeActivated("code-id", 2), 2)
  assert.equal(await service.onCodeActivated("code-id", 2), 0)
  assert.equal(discord.sent.length, 1)
  assert.equal(discord.assigned.length, 1)
  assert.equal(completed.length, 2)
  assert.equal(discord.privateMessages.length, 1)
  assert.equal(discord.privateMessages[0].content, "Nice find. Have a 🍭")

  const duplicateOnly = discordFixture()
  const duplicateService = createGiftCodeCommunityService({
    repository: { async prepareCodeEngagement() { return [] } },
    client: duplicateOnly.client,
    gameProfile: "wos",
    logger: { warn() {} }
  })
  assert.equal(await duplicateService.onCodeActivated("duplicate-code", 0), 0)
  assert.equal(duplicateOnly.privateMessages.length, 0)
})

test("unmanageable reward role and missing channel are contained", async () => {
  for (const fixture of [
    discordFixture({ roleEditable: false }),
    discordFixture({ channelAvailable: false })
  ]) {
    let failed = 0
    const type = fixture.role.editable === false ? "contributor_role" : "code_progress"
    const event = { id: "33333333-3333-4333-8333-333333333333", event_type: type }
    const repository = {
      async claimEvent() { return event },
      async getEventPayload() {
        return {
          ...event,
          guild_id: "777",
          gift_code_id: "code-id",
          gift_code_channel_id: "123",
          contributor_role_id: "456",
          discord_user_id: "999",
          code: "ABC",
          metadata: { queuedCount: 1 }
        }
      },
      async codeProgress() {
        return { successful: 0, already_redeemed: 0, account_issues: 0, restricted: 0, remaining: 1, total: 1 }
      },
      async getSettings() { return { contributor_role_id: "456" } },
      async claimContributorRoleProvision() { return {} },
      async completeContributorRoleProvision() { return {} },
      async failContributorRoleProvision() {},
      async markContributorRoleUnavailable() {},
      async failEvent() { failed += 1 },
      async completeEvent() { throw new Error("must not complete") }
    }
    const service = createGiftCodeCommunityService({
      repository,
      client: fixture.client,
      gameProfile: "wos",
      logger: { warn() {} }
    })
    assert.equal(await service.onAutoRedemptionEnabled(event), null)
    assert.equal(failed, 1)
  }
})

test("gift-code channel validation reports an actionable configuration error", async () => {
  const discord = discordFixture({ channelAvailable: false })
  const service = createGiftCodeCommunityService({
    repository: { async setChannel() { throw new Error("must not persist") } },
    client: discord.client,
    gameProfile: "wos",
    logger: { warn() {} }
  })
  await assert.rejects(
    service.configureChannel("777", "123"),
    error => error.code === "GIFT_CODE_CHANNEL_UNAVAILABLE"
      && error.giftCodeHandler === "configure_channel_resolve"
      && /visible text channel/.test(error.message)
  )
})

test("progress edits one stored message and finalizes aggregate state", async () => {
  const discord = discordFixture()
  let completion
  let resultStatus
  const event = {
    id: "44444444-4444-4444-8444-444444444444",
    gift_code_id: "code-id",
    channel_id: "123",
    message_id: "789"
  }
  const repository = {
    async claimProgressRefresh(_giftCodeId, playerAccountId, status) {
      assert.equal(playerAccountId, "account-id")
      resultStatus = status
      return {
        event,
        progress: { successful: 2, already_redeemed: 1, account_issues: 0, restricted: 0, remaining: 0, total: 3 },
        completed: 3
      }
    },
    async getEventPayload() {
      return {
        ...event,
        guild_id: "777",
        gift_code_channel_id: "123",
        discord_user_id: "999",
        code: "ABC",
        metadata: { queuedCount: 3 }
      }
    },
    async completeEvent(_id, _worker, value) { completion = value },
    async failEvent() { throw new Error("must not fail") }
  }
  const service = createGiftCodeCommunityService({
    repository,
    client: discord.client,
    gameProfile: "kingshot",
    logger: { warn() {} }
  })
  assert.equal(await service.onRedemptionUpdated("code-id", "account-id", "success"), true)
  assert.equal(discord.sent.length, 0)
  assert.equal(discord.edits.length, 1)
  assert.deepEqual(discord.fetchedMessages, ["789"])
  assert.equal(resultStatus, "success")
  assert.match(discord.edits[0].content, /Remaining: 0/)
  assert.equal(completion.finalized, true)
})

test("already-redeemed progress reuses the message while coalesced results do not edit", async () => {
  const discord = discordFixture()
  let calls = 0
  const event = {
    id: "55555555-5555-4555-8555-555555555555",
    gift_code_id: "code-id",
    channel_id: "123",
    message_id: "789"
  }
  const repository = {
    async claimProgressRefresh(_giftCodeId, playerAccountId, resultStatus) {
      calls += 1
      assert.equal(playerAccountId, "account-id")
      assert.equal(resultStatus, "already_redeemed")
      if (calls === 1) return null
      return {
        event,
        progress: { successful: 1, already_redeemed: 2, account_issues: 0, restricted: 0, remaining: 1, total: 4 },
        completed: 3
      }
    },
    async getEventPayload() {
      return {
        ...event,
        guild_id: "777",
        gift_code_channel_id: "123",
        discord_user_id: "999",
        code: "ABC",
        metadata: { queuedCount: 4 }
      }
    },
    async completeEvent() {},
    async failEvent() { throw new Error("must not fail") }
  }
  const service = createGiftCodeCommunityService({
    repository,
    client: discord.client,
    gameProfile: "wos",
    logger: { warn() {} }
  })
  assert.equal(await service.onRedemptionUpdated(
    "code-id", "account-id", "already_redeemed"
  ), false)
  assert.equal(discord.edits.length, 0, "coalesced result edited the public message")
  assert.equal(await service.onRedemptionUpdated(
    "code-id", "account-id", "already_redeemed"
  ), true)
  assert.deepEqual(discord.fetchedMessages, ["789"])
  assert.equal(discord.edits.length, 1)
  assert.match(discord.edits[0].content, /Already claimed: 2/)
  assert.match(discord.edits[0].content, /Remaining: 1/)
})

test("native selector IDs remain distinct and scoped to the consolidated panel", () => {
  assert.ok(IDS.adminChannelSelect.startsWith("gcux:"))
  assert.equal(IDS.adminRoleSelect, undefined)
  assert.equal(IDS.adminRole, undefined)
})

function panelInteraction({
  commandName = null,
  customId = null,
  modal = false,
  fields = {},
  values = [],
  options = {}
} = {}) {
  return {
    commandName,
    customId,
    user: { id: "999999999999999999" },
    guildId: "777777777777777777",
    client: {},
    deferred: false,
    replied: false,
    isChatInputCommand: () => Boolean(commandName),
    isModalSubmit: () => modal,
    options: { getString: name => options[name] },
    fields: { getTextInputValue: name => fields[name] },
    values,
    async deferReply(options) {
      assert.equal(options.flags, MessageFlags.Ephemeral)
      this.deferred = true
    },
    async editReply(payload) { this.edited = payload },
    async deferUpdate() { this.deferred = true },
    async showModal(value) { this.modal = value.toJSON() }
  }
}

function panelDependencies({
  accounts,
  submissions = [],
  authorization = async () => true,
  activePages = [],
  adminDiagnostics = null,
  community = null,
  submissionResult = null,
  sourceChannels = [],
  logger = { error() {} }
} = {}) {
  const terms = profileTerminology("wos")
  return {
    sessions: new InteractionSessionStore({ maximumSessions: 20 }),
    userCanManageServer: authorization,
    healthProvider: () => ({ available: true, gameProfile: "wos" }),
    poolProvider: () => ({}),
    playerRepositoryFactory: () => ({
      async getOwnedAccount(discordUserId, playerId) {
        return accounts.find(account => account.player_id === playerId
          && account.is_active
          && (!account.discord_user_id || account.discord_user_id === discordUserId)) || null
      }
    }),
    playerServiceFactory: () => ({
      terms,
      async register({ playerId, locationNumber }) {
        const account = {
          id: `account-${playerId}`,
          player_id: playerId,
          state_or_kingdom_number: locationNumber,
          is_primary: accounts.length === 0,
          is_active: true,
          gift_redemption_enabled: false,
          guild_gift_code_enrolled: false,
          registration_status: "new"
        }
        accounts.push(account)
        return account
      },
      async changeLocation({ playerId, locationNumber }) {
        accounts.find(account => account.player_id === playerId).state_or_kingdom_number = locationNumber
      },
      async remove({ playerId }) {
        const account = accounts.find(value => value.player_id === playerId && value.is_active)
        if (!account) throw new Error("PLAYER_NOT_FOUND")
        account.is_active = false
        account.is_primary = false
        account.gift_redemption_enabled = false
        const replacement = accounts.find(value => value.is_active) || null
        if (replacement) replacement.is_primary = true
        return { account, replacement }
      },
      async release({ discordUserId, playerId }) {
        const account = accounts.find(value => value.player_id === playerId
          && value.is_active
          && (!value.discord_user_id || value.discord_user_id === discordUserId))
        if (!account) throw new PlayerAccountError("PLAYER_OWNERSHIP_CHANGED", "You are no longer the current owner of that Player ID.")
        account.is_active = false
        account.is_primary = false
        account.gift_redemption_enabled = false
        account.discord_user_id = null
        return {
          account,
          previousOwnerDiscordUserId: discordUserId,
          guildIds: ["777777777777777777"]
        }
      },
      async operatorLookup({ playerId }) {
        return accounts.find(value => value.player_id === playerId) || null
      },
      async operatorRelease({ playerId, expectedAccountId, expectedOwnerDiscordUserId }) {
        const account = accounts.find(value => value.player_id === playerId
          && String(value.id) === expectedAccountId
          && value.discord_user_id === expectedOwnerDiscordUserId)
        if (!account) throw Object.assign(new Error("ownership changed"), { code: "PLAYER_OWNERSHIP_CHANGED" })
        const previousOwnerDiscordUserId = account.discord_user_id
        account.discord_user_id = null
        account.is_active = false
        account.gift_redemption_enabled = false
        return { account, previousOwnerDiscordUserId, guildIds: [] }
      }
    }),
    giftRepositoryFactory: () => ({}),
    sourceRepositoryFactory: () => ({
      async sourceStatus() { return { sources: [], channels: [] } },
      async configureDiscordChannel(input) { sourceChannels.push(input) }
    }),
    sourceIngestionFactory: () => ({ async ingest() {} }),
    giftServiceFactory: () => ({
      terms,
      async status({ playerId } = {}) {
        const activeAccounts = accounts.filter(account => account.is_active)
        return playerId
          ? activeAccounts.filter(account => account.player_id === playerId)
          : activeAccounts
      },
      async submit(input) {
        submissions.push(input)
        return submissionResult
          ? submissionResult(input)
          : { giftCode: { code: input.code.trim() }, outcome: "new candidate" }
      },
      async activeCodes({ page }) {
        return activePages[page] || {
          codes: [], activeCount: 0, expiredCount: 0, page, pageSize: 15
        }
      },
      async setAutomaticRedemption({ playerId, enabled }) {
        const account = accounts.find(value => value.player_id === playerId)
        account.gift_redemption_enabled = enabled
        account.guild_gift_code_enrolled = enabled
        return { ...account, engagement_event: null }
      },
      async adminStatus() {
        if (!adminDiagnostics) throw new Error("authorization must run before diagnostics")
        return adminDiagnostics
      }
    }),
    communityRepositoryFactory: () => ({}),
    communityServiceFactory: () => community || ({
      async configuration() { return { settings: null, channelAvailable: false, roleAvailable: false } }
    }),
    runtimeProvider: () => ({
      status: () => ({ verificationEnabled: true, verifierConfigured: true })
    }),
    env: { GIFT_CODE_MAX_AUTO_REDEEM_ACCOUNTS_PER_USER: "2" },
    logger
  }
}

test("new player panel performs registration and State updates through private modals", async () => {
  const accounts = []
  const refreshes = []
  const dependencies = panelDependencies({
    accounts,
    community: {
      async refreshStatusCard(guildId, userId) {
        refreshes.push([guildId, userId])
        return false
      }
    }
  })
  const command = panelInteraction({ commandName: "player-register" })
  await handleGiftCodePanelInteraction(command, dependencies)
  assert.match(command.edited.content, /Register your game account/)

  const registerButton = panelInteraction({
    customId: command.edited.components[0].components[0].data.custom_id
  })
  await handleGiftCodePanelInteraction(registerButton, dependencies)
  assert.equal(registerButton.modal.components[1].components[0].label, "State")

  const registration = panelInteraction({
    customId: registerButton.modal.custom_id,
    modal: true,
    fields: { player_id: "12345", location: "689" }
  })
  await handleGiftCodePanelInteraction(registration, dependencies)
  assert.match(registration.edited.content, /State: 689/)

  const locationButton = panelInteraction({
    customId: registration.edited.components.at(-1).components[1].data.custom_id
  })
  await handleGiftCodePanelInteraction(locationButton, dependencies)
  const locationUpdate = panelInteraction({
    customId: locationButton.modal.custom_id,
    modal: true,
    fields: { location: "700" }
  })
  await handleGiftCodePanelInteraction(locationUpdate, dependencies)
  assert.match(locationUpdate.edited.content, /State: 700/)
  assert.deepEqual(refreshes, [
    ["777777777777777777", "999999999999999999"],
    ["777777777777777777", "999999999999999999"]
  ])
})

test("persistent Register Character button launches the canonical player registration flow", async () => {
  const localAccounts = []
  const dependencies = panelDependencies({ accounts: localAccounts })
  const button = panelInteraction({ customId: IDS.publicRegister })
  await handleGiftCodePanelInteraction(button, dependencies)
  assert.match(button.modal.title, /Register/)
  assert.equal(button.modal.components[1].components[0].label, "State")

  const registration = panelInteraction({
    customId: button.modal.custom_id,
    modal: true,
    fields: { player_id: "12345", location: "689" }
  })
  await handleGiftCodePanelInteraction(registration, dependencies)
  assert.match(registration.edited.content, /Character registered\. Auto-Redeem Enabled\./)
  assert.match(registration.edited.content, /Automatic gift-code redemption: Enabled/)
  assert.equal(localAccounts[0].gift_redemption_enabled, true)
})

test("registration default, historical opt-out and account limit remain distinct", async () => {
  const calls = []
  const community = { async onAutoRedemptionEnabled() {} }
  const gifts = {
    async setAutomaticRedemption(input) {
      calls.push(input)
      return { player_id: input.playerId, engagement_event: null }
    }
  }
  const fresh = await completeCharacterRegistration({
    account: { player_id: "1", registration_status: "new" },
    gifts, community, guildId: "777", discordUserId: "999"
  })
  assert.equal(fresh.autoRedeemEnabled, true)
  assert.equal(calls[0].preferenceSource, "registration_default")

  const claimed = await completeCharacterRegistration({
    account: { player_id: "4", registration_status: "claimed" },
    gifts, community, guildId: "777", discordUserId: "999"
  })
  assert.equal(claimed.autoRedeemEnabled, true)
  assert.equal(calls[1].preferenceSource, "registration_default")

  const optedOut = await completeCharacterRegistration({
    account: {
      player_id: "2",
      registration_status: "reactivated",
      account_metadata: { autoRedeemPreference: { enabled: false, explicit: true } }
    },
    gifts, community, guildId: "777", discordUserId: "999"
  })
  assert.equal(optedOut.autoRedeemEnabled, false)
  assert.equal(calls.length, 2)

  const limit = await completeCharacterRegistration({
    account: { player_id: "3", registration_status: "new" },
    gifts: {
      async setAutomaticRedemption() {
        throw Object.assign(new Error("limit"), { code: "AUTO_REDEEM_ACCOUNT_LIMIT" })
      }
    },
    community,
    guildId: "777",
    discordUserId: "999"
  })
  assert.equal(limit.limitReached, true)
  assert.match(limit.notice, /covered-character limit/)
})

test("ordinary users can submit candidates without invoking admin authorization", async () => {
  const accounts = []
  const submissions = []
  let authorizationCalls = 0
  const dependencies = panelDependencies({
    accounts,
    submissions,
    authorization: async () => { authorizationCalls += 1; return false }
  })
  const command = panelInteraction({ commandName: "gift-codes" })
  await handleGiftCodePanelInteraction(command, dependencies)
  const submitButton = panelInteraction({
    customId: command.edited.components[0].components[0].data.custom_id
  })
  await handleGiftCodePanelInteraction(submitButton, dependencies)
  const submission = panelInteraction({
    customId: submitButton.modal.custom_id,
    modal: true,
    fields: { code: "  MixedCaseCode  " }
  })
  await handleGiftCodePanelInteraction(submission, dependencies)
  assert.equal(authorizationCalls, 0)
  assert.equal(submissions[0].code, "  MixedCaseCode  ")
  assert.equal(submissions[0].isAdmin, undefined)
  assert.match(submission.edited.content, /new candidate/)
})

test("removing the only active account immediately renders the empty state without panel failure", async () => {
  const logs = []
  const sole = { ...accounts[0], player_id: "93986200" }
  const dependencies = panelDependencies({
    accounts: [sole],
    logger: { error(value) { logs.push(value) } }
  })
  const command = panelInteraction({ commandName: "player-register" })
  await handleGiftCodePanelInteraction(command, dependencies)
  const remove = panelInteraction({
    customId: command.edited.components.at(-1).components[3].data.custom_id
  })
  await handleGiftCodePanelInteraction(remove, dependencies)
  assert.equal(sole.is_active, false)
  assert.match(remove.edited.content, /Register your game account/)
  assert.equal(remove.edited.components[0].components[0].data.label, "Register Character")
  assert.doesNotMatch(remove.edited.content, /93986200|Active: No/)
  assert.deepEqual(logs, [])
})

test("self release requires confirmation, rechecks ownership and refreshes only the previous owner", async () => {
  const ownerId = "999999999999999999"
  const account = {
    ...accounts[0],
    id: "self-release-account",
    player_id: "93986201",
    discord_user_id: ownerId
  }
  const refreshes = []
  const dependencies = panelDependencies({
    accounts: [account],
    community: {
      async refreshStatusCard(guildId, userId) {
        refreshes.push([guildId, userId])
        return true
      }
    }
  })
  const command = panelInteraction({ commandName: "player-register" })
  await handleGiftCodePanelInteraction(command, dependencies)

  const release = panelInteraction({
    customId: command.edited.components.at(-1).components[4].data.custom_id
  })
  await handleGiftCodePanelInteraction(release, dependencies)
  assert.equal(account.discord_user_id, ownerId, "opening confirmation must not release ownership")
  assert.match(release.edited.content, /RELEASE CHARACTER/)

  const confirm = panelInteraction({
    customId: release.edited.components[0].components[0].data.custom_id
  })
  await handleGiftCodePanelInteraction(confirm, dependencies)
  assert.equal(account.discord_user_id, null)
  assert.equal(account.gift_redemption_enabled, false)
  assert.match(confirm.edited.content, /has been released/)
  assert.deepEqual(refreshes, [["777777777777777777", ownerId]])
})

test("self release rechecks ownership at destructive confirmation", async () => {
  const account = {
    ...accounts[0], id: "stale-account", player_id: "93986202",
    discord_user_id: "999999999999999999"
  }
  const dependencies = panelDependencies({ accounts: [account] })
  const command = panelInteraction({ commandName: "player-register" })
  await handleGiftCodePanelInteraction(command, dependencies)
  const release = panelInteraction({
    customId: command.edited.components.at(-1).components[4].data.custom_id
  })
  await handleGiftCodePanelInteraction(release, dependencies)
  assert.match(release.edited.content, /RELEASE CHARACTER/)
  account.discord_user_id = "888888888888888888"
  const confirm = panelInteraction({
    customId: release.edited.components[0].components[0].data.custom_id
  })
  await handleGiftCodePanelInteraction(confirm, dependencies)
  assert.match(confirm.edited.content, /no longer the current owner/)
  assert.equal(account.discord_user_id, "888888888888888888")
})

test("guild authority never grants operator recovery and BOT_OWNER_IDS is rechecked", async () => {
  const operatorId = "707866087248756736"
  const target = {
    ...accounts[0], id: "operator-account", player_id: "93986203",
    discord_user_id: "888888888888888888"
  }
  let guildAuthorizationCalls = 0
  const dependencies = panelDependencies({
    accounts: [target],
    authorization: async () => { guildAuthorizationCalls += 1; return true }
  })

  for (const authority of ["member", "administrator", "guild owner", "configured guild manager", "role holder"]) {
    const denied = panelInteraction({ commandName: "player-admin", options: { player_id: target.player_id } })
    denied.member = { authority }
    denied.guild = { ownerId: authority === "guild owner" ? denied.user.id : "other" }
    await handleGiftCodePanelInteraction(denied, dependencies)
    assert.match(denied.edited.content, /only to configured bot operators/)
  }
  assert.equal(guildAuthorizationCalls, 0)

  dependencies.env.BOT_OWNER_IDS = ` ${operatorId} `
  const lookup = panelInteraction({ commandName: "player-admin", options: { player_id: target.player_id } })
  lookup.user.id = operatorId
  await handleGiftCodePanelInteraction(lookup, dependencies)
  assert.match(lookup.edited.content, /GLOBAL OPERATOR RECOVERY RELEASE/)
  assert.match(lookup.edited.content, /Current Discord owner/)
  assert.equal(target.discord_user_id, "888888888888888888", "lookup must require confirmation")

  dependencies.env.BOT_OWNER_IDS = ""
  const staleConfirm = panelInteraction({
    customId: lookup.edited.components[0].components[0].data.custom_id
  })
  staleConfirm.user.id = operatorId
  await handleGiftCodePanelInteraction(staleConfirm, dependencies)
  assert.match(staleConfirm.edited.content, /only to configured bot operators/)
  assert.equal(target.discord_user_id, "888888888888888888")

  dependencies.env.BOT_OWNER_IDS = operatorId
  const retry = panelInteraction({ commandName: "player-admin", options: { player_id: target.player_id } })
  retry.user.id = operatorId
  await handleGiftCodePanelInteraction(retry, dependencies)
  const confirmed = panelInteraction({
    customId: retry.edited.components[0].components[0].data.custom_id
  })
  confirmed.user.id = operatorId
  await handleGiftCodePanelInteraction(confirmed, dependencies)
  assert.equal(target.discord_user_id, null)
  assert.match(confirmed.edited.content, /Operator release completed/)
})

test("removing one account selects the remaining active account in both profile wordings", async () => {
  for (const [terms, locationLabel] of [[wos, "State"], [kingshot, "Kingdom"]]) {
    const first = { ...accounts[0], player_id: "111", state_or_kingdom_number: "689" }
    const second = { ...accounts[1], player_id: "222", state_or_kingdom_number: "700" }
    const localAccounts = [first, second]
    const dependencies = panelDependencies({ accounts: localAccounts })
    dependencies.giftServiceFactory = () => ({
      terms,
      async status() { return localAccounts.filter(account => account.is_active) }
    })
    const command = panelInteraction({ commandName: "player-register" })
    await handleGiftCodePanelInteraction(command, dependencies)
    const remove = panelInteraction({
      customId: command.edited.components.at(-1).components[3].data.custom_id
    })
    await handleGiftCodePanelInteraction(remove, dependencies)
    assert.match(remove.edited.content, /Selected: Player ID 222/)
    assert.match(remove.edited.content, new RegExp(`${locationLabel}: 700`))
    assert.doesNotMatch(remove.edited.content, /Player ID 111/)
  }
})

test("gift-code panel excludes inactive accounts from normal selection", async () => {
  const inactive = { ...accounts[0], player_id: "93986200", is_active: false }
  const active = { ...accounts[1], player_id: "222", is_active: true }
  const dependencies = panelDependencies({ accounts: [inactive, active] })
  const command = panelInteraction({ commandName: "gift-codes" })
  await handleGiftCodePanelInteraction(command, dependencies)
  assert.match(command.edited.content, /Selected player: 222/)
  assert.doesNotMatch(command.edited.content, /93986200|Inactive/)
})

test("direct gift-code submission preserves case and uses concise status-aware replies", async () => {
  const cases = [
    [false, "candidate", "Got it. I'll check that one."],
    [true, "active", "I've already got that one."],
    [true, "candidate", "Already checking that one."],
    [true, "verifying", "Already checking that one."],
    [true, "expired", "That one's already marked expired."],
    [true, "invalid", "I've already checked that one and it wasn't valid."],
    [true, "restricted", "That one's already waiting for review."],
    [true, "unknown", "That one's already waiting for review."]
  ]
  for (const [duplicate, status, expected] of cases) {
    const submissions = []
    const dependencies = panelDependencies({
      accounts: [],
      submissions,
      authorization: async () => { throw new Error("ordinary command checked admin access") },
      submissionResult: input => ({
        giftCode: { code: input.code.trim(), status },
        duplicate,
        outcome: duplicate ? "already known" : "new candidate"
      })
    })
    const command = panelInteraction({
      commandName: "gift-code-add",
      options: { code: "MiXeD123" }
    })
    await handleGiftCodePanelInteraction(command, dependencies)
    assert.equal(command.edited.content, expected)
    assert.equal(submissions[0].code, "MiXeD123")
    assert.equal(submissions[0].discordUserId, "999999999999999999")
  }
})

test("ordinary users can page through profile-scoped active codes", async () => {
  const dependencies = panelDependencies({
    accounts: [],
    activePages: [
      { codes: [{ code: "NEWEST" }], activeCount: 16, expiredCount: 2, page: 0, pageSize: 15 },
      { codes: [{ code: "OLDEST" }], activeCount: 16, expiredCount: 2, page: 1, pageSize: 15 }
    ]
  })
  const command = panelInteraction({ commandName: "gift-codes" })
  await handleGiftCodePanelInteraction(command, dependencies)
  const activeButton = command.edited.components[0].components.find(button => button.data.label === "Active Codes")
  const active = panelInteraction({ customId: activeButton.data.custom_id })
  await handleGiftCodePanelInteraction(active, dependencies)
  assert.match(active.edited.content, /NEWEST/)
  const next = panelInteraction({ customId: active.edited.components[0].components[2].data.custom_id })
  await handleGiftCodePanelInteraction(next, dependencies)
  assert.match(next.edited.content, /OLDEST/)
  assert.match(next.edited.content, /Page 2 of 2/)
})

test("gift-code administration retains the existing authorization gate", async () => {
  const accounts = []
  let authorizationCalls = 0
  const dependencies = panelDependencies({
    accounts,
    authorization: async () => { authorizationCalls += 1; return false }
  })
  const command = panelInteraction({ commandName: "gift-codes-admin" })
  await handleGiftCodePanelInteraction(command, dependencies)
  assert.equal(authorizationCalls, 1)
  assert.match(command.edited.content, /do not have permission/)
})

function adminDiagnostics() {
  return {
    pending_candidates: 0,
    active_codes: 0,
    expired_codes: 0,
    invalid_codes: 0,
    restricted_review_codes: 0,
    pending_redemptions: 0,
    retry_count: 0
  }
}

test("gift-code Configure Channel succeeds independently of banter Apps Script", async () => {
  const banter = createBanterConfigLookup({
    fetchAction: async () => { throw Object.assign(new Error("Apps Script 404"), { code: "ERR_BAD_REQUEST" }) },
    logger: { warn() {} }
  })
  assert.equal((await banter.get("777777777777777777")).banterChannelId, "")
  const configured = []
  const community = {
    async configuration() {
      return { settings: null, channelAvailable: false, roleAvailable: false }
    },
    async configureChannel(guildId, channelId) { configured.push({ guildId, channelId }) }
  }
  const dependencies = panelDependencies({
    accounts: [],
    adminDiagnostics: adminDiagnostics(),
    community
  })
  const command = panelInteraction({ commandName: "gift-codes-admin" })
  await handleGiftCodePanelInteraction(command, dependencies)
  const configureButton = command.edited.components
    .flatMap(row => row.components)
    .find(button => button.data.label === "Configure Channel")
  const configure = panelInteraction({ customId: configureButton.data.custom_id })
  await handleGiftCodePanelInteraction(configure, dependencies)
  const select = panelInteraction({
    customId: configure.edited.components[0].components[0].data.custom_id,
    values: ["888888888888888888"]
  })
  await handleGiftCodePanelInteraction(select, dependencies)
  assert.deepEqual(configured, [{
    guildId: "777777777777777777",
    channelId: "888888888888888888"
  }])
  assert.match(select.edited.content, /Gift Code Administration/)
})

test("gift-code source channel is configured separately from the announcement channel", async () => {
  const sourceChannels = []
  const announcementChannels = []
  const dependencies = panelDependencies({
    accounts: [],
    sourceChannels,
    adminDiagnostics: adminDiagnostics(),
    community: {
      async configuration() {
        return { settings: null, channelAvailable: false, roleAvailable: false }
      },
      async configureChannel(guildId, channelId) {
        announcementChannels.push({ guildId, channelId })
      }
    }
  })
  const command = panelInteraction({ commandName: "gift-codes-admin" })
  await handleGiftCodePanelInteraction(command, dependencies)
  const sourceButton = command.edited.components
    .flatMap(row => row.components)
    .find(button => button.data.label === "Configure Source Channel")
  const configure = panelInteraction({ customId: sourceButton.data.custom_id })
  await handleGiftCodePanelInteraction(configure, dependencies)
  assert.match(configure.edited.content, /read for candidates/)
  const select = panelInteraction({
    customId: configure.edited.components[0].components[0].data.custom_id,
    values: ["333333333333333333"]
  })
  await handleGiftCodePanelInteraction(select, dependencies)
  assert.deepEqual(sourceChannels, [{
    guildId: "777777777777777777",
    channelId: "333333333333333333"
  }])
  assert.deepEqual(announcementChannels, [])
})

test("PostgreSQL channel-configuration failure stays unavailable and logs safe SQLSTATE context", async () => {
  const logs = []
  const databaseError = Object.assign(
    new Error("connection failed ADMIN_API_KEY=secret token=discord-token https://db.internal/path"),
    { code: "08006" }
  )
  const dependencies = panelDependencies({
    accounts: [],
    adminDiagnostics: adminDiagnostics(),
    community: {
      async configuration() {
        return { settings: null, channelAvailable: false, roleAvailable: false }
      },
      async configureChannel() {
        databaseError.giftCodeHandler = "configure_channel_persist"
        throw databaseError
      }
    },
    logger: { error(value) { logs.push(value) } }
  })
  const command = panelInteraction({ commandName: "gift-codes-admin" })
  await handleGiftCodePanelInteraction(command, dependencies)
  const configureButton = command.edited.components
    .flatMap(row => row.components)
    .find(button => button.data.label === "Configure Channel")
  const configure = panelInteraction({ customId: configureButton.data.custom_id })
  await handleGiftCodePanelInteraction(configure, dependencies)
  const select = panelInteraction({
    customId: configure.edited.components[0].components[0].data.custom_id,
    values: ["888888888888888888"]
  })
  await handleGiftCodePanelInteraction(select, dependencies)
  assert.match(select.edited.content, /services are temporarily unavailable/)
  assert.equal(logs.length, 1)
  const diagnostic = JSON.parse(logs[0])
  assert.equal(diagnostic.event, "gift_code_panel_failure")
  assert.equal(diagnostic.game_profile, "wos")
  assert.equal(diagnostic.guild_id, "777777777777777777")
  assert.equal(diagnostic.interaction_category, "component")
  assert.equal(diagnostic.handler, "configure_channel_persist")
  assert.equal(diagnostic.error_code, "08006")
  assert.doesNotMatch(logs[0], /secret|discord-token|db\.internal/)
})

test("gift-code failure diagnostics classify channel selection without logging custom IDs", () => {
  const interaction = panelInteraction({ customId: `${IDS.adminChannelSelect}sensitive-session` })
  const diagnostic = giftCodePanelFailureDiagnostics(
    interaction,
    { gameProfile: "wos" },
    Object.assign(new Error("database unavailable"), { code: "57P01" })
  )
  assert.equal(diagnostic.handler, "configure_channel_select")
  assert.equal(diagnostic.error_code, "57P01")
  assert.doesNotMatch(JSON.stringify(diagnostic), /sensitive-session/)
})
