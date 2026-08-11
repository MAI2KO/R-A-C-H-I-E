const test = require("node:test")
const assert = require("node:assert/strict")
const fs = require("node:fs")
const path = require("node:path")
const { ChannelType, MessageFlags } = require("discord.js")

const { InteractionSessionStore } = require("../src/interactionSessions")
const { createBanterConfigLookup } = require("../src/banterConfig")
const { profileTerminology } = require("../src/giftCodes/terminology")
const {
  IDS,
  selectedAccount,
  playerPanel,
  giftPanel,
  activeCodesPanel,
  registrationModal,
  adminPanel,
  formatCommunityStats,
  giftCodePanelFailureDiagnostics,
  handleGiftCodePanelInteraction
} = require("../src/giftCodes/discord/panelInteractions")
const {
  buildPlayerRegisterCommand,
  buildGiftCodesCommand,
  buildGiftCodesAdminCommand,
  getGiftCommandData
} = require("../src/giftCodes/discord/commands")
const {
  codeProgressMessage,
  joinMessage,
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
  assert.equal(buildGiftCodesCommand("wos").toJSON().name, "gift-codes")
  assert.equal(buildGiftCodesAdminCommand("wos").toJSON().name, "gift-codes-admin")
  const registered = [
    buildPlayerRegisterCommand("wos").toJSON(),
    ...getGiftCommandData({ PLAYER_GIFT_CODES_ENABLED: "true", GAME_PROFILE: "wos" })
  ].map(command => command.name)
  assert.deepEqual(registered, ["player-register", "gift-codes", "gift-codes-admin"])
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
  assert.equal(empty.components[0].components[0].data.label, "Register Player")

  const selected = selectedAccount(accounts, "222")
  const panel = playerPanel({ sessionId: "s", accounts, selected, terms: wos })
  assert.match(panel.content, /State: 700/)
  assert.match(panel.content, /Primary: No/)
  const menu = panel.components[0].components[0].toJSON()
  assert.equal(menu.options.length, 2)
  assert.equal(menu.options[1].default, true)
  assert.deepEqual(panel.components[1].components.map(button => button.data.label), [
    "Add Account", "Change State", "Gift Code Settings", "Remove Account"
  ])
})

test("registration modal and player panels use State or Kingdom correctly", () => {
  const wosModal = registrationModal("s", wos).toJSON()
  const kingshotModal = registrationModal("s", kingshot).toJSON()
  assert.equal(wosModal.components[1].components[0].label, "State")
  assert.equal(kingshotModal.components[1].components[0].label, "Kingdom")
  assert.match(playerPanel({ sessionId: "s", accounts, selected: accounts[0], terms: kingshot }).content, /Kingdom: 689/)
})

test("gift panel exposes public submission, one-account toggle and cap state", () => {
  const panel = giftPanel({
    sessionId: "s",
    accounts,
    selected: accounts[0],
    terms: wos,
    maximumEnabled: 2
  })
  assert.match(panel.content, /Accounts enabled: 1 \/ 2/)
  assert.match(panel.content, /Recent result: success/)
  assert.deepEqual(panel.components.at(-1).components.map(button => button.data.label), [
    "Submit Gift Code", "Disable Auto-Redeem", "Redemption History", "Active Codes", "Change Player"
  ])
  const noAccount = giftPanel({
    sessionId: "s", accounts: [], selected: null, terms: wos, maximumEnabled: 2
  })
  assert.deepEqual(noAccount.components[0].components.map(button => button.data.label), [
    "Submit Gift Code", "Active Codes", "Register Player"
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
    "Queue Status", "Community Stats", "Refresh"
  ])
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
  assert.match(output, /Verified gift codes: 3/)
  assert.doesNotMatch(output, /\b(?:free|paid|premium|subscription|protected)\b/i)
})

function discordFixture({
  roleEditable = true,
  channelAvailable = true,
  manageRoles = true,
  persistedRole = true
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
  const message = { id: "789", async edit(payload) { edits.push(payload); return this } }
  const channel = {
    id: "123",
    type: ChannelType.GuildText,
    isSendable: () => channelAvailable,
    permissionsFor: () => ({ has: () => channelAvailable }),
    async send(payload) { sent.push(payload); return message },
    messages: { async fetch(messageId) { fetchedMessages.push(messageId); return message } }
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
  assert.match(codeMessage, /Redeemed: 57/)
  assert.ok(!codeMessage.includes("282021376"))
  const joined = joinMessage(payload, {
    registered_users: 47,
    auto_redeem_players: 40,
    enabled_accounts: 63,
    successful_redemptions: 418
  }, { enabledCount: 2, successfulRedemptions: 14 }, wos, 2)
  assert.match(joined, /State: 689/)
  assert.match(joined, /Gift Code Auto-Redeem Activated/)
  assert.match(joined, /Characters covered: 2 \/ 2/)
  assert.match(joined, /Players using Auto-Redeem: 40/)
  assert.ok(!joined.includes("Player ID"))
  assert.doesNotMatch(`${codeMessage}\n${joined}`, /\b(?:free|paid|premium|subscription)\b/i)
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
  values = []
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
  logger = { error() {} }
} = {}) {
  const terms = profileTerminology("wos")
  return {
    sessions: new InteractionSessionStore({ maximumSessions: 20 }),
    userCanManageServer: authorization,
    healthProvider: () => ({ available: true, gameProfile: "wos" }),
    poolProvider: () => ({}),
    playerRepositoryFactory: () => ({}),
    playerServiceFactory: () => ({
      async register({ playerId, locationNumber }) {
        const account = {
          id: `account-${playerId}`,
          player_id: playerId,
          state_or_kingdom_number: locationNumber,
          is_primary: accounts.length === 0,
          is_active: true,
          gift_redemption_enabled: false
        }
        accounts.push(account)
        return account
      },
      async changeLocation({ playerId, locationNumber }) {
        accounts.find(account => account.player_id === playerId).state_or_kingdom_number = locationNumber
      }
    }),
    giftRepositoryFactory: () => ({}),
    giftServiceFactory: () => ({
      terms,
      async status({ playerId } = {}) {
        return playerId ? accounts.filter(account => account.player_id === playerId) : accounts
      },
      async submit(input) {
        submissions.push(input)
        return { giftCode: { code: input.code.trim() }, outcome: "new candidate" }
      },
      async activeCodes({ page }) {
        return activePages[page] || {
          codes: [], activeCount: 0, expiredCount: 0, page, pageSize: 15
        }
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
  const dependencies = panelDependencies({ accounts })
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
