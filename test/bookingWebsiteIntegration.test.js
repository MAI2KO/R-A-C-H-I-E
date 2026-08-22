const assert = require("node:assert/strict")
const test = require("node:test")

const {
  bookingWebsiteConfig,
  canonicalRequest,
  createBookingWebsiteClient,
  signRequest
} = require("../src/bookingWebsiteClient")
const {
  deliverWork,
  discoverManagers,
  handleBookingApprovalInteraction,
  parseApprovalButton,
  renderWork
} = require("../src/bookingDiscordIntegration")

const secret = "local-placeholder-secret-value-1234567890"
const workId = "11111111-1111-4111-8111-111111111111"
const requestId = "22222222-2222-4222-8222-222222222222"

test("website client signs the exact profile-specific canonical request", async () => {
  let captured
  const config = bookingWebsiteConfig({ GAME_PROFILE: "wos", BOOKING_WEBSITE_INTEGRATION_ENABLED: "true",
    BOOKING_WEBSITE_BASE_URL: "https://wos.example.test/", BOOKING_WEBSITE_INTEGRATION_SECRET: secret })
  const api = createBookingWebsiteClient({ config, now: () => 1_800_000_000_000,
    createNonce: () => "nonce-value-123456789",
    fetchImplementation: async (url, options) => {
      captured = { url, options }
      return { ok: true, async json() { return { ok: true, work: [] } } }
    } })
  await api.claim(3)
  const path = "/api/internal/v1/discord/work/claim"
  const expected = signRequest({ secret, method: "POST", path, timestamp: "1800000000",
    nonce: "nonce-value-123456789", body: "{\"limit\":3}" })
  assert.equal(captured.url, `https://wos.example.test${path}`)
  assert.equal(captured.options.headers["x-booking-profile"], "wos")
  assert.equal(captured.options.headers["x-booking-signature"], expected)
  assert.match(canonicalRequest({ method: "POST", path, timestamp: "1800000000",
    nonce: "nonce-value-123456789", body: "{\"limit\":3}" }), /^v1\nPOST/)
})

test("integration is disabled unless complete configuration is explicitly enabled", () => {
  const off = bookingWebsiteConfig({ GAME_PROFILE: "wos" })
  assert.equal(off.enabled, false)
  assert.equal(off.disabledReason, "BOOKING_WEBSITE_INTEGRATION_ENABLED is not true")
  assert.equal(bookingWebsiteConfig({ GAME_PROFILE: "kingshot", BOOKING_WEBSITE_INTEGRATION_ENABLED: "true",
    BOOKING_WEBSITE_BASE_URL: "https://ks.example", BOOKING_WEBSITE_INTEGRATION_SECRET: secret }).enabled, true)
  assert.equal(bookingWebsiteConfig({ GAME_PROFILE: "wos", BOOKING_WEBSITE_INTEGRATION_ENABLED: " true ",
    BOOKING_WEBSITE_BASE_URL: "https://wos.example", BOOKING_WEBSITE_INTEGRATION_SECRET: secret }).enabled, true)
  assert.equal(bookingWebsiteConfig({ GAME_PROFILE: "wos", BOOKING_WEBSITE_INTEGRATION_ENABLED: "true",
    BOOKING_WEBSITE_INTEGRATION_SECRET: secret }).disabledReason, "BOOKING_WEBSITE_BASE_URL is missing")
  assert.equal(bookingWebsiteConfig({ GAME_PROFILE: "wos", BOOKING_WEBSITE_INTEGRATION_ENABLED: "true",
    BOOKING_WEBSITE_BASE_URL: "https://wos.example" }).disabledReason, "BOOKING_WEBSITE_INTEGRATION_SECRET is missing")
  assert.equal(bookingWebsiteConfig({ GAME_PROFILE: "other", BOOKING_WEBSITE_INTEGRATION_ENABLED: "true",
    BOOKING_WEBSITE_BASE_URL: "https://bad.example", BOOKING_WEBSITE_INTEGRATION_SECRET: secret }).enabled, false)
})

test("manager rendering is profile-correct and includes only supplied requirements", () => {
  const rendered = renderWork({ workId, requestId, type: "manager_request", profile: "kingshot",
    communityCode: "9999", communityName: "Test Server", serviceLabel: "Construction",
    playerName: "mark", playerId: "8008", alliance: "ABC", time: "14:30", date: "2030-08-21",
    requirements: [{ label: "Truegold", value: 100, unit: "items" }] })
  assert.match(rendered.content, /Kingdom 9999 — Test Server/)
  assert.match(rendered.content, /Truegold: 100/)
  assert.doesNotMatch(rendered.content, /Speed-ups|Fire Crystals/)
  assert.equal(rendered.components[0].components.length, 2)
})

test("manager final updates preserve request context and remove approval buttons", () => {
  const cases = [
    {
      profile: "wos", location: /State 9999 — Test Server/,
      requirements: [
        { label: "Fire Crystals", value: 100, unit: "items" },
        { label: "Speed-ups", value: 100, unit: "days" }
      ],
      status: "confirmed", decidedByDisplayName: "MAI2KO",
      outcome: /APPROVED\nApproved by MAI2KO$/
    },
    {
      profile: "kingshot", location: /Kingdom 9999 — Test Server/,
      requirements: [{ label: "Truegold", value: 200, unit: "items" }],
      status: "denied", decidedByDisplayName: "Jenn",
      outcome: /DENIED\nDenied by Jenn$/
    },
    {
      profile: "wos", location: /State 9999 — Test Server/,
      requirements: [{ label: "Fire Crystal Shards", value: 300, unit: "items" }],
      status: "expired", decidedByDisplayName: null,
      outcome: /EXPIRED\nThe temporary booking hold expired\.$/
    }
  ]

  for (const item of cases) {
    const pending = renderWork({ workId, requestId, type: "manager_request", profile: item.profile,
      communityCode: "9999", communityName: "Test Server", serviceLabel: "Construction",
      playerName: "mark", playerId: "8008", alliance: "TIT", time: "14:30", date: "2030-08-21",
      requirements: item.requirements })
    const final = renderWork({ workId, requestId, type: "manager_update", profile: item.profile,
      status: item.status, decidedByDisplayName: item.decidedByDisplayName,
      originalContent: pending.content })

    assert.match(final.content, item.location)
    assert.match(final.content, /Player: mark/)
    assert.match(final.content, /Alliance: TIT/)
    assert.match(final.content, /Player ID: 8008/)
    assert.match(final.content, /Time: 14:30 UTC/)
    assert.match(final.content, /Date: 21 August/)
    for (const requirement of item.requirements) {
      const suffix = requirement.unit === "days" ? " days" : ""
      assert.match(final.content, new RegExp(`${requirement.label}: ${requirement.value}${suffix}`))
    }
    assert.match(final.content, item.outcome)
    assert.doesNotMatch(final.content, /holding the slot temporarily/)
    assert.deepEqual(final.components, [])
    assert.deepEqual(renderWork({ workId, requestId, type: "manager_update", profile: item.profile,
      status: item.status, decidedByDisplayName: item.decidedByDisplayName,
      originalContent: final.content }), final)
  }
})

test("all manager message copies receive the same context-preserving final edit", async () => {
  const pending = renderWork({ workId, requestId, type: "manager_request", profile: "wos",
    communityCode: "9999", communityName: "Test Server", serviceLabel: "Construction",
    playerName: "mark", playerId: "8008", alliance: "TIT", time: "14:30", date: "2030-08-21",
    requirements: [{ label: "Fire Crystals", value: 100, unit: "items" }] })
  const edits = []
  const messages = new Map(["message-1", "message-2"].map(id => [id, {
    id, content: pending.content, async edit(value) { edits.push(value) }
  }]))
  const channel = { id: "channel", messages: { fetch: async id => messages.get(id) } }
  const outcomes = []
  const api = { async outcome(_work, value) { outcomes.push(value) } }
  const base = { type: "manager_update", profile: "wos", requestId, status: "confirmed",
    decidedByDisplayName: "MAI2KO", discordChannelId: channel.id }

  await deliverWork({ channels: { fetch: async () => channel } }, api,
    { ...base, workId, discordMessageId: "message-1" })
  await deliverWork({ channels: { fetch: async () => channel } }, api,
    { ...base, workId: requestId, discordMessageId: "message-2" })

  assert.equal(edits.length, 2)
  assert.deepEqual(edits[0], edits[1])
  assert.match(edits[0].content, /Fire Crystals: 100[\s\S]*APPROVED\nApproved by MAI2KO$/)
  assert.deepEqual(edits[0].components, [])
  assert.ok(outcomes.every(item => item.status === "sent"))
})

test("player and reminder rendering uses profile terminology and bounded private content", () => {
  const confirmed = renderWork({ workId, type: "player_confirmed", profile: "wos",
    communityCode: "9999", communityName: "Test Server", serviceLabel: "Research",
    date: "2030-08-21", time: "14:30" })
  assert.match(confirmed.content, /Appointment confirmed[\s\S]*State 9999/)
  assert.doesNotMatch(confirmed.content, /bookingId|booking-work|Player ID/)
  const reminder = renderWork({ workId, type: "appointment_reminder", profile: "kingshot",
    communityCode: "9999", communityName: "Test Server", serviceLabel: "Construction",
    playerName: "mark", alliance: "ABC", time: "14:30" })
  assert.match(reminder.content, /in 30 minutes[\s\S]*Kingdom 9999[\s\S]*Alliance: ABC/)
})

test("manager mutations add bounded attribution while self-service notifications remain unattributed", () => {
  for (const [profile, place] of [["wos", "State 9999"], ["kingshot", "Kingdom 9999"]]) {
    const rescheduled = renderWork({ workId, type: "player_rescheduled", profile,
      communityCode: "9999", communityName: "Test Server", serviceLabel: "Construction",
      previousDate: "2030-08-21", previousTime: "14:00", date: "2030-08-21", time: "14:30",
      attributionDisplayName: "MAI2KO" })
    assert.match(rescheduled.content, /Appointment rescheduled/)
    assert.match(rescheduled.content, /Previous: 21 August at 14:00 UTC/)
    assert.match(rescheduled.content, /New: 21 August at 14:30 UTC/)
    assert.match(rescheduled.content, /Rescheduled by MAI2KO$/)

    const cancelled = renderWork({ workId, type: "player_cancelled", profile,
      communityCode: "9999", communityName: "Test Server", serviceLabel: "Construction",
      date: "2030-08-21", time: "14:30", attributionDisplayName: "Jenn" })
    assert.match(cancelled.content, new RegExp(place))
    assert.match(cancelled.content, /Cancelled by Jenn$/)

    const selfRescheduled = renderWork({ workId, type: "player_rescheduled", profile,
      communityCode: "9999", communityName: "Test Server", serviceLabel: "Construction",
      previousDate: "2030-08-21", previousTime: "14:00", date: "2030-08-21", time: "14:30" })
    const selfCancelled = renderWork({ workId, type: "player_cancelled", profile,
      communityCode: "9999", communityName: "Test Server", serviceLabel: "Construction",
      date: "2030-08-21", time: "14:30" })
    assert.doesNotMatch(selfRescheduled.content, /Rescheduled by/)
    assert.doesNotMatch(selfCancelled.content, /Cancelled by/)
  }
})

test("manager discovery includes owner, administrators and manager role and deduplicates guilds", async () => {
  const member = (id, { owner = false, admin = false, role = false } = {}) => ({ id,
    user: { bot: false }, permissions: { has: () => admin }, roles: { cache: { has: () => role } }, owner })
  const guilds = {
    a: { ownerId: "1", members: { fetch: async () => new Map([["1", member("1")], ["2", member("2", { admin: true })], ["3", member("3", { role: true })], ["4", member("4")]]) } },
    b: { ownerId: "5", members: { fetch: async () => new Map([["2", member("2", { admin: true })], ["5", member("5")]]) } }
  }
  const recipients = await discoverManagers({ guilds: { fetch: async id => guilds[id] } }, {
    guilds: [{ guildId: "a", managerRoleId: "role" }, { guildId: "b", managerRoleId: null }]
  })
  assert.deepEqual(recipients.map(row => row.discordUserId), ["1", "2", "3", "5"])
  assert.equal(recipients.filter(row => row.discordUserId === "2").length, 1)
})

test("delivery uses a stable Discord nonce and classifies forbidden DMs permanently", async () => {
  const outcomes = []
  const sentByNonce = new Map()
  const channel = { id: "77", async send(input) {
    assert.equal(input.enforceNonce, true)
    if (!sentByNonce.has(input.nonce)) sentByNonce.set(input.nonce, { id: "88" })
    return sentByNonce.get(input.nonce)
  } }
  const client = { user: { id: "bot" }, users: { fetch: async () => ({ createDM: async () => channel }) } }
  const api = { outcome: async (_work, result) => { outcomes.push(result) } }
  await deliverWork(client, api, { workId, type: "player_confirmed", profile: "wos",
    recipientDiscordUserId: "1", communityCode: "1", communityName: "T", serviceLabel: "Troop", date: "2030-01-01", time: "10:00" })
  await deliverWork(client, api, { workId, type: "player_confirmed", profile: "wos",
    recipientDiscordUserId: "1", communityCode: "1", communityName: "T", serviceLabel: "Troop", date: "2030-01-01", time: "10:00" })
  assert.deepEqual(outcomes[0], { status: "sent", discordChannelId: "77", discordMessageId: "88" })
  assert.deepEqual(outcomes[1], outcomes[0])
  assert.equal(sentByNonce.size, 1)
  client.users.fetch = async () => { const error = new Error("closed"); error.code = 50007; throw error }
  await deliverWork(client, api, { workId: requestId, type: "player_confirmed", profile: "wos", recipientDiscordUserId: "2" })
  assert.equal(outcomes[2].status, "permanent_failure")
})

test("approval buttons call the website with canonical actor and return controlled duplicate state", async () => {
  assert.deepEqual(parseApprovalButton(`booking-approval:v1:${requestId}:deny`), { requestId, action: "deny" })
  const replies = []
  const interaction = { customId: `booking-approval:v1:${requestId}:approve`, isButton: () => true,
    user: { id: "123", username: "mark" }, member: { displayName: "Mark" },
    async deferReply(value) { assert.deepEqual(value, { flags: 64 }) }, async editReply(value) { replies.push(value) } }
  const calls = []
  const handled = await handleBookingApprovalInteraction(interaction, { async approval(...args) {
    calls.push(args); return { result: { outcome: "already_confirmed", decidedByDisplayName: "Jenn" } }
  } })
  assert.equal(handled, true)
  assert.deepEqual(calls[0], [requestId, "approve", { discordUserId: "123", displayName: "Mark" }])
  assert.equal(replies[0], "This request is already approved by Jenn.")
})

test("cross-profile work is ignored by a profile worker", async () => {
  const { createBookingWebsiteRuntime } = require("../src/bookingDiscordIntegration")
  let delivered = 0
  const api = { profile: "wos", claim: async () => ({ work: [{ profile: "kingshot", type: "player_confirmed" }] }),
    outcome: async () => { delivered++ } }
  const runtime = createBookingWebsiteRuntime({ client: {}, api, logger: { log() {}, error() {} } })
  assert.equal(await runtime.tick(), 1)
  assert.equal(delivered, 0)
})

test("polling logs first connection, non-empty claims, one failure category and recovery", async () => {
  const { createBookingWebsiteRuntime } = require("../src/bookingDiscordIntegration")
  const logs = []
  const failures = []
  const responses = [
    { work: [] },
    { work: [{ profile: "other" }] },
    Object.assign(new Error("private detail"), { code: "network_error" }),
    Object.assign(new Error("another private detail"), { code: "network_error" }),
    { work: [] }
  ]
  const api = { profile: "wos", async claim() {
    const value = responses.shift()
    if (value instanceof Error) throw value
    return value
  } }
  const runtime = createBookingWebsiteRuntime({ client: {}, api,
    logger: { log(value) { logs.push(JSON.parse(value)) }, error(value) { failures.push(JSON.parse(value)) } } })
  for (let index = 0; index < 5; index++) await runtime.tick()
  assert.equal(logs.filter(item => item.event === "booking_website_connection_established").length, 1)
  assert.equal(logs.filter(item => item.event === "booking_website_work_claimed")[0].work_count, 1)
  assert.equal(failures.length, 1)
  assert.deepEqual(failures[0], { event: "booking_website_poll_failed", game_profile: "wos",
    operation: "claim", http_status: null, error_code: "network_error" })
  assert.equal(logs.filter(item => item.event === "booking_website_connection_recovered").length, 1)
  assert.doesNotMatch(JSON.stringify([...logs, ...failures]), /private detail/)
})

test("worker polls immediately and its timer continues after failure, empty, and non-empty polls", async () => {
  const { createBookingWebsiteRuntime } = require("../src/bookingDiscordIntegration")
  let scheduled
  let interval
  let claims = 0
  const api = { profile: "wos", async claim() {
    claims++
    if (claims === 1) throw Object.assign(new Error("temporary"), { code: "network_error" })
    return { work: claims === 3 ? [{ profile: "kingshot" }] : [] }
  } }
  const runtime = createBookingWebsiteRuntime({ client: {}, api,
    logger: { log() {}, error() {} },
    setIntervalFn(handler, milliseconds) {
      scheduled = handler; interval = milliseconds
      return { unref() {} }
    }, clearIntervalFn() {} })
  assert.deepEqual(runtime.start(), { started: true })
  await new Promise(resolve => setImmediate(resolve))
  assert.equal(claims, 1)
  assert.equal(interval, 10000)
  scheduled(); await new Promise(resolve => setImmediate(resolve))
  scheduled(); await new Promise(resolve => setImmediate(resolve))
  assert.equal(claims, 3)
})
