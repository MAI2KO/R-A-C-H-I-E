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
  bookingWindowOpenMessages,
  manualGuestLinkMessage,
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
  assert.equal(api.baseUrl, "https://wos.example.test")
  const path = "/api/internal/v1/discord/work/claim"
  const expected = signRequest({ secret, method: "POST", path, timestamp: "1800000000",
    nonce: "nonce-value-123456789", body: "{\"limit\":3}" })
  assert.equal(captured.url, `https://wos.example.test${path}`)
  assert.equal(captured.options.headers["x-booking-profile"], "wos")
  assert.equal(captured.options.headers["x-booking-signature"], expected)
  assert.match(canonicalRequest({ method: "POST", path, timestamp: "1800000000",
    nonce: "nonce-value-123456789", body: "{\"limit\":3}" }), /^v1\nPOST/)
})

test("setup and canonical registration use the signed profile-specific website integration", async () => {
  const requests = []
  const config = bookingWebsiteConfig({ GAME_PROFILE: "kingshot", BOOKING_WEBSITE_INTEGRATION_ENABLED: "true",
    BOOKING_WEBSITE_BASE_URL: "https://ks.example.test", BOOKING_WEBSITE_INTEGRATION_SECRET: secret })
  const api = createBookingWebsiteClient({ config, now: () => 1_800_000_000_000,
    createNonce: () => `nonce-value-${requests.length}-123456789`,
    fetchImplementation: async (url, options) => {
      requests.push({ url, options })
      return { ok: true, async json() { return { ok: true, status: "linked" } } }
    } })
  await api.communitySetup({ guildId: "777", communityCode: "9999", dryRun: true })
  await api.registration({ guildId: "777", discordUserId: "888", playerId: "123" })
  assert.deepEqual(requests.map(request => new URL(request.url).pathname), [
    "/api/internal/v1/discord/setup/community", "/api/internal/v1/discord/registration"
  ])
  assert.equal(requests.every(request => request.options.headers["x-booking-profile"] === "kingshot"), true)
  assert.equal(requests.every(request => /^v1=[0-9a-f]{64}$/.test(
    request.options.headers["x-booking-signature"]
  )), true)
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
  assert.equal(bookingWebsiteConfig({ NODE_ENV: "production", GAME_PROFILE: "wos",
    BOOKING_WEBSITE_INTEGRATION_ENABLED: "true", BOOKING_WEBSITE_BASE_URL: "https://localhost:8080",
    BOOKING_WEBSITE_INTEGRATION_SECRET: secret }).enabled, false)
  assert.equal(bookingWebsiteConfig({ NODE_ENV: "production", GAME_PROFILE: "wos",
    BOOKING_WEBSITE_INTEGRATION_ENABLED: "true",
    BOOKING_WEBSITE_BASE_URL: "https://staging.r-a-c-h-i-e.com",
    BOOKING_WEBSITE_INTEGRATION_SECRET: secret }).enabled, false)
  assert.equal(bookingWebsiteConfig({ NODE_ENV: "production", GAME_PROFILE: "kingshot",
    BOOKING_WEBSITE_INTEGRATION_ENABLED: "true",
    BOOKING_WEBSITE_BASE_URL: "https://r-a-c-h-i-e.com",
    BOOKING_WEBSITE_INTEGRATION_SECRET: secret }).enabled, false)
  assert.equal(bookingWebsiteConfig({ NODE_ENV: "production", GAME_PROFILE: "kingshot",
    BOOKING_WEBSITE_INTEGRATION_ENABLED: "true",
    BOOKING_WEBSITE_BASE_URL: "https://peggie.r-a-c-h-i-e.com",
    BOOKING_WEBSITE_INTEGRATION_SECRET: secret }).enabled, true)
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

test("booking-open rendering is guest-only with exact resolved close time", () => {
  const guestUrl = `https://r-a-c-h-i-e.com/book/${"g".repeat(43)}`
  const rendered = bookingWindowOpenMessages({
    profile: "wos", communityCode: "9999",
    closesAt: "2030-09-06T12:30:00Z", guestUrl
  })
  assert.equal(rendered.public.content,
    `Guest booking — State 9999\n\nUse this link for players who cannot access Discord.\nCopy and paste this link into in-game chat so those players can request a booking:\n${guestUrl}\n\nCloses:\n6 September 2030, 12:30 UTC`)
  assert.deepEqual(rendered.public.components, [])
  assert.equal(rendered.manager.content, rendered.public.content)
  assert.doesNotMatch(rendered.public.content, /Member sign-up|Guest sign-up|\/booking/)
})

test("booking-open delivery posts once with stable nonces and DMs every native manager", async () => {
  const sends = []
  const channel = { id: "700", async send(payload) {
    sends.push(["public", payload])
    return { id: "701" }
  } }
  const dm = id => ({ async send(payload) { sends.push([id, payload]); return { id } } })
  const member = (id, { admin = false, role = false } = {}) => ({
    id, user: { bot: false, createDM: async () => dm(id) },
    permissions: { has: () => admin }, roles: { cache: { has: () => role } }
  })
  const guild = {
    ownerId: "1", channels: { fetch: async () => channel },
    members: { fetch: async () => new Map([
      ["1", member("1")], ["2", member("2", { admin: true })],
      ["3", member("3", { role: true })], ["4", member("4")]
    ]) }
  }
  const outcomes = []
  await deliverWork({ guilds: { fetch: async () => guild } }, {
    baseUrl: "https://r-a-c-h-i-e.com",
    async outcome(_work, outcome) { outcomes.push(outcome) }
  }, {
    workId, type: "booking_window_open", profile: "wos", communityCode: "9999",
    guilds: ["777"], closesAt: "2999-09-06T12:00:00Z",
    guestPath: `/book/${"w".repeat(43)}`
  }, { setupRepository: { async get() {
    return { minister_sign_up_channel_id: "700", bot_manager_role_id: "42" }
  } } })
  assert.deepEqual(sends.map(([kind]) => kind), ["public", "1", "2", "3"])
  assert.equal(sends.every(([, payload]) => payload.enforceNonce === true), true)
  assert.deepEqual(outcomes, [{ status: "sent", discordChannelId: "700", discordMessageId: "701" }])
  assert.equal(sends.every(([, payload]) => payload.components.length === 0), true)
  assert.equal(sends.every(([, payload]) => payload.content.includes(
    `https://r-a-c-h-i-e.com/book/${"w".repeat(43)}`)), true)
})

test("booking-open uses the Kingshot public origin and rejects missing, internal, or staging origins", async () => {
  const token = "k".repeat(43)
  const sent = []
  const outcomes = []
  const guild = { ownerId: "1", channels: { fetch: async () => ({ id: "channel", async send(payload) {
    sent.push(payload); return { id: "message" }
  } }) }, members: { fetch: async () => new Map() } }
  const setupRepository = { async get() {
    return { minister_sign_up_channel_id: "channel", bot_manager_role_id: null }
  } }
  const work = { workId, type: "booking_window_open", profile: "kingshot",
    communityCode: "1234", guilds: ["777"], closesAt: "2999-09-06T12:00:00Z",
    guestPath: `/book/${token}`, guestUrl: `https://localhost:8080/book/${token}` }
  await deliverWork({ guilds: { fetch: async () => guild } }, {
    baseUrl: "https://peggie.r-a-c-h-i-e.com",
    async outcome(_work, outcome) { outcomes.push(outcome) }
  }, work, { setupRepository })
  assert.match(sent[0].content, new RegExp(`https://peggie\\.r-a-c-h-i-e\\.com/book/${token}`))
  assert.doesNotMatch(sent[0].content, /localhost|railway|staging|\/booking/i)
  for (const baseUrl of [undefined, "https://localhost:8080", "https://127.0.0.1:8080",
    "https://web.railway.internal", "https://staging.r-a-c-h-i-e.com",
    "https://r-a-c-h-i-e.com"]) {
    await deliverWork({ guilds: { async fetch() { throw new Error("must not fetch") } } }, {
      baseUrl, async outcome(_work, outcome) { outcomes.push(outcome) }
    }, work, { setupRepository })
  }
  assert.equal(outcomes.slice(1).every((outcome) => outcome.status === "retry"
    && outcome.errorCode === "booking_website_public_url_invalid"), true)
  assert.equal(JSON.stringify(outcomes).includes(token), false)
})

test("manual guest-link delivery DMs deduplicated managers only with exact profile copy", async () => {
  const sends = []
  const member = (id, { admin = false, role = false } = {}) => ({
    id, user: { bot: false, async createDM() { return { id: `dm-${id}`, async send(payload) {
      sends.push([id, payload]); return { id: `message-${id}` }
    } } } },
    permissions: { has: () => admin }, roles: { cache: { has: () => role } }
  })
  const shared = member("1")
  const guilds = {
    "700": { ownerId: "1", members: { fetch: async () => new Map([
      ["1", shared], ["2", member("2", { admin: true })], ["4", member("4")]
    ]) } },
    "701": { ownerId: "3", members: { fetch: async () => new Map([
      ["1", shared], ["3", member("3", { role: true })]
    ]) } }
  }
  const outcomes = []
  await deliverWork({ guilds: { fetch: async id => guilds[id] } }, {
    baseUrl: "https://peggie.r-a-c-h-i-e.com",
    async outcome(_work, outcome) { outcomes.push(outcome) }
  }, {
    workId, type: "manager_guest_link", profile: "kingshot", communityCode: "1234",
    guilds: ["700", "701"], guestUrl: `https://localhost:8080/book/${"k".repeat(43)}`
  }, { setupRepository: { async get() {
    return { minister_sign_up_channel_id: "800", bot_manager_role_id: "42" }
  } } })
  assert.deepEqual(sends.map(([id]) => id), ["1", "2", "3"])
  assert.equal(sends.every(([, payload]) => payload.enforceNonce === true), true)
  assert.equal(sends[0][1].content,
    `New guest booking link — Kingdom 1234\n\nA new guest sign-up link has been created.\n\nhttps://peggie.r-a-c-h-i-e.com/book/${"k".repeat(43)}\n\nGuest bookings require manager approval.`)
  assert.deepEqual(outcomes, [{ status: "sent", discordChannelId: "dm-1",
    discordMessageId: "message-1" }])
  assert.equal(manualGuestLinkMessage({ profile: "wos", communityCode: "1234", guestUrl: "url" }).content,
    "New guest booking link — State 1234\n\nA new guest sign-up link has been created.\n\nurl\n\nGuest bookings require manager approval.")
})

test("manual guest-link retries use the same Discord-enforced nonce", async () => {
  const payloads = []
  const member = { id: "1", user: { bot: false, async createDM() { return {
    id: "dm-1", async send(payload) { payloads.push(payload); return { id: "message-1" } }
  } } }, permissions: { has: () => false }, roles: { cache: { has: () => false } } }
  const client = { guilds: { fetch: async () => ({ ownerId: "1",
    members: { fetch: async () => new Map([["1", member]]) } }) } }
  const api = { baseUrl: "https://r-a-c-h-i-e.com", async outcome() {} }
  const work = { workId, type: "manager_guest_link", profile: "wos", communityCode: "1234",
    guilds: ["700"], guestPath: `/book/${"w".repeat(43)}`,
    guestUrl: `https://staging.internal.example/book/${"x".repeat(43)}` }
  const options = { setupRepository: { async get() {
    return { minister_sign_up_channel_id: "800", bot_manager_role_id: null }
  } } }
  await deliverWork(client, api, work, options)
  await deliverWork(client, api, work, options)
  assert.equal(payloads.length, 2)
  assert.equal(payloads[0].nonce, payloads[1].nonce)
  assert.equal(payloads.every(payload => payload.enforceNonce === true), true)
  assert.match(payloads[0].content, new RegExp(`https://r-a-c-h-i-e\\.com/book/${"w".repeat(43)}`))
  assert.doesNotMatch(payloads[0].content, /localhost|railway|staging/i)
})

test("manual guest-link delivery fails safely without a valid configured website URL", async () => {
  const token = "s".repeat(43)
  const outcomes = []
  let fetched = false
  const client = { guilds: { async fetch() { fetched = true } } }
  for (const baseUrl of [undefined, "http://internal.railway", "not-a-url"]) {
    await deliverWork(client, { baseUrl, async outcome(_work, outcome) { outcomes.push(outcome) } }, {
      workId, type: "manager_guest_link", profile: "wos", communityCode: "1234",
      guilds: ["700"], guestUrl: `http://localhost:8080/book/${token}`
    }, { setupRepository: { async get() { return { bot_manager_role_id: "42" } } } })
  }
  assert.equal(fetched, false)
  assert.equal(outcomes.every(outcome => outcome.status === "retry"
    && outcome.errorCode === "booking_website_public_url_invalid"), true)
  assert.equal(JSON.stringify(outcomes).includes(token), false)
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
    date: "2030-08-21", time: "14:30", appointmentAt: "2030-08-21T14:30:00.000Z" })
  assert.match(confirmed.content, /Appointment confirmed[\s\S]*State 9999/)
  assert.match(confirmed.content, /UTC: 21 August at 14:30 UTC/)
  assert.match(confirmed.content, /Your time: <t:\d+:F>/)
  assert.doesNotMatch(confirmed.content, /bookingId|booking-work|Player ID/)
  const reminder = renderWork({ workId, type: "appointment_reminder", profile: "kingshot",
    communityCode: "9999", communityName: "Test Server", serviceLabel: "Construction",
    playerName: "mark", alliance: "ABC", date: "2030-08-21", time: "14:30",
    appointmentAt: "2030-08-21T14:30:00.000Z" })
  assert.match(reminder.content, /in 30 minutes[\s\S]*Kingdom 9999[\s\S]*Alliance: ABC/)
  assert.match(reminder.content, /UTC: 21 August at 14:30 UTC[\s\S]*Your time: <t:\d+:F>/)
})

test("manual approval, cancellation and malformed appointments use safe UTC/local presentation", () => {
  const approved = renderWork({ workId, type: "player_approved", profile: "kingshot",
    communityCode: "9999", communityName: "Test Server", serviceLabel: "Troop",
    date: "2030-08-21", time: "05:30", appointmentAt: "2030-08-21T05:30:00.000Z",
    decidedByDisplayName: "Jenn" })
  assert.match(approved.content,
    /Appointment approved[\s\S]*Kingdom 9999[\s\S]*UTC: 21 August at 05:30 UTC[\s\S]*Your time: <t:\d+:F>[\s\S]*Approved by Jenn$/)

  const cancelled = renderWork({ workId, type: "player_cancelled", profile: "wos",
    communityCode: "9999", communityName: "Test Server", serviceLabel: "Troop",
    date: "2030-08-21", time: "05:30", appointmentAt: "2030-08-20T23:30:00.000Z" })
  assert.match(cancelled.content, /UTC: 20 August at 23:30 UTC[\s\S]*Your time: <t:\d+:F>/)
  assert.doesNotMatch(cancelled.content, /Cancelled by/)

  const malformed = renderWork({ workId, type: "player_confirmed", profile: "wos",
    communityCode: "9999", communityName: "Test Server", serviceLabel: "Troop",
    date: "not-a-date", time: "99:99", appointmentAt: "invalid" })
  assert.match(malformed.content, /Your time: unavailable/)
  assert.doesNotMatch(malformed.content, /<t:(?:NaN|undefined|Invalid)/)
})

test("manager mutations add bounded attribution while self-service notifications remain unattributed", () => {
  for (const [profile, place] of [["wos", "State 9999"], ["kingshot", "Kingdom 9999"]]) {
    const rescheduled = renderWork({ workId, type: "player_rescheduled", profile,
      communityCode: "9999", communityName: "Test Server", serviceLabel: "Construction",
      previousDate: "2030-08-21", previousTime: "14:00", date: "2030-08-21", time: "14:30",
      attributionDisplayName: "MAI2KO" })
    assert.match(rescheduled.content, /Appointment rescheduled/)
    assert.match(rescheduled.content, /Previous:\nUTC: 21 August at 14:00 UTC\nYour time: <t:\d+:F>/)
    assert.match(rescheduled.content, /New:\nUTC: 21 August at 14:30 UTC\nYour time: <t:\d+:F>/)
    assert.match(rescheduled.content, new RegExp(place))
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

test("guest approval discovery includes owner, administrator and native manager role without unrelated members", async () => {
  const member = (id, { owner = false, admin = false, role = false } = {}) => ({ id,
    user: { bot: false }, permissions: { has: () => admin }, roles: { cache: { has: () => role } }, owner })
  const guilds = {
    a: { ownerId: "1", members: { fetch: async () => new Map([["1", member("1")], ["2", member("2", { admin: true })], ["3", member("3", { role: true })], ["4", member("4")]]) } },
    b: { ownerId: "5", members: { fetch: async () => new Map([["2", member("2", { admin: true })], ["5", member("5")]]) } }
  }
  let recipients
  await deliverWork({ guilds: { fetch: async id => guilds[id] } }, {
    async recipients(_work, value) { recipients = value },
    async outcome() { assert.fail("manager discovery should register recipients") }
  }, { workId, type: "manager_discovery", guilds: [{ guildId: "a" }, { guildId: "b" }] }, {
    setupRepository: { async get(guildId) {
      return { bot_manager_role_id: guildId === "a" ? "role" : null }
    } }
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
