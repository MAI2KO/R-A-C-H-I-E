const test = require("node:test")
const assert = require("node:assert/strict")

const { createPublicAllianceEventRepository } = require("../src/publicAllianceEventRepository")
const { publicAllianceEventsReadModel } = require("../src/publicAllianceEvents")
const {
  signedPublicAllianceEventsHeaders,
  verifyPublicAllianceEventsRequest
} = require("../src/publicAllianceEventsAuth")
const { handlePublicAllianceEventsRead } = require("../src/publicAllianceEventsServer")
const { handleNativeManagerAuthorization } = require("../src/nativeManagerAuthorizationServer")
const {
  publicAllianceEventsConfig,
  createPublicAllianceEventsBootstrap
} = require("../src/publicAllianceEventsBootstrap")

const secret = "local-test-secret-with-at-least-32-characters"
const now = new Date("2026-08-23T12:00:00.000Z")

function event(overrides = {}) {
  return {
    alliance_id: "20", alliance_name: "Titans", event_id: "30",
    event_name: "Bear Trap", first_occurrence_date: "2026-08-23",
    event_time_utc: "19:00", recurrence_days: 2, groups: [], ...overrides
  }
}

test("public read model reuses recurrence rules, groups by alliance, and has no private IDs", () => {
  const model = publicAllianceEventsReadModel({
    profile: "wos", communityCode: "9999", now,
    events: [
      event({ alliance_id: "2", alliance_name: "Zulu", event_id: "5", event_name: "Foundry", event_time_utc: null,
        groups: [{ group_name: "Best Bear", first_occurrence_date: "2026-08-23", event_time_utc: "18:30", sort_order: 0 }] }),
      event({ alliance_id: "1", alliance_name: "Alpha", event_id: "4" })
    ]
  })
  assert.deepEqual(model.alliances.map(alliance => alliance.name), ["Alpha", "Zulu"])
  assert.deepEqual(model.alliances[0].events[0].upcoming.map(row => row.at), [
    "2026-08-23T19:00:00.000Z", "2026-08-25T19:00:00.000Z", "2026-08-27T19:00:00.000Z"
  ])
  assert.equal(model.alliances[1].events[0].upcoming[0].group, "Best Bear")
  assert.equal(model.alliances[0].events[0].recurrence.summary, "Every 2 days")
  assert.equal(/guild|channel|message|_id|lock|claim/i.test(JSON.stringify(model)), false)
})

test("public read model returns consecutive daily occurrences", () => {
  const model = publicAllianceEventsReadModel({
    profile: "wos", communityCode: "9999", now,
    events: [event({ recurrence_days: 1 })]
  })
  assert.equal(model.alliances[0].events[0].recurrence.summary, "Every day")
  assert.deepEqual(model.alliances[0].events[0].upcoming.map(row => row.at), [
    "2026-08-23T19:00:00.000Z", "2026-08-24T19:00:00.000Z", "2026-08-25T19:00:00.000Z"
  ])
})

test("repository query is profile/community scoped and selects active alliance events only", async () => {
  const calls = []
  const repository = createPublicAllianceEventRepository({ query: async (sql, values) => {
    calls.push({ sql, values })
    return { rows: [] }
  } }, "kingshot")
  await repository.listForCommunity("9999")
  assert.deepEqual(calls[0].values, ["kingshot", "9999", 1000])
  assert.match(calls[0].sql, /destination\.game_profile = \$1/)
  assert.match(calls[0].sql, /destination\.state_number = \$2/)
  assert.match(calls[0].sql, /e\.status = 'active'/)
  assert.match(calls[0].sql, /link\.sharing_enabled = true/)
  assert.doesNotMatch(calls[0].sql, /\bstate_events\b/)
})

test("guild repository query is profile/guild scoped and independent of State delivery tables", async () => {
  const calls = []
  const repository = createPublicAllianceEventRepository({ query: async (sql, values) => {
    calls.push({ sql, values })
    return { rows: [] }
  } }, "wos")
  await repository.listForGuild("123456789012345678")
  assert.deepEqual(calls[0].values, ["wos", "123456789012345678", 1000])
  assert.match(calls[0].sql, /e\.game_profile = \$1/)
  assert.match(calls[0].sql, /e\.guild_id = \$2/)
  assert.match(calls[0].sql, /e\.status = 'active'/)
  assert.doesNotMatch(calls[0].sql, /event_state_destinations|event_state_links|state_events|publish_to_state|roundup/i)
  await assert.rejects(repository.listForGuild("9999"), /Invalid Discord guild ID/)
})

test("signed internal reads authenticate profile and return only the public model", async () => {
  const path = "/internal/v1/public-alliance-events/9999"
  const headers = signedPublicAllianceEventsHeaders({ secret, profile: "wos", method: "GET", path,
    now: () => now.getTime(), createNonce: () => "abcdefghijklmnop" })
  assert.equal(verifyPublicAllianceEventsRequest({ secret, method: "GET", path, profile: "wos",
    timestamp: headers["x-alliance-events-timestamp"], nonce: headers["x-alliance-events-nonce"],
    signature: headers["x-alliance-events-signature"], now: () => now.getTime() }), true)
  const result = await handlePublicAllianceEventsRead({ method: "GET", path, headers }, {
    config: { secret, profile: "wos" }, repository: { listForCommunity: async () => [event()] }, now: () => now
  })
  assert.equal(result.status, 200)
  assert.equal(result.body.profile, "wos")
  const wrongProfile = await handlePublicAllianceEventsRead({ method: "GET", path,
    headers: { ...headers, "x-alliance-events-profile": "kingshot" } }, {
    config: { secret, profile: "wos" }, repository: { listForCommunity: async () => [] }, now: () => now
  })
  assert.equal(wrongProfile.status, 401)
})

test("signed guild reads validate snowflakes and do not expose the requested guild ID", async () => {
  const path = "/internal/v1/public-alliance-events/guild/123456789012345678"
  const headers = signedPublicAllianceEventsHeaders({ secret, profile: "wos", method: "GET", path,
    now: () => now.getTime(), createNonce: () => "abcdefghijklmnop" })
  const result = await handlePublicAllianceEventsRead({ method: "GET", path, headers }, {
    config: { secret, profile: "wos" }, repository: { listForGuild: async guildId => {
      assert.equal(guildId, "123456789012345678")
      return [event()]
    } }, now: () => now
  })
  assert.equal(result.status, 200)
  assert.deepEqual(Object.keys(result.body), ["ok", "profile", "alliances"])
  assert.doesNotMatch(JSON.stringify(result.body), /123456789012345678/)

  const invalidPath = "/internal/v1/public-alliance-events/guild/9999"
  assert.equal((await handlePublicAllianceEventsRead({ method: "GET", path: invalidPath, headers }, {
    config: { secret, profile: "wos" }, repository: {}, now: () => now
  })).status, 404)
})

test("read endpoint fails closed and configuration is dormant by default", async () => {
  assert.equal(publicAllianceEventsConfig({ GAME_PROFILE: "wos" }).enabled, false)
  assert.equal(publicAllianceEventsConfig({ GAME_PROFILE: "wos", ALLIANCE_EVENTS_READ_ENABLED: "true",
    ALLIANCE_EVENTS_READ_SECRET: secret, ALLIANCE_EVENTS_READ_PORT: "3001" }).enabled, true)
  const path = "/internal/v1/public-alliance-events/9999"
  const headers = signedPublicAllianceEventsHeaders({ secret, profile: "wos", method: "GET", path,
    now: () => now.getTime(), createNonce: () => "abcdefghijklmnop" })
  const unavailable = await handlePublicAllianceEventsRead({ method: "GET", path, headers }, {
    config: { secret, profile: "wos" }, repository: { listForCommunity: async () => { throw new Error("db") } }, now: () => now
  })
  assert.equal(unavailable.status, 503)
})

test("signed native manager endpoint returns only a live scoped decision", async () => {
  const path = "/internal/v1/manager-authorization/guild/300000000000000003/user/200000000000000002"
  const headers = signedPublicAllianceEventsHeaders({ secret, profile: "wos", method: "GET", path,
    now: () => now.getTime(), createNonce: () => "abcdefghijklmnop" })
  let allowed = true
  const verifyManager = async input => {
    assert.deepEqual(input, {
      guildId: "300000000000000003", discordUserId: "200000000000000002"
    })
    return allowed
      ? { status: "authorized", via: "bot_manager_role" }
      : { status: "denied", reason: "insufficient_permissions" }
  }
  let result = await handleNativeManagerAuthorization({ method: "GET", path, headers }, {
    config: { secret, profile: "wos" }, verifyManager, now: () => now
  })
  assert.deepEqual(result, { status: 200,
    body: { ok: true, canManage: true, via: "bot_manager_role" } })
  allowed = false
  result = await handleNativeManagerAuthorization({ method: "GET", path, headers }, {
    config: { secret, profile: "wos" }, verifyManager, now: () => now
  })
  assert.deepEqual(result, { status: 200, body: { ok: true, canManage: false } })
  assert.doesNotMatch(JSON.stringify(result), /role|guild|user/i)
})

test("native manager endpoint rejects cross-profile signatures and controls failures", async () => {
  const path = "/internal/v1/manager-authorization/guild/300000000000000003/user/200000000000000002"
  const headers = signedPublicAllianceEventsHeaders({ secret, profile: "kingshot", method: "GET", path,
    now: () => now.getTime(), createNonce: () => "abcdefghijklmnop" })
  assert.equal((await handleNativeManagerAuthorization({ method: "GET", path, headers }, {
    config: { secret, profile: "wos" }, verifyManager: async () => ({ status: "authorized" }),
    now: () => now
  })).status, 401)
  const wosHeaders = signedPublicAllianceEventsHeaders({ secret, profile: "wos", method: "GET", path,
    now: () => now.getTime(), createNonce: () => "abcdefghijklmnop" })
  const failed = await handleNativeManagerAuthorization({ method: "GET", path, headers: wosHeaders }, {
    config: { secret, profile: "wos" },
    verifyManager: async () => ({ status: "unavailable", reason: "database_unavailable" }),
    now: () => now
  })
  assert.deepEqual(failed, { status: 503,
    body: { ok: false, code: "manager_verification_unavailable" } })
})

test("private listener can serve manager authorization without scheduler availability", async () => {
  let serverInput
  const bootstrap = createPublicAllianceEventsBootstrap({
    initializationPromise: Promise.resolve({ available: false, gameProfile: null }),
    managerInitializationPromise: Promise.resolve({ available: true }),
    client: { guilds: { cache: new Map(), fetch: async () => null } },
    config: { enabled: true, requested: true, profile: "wos", secret,
      port: 3001, secretConfigured: true, disabledReason: null },
    getPoolFn: () => ({ query: async () => ({ rows: [] }) }),
    createServer: input => {
      serverInput = input
      return { start: async () => ({ started: true }), stop: async () => {} }
    },
    processRef: { once() {} },
    logger: { log() {}, error() {} }
  })
  assert.deepEqual(await bootstrap.start(), { started: true })
  assert.equal(serverInput.repository, null)
  assert.equal(typeof serverInput.verifyManager, "function")
})
