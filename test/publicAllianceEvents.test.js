const test = require("node:test")
const assert = require("node:assert/strict")

const { createPublicAllianceEventRepository } = require("../src/publicAllianceEventRepository")
const { publicAllianceEventsReadModel } = require("../src/publicAllianceEvents")
const {
  signedPublicAllianceEventsHeaders,
  verifyPublicAllianceEventsRequest
} = require("../src/publicAllianceEventsAuth")
const { handlePublicAllianceEventsRead } = require("../src/publicAllianceEventsServer")
const { publicAllianceEventsConfig } = require("../src/publicAllianceEventsBootstrap")

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
