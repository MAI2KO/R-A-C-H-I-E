const test = require("node:test")
const assert = require("node:assert/strict")
const { MessageFlags } = require("discord.js")

const { profileTerminology } = require("../src/giftCodes/terminology")
const { buildPlayerRegisterCommand, getPlayerCommandData } = require("../src/giftCodes/discord/commands")
const { handlePlayerInteraction } = require("../src/giftCodes/discord/interactions")
const { createPlayerService, PlayerAccountError } = require("../src/giftCodes/playerService")
const { signingMaterial, signRequestFields } = require("../src/giftCodes/signing")
const { centuryAdapter } = require("../src/giftCodes/adapters")
const { classifyCenturyResponse } = require("../src/giftCodes/responseClassifier")
const {
  ConservativeRateLimiter,
  retryAfterMilliseconds
} = require("../src/giftCodes/rateLimiter")
const {
  createCenturyGameClient,
  responseDiagnostics
} = require("../src/giftCodes/centuryGameClient")
const { postgresIsEnabled } = require("../src/db")

test("profile terminology consistently distinguishes State and Kingdom", () => {
  assert.deepEqual(
    [profileTerminology("wos").locationLabel, profileTerminology("kingshot").locationLabel],
    ["State", "Kingdom"]
  )
  assert.equal(profileTerminology("wos").gameName, "Whiteout Survival")
  assert.equal(profileTerminology("kingshot").gameName, "Kingshot")
  assert.throws(() => profileTerminology("other"), /Unsupported/)
})

test("player registration panel command is profile aware and feature gated", () => {
  const wos = buildPlayerRegisterCommand("wos").toJSON()
  const kingshot = buildPlayerRegisterCommand("kingshot").toJSON()
  assert.equal(wos.name, "player-register")
  assert.match(wos.description, /Whiteout Survival/)
  assert.match(kingshot.description, /Kingshot/)
  assert.equal(getPlayerCommandData({ PLAYER_GIFT_CODES_ENABLED: "false" }), null)
  assert.equal(getPlayerCommandData({
    PLAYER_GIFT_CODES_ENABLED: "true",
    GAME_PROFILE: "invalid"
  }), null)
  assert.equal(postgresIsEnabled({ PLAYER_GIFT_CODES_ENABLED: "true" }), true)
  assert.equal(postgresIsEnabled({ EVENT_SCHEDULER_ENABLED: "true" }), true)
  assert.equal(postgresIsEnabled({}), false)
})

test("player service validates ownership inputs and does not roll back after mirror failure", async () => {
  const calls = []
  const account = {
    id: "account",
    player_id: "12345",
    state_or_kingdom_number: "689",
    is_primary: true
  }
  const service = createPlayerService({
    gameProfile: "wos",
    repository: {
      async registerAccount(input) { calls.push(input); return account },
      async getOwnedAccount(owner, playerId) { calls.push({ owner, playerId }); return account },
      async listOwnedAccounts(owner) { calls.push({ owner }); return [account] },
      async updateLocation() { return null },
      async deactivateAccount() { return null }
    },
    mirror: { async mirrorRegistration() { throw Object.assign(new Error("offline"), { code: "OFFLINE" }) } },
    logger: { warn: message => calls.push(message) }
  })
  assert.equal((await service.register({
    discordUserId: "999",
    playerId: " 12345 ",
    locationNumber: " 689 "
  })).id, "account")
  assert.deepEqual(calls[0], {
    discordUserId: "999",
    playerId: "12345",
    locationNumber: "689",
    guildId: null
  })
  assert.match(calls[1], /Optional mirror failed/)
  assert.equal((await service.view({ discordUserId: "999", playerId: "12345" })).length, 1)
  await assert.rejects(
    service.changeLocation({ discordUserId: "999", playerId: "12345", locationNumber: "700" }),
    PlayerAccountError
  )
})

function playerInteraction(gameProfile, subcommand, values = {}) {
  return {
    commandName: "player",
    user: { id: "999" },
    deferred: false,
    isChatInputCommand: () => true,
    options: {
      getSubcommand: () => subcommand,
      getString: name => values[name] ?? null
    },
    async deferReply(options) {
      assert.equal(options.flags, MessageFlags.Ephemeral)
      this.deferred = true
    },
    async editReply(payload) {
      assert.equal(this.deferred, true)
      this.edited = payload
    },
    health: { available: true, gameProfile }
  }
}

test("player interactions remain private and use correct profile language", async () => {
  for (const [gameProfile, locationOption, expected] of [
    ["wos", "state", /State has been changed from 689 to 700/],
    ["kingshot", "kingdom", /Kingdom has been changed from 521 to 540/]
  ]) {
    const values = gameProfile === "wos"
      ? { player_id: "123", [locationOption]: "700" }
      : { player_id: "123", [locationOption]: "540" }
    const interaction = playerInteraction(gameProfile, "location", values)
    await handlePlayerInteraction(interaction, {
      healthProvider: () => interaction.health,
      poolProvider: () => ({}),
      repositoryFactory: () => ({}),
      serviceFactory: ({ gameProfile }) => ({
        terms: profileTerminology(gameProfile),
        async changeLocation() {
          return {
            changed: true,
            previousNumber: gameProfile === "wos" ? "689" : "521",
            account: { state_or_kingdom_number: values[locationOption] }
          }
        }
      })
    })
    assert.match(interaction.edited, expected)
    assert.doesNotMatch(interaction.edited, /state_or_kingdom/)
  }
})

test("player command degrades without touching repository storage", async () => {
  let repositoryCreated = false
  const interaction = playerInteraction("wos", "view")
  await handlePlayerInteraction(interaction, {
    healthProvider: () => ({ available: false, gameProfile: "wos" }),
    repositoryFactory: () => { repositoryCreated = true }
  })
  assert.equal(repositoryCreated, false)
  assert.match(interaction.edited, /temporarily unavailable/)
})

test("signing is deterministic, sorted and preserves gift-code case", () => {
  const fields = { fid: "123", cdk: "GiftCode", kid: "689", time: "1700000000000" }
  assert.equal(
    signingMaterial(fields),
    "cdk=GiftCode&fid=123&kid=689&time=1700000000000"
  )
  assert.equal(signRequestFields(fields, "test-suffix"), "c84cd3058471383a8dddb115594d92d6")
  assert.notEqual(
    signRequestFields(fields, "test-suffix"),
    signRequestFields({ ...fields, cdk: "giftcode" }, "test-suffix")
  )
})

test("Century adapters use current public-client defaults and environment overrides", () => {
  const defaultWos = centuryAdapter("wos", {})
  const defaultKingshot = centuryAdapter("kingshot", {})
  const wos = centuryAdapter("wos", { CENTURY_WOS_SIGNING_SUFFIX: "wos-override" })
  const kingshot = centuryAdapter("kingshot", { CENTURY_KINGSHOT_SIGNING_SUFFIX: "ks-override" })
  assert.equal(wos.apiBaseUrl, "https://wos-giftcode-api.centurygame.com/api")
  assert.equal(kingshot.apiBaseUrl, "https://kingshot-giftcode.centurygame.com/api")
  assert.equal(defaultWos.signingSuffix, "tB87#kPtkxqOS2")
  assert.equal(defaultKingshot.signingSuffix, "mN4!pQs6JrYwV9")
  assert.equal(wos.signingSuffix, "wos-override")
  assert.equal(kingshot.signingSuffix, "ks-override")
})

test("Whiteout Survival signing matches a validated official request", () => {
  const fields = {
    cdk: "WOS0804",
    fid: "282021376",
    kid: "689",
    time: "1786379187"
  }
  const adapter = centuryAdapter("wos", {})
  assert.equal(signingMaterial(fields), "cdk=WOS0804&fid=282021376&kid=689&time=1786379187")
  assert.equal(signRequestFields(fields, adapter.signingSuffix), "e36b7c41cb61ba90e8826fa9f75f6165")
})

test("Kingshot signing matches a validated official request", () => {
  const fields = {
    cdk: "KS0810",
    fid: "368775177",
    kid: "521",
    time: "1786380971"
  }
  const adapter = centuryAdapter("kingshot", {})
  assert.equal(signingMaterial(fields), "cdk=KS0810&fid=368775177&kid=521&time=1786380971")
  assert.equal(signRequestFields(fields, adapter.signingSuffix), "25248dee52bf7b2926624f16106b3989")
})

test("Century responses classify known, temporary, rate-limited and unknown results", () => {
  assert.equal(classifyCenturyResponse({
    httpStatus: 200,
    data: { code: 0, err_code: 20000, msg: "SUCCESS" }
  }).state, "success")
  assert.equal(classifyCenturyResponse({
    httpStatus: 200,
    data: { code: 1, err_code: 40008, msg: "RECEIVED." },
    profileMappings: centuryAdapter("wos", {}).responseMappings
  }).state, "already_redeemed")
  assert.equal(classifyCenturyResponse({ httpStatus: 429, data: {} }).state, "rate_limited")
  assert.equal(classifyCenturyResponse({ httpStatus: 503, data: {} }).state, "temporary_error")
  assert.equal(classifyCenturyResponse({ httpStatus: 403, data: {} }).state, "upstream_rejection")
  assert.equal(classifyCenturyResponse({
    httpStatus: 403,
    data: { err_code: 40008, msg: "RECEIVED." },
    profileMappings: centuryAdapter("wos", {}).responseMappings
  }).state, "already_redeemed")
  const unknown = classifyCenturyResponse({
    httpStatus: 200,
    data: { code: 9, err_code: 49999, msg: "NEW RESPONSE" }
  })
  assert.equal(unknown.state, "unknown_response")
  assert.deepEqual(unknown.raw, { code: 9, errCode: 49999, message: "NEW RESPONSE" })
})

test("conservative limiter serializes work, respects Retry-After and avoids permanent retries", async () => {
  let now = 0
  const sleeps = []
  const starts = []
  const limiter = new ConservativeRateLimiter({
    gameProfile: "wos",
    minimumDelayMs: 50,
    maximumRetries: 2,
    baseBackoffMs: 100,
    maximumBackoffMs: 5000,
    now: () => now,
    sleep: async delay => { sleeps.push(delay); now += delay }
  })
  let firstAttempts = 0
  const first = limiter.schedule(async () => {
    starts.push(`first:${now}`)
    firstAttempts += 1
    return firstAttempts === 1
      ? { httpStatus: 429, headers: { "retry-after": "2", "x-ratelimit-limit": "30" }, retryable: true }
      : { httpStatus: 200, headers: { "x-ratelimit-remaining": "29" }, retryable: false }
  })
  const second = limiter.schedule(async () => {
    starts.push(`second:${now}`)
    return { httpStatus: 200, headers: {}, retryable: false }
  })
  assert.equal((await first).attempts, 2)
  assert.equal((await second).attempts, 1)
  assert.deepEqual(starts, ["first:0", "first:2000", "second:2050"])
  assert.ok(sleeps.includes(2000))
  assert.equal(limiter.getObservations().length, 3)

  let permanentAttempts = 0
  const permanent = await limiter.schedule(async () => {
    permanentAttempts += 1
    return { httpStatus: 200, headers: {}, retryable: false }
  })
  assert.equal(permanent.attempts, 1)
  assert.equal(permanentAttempts, 1)
  assert.equal(retryAfterMilliseconds({ "retry-after": "3" }, 0), 3000)
})

test("Century client sends one URL-encoded signed request through injected transport", async () => {
  let request
  const limiter = new ConservativeRateLimiter({
    gameProfile: "wos",
    minimumDelayMs: 0,
    maximumRetries: 0
  })
  const client = createCenturyGameClient({
    gameProfile: "wos",
    adapter: {
      frontendUrl: "https://wos-giftcode.centurygame.com",
      apiBaseUrl: "https://wos-giftcode-api.centurygame.com/api",
      redemptionPath: "/gift_code",
      signingSuffix: "test-suffix",
      responseMappings: {}
    },
    limiter,
    now: () => 1700000000000,
    transport: {
      async post(url, body, options) {
        request = { url, body, options }
        return {
          status: 200,
          headers: { "x-ratelimit-limit": "30", "x-ratelimit-remaining": "29" },
          data: { code: 0, err_code: 20000, msg: "SUCCESS" }
        }
      }
    }
  })
  const result = await client.redeem({ playerId: "123", code: "GiftCode", locationNumber: "689" })
  assert.equal(request.url, "https://wos-giftcode-api.centurygame.com/api/gift_code")
  assert.equal(request.options.headers["Content-Type"], "application/x-www-form-urlencoded")
  assert.equal(request.options.headers.Accept, "application/json, text/plain, */*")
  assert.equal(request.options.headers.Origin, "https://wos-giftcode.centurygame.com")
  assert.equal(request.options.headers.Referer, "https://wos-giftcode.centurygame.com/")
  assert.equal(request.options.headers["User-Agent"], "R.A.C.H.I.E gift-code client")
  assert.equal(request.options.headers.Authorization, undefined)
  assert.equal(request.options.headers.Cookie, undefined)
  assert.equal(request.options.maxRedirects, 5)
  assert.equal(request.options.decompress, true)
  const form = new URLSearchParams(request.body)
  assert.deepEqual([...form.keys()].sort(), ["cdk", "fid", "kid", "sign", "time"])
  assert.equal(form.get("fid"), "123")
  assert.equal(form.get("cdk"), "GiftCode")
  assert.equal(form.get("kid"), "689")
  assert.equal(form.get("time"), "1700000000")
  assert.equal(form.get("sign"), "15cba9d17960a3c42fe6366ac6ba0663")
  assert.equal(result.classification.state, "success")
})

test("403 HTML and JSON edge responses remain reviewable upstream rejections", async () => {
  for (const [expectedType, response] of [
    ["html", {
      status: 403,
      headers: { "content-type": "text/html", server: "cloud-edge", "cf-ray": "ray-id" },
      data: "<!doctype html><html><body>Access denied</body></html>"
    }],
    ["json", {
      status: 403,
      headers: { "content-type": "application/json", server: "nginx" },
      data: { error: "Forbidden" }
    }]
  ]) {
    const client = createCenturyGameClient({
      gameProfile: "wos",
      adapter: centuryAdapter("wos", {}),
      now: () => 1786380000000,
      limiter: new ConservativeRateLimiter({ gameProfile: "wos", minimumDelayMs: 0, maximumRetries: 0 }),
      transport: { async post() { return response } }
    })
    const result = await client.redeem({ playerId: "282021376", code: "gogoWOS", locationNumber: "689" })
    assert.equal(result.classification.state, "upstream_rejection")
    assert.equal(result.classification.raw.errCode, null)
    assert.equal(result.responseDiagnostics.responseType, expectedType)
  }
})

test("edge diagnostics are bounded, allowlisted and redact request material", () => {
  const diagnostics = responseDiagnostics(
    `<html><script>token=script-secret</script><body>fid=282021376 sign=abc123 ${"x".repeat(3000)}</body></html>`,
    {
      "content-type": "text/html; charset=utf-8",
      server: "cloud-edge",
      "cf-ray": "safe-ray",
      "x-ratelimit-remaining": "4",
      "set-cookie": "session=never-store"
    },
    1024,
    ["282021376", "abc123", "public-client-suffix"]
  )
  assert.equal(diagnostics.responseType, "html")
  assert.equal(diagnostics.bodySummary.length, 1024)
  assert.equal(diagnostics.bodyTruncated, true)
  assert.deepEqual(diagnostics.edgeHeaders, { "cf-ray": "safe-ray" })
  assert.doesNotMatch(JSON.stringify(diagnostics), /282021376|abc123|script-secret|never-store|set-cookie/)
})
