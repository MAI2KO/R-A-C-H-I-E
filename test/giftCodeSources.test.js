const test = require("node:test")
const assert = require("node:assert/strict")

const {
  parseSourceExpiry,
  parseDiscordMirrorMessage,
  parseActiveCatalogueHtml
} = require("../src/giftCodes/sourceParsers")
const {
  booleanFlag,
  sourcePollingConfig,
  effectiveSourcePollingConfig
} = require("../src/giftCodes/sourceConfig")
const {
  sourceStartupDiagnostic,
  startCatalogueSourceRuntime
} = require("../src/giftCodes/workflowRuntime")
const {
  CATALOGUES,
  createCatalogueAdapter,
  createCataloguePoller
} = require("../src/giftCodes/catalogueSources")
const { createGiftCodeSourceIngestionService } = require("../src/giftCodes/sourceIngestion")

test("Discord mirror parser recognises explicit labels, preserves case and handles expiry", () => {
  const observedAt = new Date("2026-07-29T12:00:00Z")
  assert.deepEqual(
    parseDiscordMirrorMessage(
      "📌 Code: 47ar5vzKz\n⏰ Valid Until: August 3, 23:59 (UTC+0)\n🎉 Redemption page: example",
      observedAt
    ),
    {
      code: "47ar5vzKz",
      sourceReportedExpiryAt: new Date("2026-08-03T23:59:00Z")
    }
  )
  assert.equal(
    parseDiscordMirrorMessage("Gift code: FB4Million\nValid until: July 30, 23:59 (UTC+0)", observedAt).code,
    "FB4Million"
  )
  assert.deepEqual(parseDiscordMirrorMessage("Redeem Code: `KingShot42`", observedAt), {
    code: "KingShot42",
    sourceReportedExpiryAt: null
  })
  assert.equal(parseDiscordMirrorMessage("Please remember tonight's event", observedAt), null)
  assert.equal(parseDiscordMirrorMessage("Some arbitrary value: ABC123", observedAt), null)
})

test("source expiry chooses a sensible upcoming year without changing canonical state", () => {
  assert.equal(
    parseSourceExpiry("Expires: January 2, 01:30 (GMT)", new Date("2026-12-20T00:00:00Z")).toISOString(),
    "2027-01-02T01:30:00.000Z"
  )
  assert.equal(
    parseSourceExpiry("Expires: January 2, 2025 01:30 (UTC+0)", new Date("2026-12-20T00:00:00Z")).toISOString(),
    "2025-01-02T01:30:00.000Z"
  )
  assert.equal(parseSourceExpiry("Expires: February 30, 23:59 (UTC+0)"), null)
  assert.equal(parseSourceExpiry("Expires: August 3, 25:00 (UTC+0)"), null)
})

test("catalogue parser extracts only explicitly active catalogue entries", () => {
  const wosFixture = `
    <div data-status="active" data-code="WosCase1"></div>
    <article class="gift active-code"><strong>WOS2026</strong></article>
    <div data-status="expired" data-code="OLDWOS"></div>
    <p>RandomCode999</p>`
  const kingshotFixture = `
    <li data-code="KsLive7" data-state="active"></li>
    <li data-code="KsOld7" data-state="expired"></li>`
  assert.deepEqual(parseActiveCatalogueHtml(wosFixture), ["WosCase1", "WOS2026"])
  assert.deepEqual(parseActiveCatalogueHtml(kingshotFixture), ["KsLive7"])
})

test("catalogue adapters use bounded, timed ordinary GET requests and fail closed on markup", async () => {
  const calls = []
  const transport = {
    async get(url, options) {
      calls.push({ url, options })
      return { data: '<div data-status="active" data-code="ExactCase"></div>' }
    }
  }
  const adapter = createCatalogueAdapter({
    gameProfile: "wos", transport, timeoutMs: 4321, maximumBodyBytes: 2048
  })
  assert.deepEqual(await adapter.fetchActiveCodes(), ["ExactCase"])
  assert.equal(calls[0].url, CATALOGUES.wos.url)
  assert.equal(calls[0].options.timeout, 4321)
  assert.equal(calls[0].options.maxContentLength, 2048)
  assert.match(calls[0].options.headers["User-Agent"], /GiftCodeDiscovery/)

  const malformed = createCatalogueAdapter({
    gameProfile: "kingshot",
    transport: { async get() { return { data: "<html>changed</html>" } } }
  })
  await assert.rejects(malformed.fetchActiveCodes(), { code: "SOURCE_MARKUP_UNRECOGNISED" })

  const oversized = createCatalogueAdapter({
    gameProfile: "kingshot",
    maximumBodyBytes: 10,
    transport: { async get() { return { data: "this response is too large" } } }
  })
  await assert.rejects(oversized.fetchActiveCodes(), { code: "SOURCE_BODY_TOO_LARGE" })
})

test("polling flags default off and enable profiles independently", () => {
  assert.deepEqual(sourcePollingConfig({}), {
    pollingEnabled: false,
    wosEnabled: false,
    kingshotEnabled: false,
    intervalMs: 900000,
    timeoutMs: 10000,
    maximumBodyBytes: 1048576
  })
  const wosOnly = sourcePollingConfig({
    GIFT_CODE_SOURCE_POLLING_ENABLED: "true",
    WOS_REWARDS_SOURCE_ENABLED: "true"
  })
  assert.equal(wosOnly.pollingEnabled && wosOnly.wosEnabled, true)
  assert.equal(wosOnly.pollingEnabled && wosOnly.kingshotEnabled, false)
})

test("effective source polling uses the global gate and only the matching profile gate", () => {
  const cases = [
    [{ GIFT_CODE_SOURCE_POLLING_ENABLED: "true", WOS_REWARDS_SOURCE_ENABLED: "true" }, "wos", true],
    [{ GIFT_CODE_SOURCE_POLLING_ENABLED: "true", WOS_REWARDS_SOURCE_ENABLED: "false" }, "wos", false],
    [{ GIFT_CODE_SOURCE_POLLING_ENABLED: "false", WOS_REWARDS_SOURCE_ENABLED: "true" }, "wos", false],
    [{
      GIFT_CODE_SOURCE_POLLING_ENABLED: "true",
      WOS_REWARDS_SOURCE_ENABLED: "true",
      KINGSHOT_REWARDS_SOURCE_ENABLED: "false"
    }, "wos", true],
    [{
      GIFT_CODE_SOURCE_POLLING_ENABLED: "true",
      WOS_REWARDS_SOURCE_ENABLED: "false",
      KINGSHOT_REWARDS_SOURCE_ENABLED: "true"
    }, "kingshot", true],
    [{
      GIFT_CODE_SOURCE_POLLING_ENABLED: "true",
      WOS_REWARDS_SOURCE_ENABLED: "true",
      KINGSHOT_REWARDS_SOURCE_ENABLED: "false"
    }, "kingshot", false]
  ]
  for (const [env, profile, expected] of cases) {
    assert.equal(effectiveSourcePollingConfig(profile, env).publicCatalogueEnabled, expected)
  }
  assert.equal(booleanFlag("true"), true)
  assert.equal(booleanFlag(" TRUE "), true)
  assert.equal(booleanFlag("1"), true)
  assert.equal(booleanFlag("false"), false)
  assert.equal(booleanFlag("0"), false)
  assert.equal(booleanFlag("yes"), false)
})

test("source runtime logs effective startup state and starts one poller exactly once", () => {
  const logs = []
  let pollerConstructions = 0
  let workerConstructions = 0
  let starts = 0
  const runtime = startCatalogueSourceRuntime({
    gameProfile: "wos",
    env: {
      GIFT_CODE_SOURCE_POLLING_ENABLED: "true",
      WOS_REWARDS_SOURCE_ENABLED: "true",
      GIFT_CODE_SOURCE_POLL_INTERVAL_SECONDS: "900"
    },
    sourceRepository: {},
    sourceIngestion: {},
    logger: { log(value) { logs.push(value) }, error(value) { logs.push(value) } },
    adapterFactory: input => ({ name: "fixture", ...input }),
    pollerFactory: input => {
      pollerConstructions += 1
      assert.equal(input.enabled, true)
      return { enabled: true, async poll() {} }
    },
    workerFactory: input => {
      workerConstructions += 1
      assert.equal(input.enabled, true)
      return {
        start() { starts += 1; return { started: true } },
        isRunning() { return true },
        async stop() {}
      }
    }
  })
  assert.equal(runtime.config.publicCatalogueEnabled, true)
  assert.equal(runtime.start.started, true)
  assert.equal(pollerConstructions, 1)
  assert.equal(workerConstructions, 1)
  assert.equal(starts, 1)
  assert.deepEqual(logs, [
    "[Gift code sources] wos: polling=true, profile_source=true, " +
      "public_catalogue=true, subsystem=true, interval=900s"
  ])
  assert.equal(sourceStartupDiagnostic(runtime.config, false),
    "[Gift code sources] wos: polling=true, profile_source=true, " +
    "public_catalogue=true, subsystem=false, interval=900s")
})

test("source poller startup failure is contained and logged without environment data", () => {
  const errors = []
  const runtime = startCatalogueSourceRuntime({
    gameProfile: "wos",
    env: {
      GIFT_CODE_SOURCE_POLLING_ENABLED: "true",
      WOS_REWARDS_SOURCE_ENABLED: "true"
    },
    sourceRepository: {},
    sourceIngestion: {},
    logger: { log() {}, error(value) { errors.push(JSON.parse(value)) } },
    adapterFactory: () => ({}),
    pollerFactory: () => ({ async poll() {} }),
    workerFactory: () => ({
      start() { throw Object.assign(new Error("contains environment details"), { code: "START_FAILED" }) }
    })
  })
  assert.equal(runtime.start.started, false)
  assert.equal(runtime.error, "START_FAILED")
  assert.deepEqual(errors, [{
    event: "gift_code_source_poller_start_failed",
    game_profile: "wos",
    error_code: "START_FAILED"
  }])
})

test("catalogue poller coalesces concurrent runs, records duplicates and contains failures", async () => {
  let release
  let fetches = 0
  const observed = []
  const health = []
  const sourceRepository = {
    async ensureSource() { return { id: "source-1" } },
    async markCataloguePoll(id, result) { health.push({ id, ...result }) }
  }
  const ingestion = {
    async ingest(input) {
      observed.push(input)
      return { duplicate: input.code === "KNOWN" }
    }
  }
  const adapter = {
    name: "WOSRewards",
    url: CATALOGUES.wos.url,
    async fetchActiveCodes() {
      fetches += 1
      await new Promise(resolve => { release = resolve })
      return ["NEW", "KNOWN"]
    }
  }
  const poller = createCataloguePoller({
    gameProfile: "wos", sourceRepository, ingestion, adapter, enabled: true
  })
  const first = poller.poll()
  const second = poller.poll()
  await new Promise(resolve => setImmediate(resolve))
  release()
  assert.deepEqual(await Promise.all([first, second]), [
    { polled: true, observed: 2, candidates: 1 },
    { polled: true, observed: 2, candidates: 1 }
  ])
  assert.equal(fetches, 1)
  assert.equal(observed.length, 2)
  assert.deepEqual(health[0].observedCodes, ["NEW", "KNOWN"])

  const warnings = []
  const failed = createCataloguePoller({
    gameProfile: "kingshot",
    sourceRepository,
    ingestion,
    adapter: {
      name: "KingshotRewards",
      url: CATALOGUES.kingshot.url,
      async fetchActiveCodes() { throw Object.assign(new Error("timeout with secret body"), { code: "ETIMEDOUT" }) }
    },
    enabled: true,
    logger: { warn(value) { warnings.push(JSON.parse(value)) } }
  })
  assert.deepEqual(await failed.poll(), { polled: true, failed: true, errorCode: "ETIMEDOUT" })
  assert.equal(warnings[0].error_code, "ETIMEDOUT")
  assert.equal(JSON.stringify(warnings).includes("secret body"), false)

  const unavailableSource = createCataloguePoller({
    gameProfile: "wos",
    sourceRepository: {
      async ensureSource() { throw Object.assign(new Error("database unavailable"), { code: "08006" }) }
    },
    ingestion,
    adapter,
    enabled: true,
    logger: { warn(value) { warnings.push(JSON.parse(value)) } }
  })
  assert.deepEqual(await unavailableSource.poll(), {
    polled: true,
    failed: true,
    errorCode: "08006"
  })
})

test("Discord source ingestion accepts webhooks, stores scoped provenance and ignores replays", async () => {
  const submissions = []
  const observations = []
  let prior = null
  const sourceRepository = {
    async discordChannel(channelId, guildId) {
      assert.equal(channelId, "22")
      assert.equal(guildId, "11")
      return { source_name: "Official mirror", require_webhook: true }
    },
    async ensureSource(source) { return { id: "source-1", ...source } },
    async observation() { return prior },
    async recordObservation(input) { observations.push(input); prior = { id: "code-1", code: input.code } }
  }
  const ingestion = createGiftCodeSourceIngestionService({
    giftRepository: {
      async recordSubmission(input) {
        submissions.push(input)
        return { giftCode: { id: "code-1", code: input.code }, submission: {}, duplicate: false }
      }
    },
    sourceRepository,
    gameProfile: "wos"
  })
  const message = {
    guildId: "11",
    channelId: "22",
    id: "33",
    webhookId: "44",
    content: "Code: MiXeD123",
    createdTimestamp: Date.parse("2026-08-01T12:00:00Z"),
    author: { username: "Official WOS" }
  }
  const first = await ingestion.ingestDiscordMessage(message)
  const replay = await ingestion.ingestDiscordMessage(message)
  assert.equal(first.duplicate, false)
  assert.equal(replay.duplicateObservation, true)
  assert.equal(submissions.length, 1)
  assert.equal(submissions[0].submittedByDiscordUserId, null)
  assert.deepEqual(observations[0].provenance, {
    transport: "discord",
    guildId: "11",
    channelId: "22",
    messageId: "33",
    webhookId: "44",
    sourceDisplayName: "Official WOS"
  })
})

test("missing, inaccessible and failing Discord source configuration is contained", async () => {
  const warnings = []
  const base = {
    giftRepository: { async recordSubmission() { throw new Error("must not submit") } },
    gameProfile: "kingshot",
    logger: { warn(value) { warnings.push(JSON.parse(value)) } }
  }
  const missing = createGiftCodeSourceIngestionService({
    ...base,
    sourceRepository: { async discordChannel() { return null } }
  })
  assert.equal(await missing.ingestDiscordMessage({ guildId: "1", channelId: "2", id: "3" }), null)
  const failed = createGiftCodeSourceIngestionService({
    ...base,
    sourceRepository: { async discordChannel() { throw Object.assign(new Error("database detail"), { code: "08006" }) } }
  })
  assert.equal(await failed.ingestDiscordMessage({ guildId: "1", channelId: "2", id: "3" }), null)
  assert.equal(warnings[0].error_code, "08006")
  assert.equal(JSON.stringify(warnings).includes("database detail"), false)
})
