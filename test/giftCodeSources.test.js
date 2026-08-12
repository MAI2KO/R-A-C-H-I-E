const test = require("node:test")
const assert = require("node:assert/strict")
const fs = require("node:fs")
const path = require("node:path")
const { Collection, MessageReferenceType } = require("discord.js")

const {
  parseSourceExpiry,
  parseDiscordMirrorMessage,
  parseActiveCatalogueHtml,
  parseWosActiveCatalogueHtml
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
const {
  parseDiscordSourceMessage,
  createGiftCodeSourceIngestionService
} = require("../src/giftCodes/sourceIngestion")

function forwardedMessage({
  id = "forward-1",
  content = "",
  embeds = [],
  createdTimestamp = Date.parse("2026-08-12T12:00:00Z")
} = {}) {
  return {
    guildId: "11",
    channelId: "22",
    id,
    webhookId: "44",
    content: "Outer forwarding commentary is not source content",
    createdTimestamp,
    author: { username: "Mirror relay" },
    reference: {
      type: MessageReferenceType.Forward,
      guildId: "source-guild",
      channelId: "source-channel",
      messageId: "source-message"
    },
    messageSnapshots: new Collection([
      ["source-message", { content, embeds }]
    ])
  }
}

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

test("forwarded Discord snapshots reuse labelled parsing for content and useful embed text", () => {
  const observedAt = new Date("2026-07-29T12:00:00Z")
  assert.deepEqual(parseDiscordSourceMessage(forwardedMessage({
    content: "📌 Code: 47ar5vzKz\n⏰ Valid Until: August 3, 23:59 (UTC+0)"
  }), observedAt), {
    code: "47ar5vzKz",
    sourceReportedExpiryAt: new Date("2026-08-03T23:59:00Z"),
    sourceMessageKind: "forwarded_snapshot"
  })
  assert.deepEqual(parseDiscordSourceMessage(forwardedMessage({
    content: "Gift Code: ExactCase7"
  }), observedAt), {
    code: "ExactCase7",
    sourceReportedExpiryAt: null,
    sourceMessageKind: "forwarded_snapshot"
  })
  assert.deepEqual(parseDiscordSourceMessage(forwardedMessage({
    embeds: [{
      title: "Official announcement",
      description: "A new reward is available.",
      fields: [
        { name: "Redeem Code", value: "EmbedCase9" },
        { name: "Valid Until", value: "August 4, 23:59 (UTC+0)" }
      ],
      url: "https://example.invalid/RandomUrlCode999",
      footer: { text: "FooterCode888" }
    }]
  }), observedAt), {
    code: "EmbedCase9",
    sourceReportedExpiryAt: new Date("2026-08-04T23:59:00Z"),
    sourceMessageKind: "forwarded_snapshot"
  })
})

test("empty, malformed and arbitrary forwarded snapshots are ignored", () => {
  const observedAt = new Date("2026-08-12T12:00:00Z")
  assert.equal(parseDiscordSourceMessage(forwardedMessage(), observedAt), null)
  assert.equal(parseDiscordSourceMessage(forwardedMessage({
    content: "Tonight's event starts at 20:00. RandomCode999"
  }), observedAt), null)
  assert.equal(parseDiscordSourceMessage({
    ...forwardedMessage({ content: "Code: ShouldNotParse" }),
    messageSnapshots: []
  }, observedAt), null)
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

test("current WOS catalogue section extracts active cards only and preserves exact case", () => {
  const fixture = fs.readFileSync(
    path.join(__dirname, "fixtures", "wosRewardsActiveCatalogue.html"),
    "utf8"
  )
  assert.deepEqual(parseWosActiveCatalogueHtml(fixture), [
    "100gomYTKOR",
    "GuDokYTKOR",
    "2ndYoutubeKR"
  ])
  for (const excluded of ["GIFT2026", "ExpiredCode7", "FooterCode999"]) {
    assert.equal(parseWosActiveCatalogueHtml(fixture).includes(excluded), false)
  }
})

test("WOS catalogue parsing ignores arbitrary text and fails closed on structural drift", () => {
  assert.deepEqual(parseWosActiveCatalogueHtml(`
    <h2>Latest Whiteout Survival Codes</h2>
    <p>RandomCode999</p><p>ABC123 Added yesterday Copy</p>
  `), [])
  assert.deepEqual(parseWosActiveCatalogueHtml(`
    <h2>Active Gift Codes</h2>
    <p>RandomCode999</p>
    <article><strong>AlmostCode1</strong><span>Available now</span><button>Copy</button></article>
  `), [])
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

  const malformedWos = createCatalogueAdapter({
    gameProfile: "wos",
    transport: {
      async get() {
        return {
          status: 200,
          headers: { "content-type": "text/html; charset=utf-8" },
          data: "<html><h2>Changed Catalogue</h2></html>"
        }
      }
    }
  })
  await assert.rejects(malformedWos.fetchActiveCodes(), error => {
    assert.equal(error.code, "SOURCE_MARKUP_UNRECOGNISED")
    assert.deepEqual(error.sourceDiagnostics, {
      httpStatus: 200,
      contentType: "text/html; charset=utf-8",
      responseBytes: 39,
      expectedStructure: "active_gift_codes_section"
    })
    assert.equal(error.message.includes("<html>"), false)
    return true
  })

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

test("forwarded source observations retain outer provenance and deduplicate canonical work", async () => {
  const canonical = new Map()
  const observations = new Map()
  const recordedObservations = []
  const health = { lastObservation: null, lastCandidate: null }
  let submissions = 0
  let contributorRewards = 0
  const sourceRepository = {
    async discordChannel(channelId, guildId) {
      return channelId === "22" && guildId === "11"
        ? { source_name: "Official mirror", require_webhook: false }
        : null
    },
    async ensureSource(source) { return { id: "source-1", ...source } },
    async observation(_sourceId, observationKey) { return observations.get(observationKey) || null },
    async recordObservation(input) {
      recordedObservations.push(input)
      observations.set(input.observationKey, canonical.get(input.code))
      health.lastObservation = input.observedAt
      if (input.candidateCreated) health.lastCandidate = input.observedAt
    }
  }
  const ingestion = createGiftCodeSourceIngestionService({
    giftRepository: {
      async recordSubmission(input) {
        submissions += 1
        const existing = canonical.get(input.code)
        if (existing) return { giftCode: existing, submission: {}, duplicate: true }
        const giftCode = { id: `code-${canonical.size + 1}`, code: input.code }
        canonical.set(input.code, giftCode)
        if (input.submittedByDiscordUserId) contributorRewards += 1
        return { giftCode, submission: {}, duplicate: false }
      }
    },
    sourceRepository,
    gameProfile: "wos"
  })

  const first = await ingestion.ingestDiscordMessage(forwardedMessage({
    id: "forward-1",
    content: "Code: ForwardCase7",
    createdTimestamp: Date.parse("2026-08-12T12:00:00Z")
  }))
  const duplicate = await ingestion.ingestDiscordMessage(forwardedMessage({
    id: "forward-2",
    content: "Gift Code: ForwardCase7",
    createdTimestamp: Date.parse("2026-08-12T12:05:00Z")
  }))

  assert.equal(first.duplicate, false)
  assert.equal(duplicate.duplicate, true)
  assert.equal(canonical.size, 1)
  assert.equal(submissions, 2)
  assert.equal(contributorRewards, 0)
  assert.equal(recordedObservations[0].candidateCreated, true)
  assert.equal(recordedObservations[1].candidateCreated, false)
  assert.equal(health.lastObservation.toISOString(), "2026-08-12T12:05:00.000Z")
  assert.equal(health.lastCandidate.toISOString(), "2026-08-12T12:00:00.000Z")
  assert.deepEqual(recordedObservations[0].provenance, {
    transport: "discord",
    guildId: "11",
    channelId: "22",
    messageId: "forward-1",
    webhookId: "44",
    sourceDisplayName: "Mirror relay",
    sourceMessageKind: "forwarded_snapshot",
    forwardedSourceGuildId: "source-guild",
    forwardedSourceChannelId: "source-channel",
    forwardedSourceMessageId: "source-message"
  })
})

test("forwarded source ingestion remains profile isolated", async () => {
  function profileHarness(profile) {
    const codes = new Map()
    const ingestion = createGiftCodeSourceIngestionService({
      gameProfile: profile,
      sourceRepository: {
        async discordChannel() { return { source_name: `${profile} mirror`, require_webhook: false } },
        async ensureSource(source) { return { id: `${profile}-source`, ...source } },
        async observation() { return null },
        async recordObservation() {}
      },
      giftRepository: {
        async recordSubmission(input) {
          const giftCode = { id: `${profile}-code`, code: input.code, game_profile: profile }
          codes.set(input.code, giftCode)
          return { giftCode, submission: {}, duplicate: false }
        }
      }
    })
    return { ingestion, codes }
  }
  const wos = profileHarness("wos")
  const kingshot = profileHarness("kingshot")
  await wos.ingestion.ingestDiscordMessage(forwardedMessage({ content: "Code: SharedCase7" }))
  await kingshot.ingestion.ingestDiscordMessage(forwardedMessage({ content: "Code: SharedCase7" }))
  assert.equal(wos.codes.get("SharedCase7").game_profile, "wos")
  assert.equal(kingshot.codes.get("SharedCase7").game_profile, "kingshot")
  assert.notEqual(wos.codes.get("SharedCase7").id, kingshot.codes.get("SharedCase7").id)
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
