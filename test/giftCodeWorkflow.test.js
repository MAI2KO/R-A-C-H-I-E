const test = require("node:test")
const assert = require("node:assert/strict")

const { centuryAdapter } = require("../src/giftCodes/adapters")
const { boundedResponseData, createCenturyGameClient } = require("../src/giftCodes/centuryGameClient")
const { classifyCenturyResponse } = require("../src/giftCodes/responseClassifier")
const { ConservativeRateLimiter } = require("../src/giftCodes/rateLimiter")
const {
  getVerifierAccount,
  verificationWorkerIsEnabled,
  redemptionWorkerIsEnabled,
  giftWorkerConfig
} = require("../src/giftCodes/config")
const {
  verificationTransition,
  redemptionTransition,
  createVerificationProcessor,
  createRedemptionProcessor,
  createPollingWorker
} = require("../src/giftCodes/workers")
const {
  createGiftCodeCommunityService,
  verificationResultMessage
} = require("../src/giftCodes/communityService")
const { createGiftCodeService } = require("../src/giftCodes/service")
const { notificationMessage, createDiscordGiftNotifier } = require("../src/giftCodes/notifications")
const {
  buildGiftCodesCommand,
  buildGiftCodesAdminCommand,
  getGiftCommandData
} = require("../src/giftCodes/discord/commands")
const {
  formatPlayerGiftStatus,
  formatCodeDiagnostics,
  handleGiftInteraction
} = require("../src/giftCodes/discord/giftInteractions")
const { machineFields } = require("../src/giftCodes/repository")
const { runControlledCenturyVerification } = require("../scripts/controlledCenturyVerification")
const { profileTerminology } = require("../src/giftCodes/terminology")

const config = {
  leaseSeconds: 60,
  maximumAttempts: 3,
  retryBaseSeconds: 10,
  retryCapSeconds: 60
}

function centuryResult(state, overrides = {}) {
  return {
    httpStatus: overrides.httpStatus ?? 200,
    headers: overrides.headers || {},
    classification: {
      state,
      retryable: ["rate_limited", "temporary_error"].includes(state),
      raw: {
        code: overrides.code ?? 1,
        errCode: overrides.errCode ?? 49999,
        message: overrides.message || state
      }
    },
    requestStartedAt: new Date("2026-08-11T10:00:00Z"),
    responseReceivedAt: new Date("2026-08-11T10:00:01Z"),
    rateLimit: { limit: "30", remaining: "29", reset: null, retryAfter: null }
  }
}

test("profile classifiers prioritize verified err_code mappings", () => {
  const wos = centuryAdapter("wos", {})
  const kingshot = centuryAdapter("kingshot", {})
  assert.deepEqual(wos.responseMappings.errCodes, {
    20000: "success",
    40008: "already_redeemed",
    40011: "redemption_limit",
    40007: "expired",
    40005: "claim_limit",
    40006: "level_restriction",
    40012: "account_age_restriction",
    40014: "invalid_code",
    40017: "account_restriction",
    40019: "simultaneous_action_throttle",
    40020: "invalid_player",
    40100: "verification_throttle",
    40102: "verification_error",
    40103: "verification_error"
  })
  assert.deepEqual(kingshot.responseMappings.errCodes, {
    20000: "success",
    40008: "already_redeemed",
    40011: "redemption_limit",
    40007: "expired",
    40005: "claim_limit",
    40006: "level_restriction",
    40012: "account_age_restriction",
    40014: "invalid_code",
    40017: "account_restriction",
    40019: "simultaneous_action_throttle",
    40020: "invalid_player",
    40100: "verification_throttle",
    40102: "verification_error",
    40103: "verification_error"
  })
  assert.equal(classifyCenturyResponse({
    httpStatus: 200,
    data: { code: 20000, err_code: 40008, msg: "RECEIVED." },
    profileMappings: wos.responseMappings
  }).state, "already_redeemed")
  assert.equal(classifyCenturyResponse({
    httpStatus: 200,
    data: { code: 1, err_code: 40008, msg: "RECEIVED." },
    profileMappings: kingshot.responseMappings
  }).state, "already_redeemed")
  assert.equal(classifyCenturyResponse({
    httpStatus: 200,
    data: { code: 0, err_code: 20000, msg: "SUCCESS" },
    profileMappings: kingshot.responseMappings
  }).state, "success")
  for (const profile of ["wos", "kingshot"]) {
    const limited = classifyCenturyResponse({
      httpStatus: 200,
      data: {
        code: 1,
        data: [],
        msg: "The same type of Gift Code can only be redeemed once.",
        err_code: 40011
      },
      profileMappings: centuryAdapter(profile, {}).responseMappings
    })
    assert.equal(limited.state, "redemption_limit")
    assert.equal(limited.retryable, false)
    assert.equal(limited.permanent, true)
  }
})

test("verified profile mappings have explicit retry and canonical-code behavior", () => {
  const shared = [
    [20000, "success", { codeStatus: "active", verificationState: "complete" }],
    [40007, "expired", { codeStatus: "expired", verificationState: "complete" }],
    [40008, "already_redeemed", { codeStatus: "active", verificationState: "complete" }],
    [40011, "redemption_limit", { codeStatus: "active", verificationState: "complete" }],
    [40014, "invalid_code", { codeStatus: "invalid", verificationState: "complete" }]
  ]
  const now = new Date("2026-08-12T12:00:00Z")
  for (const profile of ["wos", "kingshot"]) {
    for (const [errCode, expected, transition] of shared) {
      const classified = classifyCenturyResponse({
        httpStatus: 200,
        data: { code: 1, err_code: errCode, msg: "message does not drive classification" },
        profileMappings: centuryAdapter(profile, {}).responseMappings
      })
      assert.equal(classified.state, expected, `${profile} ${errCode}`)
      assert.equal(classified.retryable, false, `${profile} ${errCode}`)
      assert.equal(classified.permanent, true, `${profile} ${errCode}`)
      assert.deepEqual(verificationTransition(classified.state, 1, now, config), transition)
    }
  }

  const wosPlayerError = classifyCenturyResponse({
    httpStatus: 200,
    data: { code: 1, err_code: 40020, msg: "USER INFO ERROR." },
    profileMappings: centuryAdapter("wos", {}).responseMappings
  })
  assert.equal(wosPlayerError.state, "invalid_player")
  assert.equal(wosPlayerError.retryable, false)
  assert.deepEqual(verificationTransition(wosPlayerError.state, 1, now, config), {
    codeStatus: "candidate",
    verificationState: "blocked"
  })
})

test("official frontend mappings classify proven account and transient outcomes", () => {
  const mapped = [
    [40005, "claim_limit", false, true],
    [40006, "level_restriction", false, true],
    [40012, "account_age_restriction", false, true],
    [40017, "account_restriction", false, true],
    [40019, "simultaneous_action_throttle", true, false],
    [40020, "invalid_player", false, true],
    [40100, "verification_throttle", true, false],
    [40102, "verification_error", false, true],
    [40103, "verification_error", false, true]
  ]
  for (const profile of ["wos", "kingshot"]) {
    for (const [errCode, state, retryable, permanent] of mapped) {
      const classified = classifyCenturyResponse({
        httpStatus: 200,
        data: { code: 1, err_code: errCode, msg: "message does not drive classification" },
        profileMappings: centuryAdapter(profile, {}).responseMappings
      })
      assert.equal(classified.state, state, `${profile} ${errCode}`)
      assert.equal(classified.retryable, retryable, `${profile} ${errCode}`)
      assert.equal(classified.permanent, permanent, `${profile} ${errCode}`)
    }
  }
})

test("unresolved frontend mappings remain profile-scoped and fail closed", () => {
  const sharedUnresolved = [40001, 40002, 40003, 40004, 40009, 40015, 40016, 40101]
  const now = new Date("2026-08-12T12:00:00Z")
  for (const [profile, unresolved] of [
    ["wos", [...sharedUnresolved, 40018]],
    ["kingshot", sharedUnresolved]
  ]) {
    for (const errCode of unresolved) {
      const classified = classifyCenturyResponse({
        httpStatus: 200,
        data: { code: 1, err_code: errCode, msg: `UNRESOLVED ${errCode}` },
        profileMappings: centuryAdapter(profile, {}).responseMappings
      })
      assert.equal(classified.state, "unknown_response", `${profile} ${errCode}`)
      assert.equal(classified.retryable, false, `${profile} ${errCode}`)
      assert.equal(classified.permanent, false, `${profile} ${errCode}`)
      assert.deepEqual(verificationTransition(classified.state, 1, now, config), {
        codeStatus: "unknown",
        verificationState: "review"
      })
    }
  }
})

test("Kingshot classifies exact verified expired and invalid-code payloads", () => {
  const mappings = centuryAdapter("kingshot", {}).responseMappings
  const expired = classifyCenturyResponse({
    httpStatus: 200,
    data: { code: 1, data: [], msg: "TIME ERROR.", err_code: 40007 },
    profileMappings: mappings
  })
  const invalid = classifyCenturyResponse({
    httpStatus: 200,
    data: { code: 1, data: [], msg: "CDK NOT FOUND.", err_code: 40014 },
    profileMappings: mappings
  })
  const errCodeWins = classifyCenturyResponse({
    httpStatus: 200,
    data: { code: 1, data: [], msg: "SUCCESS", err_code: 40014 },
    profileMappings: mappings
  })
  const unknown = classifyCenturyResponse({
    httpStatus: 200,
    data: { code: 1, data: [], msg: "NEW KINGSHOT RESPONSE", err_code: 49999 },
    profileMappings: mappings
  })

  assert.equal(expired.state, "expired")
  assert.equal(expired.retryable, false)
  assert.equal(expired.permanent, true)
  assert.equal(invalid.state, "invalid_code")
  assert.equal(invalid.retryable, false)
  assert.equal(invalid.permanent, true)
  assert.equal(errCodeWins.state, "invalid_code")
  assert.equal(unknown.state, "unknown_response")
  assert.equal(unknown.retryable, false)
  assert.equal(unknown.permanent, false)
})

test("unknown response metadata is retained when bounded and safely truncated when large", () => {
  assert.deepEqual(boundedResponseData({ code: 9, detail: { reason: "new" } }), {
    summary: '{"code":9,"detail":{"reason":"new"}}',
    truncated: false,
    originalCharacters: 36
  })
  assert.deepEqual(boundedResponseData({ detail: "x".repeat(100) }, 20), {
    summary: '{"detail":"xxxxxxxxx',
    truncated: true,
    originalCharacters: 113
  })
})

test("verifier configuration is optional, profile scoped and validated without fallback players", () => {
  assert.equal(getVerifierAccount("wos", {}), null)
  assert.deepEqual(getVerifierAccount("wos", { WOS_GIFT_VERIFY_FID: "123" }), {
    configured: false,
    reason: "incomplete verifier configuration"
  })
  assert.deepEqual(getVerifierAccount("wos", {
    WOS_GIFT_VERIFY_FID: "123",
    WOS_GIFT_VERIFY_KID: "689",
    KINGSHOT_GIFT_VERIFY_FID: "999",
    KINGSHOT_GIFT_VERIFY_KID: "521"
  }), { configured: true, playerId: "123", locationNumber: "689" })
  assert.equal(verificationWorkerIsEnabled({}), false)
  assert.equal(redemptionWorkerIsEnabled({}), false)
  assert.equal(verificationWorkerIsEnabled({ GIFT_CODE_VERIFICATION_ENABLED: "true" }), true)
  assert.equal(redemptionWorkerIsEnabled({ GIFT_CODE_REDEMPTION_WORKER_ENABLED: "true" }), true)
  assert.equal(giftWorkerConfig({}).maximumAttempts, 5)
})

test("verification state machine distinguishes valid, invalid, blocked, restricted, retry and unknown", () => {
  const now = new Date("2026-08-11T10:00:00Z")
  assert.deepEqual(verificationTransition("success", 1, now, config), {
    codeStatus: "active", verificationState: "complete"
  })
  assert.deepEqual(verificationTransition("already_redeemed", 1, now, config), {
    codeStatus: "active", verificationState: "complete"
  })
  assert.deepEqual(verificationTransition("redemption_limit", 1, now, config), {
    codeStatus: "active", verificationState: "complete"
  })
  for (const state of [
    "claim_limit", "level_restriction", "account_restriction", "account_age_restriction"
  ]) {
    assert.deepEqual(verificationTransition(state, 1, now, config), {
      codeStatus: "active", verificationState: "complete"
    })
  }
  assert.deepEqual(verificationTransition("expired", 1, now, config), {
    codeStatus: "expired", verificationState: "complete"
  })
  assert.deepEqual(verificationTransition("invalid_code", 1, now, config), {
    codeStatus: "invalid", verificationState: "complete"
  })
  assert.deepEqual(verificationTransition("invalid_player", 1, now, config), {
    codeStatus: "candidate", verificationState: "blocked"
  })
  assert.deepEqual(verificationTransition("eligibility_restriction", 1, now, config), {
    codeStatus: "restricted", verificationState: "review"
  })
  assert.equal(verificationTransition("rate_limited", 1, now, config).verificationState, "retry")
  assert.equal(verificationTransition("verification_throttle", 1, now, config).verificationState, "retry")
  assert.equal(
    verificationTransition("simultaneous_action_throttle", 1, now, config).verificationState,
    "retry"
  )
  assert.equal(
    verificationTransition("rate_limited", 1, now, config, 120000).nextRetryAt.toISOString(),
    "2026-08-11T10:02:00.000Z"
  )
  assert.deepEqual(verificationTransition("temporary_error", 3, now, config), {
    codeStatus: "candidate", verificationState: "blocked"
  })
  assert.deepEqual(verificationTransition("unknown_response", 1, now, config), {
    codeStatus: "unknown", verificationState: "review"
  })
  assert.deepEqual(verificationTransition("upstream_rejection", 1, now, config), {
    codeStatus: "unknown", verificationState: "review"
  })
  assert.deepEqual(verificationTransition("verification_error", 1, now, config), {
    codeStatus: "unknown", verificationState: "review"
  })
})

test("upstream rejection does not activate or fan out a candidate", async () => {
  let finished
  let activationCalls = 0
  let feedbackCalls = 0
  const processor = createVerificationProcessor({
    repository: {
      gameProfile: "wos",
      async claimVerification() { return { id: "code-id", code: "ABC", verification_attempt_count: 1 } },
      async finishVerification(input) {
        finished = input
        return { giftCode: { id: "code-id", status: input.codeStatus }, queued: [] }
      }
    },
    client: { async redeem() { return centuryResult("upstream_rejection", { httpStatus: 403, errCode: null }) } },
    verifier: { configured: true, playerId: "123", locationNumber: "689" },
    community: {
      async onCodeActivated() { activationCalls += 1 },
      async onVerificationResult() { feedbackCalls += 1 }
    },
    config,
    botInstanceName: "test",
    logger: { log() {}, warn() {}, error() {} },
    workerId: "verify-worker"
  })
  assert.equal(await processor.tick(), 1)
  assert.equal(finished.codeStatus, "unknown")
  assert.equal(finished.verificationState, "review")
  assert.equal(activationCalls, 0)
  assert.equal(feedbackCalls, 1)
})

test("verifier redemption limit recognises and activates codes for both profiles", async () => {
  for (const gameProfile of ["wos", "kingshot"]) {
    let finished
    let activationCalls = 0
    let feedbackCalls = 0
    const processor = createVerificationProcessor({
      repository: {
        gameProfile,
        async claimVerification() {
          return { id: `${gameProfile}-code`, code: "TYPELIMIT", verification_attempt_count: 1 }
        },
        async finishVerification(input) {
          finished = input
          return {
            giftCode: { id: `${gameProfile}-code`, status: input.codeStatus },
            queued: [{ id: `${gameProfile}-redemption` }]
          }
        }
      },
      client: {
        async redeem() {
          return centuryResult("redemption_limit", {
            errCode: 40011,
            message: "The same type of Gift Code can only be redeemed once."
          })
        }
      },
      verifier: { configured: true, playerId: "123", locationNumber: "521" },
      community: {
        async onCodeActivated(_id, queuedCount) {
          activationCalls += 1
          assert.equal(queuedCount, 1)
        },
        async onVerificationResult() { feedbackCalls += 1 }
      },
      config,
      botInstanceName: `test-${gameProfile}`,
      logger: { log() {}, warn() {}, error() {} },
      workerId: `verify-${gameProfile}`
    })
    assert.equal(await processor.tick(), 1)
    assert.equal(finished.result.classification.state, "redemption_limit")
    assert.equal(finished.codeStatus, "active")
    assert.equal(finished.verificationState, "complete")
    assert.equal(finished.nextRetryAt, undefined)
    assert.equal(activationCalls, 1)
    assert.equal(feedbackCalls, 0)
  }
})

test("fresh Kingshot 40011 follows the real client and worker path", async () => {
  let requests = 0
  let finished
  const client = createCenturyGameClient({
    gameProfile: "kingshot",
    env: {},
    transport: {
      async post() {
        requests += 1
        return {
          status: 200,
          headers: { "content-type": "application/json" },
          data: { code: 1, data: [], msg: "SAME TYPE EXCHANGE.", err_code: 40011 }
        }
      }
    },
    limiter: new ConservativeRateLimiter({
      gameProfile: "kingshot",
      minimumDelayMs: 0,
      maximumRetries: 0
    }),
    now: () => Date.parse("2026-08-12T12:00:00Z")
  })
  const processor = createVerificationProcessor({
    repository: {
      gameProfile: "kingshot",
      async claimVerification() {
        return { id: "fresh-40011", code: "Kingshot888", verification_attempt_count: 1 }
      },
      async finishVerification(input) {
        finished = input
        return { giftCode: { id: "fresh-40011", status: input.codeStatus }, queued: [{}] }
      }
    },
    client,
    verifier: { configured: true, playerId: "368775177", locationNumber: "521" },
    config,
    botInstanceName: "peggie-kingshot",
    logger: { log() {}, warn() {}, error() {} },
    workerId: "fresh-40011-worker"
  })
  assert.equal(await processor.tick(), 1)
  assert.equal(requests, 1)
  assert.equal(finished.result.classification.state, "redemption_limit")
  assert.equal(finished.codeStatus, "active")
  assert.equal(finished.verificationState, "complete")
})

test("stored 40011 recovery uses the current adapter without a Century request", async () => {
  let claimed = true
  let finished
  let requests = 0
  const processor = createVerificationProcessor({
    repository: {
      gameProfile: "kingshot",
      async storedVerificationReview(input) {
        assert.deepEqual(input.errCodes, [40011])
        if (!claimed) return null
        claimed = false
        return {
          id: "stored-code",
          code: "Kingshot888",
          verification_attempt_count: 1,
          stored_attempt_id: "stored-attempt",
          stored_http_status: 200,
          stored_api_code: 1,
          stored_err_code: 40011,
          stored_api_message: "SAME TYPE EXCHANGE."
        }
      },
      async finishStoredVerificationRecovery(input) {
        finished = input
        return { giftCode: { id: "stored-code", status: input.codeStatus }, queued: [{}] }
      },
      async claimVerification() { return null }
    },
    client: {
      adapter: centuryAdapter("kingshot", {}),
      async redeem() { requests += 1; throw new Error("must not request Century") }
    },
    verifier: { configured: true, playerId: "368775177", locationNumber: "521" },
    config,
    botInstanceName: "peggie-kingshot",
    logger: { log() {}, warn() {}, error() {} },
    workerId: "stored-40011-worker"
  })
  assert.equal(await processor.tick(), 1)
  assert.equal(requests, 0)
  assert.equal(finished.classification, "redemption_limit")
  assert.equal(finished.codeStatus, "active")
  assert.equal(finished.verificationState, "complete")
  assert.equal(await processor.tick(), 0, "restart-style repeat must not recover twice")
})

test("account-specific verifier restrictions activate and fan out without global invalidation", async () => {
  for (const classification of [
    "claim_limit", "level_restriction", "account_restriction", "account_age_restriction"
  ]) {
    let finished
    let activations = 0
    const processor = createVerificationProcessor({
      repository: {
        gameProfile: "wos",
        async claimVerification() {
          return { id: classification, code: "ACCOUNTBOUND", verification_attempt_count: 1 }
        },
        async finishVerification(input) {
          finished = input
          return {
            giftCode: { id: classification, status: input.codeStatus },
            queued: [{ id: `${classification}-redemption` }]
          }
        }
      },
      client: { async redeem() { return centuryResult(classification) } },
      verifier: { configured: true, playerId: "123", locationNumber: "689" },
      community: { async onCodeActivated() { activations += 1 } },
      config,
      botInstanceName: "test",
      logger: { log() {}, warn() {}, error() {} },
      workerId: `${classification}-worker`
    })
    assert.equal(await processor.tick(), 1)
    assert.equal(finished.codeStatus, "active")
    assert.equal(finished.verificationState, "complete")
    assert.equal(activations, 1)
  }
})

test("verification-result feedback is concise, truthful and contains no API details", () => {
  assert.equal(verificationResultMessage("invalid_code"), "That code doesn't look valid.")
  assert.equal(verificationResultMessage("expired"), "That code has expired.")
  assert.equal(
    verificationResultMessage("eligibility_restriction"),
    "I couldn't confirm that code, so I haven't added it."
  )
  assert.equal(
    verificationResultMessage("upstream_rejection"),
    "I couldn't verify that code right now. I'll leave it for review."
  )
  assert.equal(
    verificationResultMessage("temporary_error"),
    "I couldn't verify that code right now. I'll leave it for review."
  )
  assert.equal(
    verificationResultMessage("unknown_response"),
    "I couldn't verify that code, so I've left it for review."
  )
  for (const classification of [
    "invalid_code", "expired", "eligibility_restriction",
    "upstream_rejection", "temporary_error", "unknown_response"
  ]) {
    assert.doesNotMatch(verificationResultMessage(classification), /HTTP|err_code|Century|403|20000/i)
  }
})

test("verification-result notification is not duplicated after service restart", async () => {
  const sent = []
  let notification = {
    id: "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa",
    submitted_by_discord_user_id: "999",
    submitted_code: "ABC",
    classification: "upstream_rejection"
  }
  const repository = {
    async claimVerificationResultNotification() {
      const claimed = notification
      notification = null
      return claimed
    },
    async finishVerificationResultNotification(id, workerId, result) {
      assert.equal(id, "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa")
      assert.equal(result.sent, true)
    },
    async claimNextPending() { return null }
  }
  const client = {
    users: {
      async fetch() {
        return { async send(payload) { sent.push(payload) } }
      }
    }
  }
  const first = createGiftCodeCommunityService({
    repository, client, gameProfile: "wos", workerId: "worker-a", logger: { warn() {} }
  })
  assert.equal(await first.recoverOne(), 1)
  const restarted = createGiftCodeCommunityService({
    repository, client, gameProfile: "wos", workerId: "worker-b", logger: { warn() {} }
  })
  assert.equal(await restarted.recoverOne(), 0)
  assert.equal(sent.length, 1)
  assert.equal(sent[0].content, "I couldn't verify that code right now. I'll leave it for review.")
  assert.equal(sent[0].enforceNonce, true)
})

test("meaningful non-active verification outcomes are delivered privately", async () => {
  const outcomes = [
    ["invalid_code", "That code doesn't look valid."],
    ["expired", "That code has expired."],
    ["redemption_limit", "I couldn't confirm that code, so I haven't added it."],
    ["upstream_rejection", "I couldn't verify that code right now. I'll leave it for review."],
    ["unknown_response", "I couldn't verify that code, so I've left it for review."]
  ]
  const sent = []
  let index = 0
  const repository = {
    async claimVerificationResultNotification() {
      const classification = outcomes[index]?.[0]
      if (!classification) return null
      index += 1
      return {
        id: `${String(index).padStart(8, "0")}-0000-4000-8000-000000000000`,
        submitted_by_discord_user_id: "999",
        classification
      }
    },
    async finishVerificationResultNotification() {}
  }
  const service = createGiftCodeCommunityService({
    repository,
    client: {
      users: {
        async fetch(userId) {
          assert.equal(userId, "999")
          return { async send(payload) { sent.push(payload) } }
        }
      }
    },
    gameProfile: "wos",
    workerId: "feedback-worker",
    logger: { warn() {} }
  })
  for (const outcome of outcomes) assert.equal(await service.onVerificationResult(outcome[0]), true)
  assert.deepEqual(sent.map(message => message.content), outcomes.map(([, message]) => message))
  assert.ok(sent.every(message => message.enforceNonce === true))
  assert.doesNotMatch(JSON.stringify(sent), /HTTP|err_code|Century|403|20000/i)
})

test("stored edge metadata and Inspect Code remain useful without exposing raw response data", () => {
  const fields = machineFields({
    profile: "wos",
    endpoint: "/gift_code",
    httpStatus: 403,
    classification: { state: "upstream_rejection", raw: { code: null, errCode: null, message: "" } },
    responseDiagnostics: { responseType: "html", bodySummary: "Access denied" }
  })
  assert.equal(fields.metadata.classification, "upstream_rejection")
  assert.equal(fields.metadata.response.responseType, "html")
  assert.equal(fields.metadata.rawResponse, undefined)
  const output = formatCodeDiagnostics({
    code: "gogoWOS",
    status: "unknown",
    verification_state: "review",
    verification_http_status: 403,
    last_err_code: null,
    verification_metadata: fields.metadata,
    pending_count: 0,
    success_count: 0,
    already_redeemed_count: 0,
    failed_count: 0
  })
  assert.match(output, /HTTP status: 403/)
  assert.match(output, /Current status: unknown/)
  assert.match(output, /Current verification state: review/)
  assert.match(output, /Latest recorded verification attempt/)
  assert.match(output, /Century err_code: None/)
  assert.match(output, /Response type: HTML/)
  assert.match(output, /Classification: Upstream Rejection/)
  assert.match(output, /Historical verification failed before Century returned an error code/)
  assert.match(output, /Re-verification with the current client is required/)
  assert.doesNotMatch(output, /Access denied/)
})

test("Inspect Code does not recommend transport recovery for semantic unknown responses", () => {
  const output = formatCodeDiagnostics({
    code: "SEMANTICUNKNOWN",
    status: "unknown",
    verification_state: "review",
    latest_verification_attempt_id: "semantic-attempt",
    latest_verification_http_status: 200,
    latest_verification_err_code: 49999,
    latest_verification_classification: "unknown_response",
    pending_count: 0,
    success_count: 0,
    already_redeemed_count: 0,
    failed_count: 0
  })
  assert.doesNotMatch(output, /Re-verification with the current client is required/)
})

test("Inspect Code shows sanitized verification and latest redemption messages", () => {
  const output = formatCodeDiagnostics({
    code: "KSUNKNOWN",
    status: "unknown",
    verification_state: "review",
    latest_verification_attempt_id: "verification-attempt",
    latest_verification_http_status: 200,
    latest_verification_err_code: 40011,
    latest_verification_api_message: "VERIFY DETAIL 987654321 token=secret-value",
    latest_verification_classification: "redemption_limit",
    latest_verification_metadata: {
      response: { responseType: "json", bodySummary: "RAW RESPONSE MUST NOT LEAK" }
    },
    latest_redemption_attempt_id: "redemption-attempt",
    latest_redemption_http_status: 200,
    latest_redemption_err_code: 40011,
    latest_redemption_api_message: "Player 368775177 result sign=secret-signature @everyone",
    latest_redemption_classification: "redemption_limit",
    latest_redemption_player_id_snapshot: "368775177",
    pending_count: 1,
    success_count: 2,
    already_redeemed_count: 3,
    failed_count: 4
  })

  assert.match(output, /Century message: VERIFY DETAIL \[identifier redacted\] token=\[redacted\]/)
  assert.match(output, /\*\*Latest player redemption\*\*/)
  assert.match(output, /Century message: Player \[redacted\] result sign=\[redacted\] \[mention redacted\]/)
  assert.equal((output.match(/Classification: Redemption Limit/g) || []).length, 2)
  assert.equal((output.match(/Meaning: Another gift code of this type was already redeemed on this character/g) || []).length, 2)
  assert.match(output, /Queue counts: pending 1, success 2, already redeemed 3, review\/failed 4/)
  assert.doesNotMatch(output, /987654321|368775177|secret-value|secret-signature|RAW RESPONSE MUST NOT LEAK/)
})

test("Inspect Code reports historical messages that were not recorded", () => {
  const output = formatCodeDiagnostics({
    code: "OLDKS",
    status: "unknown",
    verification_state: "review",
    latest_verification_attempt_id: "old-verification",
    latest_verification_http_status: 200,
    latest_verification_err_code: 40011,
    latest_verification_api_message: null,
    latest_verification_classification: "unknown_response",
    latest_verification_metadata: { response: { responseType: "json" } },
    latest_redemption_attempt_id: "old-redemption",
    latest_redemption_http_status: 200,
    latest_redemption_err_code: 40011,
    latest_redemption_api_message: null,
    latest_redemption_classification: "unknown_response",
    pending_count: 0,
    success_count: 0,
    already_redeemed_count: 0,
    failed_count: 1
  })
  assert.equal((output.match(/Century message: Not recorded/g) || []).length, 2)
})

test("Inspect Code preserves meanings for newly resolved frontend classifications", () => {
  for (const [classification, meaning] of [
    ["claim_limit", "This character has reached the claim limit"],
    ["level_restriction", "This character's City or Town Center level is too low"],
    ["account_restriction", "This account does not meet the redemption requirements"],
    ["account_age_restriction", "This account does not meet the required account age"],
    ["verification_throttle", "Verification requests were made too frequently"],
    ["verification_error", "The frontend verification code was incorrect or expired"],
    ["simultaneous_action_throttle", "Too many simultaneous actions were in progress"]
  ]) {
    const output = formatCodeDiagnostics({
      code: "MAPPED",
      status: "unknown",
      verification_state: "review",
      latest_verification_attempt_id: classification,
      latest_verification_http_status: 200,
      latest_verification_err_code: 49999,
      latest_verification_classification: classification,
      pending_count: 0,
      success_count: 0,
      already_redeemed_count: 0,
      failed_count: 0
    })
    assert.ok(output.includes(`Meaning: ${meaning}`), classification)
  }
})

test("unknown numeric responses remain unknown after resolving 40011", () => {
  for (const profile of ["wos", "kingshot"]) {
    const classified = classifyCenturyResponse({
      httpStatus: 200,
      data: { code: 1, data: [], msg: "UNVERIFIED RESPONSE", err_code: 49998 },
      profileMappings: centuryAdapter(profile, {}).responseMappings
    })
    assert.equal(classified.state, "unknown_response")
    assert.equal(classified.raw.message, "UNVERIFIED RESPONSE")
    assert.deepEqual(verificationTransition(classified.state, 1, new Date(), config), {
      codeStatus: "unknown",
      verificationState: "review"
    })
  }
})

test("controlled live WOS harness is opt-in and makes exactly one injected request", async () => {
  let requests = 0
  await assert.rejects(runControlledCenturyVerification({ env: {} }), /blocked/)
  const report = await runControlledCenturyVerification({
    env: {
      ALLOW_ONE_LIVE_CENTURY_REQUEST: "true",
      GAME_PROFILE: "wos",
      LIVE_CENTURY_GIFT_CODE: "gogoWOS",
      WOS_GIFT_VERIFY_FID: "123",
      WOS_GIFT_VERIFY_KID: "689"
    },
    clientFactory: () => ({
      adapter: centuryAdapter("wos", {}),
      async redeem() {
        requests += 1
        return centuryResult("upstream_rejection", { httpStatus: 403, errCode: null })
      }
    }),
    write() {}
  })
  assert.equal(requests, 1)
  assert.equal(report.requestCount, 1)
  assert.equal(report.httpStatus, 403)
})

test("controlled live Kingshot harness uses its profile verifier and makes exactly one request", async () => {
  let requests = 0
  let clientOptions = null
  const output = []
  const report = await runControlledCenturyVerification({
    env: {
      ALLOW_ONE_LIVE_CENTURY_REQUEST: "true",
      GAME_PROFILE: "kingshot",
      LIVE_CENTURY_GIFT_CODE: "KS0810",
      KINGSHOT_GIFT_VERIFY_FID: "368775177",
      KINGSHOT_GIFT_VERIFY_KID: "521",
      WOS_GIFT_VERIFY_FID: "ignored",
      WOS_GIFT_VERIFY_KID: "999"
    },
    clientFactory(options) {
      clientOptions = options
      return {
        adapter: centuryAdapter("kingshot", {}),
        async redeem(input) {
          requests += 1
          assert.deepEqual(input, {
            playerId: "368775177",
            locationNumber: "521",
            code: "KS0810"
          })
          return {
            ...centuryResult("success", { errCode: 20000 }),
            responseDiagnostics: {
              responseType: "json",
              bodySummary: "bounded",
              bodyTruncated: false,
              originalBodyCharacters: 7
            }
          }
        }
      }
    },
    write(value) { output.push(value) }
  })

  assert.equal(requests, 1)
  assert.equal(clientOptions.gameProfile, "kingshot")
  assert.equal(clientOptions.limiter.maximumRetries, 0)
  assert.equal(report.requestCount, 1)
  assert.equal(report.method, "POST")
  assert.match(report.endpoint, /kingshot-giftcode\.centurygame\.com\/api\/gift_code$/)
  assert.equal(report.httpStatus, 200)
  assert.equal(report.classification, "success")
  assert.equal(report.centuryErrCode, 20000)
  assert.deepEqual(report.rateLimit, { limit: "30", remaining: "29", reset: null, retryAfter: null })
  assert.equal(report.response.bodySummary, "bounded")
  assert.equal(output.length, 1)
  assert.doesNotMatch(output[0], /mN4!pQs6JrYwV9|fid=|kid=|sign=/)
})

test("redemption state machine has bounded durable retries and manual-review terminals", () => {
  const now = new Date("2026-08-11T10:00:00Z")
  for (const state of ["success", "already_redeemed", "expired", "invalid_code", "invalid_player"]) {
    assert.deepEqual(redemptionTransition(state, 1, now, config), { status: state, retryable: false })
  }
  assert.deepEqual(redemptionTransition("redemption_limit", 1, now, config), {
    status: "restricted", retryable: false
  })
  for (const state of [
    "claim_limit", "level_restriction", "account_restriction", "account_age_restriction"
  ]) {
    assert.deepEqual(redemptionTransition(state, 1, now, config), {
      status: "restricted", retryable: false
    })
  }
  assert.equal(redemptionTransition("rate_limited", 1, now, config).nextRetryAt.toISOString(), "2026-08-11T10:00:10.000Z")
  assert.equal(redemptionTransition("verification_throttle", 1, now, config).status, "rate_limited")
  assert.equal(
    redemptionTransition("simultaneous_action_throttle", 1, now, config).status,
    "temporary_error"
  )
  assert.equal(
    redemptionTransition("rate_limited", 1, now, config, 90000).nextRetryAt.toISOString(),
    "2026-08-11T10:01:30.000Z"
  )
  assert.deepEqual(redemptionTransition("temporary_error", 3, now, config), {
    status: "retry_exhausted", retryable: false
  })
  assert.deepEqual(redemptionTransition("unknown_response", 1, now, config), {
    status: "unknown", retryable: false
  })
})

test("gift settings target one owned account and preserve profile terminology", async () => {
  const calls = []
  const repository = {
    async activeAccountStatuses(owner, playerId) {
      calls.push({ owner, playerId })
      return [{
        player_id: playerId || "111",
        state_or_kingdom_number: "689",
        is_active: true,
        is_primary: true
      }]
    },
    async setAutoRedemption(input) {
      calls.push(input)
      return {
        account: {
          player_id: input.playerId,
          state_or_kingdom_number: "689",
          gift_redemption_enabled: input.enabled
        },
        limitReached: false,
        enabledCount: 1,
        engagementEvent: null
      }
    }
  }
  const service = createGiftCodeService({ repository, gameProfile: "wos" })
  const account = await service.setAutomaticRedemption({
    discordUserId: "999",
    playerId: "222",
    enabled: true
  })
  assert.equal(service.terms.locationLabel, "State")
  assert.equal(account.player_id, "222")
  assert.deepEqual(calls.at(-1), {
    discordUserId: "999",
    playerId: "222",
    enabled: true,
    guildId: null,
    maximumEnabledAccounts: 2
  })
})

test("gift panel commands register for both profiles and status uses State or Kingdom wording", () => {
  const wos = buildGiftCodesCommand("wos").toJSON()
  const kingshot = buildGiftCodesCommand("kingshot").toJSON()
  const admin = buildGiftCodesAdminCommand("wos").toJSON()
  assert.equal(wos.name, "gift-codes")
  assert.equal(admin.name, "gift-codes-admin")
  assert.match(JSON.stringify(wos), /Whiteout Survival/)
  assert.match(JSON.stringify(kingshot), /Kingshot/)
  const account = {
    player_id: "123",
    state_or_kingdom_number: "689",
    gift_redemption_enabled: true,
    is_active: true,
    verification_status: "verified",
    successful_redemptions: 1,
    already_redeemed: 2,
    completed_redemption_checks: 3
  }
  const wosStatus = formatPlayerGiftStatus(account, profileTerminology("wos"))
  assert.match(wosStatus, /State: 689/)
  assert.match(wosStatus, /Successful redemptions: 1/)
  assert.match(wosStatus, /Already claimed: 2/)
  assert.match(wosStatus, /Completed redemption checks: 3/)
  assert.match(formatPlayerGiftStatus(account, profileTerminology("kingshot")), /Kingdom: 689/)
  assert.equal(getGiftCommandData({ PLAYER_GIFT_CODES_ENABLED: "false" }).length, 0)
  assert.equal(getGiftCommandData({ PLAYER_GIFT_CODES_ENABLED: "true", GAME_PROFILE: "wos" }).length, 3)
})

test("verification processor leaves candidates pending when verifier is absent", async () => {
  let claimed = false
  const processor = createVerificationProcessor({
    repository: { gameProfile: "wos", async claimVerification() { claimed = true } },
    client: { async redeem() { throw new Error("must not call") } },
    verifier: null,
    config,
    botInstanceName: "test",
    logger: { log() {}, error() {} }
  })
  assert.equal(await processor.tick(), 0)
  assert.equal(claimed, false)
  assert.deepEqual(await processor.verifyCode("ABC"), {
    processed: false,
    reason: "verifier not configured"
  })
})

test("verification processor persists classification and never exposes verifier identifiers in logs", async () => {
  const logs = []
  const warnings = []
  let finished
  const claim = { id: "code-id", code: "ABC", verification_attempt_count: 1 }
  const processor = createVerificationProcessor({
    repository: {
      gameProfile: "wos",
      async claimVerification() { return claim },
      async finishVerification(input) {
        finished = input
        return { giftCode: { status: input.codeStatus }, queued: [] }
      }
    },
    client: { async redeem() { return centuryResult("already_redeemed", { errCode: 40008 }) } },
    verifier: { configured: true, playerId: "987654321", locationNumber: "689" },
    community: {
      async onCodeActivated() {
        throw Object.assign(new Error("role hierarchy"), { code: "CONTRIBUTOR_ROLE_HIERARCHY" })
      }
    },
    config,
    botInstanceName: "test",
    logger: {
      log(value) { logs.push(value) },
      warn(value) { warnings.push(value) },
      error() {}
    },
    now: () => new Date("2026-08-11T10:00:02Z"),
    workerId: "verify-worker"
  })
  assert.equal(await processor.tick(), 1)
  assert.equal(finished.codeStatus, "active")
  assert.equal(finished.verificationState, "complete")
  assert.ok(logs.every(line => !line.includes("987654321") && !line.includes("689")))
  assert.match(warnings[0], /Community activation failed/)
})

test("redemption processor records result before DM and DM failure does not alter it", async () => {
  const order = []
  const claim = {
    id: "redemption-id",
    gift_code_id: "code-id",
    code: "ABC",
    player_id_snapshot: "123",
    location_number_snapshot: "521",
    discord_user_id: "999",
    attempt_number: 1
  }
  const processor = createRedemptionProcessor({
    repository: {
      gameProfile: "kingshot",
      async claimRedemption() { return claim },
      async finishRedemption(input) { order.push(`persist:${input.status}`) },
      async claimNotification() { order.push("claim-dm"); return {} },
      async finishNotification(_id, outcome) { order.push(`dm:${outcome.sent}`) }
    },
    client: { async redeem() { return centuryResult("success", { errCode: 20000 }) } },
    notifier: async () => { order.push("send-dm"); return { sent: false, errorCode: "CANNOT_MESSAGE" } },
    config,
    logger: { log() {}, error() {} },
    now: () => new Date("2026-08-11T10:00:02Z"),
    workerId: "redeem-worker"
  })
  assert.equal(await processor.tick(), 1)
  assert.deepEqual(order, ["persist:success", "claim-dm", "send-dm", "dm:false"])
})

test("player redemption limit persists terminally and uses its specific notification", async () => {
  const calls = []
  const claim = {
    id: "redemption-limit-id",
    gift_code_id: "code-id",
    player_account_id: "account-id",
    code: "TYPELIMIT",
    player_id_snapshot: "368775177",
    location_number_snapshot: "521",
    discord_user_id: "999",
    attempt_number: 1
  }
  const processor = createRedemptionProcessor({
    repository: {
      gameProfile: "kingshot",
      async claimRedemption() { return claim },
      async finishRedemption(input) {
        calls.push({ type: "persist", status: input.status, retryable: input.retryable })
      },
      async claimNotification() { return {} },
      async finishNotification() {}
    },
    client: {
      async redeem() {
        return centuryResult("redemption_limit", {
          errCode: 40011,
          message: "The same type of Gift Code can only be redeemed once."
        })
      }
    },
    notifier: async (_claim, status) => {
      calls.push({ type: "notify", status })
      return { sent: true }
    },
    config,
    logger: { log() {}, warn() {}, error() {} },
    now: () => new Date("2026-08-11T10:00:02Z"),
    workerId: "redeem-limit-worker"
  })
  assert.equal(await processor.tick(), 1)
  assert.deepEqual(calls, [
    { type: "persist", status: "restricted", retryable: false },
    { type: "notify", status: "redemption_limit" }
  ])
})

test("player redemption DMs are concise, profile aware, masked when needed and contain failures", async () => {
  const claim = {
    code: "ABC",
    player_id_snapshot: "282021376",
    location_number_snapshot: "521",
    discord_user_id: "999",
    owner_account_count: 1,
    response_metadata: { rawResponse: "must-not-leak" },
    api_message: "Century internal response"
  }
  assert.equal(
    notificationMessage(claim, "success", "wos"),
    "ABC redeemed. Check your in-game mail."
  )
  assert.equal(
    notificationMessage(claim, "already_redeemed", "wos"),
    "ABC was already claimed on this character. You're good."
  )
  assert.equal(
    notificationMessage(claim, "redemption_limit", "kingshot"),
    "ABC can't be claimed on this character - you've already used another code of this type."
  )
  assert.equal(
    notificationMessage(claim, "claim_limit", "kingshot"),
    "ABC can't be claimed on this character - the claim limit has been reached."
  )
  assert.equal(
    notificationMessage(claim, "level_restriction", "wos"),
    "ABC can't be claimed on this character - its Furnace level is too low."
  )
  assert.equal(
    notificationMessage(claim, "level_restriction", "kingshot"),
    "ABC can't be claimed on this character - its Town level is too low."
  )
  assert.equal(
    notificationMessage({ ...claim, location_number_snapshot: "689" }, "invalid_player", "wos"),
    "I couldn't redeem ABC. Check that this character is still in State 689."
  )
  assert.equal(
    notificationMessage(claim, "invalid_player", "kingshot"),
    "I couldn't redeem ABC. Check that this character is still in Kingdom 521."
  )
  const multiple = { ...claim, location_number_snapshot: "689", owner_account_count: 2 }
  for (const status of ["success", "already_redeemed", "invalid_player"]) {
    const message = notificationMessage(multiple, status, "wos")
    assert.match(message, /Player ID \*\*\*1376/)
    assert.doesNotMatch(message, /282021376|must-not-leak|Century internal response/)
  }
  const notifier = createDiscordGiftNotifier({
    client: { users: { async fetch() { throw Object.assign(new Error("closed"), { code: 50007 }) } } },
    gameProfile: "wos",
    logger: { warn() {} }
  })
  assert.deepEqual(await notifier(claim, "success"), { sent: false, errorCode: "50007" })
})

test("low remaining rate-limit observations slow subsequent serial work without guessing reset", async () => {
  let now = 0
  const sleeps = []
  const limiter = new ConservativeRateLimiter({
    gameProfile: "wos",
    minimumDelayMs: 100,
    maximumRetries: 0,
    baseBackoffMs: 1000,
    maximumBackoffMs: 5000,
    now: () => now,
    sleep: async delay => { sleeps.push(delay); now += delay }
  })
  await limiter.schedule(async () => ({
    httpStatus: 200,
    headers: { "x-ratelimit-limit": "30", "x-ratelimit-remaining": "1" },
    retryable: false
  }))
  await limiter.schedule(async () => ({ httpStatus: 200, headers: {}, retryable: false }))
  assert.deepEqual(sleeps, [1000])
  assert.equal(limiter.getObservations()[0].limit, "30")
  assert.equal(limiter.getObservations()[0].remaining, "1")
})

test("polling worker is disabled by default and never invokes live work", async () => {
  let calls = 0
  const worker = createPollingWorker({ tick: async () => { calls += 1 }, intervalMs: 1000, enabled: false })
  assert.deepEqual(worker.start(), { started: false, reason: "disabled" })
  assert.equal(await worker.tick(), 0)
  assert.equal(calls, 0)
})

test("gift-admin enforces the existing authorization callback before repository access", async () => {
  let repositoryCreated = false
  const interaction = {
    commandName: "gift-admin",
    user: { id: "999" },
    isChatInputCommand: () => true,
    options: { getSubcommand: () => "status", getString: () => null },
    async deferReply() {},
    async editReply(message) { this.reply = message }
  }
  assert.equal(await handleGiftInteraction(interaction, {
    userCanManageServer: async () => false,
    repositoryFactory() { repositoryCreated = true },
    healthProvider: () => ({ available: true, gameProfile: "wos" })
  }), true)
  assert.equal(repositoryCreated, false)
  assert.match(interaction.reply, /do not have permission/)
})
