const test = require("node:test")
const assert = require("node:assert/strict")
const { Pool } = require("pg")

const { runMigrations } = require("../src/migrate")
const { createPlayerRepository } = require("../src/giftCodes/playerRepository")
const { createPlayerService } = require("../src/giftCodes/playerService")
const { createGiftCodeRepository } = require("../src/giftCodes/repository")
const { centuryAdapter } = require("../src/giftCodes/adapters")
const { classifyCenturyResponse } = require("../src/giftCodes/responseClassifier")
const { verificationTransition, redemptionTransition } = require("../src/giftCodes/workers")

const databaseUrl = process.env.TEST_DATABASE_URL
const silentLogger = { log() {}, error() {} }

function apiResult(state, { code = 1, errCode = 49999, message = state, httpStatus = 200 } = {}) {
  return {
    httpStatus,
    classification: { state, raw: { code, errCode, message } },
    requestStartedAt: new Date("2026-08-11T10:00:00Z"),
    responseReceivedAt: new Date("2026-08-11T10:00:01Z"),
    rateLimit: { limit: "30", remaining: "29", reset: null, retryAfter: null }
  }
}

test("gift-code workers are profile scoped, concurrency safe, durable and location current", {
  skip: databaseUrl ? false : "TEST_DATABASE_URL is not configured"
}, async () => {
  const admin = new Pool({ connectionString: databaseUrl, max: 1 })
  const schema = `gift_workflow_${process.pid}_${Date.now()}`
  let pool
  try {
    await admin.query(`CREATE SCHEMA "${schema}"`)
    pool = new Pool({
      connectionString: databaseUrl,
      max: 10,
      options: `-c search_path=${schema}`
    })
    const first = await runMigrations({ pool, logger: silentLogger })
    const second = await runMigrations({ pool, logger: silentLogger })
    assert.equal(first.applied.length, 16)
    assert.equal(first.applied.at(-1), "016_gift_code_sources.sql")
    assert.deepEqual(second.applied, [])

    const attemptColumns = (await pool.query(
      `SELECT column_name FROM information_schema.columns
        WHERE table_schema = $1 AND table_name = 'gift_code_attempts'`,
      [schema]
    )).rows.map(row => row.column_name)
    assert.ok(attemptColumns.includes("classification"))
    assert.ok(attemptColumns.includes("location_number_snapshot"))
    const indexes = (await pool.query(
      `SELECT indexname FROM pg_indexes
        WHERE schemaname = $1 AND tablename IN ('gift_codes', 'gift_code_attempts')`,
      [schema]
    )).rows.map(row => row.indexname)
    assert.ok(indexes.includes("gift_codes_verification_work_idx"))
    assert.ok(indexes.includes("gift_code_attempts_redemption_number_idx"))

    const lockRepositoryA = createGiftCodeRepository(pool, "wos")
    const lockRepositoryB = createGiftCodeRepository(pool, "wos")
    let activeRequests = 0
    let maximumConcurrentRequests = 0
    const requestStarts = []
    const operation = async () => {
      activeRequests += 1
      maximumConcurrentRequests = Math.max(maximumConcurrentRequests, activeRequests)
      requestStarts.push(Date.now())
      await new Promise(resolve => setTimeout(resolve, 20))
      activeRequests -= 1
    }
    await Promise.all([
      lockRepositoryA.withProfileRequestLock(operation, { minimumDelayMs: 30 }),
      lockRepositoryB.withProfileRequestLock(operation, { minimumDelayMs: 30 })
    ])
    assert.equal(maximumConcurrentRequests, 1)
    assert.ok(requestStarts[1] - requestStarts[0] >= 25)
    assert.equal((await pool.query(
      "SELECT COUNT(*)::integer AS count FROM gift_code_rate_limit_state WHERE game_profile = 'wos'"
    )).rows[0].count, 1)

    const owner = "111111111111111111"
    const otherOwner = "222222222222222222"
    const wosPlayers = createPlayerService({
      repository: createPlayerRepository(pool, "wos"),
      gameProfile: "wos"
    })
    const kingshotPlayers = createPlayerService({
      repository: createPlayerRepository(pool, "kingshot"),
      gameProfile: "kingshot"
    })
    const opted = await wosPlayers.register({
      discordUserId: owner,
      playerId: "10001",
      locationNumber: "689"
    })
    const optedOut = await wosPlayers.register({
      discordUserId: owner,
      playerId: "10002",
      locationNumber: "690"
    })
    const inactive = await wosPlayers.register({
      discordUserId: otherOwner,
      playerId: "10003",
      locationNumber: "691"
    })
    const kingshot = await kingshotPlayers.register({
      discordUserId: owner,
      playerId: "10001",
      locationNumber: "521"
    })
    const kingshotB = await kingshotPlayers.register({
      discordUserId: "333333333333333333",
      playerId: "20002",
      locationNumber: "522"
    })
    const kingshotC = await kingshotPlayers.register({
      discordUserId: "444444444444444444",
      playerId: "30003",
      locationNumber: "523"
    })
    const wos = createGiftCodeRepository(pool, "wos")
    const ks = createGiftCodeRepository(pool, "kingshot")
    await wos.setAutoRedemption({ discordUserId: owner, playerId: opted.player_id, enabled: true })
    await wos.setAutoRedemption({ discordUserId: otherOwner, playerId: inactive.player_id, enabled: true })
    await wosPlayers.remove({ discordUserId: otherOwner, playerId: inactive.player_id })
    await ks.setAutoRedemption({ discordUserId: owner, playerId: kingshot.player_id, enabled: true })
    await ks.setAutoRedemption({
      discordUserId: "333333333333333333", playerId: kingshotB.player_id, enabled: true
    })
    await ks.setAutoRedemption({
      discordUserId: "444444444444444444", playerId: kingshotC.player_id, enabled: true
    })

    const source = await wos.createSource({ sourceType: "discord_user", sourceName: "Discord submissions" })
    const submission = await wos.recordSubmission({
      code: "CaseSensitive",
      submittedByDiscordUserId: owner,
      sourceId: source.id,
      metadata: { submissionKind: "user" }
    })
    const duplicate = await wos.recordSubmission({ code: "CaseSensitive", submittedByDiscordUserId: owner })
    const crossProfile = await ks.recordSubmission({ code: "CaseSensitive", submittedByDiscordUserId: owner })
    assert.equal(submission.giftCode.code, "CaseSensitive")
    assert.equal(duplicate.duplicate, true)
    assert.equal(crossProfile.duplicate, false)
    assert.notEqual(crossProfile.giftCode.id, submission.giftCode.id)
    assert.equal((await pool.query(
      "SELECT raw_metadata->>'submissionKind' AS kind FROM gift_code_submissions WHERE id = $1",
      [submission.submission.id]
    )).rows[0].kind, "user")

    const claimTime = new Date("2026-08-11T10:00:00Z")
    const verificationClaims = await Promise.all([
      wos.claimVerification({ workerId: "verify-a", now: claimTime, leaseSeconds: 60 }),
      wos.claimVerification({ workerId: "verify-b", now: claimTime, leaseSeconds: 60 })
    ])
    assert.equal(verificationClaims.filter(Boolean).length, 1)
    const verification = verificationClaims.find(Boolean)
    const verified = await wos.finishVerification({
      claim: verification,
      workerId: verification.verification_claimed_by_worker,
      result: apiResult("already_redeemed", { errCode: 40008, message: "RECEIVED." }),
      now: new Date("2026-08-11T10:00:01Z"),
      codeStatus: "active",
      verificationState: "complete",
      botInstanceName: "rachie-wos"
    })
    assert.equal(verified.giftCode.status, "active")
    assert.equal(verified.queued.length, 1)
    assert.equal(verified.queued[0].player_account_id, opted.id)
    assert.equal(await wos.fanOutActiveCode({ giftCodeId: submission.giftCode.id }).then(rows => rows.length), 0)
    assert.equal((await pool.query(
      "SELECT COUNT(*)::integer AS count FROM gift_code_redemptions WHERE game_profile = 'kingshot'"
    )).rows[0].count, 0)

    const kingshotMappings = centuryAdapter("kingshot", {}).responseMappings
    const kingshotTerminalCases = [
      {
        code: "ExpiredKS",
        payload: { code: 1, data: [], msg: "TIME ERROR.", err_code: 40007 },
        expectedClassification: "expired",
        expectedStatus: "expired"
      },
      {
        code: "InvalidKS",
        payload: { code: 1, data: [], msg: "CDK NOT FOUND.", err_code: 40014 },
        expectedClassification: "invalid_code",
        expectedStatus: "invalid"
      }
    ]
    for (const terminalCase of kingshotTerminalCases) {
      const candidate = await ks.recordSubmission({
        code: terminalCase.code,
        submittedByDiscordUserId: owner
      })
      const claim = await ks.claimVerification({
        workerId: `verify-${terminalCase.code}`,
        now: claimTime,
        leaseSeconds: 60,
        code: terminalCase.code,
        manual: true
      })
      const classification = classifyCenturyResponse({
        httpStatus: 200,
        data: terminalCase.payload,
        profileMappings: kingshotMappings
      })
      const transition = verificationTransition(classification.state, 1, claimTime, {
        maximumAttempts: 5,
        retryBaseSeconds: 60,
        retryCapSeconds: 3600
      })
      const terminal = await ks.finishVerification({
        claim,
        workerId: claim.verification_claimed_by_worker,
        result: apiResult(classification.state, {
          errCode: classification.raw.errCode,
          message: classification.raw.message
        }),
        now: new Date("2026-08-11T10:00:01Z"),
        ...transition,
        botInstanceName: "peggie-kingshot"
      })
      assert.equal(classification.state, terminalCase.expectedClassification)
      assert.equal(classification.retryable, false)
      assert.equal(terminal.giftCode.status, terminalCase.expectedStatus)
      assert.equal(terminal.giftCode.verification_state, "complete")
      assert.equal(terminal.giftCode.verification_next_retry_at_utc, null)
      assert.deepEqual(terminal.queued, [])
      assert.equal(await ks.claimVerification({
        workerId: `retry-${terminalCase.code}`,
        now: new Date("2026-08-11T11:00:00Z"),
        leaseSeconds: 60,
        code: terminalCase.code,
        manual: true
      }), null)
      assert.equal((await pool.query(
        `SELECT COUNT(*)::integer AS count FROM gift_code_attempts
          WHERE game_profile = 'kingshot' AND gift_code_id = $1`,
        [candidate.giftCode.id]
      )).rows[0].count, 1)
    }
    assert.equal((await pool.query(
      "SELECT COUNT(*)::integer AS count FROM gift_code_redemptions WHERE game_profile = 'kingshot'"
    )).rows[0].count, 0)

    const limitedCandidate = await ks.recordSubmission({
      code: "MixedTypeKS",
      submittedByDiscordUserId: owner
    })
    const limitedClaim = await ks.claimVerification({
      workerId: "verify-type-limit",
      now: claimTime,
      leaseSeconds: 60,
      code: limitedCandidate.giftCode.code,
      manual: true
    })
    const limitedClassification = classifyCenturyResponse({
      httpStatus: 200,
      data: {
        code: 1,
        data: [],
        msg: "The same type of Gift Code can only be redeemed once.",
        err_code: 40011
      },
      profileMappings: kingshotMappings
    })
    const limitedVerification = verificationTransition(
      limitedClassification.state,
      limitedClaim.verification_attempt_count,
      claimTime,
      { maximumAttempts: 5, retryBaseSeconds: 60, retryCapSeconds: 3600 }
    )
    const activatedByLimit = await ks.finishVerification({
      claim: limitedClaim,
      workerId: limitedClaim.verification_claimed_by_worker,
      result: apiResult("redemption_limit", {
        errCode: 40011,
        message: limitedClassification.raw.message
      }),
      now: new Date("2026-08-11T10:00:01Z"),
      ...limitedVerification,
      botInstanceName: "peggie-kingshot"
    })
    assert.equal(activatedByLimit.giftCode.status, "active")
    assert.equal(activatedByLimit.giftCode.verification_state, "complete")
    assert.equal(activatedByLimit.queued.length, 3)

    const mixedResults = [
      { classification: "redemption_limit", errCode: 40011, durableStatus: "restricted" },
      { classification: "success", errCode: 20000, durableStatus: "success" },
      { classification: "already_redeemed", errCode: 40008, durableStatus: "already_redeemed" }
    ]
    for (let index = 0; index < mixedResults.length; index += 1) {
      const expected = mixedResults[index]
      const claim = await ks.claimRedemption({
        workerId: `mixed-redeem-${index}`,
        now: new Date(`2026-08-11T10:01:0${index}Z`),
        leaseSeconds: 60
      })
      assert.ok(claim)
      const transition = redemptionTransition(
        expected.classification,
        claim.attempt_number,
        claimTime,
        { maximumAttempts: 5, retryBaseSeconds: 60, retryCapSeconds: 3600 }
      )
      assert.equal(transition.retryable, false)
      assert.equal(transition.nextRetryAt, undefined)
      await ks.finishRedemption({
        claim,
        workerId: claim.claimed_by_worker,
        result: apiResult(expected.classification, {
          errCode: expected.errCode,
          message: expected.classification
        }),
        now: new Date(`2026-08-11T10:01:1${index}Z`),
        ...transition
      })
      assert.equal(transition.status, expected.durableStatus)
    }
    assert.equal(await ks.claimRedemption({
      workerId: "after-mixed-restart",
      now: new Date("2026-08-11T11:00:00Z"),
      leaseSeconds: 60
    }), null)
    const mixedRows = (await pool.query(
      `SELECT r.status, a.classification
         FROM gift_code_redemptions r
         JOIN gift_code_attempts a
           ON a.redemption_id = r.id AND a.game_profile = r.game_profile
        WHERE r.game_profile = 'kingshot' AND r.gift_code_id = $1
        ORDER BY a.created_at_utc, a.id`,
      [limitedCandidate.giftCode.id]
    )).rows
    assert.deepEqual(mixedRows.map(row => row.classification), [
      "redemption_limit", "success", "already_redeemed"
    ])
    assert.deepEqual(mixedRows.map(row => row.status), [
      "restricted", "success", "already_redeemed"
    ])
    const mixedCounts = (await pool.query(
      `SELECT
         COUNT(*) FILTER (WHERE status = 'success')::integer AS success,
         COUNT(*) FILTER (WHERE status = 'already_redeemed')::integer AS already,
         COUNT(*) FILTER (WHERE status = 'restricted')::integer AS restricted
       FROM gift_code_redemptions
       WHERE game_profile = 'kingshot' AND gift_code_id = $1`,
      [limitedCandidate.giftCode.id]
    )).rows[0]
    assert.deepEqual(mixedCounts, { success: 1, already: 1, restricted: 1 })
    assert.equal((await ks.getCode("MixedTypeKS")).status, "active")

    await wosPlayers.changeLocation({
      discordUserId: owner,
      playerId: opted.player_id,
      locationNumber: "700"
    })
    const redemptionClaims = await Promise.all([
      wos.claimRedemption({ workerId: "redeem-a", now: claimTime, leaseSeconds: 60 }),
      wos.claimRedemption({ workerId: "redeem-b", now: claimTime, leaseSeconds: 60 })
    ])
    assert.equal(redemptionClaims.filter(Boolean).length, 1)
    const firstRedemption = redemptionClaims.find(Boolean)
    assert.equal(firstRedemption.location_number_snapshot, "700")
    assert.equal(firstRedemption.owner_account_count, 2)
    const retryAt = new Date("2026-08-11T10:01:00Z")
    await wos.finishRedemption({
      claim: firstRedemption,
      workerId: firstRedemption.claimed_by_worker,
      result: apiResult("rate_limited", { httpStatus: 429 }),
      now: new Date("2026-08-11T10:00:02Z"),
      status: "rate_limited",
      retryable: true,
      nextRetryAt: retryAt
    })
    await wosPlayers.changeLocation({
      discordUserId: owner,
      playerId: opted.player_id,
      locationNumber: "701"
    })
    assert.equal(await wos.claimRedemption({
      workerId: "too-early",
      now: new Date("2026-08-11T10:00:59Z"),
      leaseSeconds: 60
    }), null)
    const secondRedemption = await wos.claimRedemption({
      workerId: "redeem-c",
      now: retryAt,
      leaseSeconds: 60
    })
    assert.equal(secondRedemption.id, firstRedemption.id)
    assert.equal(secondRedemption.location_number_snapshot, "701")
    const completed = await wos.finishRedemption({
      claim: secondRedemption,
      workerId: "redeem-c",
      result: apiResult("success", { code: 0, errCode: 20000, message: "SUCCESS" }),
      now: new Date("2026-08-11T10:01:01Z"),
      status: "success"
    })
    assert.equal(completed.status, "success")
    const attempts = (await pool.query(
      `SELECT attempt_number, classification, location_number_snapshot
         FROM gift_code_attempts
        WHERE redemption_id = $1 ORDER BY attempt_number`,
      [completed.id]
    )).rows
    assert.deepEqual(attempts, [
      { attempt_number: 1, classification: "rate_limited", location_number_snapshot: "700" },
      { attempt_number: 2, classification: "success", location_number_snapshot: "701" }
    ])
    const diagnostics = await wos.codeDiagnostics(submission.giftCode.code)
    assert.equal(diagnostics.latest_verification_api_message, "RECEIVED.")
    assert.equal(diagnostics.latest_verification_classification, "already_redeemed")
    assert.equal(diagnostics.latest_redemption_api_message, "SUCCESS")
    assert.equal(diagnostics.latest_redemption_classification, "success")
    assert.equal(diagnostics.latest_redemption_player_id_snapshot, "10001")
    assert.equal(await wos.claimRedemption({
      workerId: "after-restart",
      now: new Date("2026-08-11T11:00:00Z"),
      leaseSeconds: 60
    }), null)

    const notificationClaims = await Promise.all([
      wos.claimNotification(completed.id),
      wos.claimNotification(completed.id)
    ])
    assert.equal(notificationClaims.filter(Boolean).length, 1)
    await wos.finishNotification(completed.id, { sent: false, errorCode: "50007" })
    const afterDmFailure = (await pool.query(
      "SELECT status, notification_status FROM gift_code_redemptions WHERE id = $1",
      [completed.id]
    )).rows[0]
    assert.deepEqual(afterDmFailure, { status: "success", notification_status: "failed" })

    await wos.setAutoRedemption({ discordUserId: owner, playerId: optedOut.player_id, enabled: true })
    assert.equal((await pool.query(
      `SELECT COUNT(*)::integer AS count FROM gift_code_redemptions
        WHERE game_profile = 'wos' AND player_account_id = $1 AND status = 'queued'`,
      [optedOut.id]
    )).rows[0].count, 1, "late opt-in did not queue an already-active code")
    await wos.setAutoRedemption({ discordUserId: owner, playerId: optedOut.player_id, enabled: false })

    const staleSubmission = await wos.recordSubmission({ code: "StaleClaim", submittedByDiscordUserId: owner })
    const staleVerify = await wos.claimVerification({
      workerId: "verify-stale",
      now: claimTime,
      leaseSeconds: 30,
      code: staleSubmission.giftCode.code
    })
    await wos.finishVerification({
      claim: staleVerify,
      workerId: "verify-stale",
      result: apiResult("success", { code: 0, errCode: 20000 }),
      now: new Date("2026-08-11T10:00:01Z"),
      codeStatus: "active",
      verificationState: "complete"
    })
    const abandoned = await wos.claimRedemption({
      workerId: "crashed-worker",
      now: new Date("2026-08-11T12:00:00Z"),
      leaseSeconds: 30
    })
    const reclaimed = await wos.claimRedemption({
      workerId: "replacement-worker",
      now: new Date("2026-08-11T12:00:31Z"),
      leaseSeconds: 30
    })
    assert.equal(reclaimed.id, abandoned.id)
    assert.equal(reclaimed.attempt_number, abandoned.attempt_number + 1)

    await wos.setAutoRedemption({ discordUserId: owner, playerId: opted.player_id, enabled: false })
    const disabled = (await pool.query(
      "SELECT status FROM gift_code_redemptions WHERE id = $1",
      [reclaimed.id]
    )).rows[0]
    assert.equal(disabled.status, "disabled")
    assert.equal((await wos.accountStatuses(owner, optedOut.player_id))[0].gift_redemption_enabled, false)
  } finally {
    await pool?.end().catch(() => {})
    await admin.query(`DROP SCHEMA IF EXISTS "${schema}" CASCADE`).catch(() => {})
    await admin.end()
  }
})
