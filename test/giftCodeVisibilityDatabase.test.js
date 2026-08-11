const test = require("node:test")
const assert = require("node:assert/strict")
const { Pool } = require("pg")

const { runMigrations } = require("../src/migrate")
const { createGiftCodeRepository } = require("../src/giftCodes/repository")
const { createGiftCodeCommunityRepository } = require("../src/giftCodes/communityRepository")

const databaseUrl = process.env.TEST_DATABASE_URL
const logger = { log() {}, error() {} }

function apiResult(classification, { httpStatus = 200, errCode = null } = {}) {
  return {
    httpStatus,
    headers: {},
    classification: {
      state: classification,
      raw: { code: null, errCode, message: null }
    }
  }
}

async function completeVerification(repository, code, classification, codeStatus, verificationState, at) {
  const claim = await repository.claimVerification({
    workerId: `verify-${code}`,
    now: at,
    leaseSeconds: 60,
    code,
    manual: true
  })
  return repository.finishVerification({
    claim,
    workerId: `verify-${code}`,
    result: apiResult(classification, {
      httpStatus: classification === "upstream_rejection" ? 403 : 200
    }),
    now: at,
    codeStatus,
    verificationState
  })
}

test("active-code visibility and verification feedback are durable and profile scoped", {
  skip: databaseUrl ? false : "TEST_DATABASE_URL is not configured"
}, async () => {
  const admin = new Pool({ connectionString: databaseUrl, max: 1 })
  const schema = `gift_visibility_${process.pid}_${Date.now()}`
  let pool
  try {
    await admin.query(`CREATE SCHEMA "${schema}"`)
    pool = new Pool({
      connectionString: databaseUrl,
      max: 8,
      options: `-c search_path=${schema}`
    })
    await runMigrations({ pool, logger })
    const wos = createGiftCodeRepository(pool, "wos")
    const kingshot = createGiftCodeRepository(pool, "kingshot")
    const community = createGiftCodeCommunityRepository(pool, "wos")
    const owner = "111111111111111111"
    const guild = "777777777777777777"

    const statuses = [
      ["ACTIVE_OLD", "active", "2026-08-10T10:00:00Z"],
      ["ACTIVE_NEW", "active", "2026-08-11T10:00:00Z"],
      ["EXPIRED", "expired", null],
      ["INVALID", "invalid", null],
      ["RESTRICTED", "restricted", null],
      ["REVIEW", "unknown", null]
    ]
    for (const [code, status, verifiedAt] of statuses) {
      const recorded = await wos.recordSubmission({ code })
      await pool.query(
        `UPDATE gift_codes
            SET status = $2, verification_state = $3, verified_at_utc = $4
          WHERE id = $1`,
        [recorded.giftCode.id, status, status === "unknown" ? "review" : "complete", verifiedAt]
      )
    }
    const ksActive = await kingshot.recordSubmission({ code: "KINGSHOT_ONLY" })
    await pool.query(
      "UPDATE gift_codes SET status = 'active', verification_state = 'complete' WHERE id = $1",
      [ksActive.giftCode.id]
    )

    const firstPage = await wos.activeCodeVisibility({ page: 0, pageSize: 1 })
    const secondPage = await wos.activeCodeVisibility({ page: 1, pageSize: 1 })
    assert.deepEqual(firstPage.codes.map(row => row.code), ["ACTIVE_NEW"])
    assert.deepEqual(secondPage.codes.map(row => row.code), ["ACTIVE_OLD"])
    assert.equal(firstPage.activeCount, 2)
    assert.equal(firstPage.expiredCount, 1)
    assert.doesNotMatch(JSON.stringify(firstPage), /EXPIRED|INVALID|RESTRICTED|REVIEW|KINGSHOT_ONLY/)
    assert.deepEqual((await kingshot.activeCodeVisibility()).codes.map(row => row.code), ["KINGSHOT_ONLY"])
    const diagnostics = await wos.diagnostics()
    assert.equal(diagnostics.active_codes, 2)
    assert.equal(diagnostics.expired_codes, 1)
    assert.equal(diagnostics.invalid_codes, 1)
    assert.equal(diagnostics.restricted_review_codes, 2)

    const invalid = await wos.recordSubmission({
      code: "USER_INVALID",
      submittedByDiscordUserId: owner,
      metadata: { submissionKind: "user", guildId: guild }
    })
    await wos.recordSubmission({
      code: "USER_INVALID",
      submittedByDiscordUserId: "222222222222222222",
      metadata: { submissionKind: "user", guildId: guild }
    })
    await completeVerification(
      wos,
      invalid.giftCode.code,
      "invalid_code",
      "invalid",
      "complete",
      new Date("2026-08-11T11:00:00Z")
    )
    const claims = await Promise.all([
      community.claimVerificationResultNotification("notify-a", new Date("2026-08-11T11:00:01Z")),
      community.claimVerificationResultNotification("notify-b", new Date("2026-08-11T11:00:01Z"))
    ])
    assert.equal(claims.filter(Boolean).length, 1)
    const notification = claims.find(Boolean)
    assert.equal(notification.classification, "invalid_code")
    assert.equal(notification.submitted_by_discord_user_id, owner)
    assert.equal((await pool.query(
      `SELECT COUNT(*)::integer AS count FROM gift_code_submissions
        WHERE game_profile = 'wos' AND duplicate_of_gift_code_id IS NOT NULL
          AND raw_metadata ? 'verificationResultNotification'`
    )).rows[0].count, 0)
    const notificationWorker = claims[0] ? "notify-a" : "notify-b"
    await community.finishVerificationResultNotification(notification.id, notificationWorker, {
      sent: true,
      now: new Date("2026-08-11T11:00:02Z")
    })
    assert.equal(
      await community.claimVerificationResultNotification("after-restart", new Date("2026-08-11T12:00:00Z")),
      null
    )

    const review = await wos.recordSubmission({
      code: "REVIEW_THEN_ACTIVE",
      submittedByDiscordUserId: owner,
      metadata: { submissionKind: "user", guildId: guild }
    })
    await completeVerification(
      wos,
      review.giftCode.code,
      "upstream_rejection",
      "unknown",
      "review",
      new Date("2026-08-11T12:00:00Z")
    )
    const reviewNotice = await community.claimVerificationResultNotification(
      "review-notify",
      new Date("2026-08-11T12:00:01Z")
    )
    assert.equal(reviewNotice.classification, "upstream_rejection")
    await community.finishVerificationResultNotification(reviewNotice.id, "review-notify", {
      sent: true,
      now: new Date("2026-08-11T12:00:02Z")
    })
    await completeVerification(
      wos,
      review.giftCode.code,
      "success",
      "active",
      "complete",
      new Date("2026-08-11T13:00:00Z")
    )
    const engagement = await community.prepareCodeEngagement(review.giftCode.id, 0)
    assert.equal(engagement.filter(event => event.event_type === "contributor_role").length, 1)
    assert.deepEqual(await community.prepareCodeEngagement(review.giftCode.id, 0), [])
  } finally {
    await pool?.end().catch(() => {})
    await admin.query(`DROP SCHEMA IF EXISTS "${schema}" CASCADE`).catch(() => {})
    await admin.end()
  }
})
