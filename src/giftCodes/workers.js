const { randomUUID } = require("node:crypto")
const { retryAfterMilliseconds } = require("./rateLimiter")
const { classifyCenturyResponse } = require("./responseClassifier")

const RECOVERABLE_STORED_VERIFICATION_CODES = Object.freeze([40011])

function retryAt(now, attemptNumber, config) {
  const seconds = Math.min(
    config.retryCapSeconds,
    config.retryBaseSeconds * (2 ** Math.max(0, attemptNumber - 1))
  )
  return new Date(now.getTime() + seconds * 1000)
}

function retryDate(now, attemptNumber, config, retryAfterMs = 0) {
  const exponential = retryAt(now, attemptNumber, config)
  return new Date(Math.max(exponential.getTime(), now.getTime() + Math.max(0, retryAfterMs || 0)))
}

function verificationTransition(classification, attemptNumber, now, config, retryAfterMs = 0) {
  if ([
    "success", "already_redeemed", "redemption_limit", "claim_limit",
    "level_restriction", "account_restriction", "account_age_restriction"
  ].includes(classification)) {
    return { codeStatus: "active", verificationState: "complete" }
  }
  if (classification === "expired") return { codeStatus: "expired", verificationState: "complete" }
  if (classification === "invalid_code") return { codeStatus: "invalid", verificationState: "complete" }
  if (classification === "invalid_player") return { codeStatus: "candidate", verificationState: "blocked" }
  if (classification === "eligibility_restriction") {
    return { codeStatus: "restricted", verificationState: "review" }
  }
  if ([
    "rate_limited", "temporary_error", "verification_throttle",
    "simultaneous_action_throttle"
  ].includes(classification)) {
    if (attemptNumber >= config.maximumAttempts) {
      return { codeStatus: "candidate", verificationState: "blocked" }
    }
    return {
      codeStatus: "candidate",
      verificationState: "retry",
      nextRetryAt: retryDate(now, attemptNumber, config, retryAfterMs)
    }
  }
  return { codeStatus: "unknown", verificationState: "review" }
}

function redemptionTransition(classification, attemptNumber, now, config, retryAfterMs = 0) {
  const terminal = new Set([
    "success", "already_redeemed", "expired", "invalid_code", "invalid_player"
  ])
  if (terminal.has(classification)) return { status: classification, retryable: false }
  if ([
    "eligibility_restriction", "redemption_limit", "claim_limit", "level_restriction",
    "account_restriction", "account_age_restriction"
  ].includes(classification)) {
    return { status: "restricted", retryable: false }
  }
  if ([
    "rate_limited", "temporary_error", "verification_throttle",
    "simultaneous_action_throttle"
  ].includes(classification)) {
    if (attemptNumber >= config.maximumAttempts) {
      return { status: "retry_exhausted", retryable: false }
    }
    return {
      status: classification === "verification_throttle" ? "rate_limited" :
        classification === "simultaneous_action_throttle" ? "temporary_error" : classification,
      retryable: true,
      nextRetryAt: retryDate(now, attemptNumber, config, retryAfterMs)
    }
  }
  return { status: "unknown", retryable: false }
}

function sanitizeWorkerError(error) {
  return String(error?.code || error?.name || "worker_error").slice(0, 100)
}

function redeemWithProfileLock(repository, client, input) {
  const operation = () => client.redeem(input)
  if (typeof repository.withProfileRequestLock !== "function") return operation()
  return repository.withProfileRequestLock(operation, {
    minimumDelayMs: client.limiter?.minimumDelayMs || 0
  })
}

function createVerificationProcessor({
  repository,
  client,
  verifier,
  config,
  botInstanceName,
  community = null,
  logger = console,
  now = () => new Date(),
  workerId = `verify-${process.pid}-${randomUUID()}`
}) {
  async function recoverStoredReview(code = null) {
    if (typeof repository.storedVerificationReview !== "function" || !client.adapter) return null
    const claim = await repository.storedVerificationReview({
      errCodes: RECOVERABLE_STORED_VERIFICATION_CODES,
      code
    })
    if (!claim) return null
    const classified = classifyCenturyResponse({
      httpStatus: claim.stored_http_status,
      data: {
        code: claim.stored_api_code,
        err_code: claim.stored_err_code,
        msg: claim.stored_api_message
      },
      profileMappings: client.adapter.responseMappings || {}
    })
    if (classified.state !== "redemption_limit") return null
    const completedAt = now()
    const transition = verificationTransition(
      classified.state,
      claim.verification_attempt_count,
      completedAt,
      config
    )
    const completed = await repository.finishStoredVerificationRecovery({
      claim,
      classification: classified.state,
      now: completedAt,
      ...transition,
      botInstanceName
    })
    if (!completed) return null
    if (completed.giftCode.status === "active" && community) {
      try {
        await community.onCodeActivated(completed.giftCode.id, completed.queued.length)
      } catch (error) {
        logger.warn(`[Gift codes] Community activation failed: ${sanitizeWorkerError(error)}`)
      }
    }
    logger.log(JSON.stringify({
      event: "gift_code_verification_reclassified",
      game_profile: repository.gameProfile,
      gift_code_id: claim.id,
      classification: classified.state,
      err_code: classified.raw.errCode
    }))
    return completed
  }

  async function processClaim(claim) {
    const result = await redeemWithProfileLock(repository, client, {
      playerId: verifier.playerId,
      locationNumber: verifier.locationNumber,
      code: claim.code
    })
    const completedAt = now()
    const retryAfterMs = retryAfterMilliseconds(result.headers, completedAt.getTime()) || 0
    const transition = verificationTransition(
      result.classification.state,
      claim.verification_attempt_count,
      completedAt,
      config,
      retryAfterMs
    )
    const completed = await repository.finishVerification({
      claim,
      workerId,
      result,
      now: completedAt,
      ...transition,
      botInstanceName
    })
    if (completed.giftCode.status === "active" && community) {
      try {
        await community.onCodeActivated(completed.giftCode.id, completed.queued.length)
      } catch (error) {
        logger.warn(`[Gift codes] Community activation failed: ${sanitizeWorkerError(error)}`)
      }
    } else if (community && transition.verificationState !== "retry") {
      try {
        await community.onVerificationResult(completed.giftCode.id)
      } catch (error) {
        logger.warn(`[Gift codes] Verification feedback failed: ${sanitizeWorkerError(error)}`)
      }
    }
    if (result.classification.state === "invalid_player") {
      logger.error(JSON.stringify({
        event: "gift_code_verifier_configuration_error",
        game_profile: repository.gameProfile,
        gift_code_id: claim.id,
        classification: result.classification.state,
        http_status: result.httpStatus,
        err_code: result.classification.raw.errCode,
        attempt_number: claim.verification_attempt_count
      }))
    }
    logger.log(JSON.stringify({
      event: "gift_code_verification",
      game_profile: repository.gameProfile,
      gift_code_id: claim.id,
      classification: result.classification.state,
      http_status: result.httpStatus,
      err_code: result.classification.raw.errCode,
      attempt_number: claim.verification_attempt_count
    }))
    return completed
  }

  async function tick() {
    if (!verifier?.configured) return 0
    try {
      const recovered = await recoverStoredReview()
      if (recovered) return 1
      const claim = await repository.claimVerification({
        workerId,
        now: now(),
        leaseSeconds: config.leaseSeconds
      })
      if (!claim) return 0
      await processClaim(claim)
      return 1
    } catch (error) {
      logger.error(`[Gift codes] Verification job failed: ${sanitizeWorkerError(error)}`)
      return 0
    }
  }

  async function verifyCode(code) {
    if (!verifier?.configured) return { processed: false, reason: verifier?.reason || "verifier not configured" }
    const recovered = await recoverStoredReview(code)
    if (recovered) return { processed: true, result: recovered, recovered: true }
    const claim = await repository.claimVerification({
      workerId,
      now: now(),
      leaseSeconds: config.leaseSeconds,
      code,
      manual: true
    })
    if (!claim) return { processed: false, reason: "code is not available for verification" }
    return { processed: true, result: await processClaim(claim) }
  }

  return Object.freeze({ workerId, tick, verifyCode, recoverStoredReview })
}

function createRedemptionProcessor({
  repository,
  client,
  notifier,
  community = null,
  config,
  logger = console,
  now = () => new Date(),
  workerId = `redeem-${process.pid}-${randomUUID()}`
}) {
  async function notify(claim, status) {
    const notification = await repository.claimNotification(claim.id, now())
    if (!notification) return
    const outcome = await notifier({ ...claim, ...notification }, status)
    await repository.finishNotification(claim.id, {
      sent: outcome.sent,
      now: now(),
      errorCode: outcome.errorCode || null
    })
  }

  async function tick() {
    try {
      const claim = await repository.claimRedemption({
        workerId,
        now: now(),
        leaseSeconds: config.leaseSeconds
      })
      if (!claim) return 0
      const result = await redeemWithProfileLock(repository, client, {
        playerId: claim.player_id_snapshot,
        locationNumber: claim.location_number_snapshot,
        code: claim.code
      })
      const completedAt = now()
      const retryAfterMs = retryAfterMilliseconds(result.headers, completedAt.getTime()) || 0
      const transition = redemptionTransition(
        result.classification.state,
        claim.attempt_number,
        completedAt,
        config,
        retryAfterMs
      )
      await repository.finishRedemption({
        claim,
        workerId,
        result,
        now: completedAt,
        ...transition
      })
      await notify(
        claim,
        transition.status === "restricted" && [
          "redemption_limit", "claim_limit", "level_restriction",
          "account_restriction", "account_age_restriction"
        ].includes(result.classification.state)
          ? result.classification.state
          : transition.status
      )
      if (community && !transition.retryable) {
        try {
          await community.onRedemptionUpdated(
            claim.gift_code_id,
            claim.player_account_id,
            transition.status
          )
        } catch (error) {
          logger.warn(`[Gift codes] Community progress failed: ${sanitizeWorkerError(error)}`)
        }
      }
      logger.log(JSON.stringify({
        event: "gift_code_redemption",
        game_profile: repository.gameProfile,
        gift_code_id: claim.gift_code_id,
        redemption_id: claim.id,
        classification: result.classification.state,
        http_status: result.httpStatus,
        err_code: result.classification.raw.errCode,
        attempt_number: claim.attempt_number
      }))
      return 1
    } catch (error) {
      logger.error(`[Gift codes] Redemption job failed: ${sanitizeWorkerError(error)}`)
      return 0
    }
  }

  return Object.freeze({ workerId, tick })
}

function createPollingWorker({ tick, intervalMs, enabled, setIntervalFn = setInterval, clearIntervalFn = clearInterval }) {
  let timer = null
  let active = null
  let stopped = false
  function run() {
    if (!enabled || stopped || active) return active || Promise.resolve(0)
    active = Promise.resolve().then(tick).catch(() => 0).finally(() => { active = null })
    return active
  }
  return Object.freeze({
    start() {
      if (!enabled) return { started: false, reason: "disabled" }
      if (timer || stopped) return { started: false, reason: stopped ? "stopped" : "already started" }
      timer = setIntervalFn(() => { void run() }, intervalMs)
      timer.unref?.()
      void run()
      return { started: true }
    },
    async stop() {
      stopped = true
      if (timer) clearIntervalFn(timer)
      timer = null
      await active?.catch(() => {})
      return { stopped: true }
    },
    tick: run,
    isRunning: () => Boolean(timer) && !stopped
  })
}

module.exports = {
  retryAt,
  retryDate,
  verificationTransition,
  redemptionTransition,
  sanitizeWorkerError,
  redeemWithProfileLock,
  createVerificationProcessor,
  createRedemptionProcessor,
  createPollingWorker
}
