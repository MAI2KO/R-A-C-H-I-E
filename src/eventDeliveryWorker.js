const { randomUUID } = require("crypto")

const { schedulerIsEnabled } = require("./db")
const { SUPPORTED_PROFILES } = require("./eventSchedulerRepository")
const { getEventDeliveryConfig } = require("./eventDeliveryConfig")
const { generateMissingDeliveryClaims } = require("./eventDeliveryGeneration")
const { MAX_DELIVERY_ATTEMPTS } = require("./eventDeliveryRepository")

const RETRY_BACKOFF_MINUTES = Object.freeze([1, 5, 15, 30, 60])
const MAX_ERROR_LENGTH = 500

class RetryableDeliveryError extends Error {
  constructor(message) {
    super(message)
    this.name = "RetryableDeliveryError"
  }
}

class PermanentDeliveryError extends Error {
  constructor(message) {
    super(message)
    this.name = "PermanentDeliveryError"
  }
}

function sanitizeDeliveryError(error) {
  const message = String(error?.message || error || "Delivery failed")
    .replace(/\b(?:https?|postgres(?:ql)?):\/\/\S+/gi, "[redacted URL]")
    .replace(
      /\b(?:BOT_TOKEN|DATABASE_URL|APPS_SCRIPT_URL|TOKEN|PASSWORD|SECRET)\s*[:=]\s*\S+/gi,
      "[redacted credential]"
    )
    .replace(/[\r\n\t]+/g, " ")
    .replace(/\s{2,}/g, " ")
    .trim()
  return (message || "Delivery failed").slice(0, MAX_ERROR_LENGTH)
}

function retryDelayMinutes(attemptCount) {
  if (!Number.isInteger(attemptCount) || attemptCount < 1) {
    throw new Error("Attempt count must be a positive integer")
  }
  return RETRY_BACKOFF_MINUTES[Math.min(
    attemptCount - 1,
    RETRY_BACKOFF_MINUTES.length - 1
  )]
}

function timeoutPromise(promise, timeoutMs) {
  let timer
  const timeout = new Promise((_resolve, reject) => {
    timer = setTimeout(() => reject(new RetryableDeliveryError("Delivery handler timed out.")), timeoutMs)
  })
  return Promise.race([Promise.resolve(promise), timeout]).finally(() => clearTimeout(timer))
}

function waitFor(promise, timeoutMs) {
  if (!promise) return Promise.resolve(true)
  let timer
  const timeout = new Promise(resolve => {
    timer = setTimeout(() => resolve(false), timeoutMs)
  })
  return Promise.race([
    Promise.resolve(promise).then(() => true, () => true),
    timeout
  ]).finally(() => clearTimeout(timer))
}

function workerStartReason({
  env,
  health,
  repository,
  gameProfile,
  botInstanceName,
  deliveryHandler
}) {
  if (!schedulerIsEnabled(env)) return "disabled"
  if (!health?.available) return "database unavailable"
  if (!SUPPORTED_PROFILES.has(gameProfile)) return "invalid game profile"
  if (health.gameProfile && health.gameProfile !== gameProfile) return "health profile mismatch"
  if (!String(botInstanceName || "").trim()) return "missing bot instance"
  if (health.botInstanceName && health.botInstanceName !== botInstanceName) {
    return "health bot instance mismatch"
  }
  if (!repository) return "missing delivery repository"
  if (repository.gameProfile !== gameProfile) return "repository profile mismatch"
  if (typeof deliveryHandler !== "function") return "missing delivery handler"
  return null
}

function createEventDeliveryWorker({
  env = process.env,
  health,
  repository,
  gameProfile,
  botInstanceName,
  deliveryHandler,
  additionalTick = null,
  logger = console,
  now = () => new Date(),
  workerId = `${process.pid}-${randomUUID()}`,
  config = getEventDeliveryConfig(env),
  setIntervalFn = setInterval,
  clearIntervalFn = clearInterval
}) {
  let timer = null
  let activeTick = null
  let stopped = false
  const startReason = workerStartReason({
    env,
    health,
    repository,
    gameProfile,
    botInstanceName,
    deliveryHandler
  })
  const handlerTimeoutMs = Math.min(
    config.handlerTimeoutMs,
    Math.max(1000, config.claimLeaseSeconds * 1000 - 1000)
  )

  async function processClaim(claim) {
    const ownership = {
      claimId: claim.id,
      botInstanceName,
      workerId
    }
    const payload = await repository.getClaimPayload(ownership)
    if (!payload) return false

    try {
      const result = await timeoutPromise(
        deliveryHandler(payload),
        handlerTimeoutMs
      )
      return repository.markClaimSent({
        ...ownership,
        sentAt: now(),
        sentMessageId: result?.sentMessageId || null
      })
    } catch (error) {
      const failedAt = now()
      const lastError = sanitizeDeliveryError(error)
      const permanent = error instanceof PermanentDeliveryError
        || claim.attempt_count >= MAX_DELIVERY_ATTEMPTS
      if (permanent) {
        return repository.markClaimPermanentlyFailed({
          ...ownership,
          failedAt,
          lastError
        })
      }
      const nextAttemptAt = new Date(
        failedAt.getTime() + retryDelayMinutes(claim.attempt_count) * 60000
      )
      return repository.markClaimFailed({
        ...ownership,
        failedAt,
        lastError,
        nextAttemptAt
      })
    }
  }

  async function runTick() {
    const tickNow = now()
    const roundupTick = typeof additionalTick === "function"
      ? Promise.resolve().then(() => additionalTick(tickNow)).catch(error => {
        logger.error(`[Event scheduler] Roundup tick failed: ${sanitizeDeliveryError(error)}`)
      })
      : Promise.resolve()
    try {
      await generateMissingDeliveryClaims({
        repository,
        gameProfile,
        now: tickNow,
        config
      })
      const claims = await repository.claimDueDeliveries({
        now: tickNow,
        batchSize: config.batchSize,
        leaseSeconds: config.claimLeaseSeconds,
        botInstanceName,
        workerId
      })
      await Promise.all(claims.map(processClaim))
      return claims.length
    } finally {
      await roundupTick
    }
  }

  function tick() {
    if (stopped || startReason || activeTick) return activeTick || Promise.resolve(0)
    activeTick = runTick()
      .catch(error => {
        logger.error(`[Event scheduler] Delivery tick failed: ${sanitizeDeliveryError(error)}`)
        return 0
      })
      .finally(() => {
        activeTick = null
      })
    return activeTick
  }

  function start() {
    if (timer || stopped) return { started: false, reason: stopped ? "stopped" : "already started" }
    if (startReason) return { started: false, reason: startReason }
    timer = setIntervalFn(() => { void tick() }, config.pollIntervalMs)
    timer.unref?.()
    void tick()
    return { started: true, workerId }
  }

  async function stop({ timeoutMs = 5000 } = {}) {
    stopped = true
    if (timer) clearIntervalFn(timer)
    timer = null
    return { drained: await waitFor(activeTick, timeoutMs) }
  }

  return Object.freeze({
    workerId,
    config,
    start,
    tick,
    stop,
    isRunning: () => Boolean(timer) && !stopped,
    hasActiveTick: () => Boolean(activeTick)
  })
}

module.exports = {
  RETRY_BACKOFF_MINUTES,
  MAX_ERROR_LENGTH,
  RetryableDeliveryError,
  PermanentDeliveryError,
  sanitizeDeliveryError,
  retryDelayMinutes,
  timeoutPromise,
  waitFor,
  workerStartReason,
  createEventDeliveryWorker
}
