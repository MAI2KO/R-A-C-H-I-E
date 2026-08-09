const { generateMissingStateEventDeliveryClaims } = require("./stateEventDeliveryGeneration")
const { MAX_DELIVERY_ATTEMPTS } = require("./eventDeliveryRepository")
const {
  PermanentDeliveryError,
  sanitizeDeliveryError,
  retryDelayMinutes,
  timeoutPromise
} = require("./eventDeliveryWorker")

function createStateEventDeliveryProcessor({
  repository,
  gameProfile,
  botInstanceName,
  deliveryHandler,
  config,
  logger = console,
  now = () => new Date(),
  workerId = `state-${process.pid}-${Date.now()}`
}) {
  const handlerTimeoutMs = Math.min(
    config.handlerTimeoutMs,
    Math.max(1000, config.claimLeaseSeconds * 1000 - 1000)
  )

  async function processClaim(claim) {
    const ownership = { claimId: claim.id, botInstanceName, workerId }
    const payload = await repository.getClaimPayload(ownership)
    if (!payload) return false
    try {
      const result = await timeoutPromise(deliveryHandler(payload), handlerTimeoutMs)
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
        return repository.markClaimPermanentlyFailed({ ...ownership, failedAt, lastError })
      }
      const nextAttemptAt = new Date(
        failedAt.getTime() + retryDelayMinutes(claim.attempt_count) * 60000
      )
      return repository.markClaimFailed({ ...ownership, failedAt, lastError, nextAttemptAt })
    }
  }

  async function tick(tickNow = now()) {
    await generateMissingStateEventDeliveryClaims({
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
  }

  return Object.freeze({
    workerId,
    tick: tickNow => tick(tickNow).catch(error => {
      logger.error(`[Event scheduler] State-event tick failed: ${sanitizeDeliveryError(error)}`)
      return 0
    })
  })
}

module.exports = {
  createStateEventDeliveryProcessor
}
