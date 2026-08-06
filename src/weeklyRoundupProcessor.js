const { createHash, randomUUID } = require("crypto")
const { generateMissingRoundupClaims } = require("./weeklyRoundupGeneration")
const {
  MAX_DELIVERY_ATTEMPTS
} = require("./eventDeliveryRepository")
const {
  PermanentDeliveryError,
  retryDelayMinutes,
  sanitizeDeliveryError,
  timeoutPromise
} = require("./eventDeliveryWorker")

function createWeeklyRoundupProcessor({
  repository,
  gameProfile,
  botInstanceName,
  delivery,
  config,
  logger = console,
  now = () => new Date(),
  workerId = `${process.pid}-roundup-${randomUUID()}`
}) {
  const handlerTimeoutMs = Math.min(
    config.handlerTimeoutMs,
    Math.max(1000, config.claimLeaseSeconds * 1000 - 1000)
  )

  function payloadHash(message) {
    return createHash("sha256").update(JSON.stringify(message)).digest("hex")
  }

  async function processClaim(claim) {
    const ownership = { claimId: claim.id, botInstanceName, workerId }
    const payload = await repository.getClaimPayload(ownership)
    if (!payload) return false
    try {
      const prepared = await timeoutPromise(delivery(payload), handlerTimeoutMs)
      const partCount = prepared.messages.length
      if (!await repository.setPartCount({ ...ownership, partCount })) {
        throw new PermanentDeliveryError("Roundup parts changed after partial delivery.")
      }
      for (let index = 0; index < partCount; index += 1) {
        const hash = payloadHash(prepared.messages[index])
        const prior = payload.sentMessages.get(index)
        if (prior) {
          if (prior.payloadHash !== hash) {
            throw new PermanentDeliveryError("Roundup content changed after partial delivery.")
          }
          continue
        }
        if (!await repository.renewLease({
          ...ownership,
          now: now(),
          leaseSeconds: config.claimLeaseSeconds
        })) return false
        const sentMessageId = await timeoutPromise(
          prepared.sendPart(index),
          handlerTimeoutMs
        )
        if (!await repository.recordSentMessage({
          ...ownership,
          messageIndex: index,
          sentMessageId,
          payloadHash: hash
        })) return false
      }
      return repository.markSent({ ...ownership, sentAt: now() })
    } catch (error) {
      const failedAt = now()
      const failure = {
        ...ownership,
        failedAt,
        lastError: sanitizeDeliveryError(error)
      }
      if (error instanceof PermanentDeliveryError || claim.attempt_count >= MAX_DELIVERY_ATTEMPTS) {
        return repository.markPermanentlyFailed(failure)
      }
      return repository.markFailed({
        ...failure,
        nextAttemptAt: new Date(
          failedAt.getTime() + retryDelayMinutes(claim.attempt_count) * 60000
        )
      })
    }
  }

  async function tick(tickNow = now()) {
    await generateMissingRoundupClaims({
      repository,
      gameProfile,
      now: tickNow,
      graceMinutes: config.graceMinutes
    })
    const claims = await repository.claimDue({
      now: tickNow,
      batchSize: config.batchSize,
      leaseSeconds: config.claimLeaseSeconds,
      botInstanceName,
      workerId
    })
    await Promise.all(claims.map(processClaim))
    return claims.length
  }

  return Object.freeze({ workerId, tick })
}

module.exports = { createWeeklyRoundupProcessor }
