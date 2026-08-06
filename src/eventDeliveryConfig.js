const DEFAULT_LOOKAHEAD_MINUTES = 1440
const DEFAULT_GRACE_MINUTES = 60
const DEFAULT_POLL_INTERVAL_MS = 30000
const DEFAULT_BATCH_SIZE = 10
const DEFAULT_CLAIM_LEASE_SECONDS = 60
const DEFAULT_HANDLER_TIMEOUT_MS = 30000

function boundedInteger(value, fallback, minimum, maximum) {
  const text = String(value ?? "").trim()
  if (!/^\d+$/.test(text)) return fallback
  const parsed = Number(text)
  return Number.isSafeInteger(parsed) && parsed >= minimum && parsed <= maximum
    ? parsed
    : fallback
}

function getEventDeliveryConfig(env = process.env) {
  return Object.freeze({
    lookaheadMinutes: boundedInteger(
      env.EVENT_SCHEDULER_LOOKAHEAD_MINUTES,
      DEFAULT_LOOKAHEAD_MINUTES,
      1,
      10080
    ),
    graceMinutes: boundedInteger(
      env.EVENT_SCHEDULER_GRACE_MINUTES,
      DEFAULT_GRACE_MINUTES,
      0,
      1440
    ),
    pollIntervalMs: boundedInteger(
      env.EVENT_SCHEDULER_POLL_INTERVAL_MS,
      DEFAULT_POLL_INTERVAL_MS,
      5000,
      3600000
    ),
    batchSize: boundedInteger(
      env.EVENT_SCHEDULER_BATCH_SIZE,
      DEFAULT_BATCH_SIZE,
      1,
      100
    ),
    claimLeaseSeconds: boundedInteger(
      env.EVENT_SCHEDULER_CLAIM_LEASE_SECONDS,
      DEFAULT_CLAIM_LEASE_SECONDS,
      10,
      600
    ),
    handlerTimeoutMs: boundedInteger(
      env.EVENT_SCHEDULER_HANDLER_TIMEOUT_MS,
      DEFAULT_HANDLER_TIMEOUT_MS,
      1000,
      300000
    )
  })
}

module.exports = {
  DEFAULT_LOOKAHEAD_MINUTES,
  DEFAULT_GRACE_MINUTES,
  DEFAULT_POLL_INTERVAL_MS,
  DEFAULT_BATCH_SIZE,
  DEFAULT_CLAIM_LEASE_SECONDS,
  DEFAULT_HANDLER_TIMEOUT_MS,
  boundedInteger,
  getEventDeliveryConfig
}
