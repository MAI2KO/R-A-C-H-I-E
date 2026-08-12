function booleanFlag(value) {
  return String(value || "").toLowerCase() === "true"
}

function sourcePollingConfig(env = process.env) {
  const seconds = Number.parseInt(env.GIFT_CODE_SOURCE_POLL_INTERVAL_SECONDS || "900", 10)
  return Object.freeze({
    pollingEnabled: booleanFlag(env.GIFT_CODE_SOURCE_POLLING_ENABLED),
    wosEnabled: booleanFlag(env.WOS_REWARDS_SOURCE_ENABLED),
    kingshotEnabled: booleanFlag(env.KINGSHOT_REWARDS_SOURCE_ENABLED),
    intervalMs: Math.max(300, Number.isFinite(seconds) ? seconds : 900) * 1000,
    timeoutMs: 10000,
    maximumBodyBytes: 1024 * 1024
  })
}

module.exports = { booleanFlag, sourcePollingConfig }
