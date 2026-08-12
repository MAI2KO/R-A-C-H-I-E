function booleanFlag(value) {
  return ["true", "1"].includes(String(value ?? "").trim().toLowerCase())
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

function effectiveSourcePollingConfig(gameProfile, env = process.env) {
  const config = sourcePollingConfig(env)
  const profileSourceEnabled = gameProfile === "wos"
    ? config.wosEnabled
    : gameProfile === "kingshot" ? config.kingshotEnabled : false
  return Object.freeze({
    ...config,
    gameProfile,
    profileSourceEnabled,
    publicCatalogueEnabled: config.pollingEnabled && profileSourceEnabled
  })
}

module.exports = { booleanFlag, sourcePollingConfig, effectiveSourcePollingConfig }
