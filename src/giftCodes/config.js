const { boundedInteger } = require("./rateLimiter")
const { normalizePlayerId, normalizeLocationNumber } = require("./validation")
const { profileTerminology } = require("./terminology")

function verificationWorkerIsEnabled(env = process.env) {
  return env.GIFT_CODE_VERIFICATION_ENABLED === "true"
}

function redemptionWorkerIsEnabled(env = process.env) {
  return env.GIFT_CODE_REDEMPTION_WORKER_ENABLED === "true"
}

function verifierEnvironmentNames(gameProfile) {
  if (gameProfile === "wos") {
    return { player: "WOS_GIFT_VERIFY_FID", location: "WOS_GIFT_VERIFY_KID" }
  }
  if (gameProfile === "kingshot") {
    return { player: "KINGSHOT_GIFT_VERIFY_FID", location: "KINGSHOT_GIFT_VERIFY_KID" }
  }
  throw new Error("Unsupported game profile")
}

function getVerifierAccount(gameProfile, env = process.env) {
  const names = verifierEnvironmentNames(gameProfile)
  const playerId = String(env[names.player] || "").trim()
  const locationNumber = String(env[names.location] || "").trim()
  if (!playerId && !locationNumber) return null
  if (!playerId || !locationNumber) return { configured: false, reason: "incomplete verifier configuration" }
  const terms = profileTerminology(gameProfile)
  try {
    return Object.freeze({
      configured: true,
      playerId: normalizePlayerId(playerId, terms.playerLabel),
      locationNumber: normalizeLocationNumber(locationNumber, terms.locationLabel)
    })
  } catch {
    return { configured: false, reason: "invalid verifier configuration" }
  }
}

function giftWorkerConfig(env = process.env) {
  return Object.freeze({
    pollIntervalMs: boundedInteger(env.GIFT_CODE_WORKER_POLL_INTERVAL_MS, 10000, 1000, 300000),
    leaseSeconds: boundedInteger(env.GIFT_CODE_CLAIM_LEASE_SECONDS, 120, 15, 900),
    maximumAttempts: boundedInteger(env.GIFT_CODE_MAX_ATTEMPTS, 5, 1, 20),
    retryBaseSeconds: boundedInteger(env.GIFT_CODE_RETRY_BASE_SECONDS, 60, 10, 3600),
    retryCapSeconds: boundedInteger(env.GIFT_CODE_RETRY_CAP_SECONDS, 3600, 60, 86400)
  })
}

module.exports = {
  verificationWorkerIsEnabled,
  redemptionWorkerIsEnabled,
  verifierEnvironmentNames,
  getVerifierAccount,
  giftWorkerConfig
}
