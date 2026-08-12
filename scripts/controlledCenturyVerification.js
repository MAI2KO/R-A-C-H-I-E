const { createCenturyGameClient } = require("../src/giftCodes/centuryGameClient")
const { getVerifierAccount } = require("../src/giftCodes/config")
const { ConservativeRateLimiter } = require("../src/giftCodes/rateLimiter")

async function runControlledCenturyVerification({
  env = process.env,
  clientFactory = createCenturyGameClient,
  write = value => process.stdout.write(`${value}\n`)
} = {}) {
  if (env.ALLOW_ONE_LIVE_CENTURY_REQUEST !== "true") {
    throw new Error("Live Century request blocked: set ALLOW_ONE_LIVE_CENTURY_REQUEST=true explicitly")
  }
  const gameProfile = String(env.GAME_PROFILE || "").trim()
  if (!["wos", "kingshot"].includes(gameProfile)) {
    throw new Error("Controlled harness requires GAME_PROFILE=wos or GAME_PROFILE=kingshot")
  }
  const verifier = getVerifierAccount(gameProfile, env)
  if (!verifier?.configured) throw new Error(verifier?.reason || `${gameProfile} verifier is not configured`)
  const code = String(env.LIVE_CENTURY_GIFT_CODE || "").trim()
  if (!code) throw new Error("LIVE_CENTURY_GIFT_CODE is required")

  const limiter = new ConservativeRateLimiter({
    gameProfile,
    minimumDelayMs: 0,
    maximumRetries: 0
  })
  const client = clientFactory({ gameProfile, env, limiter })
  const result = await client.redeem({
    playerId: verifier.playerId,
    locationNumber: verifier.locationNumber,
    code
  })
  const report = {
    requestCount: 1,
    method: "POST",
    endpoint: `${client.adapter.apiBaseUrl}${client.adapter.redemptionPath}`,
    httpStatus: result.httpStatus,
    classification: result.classification.state,
    centuryErrCode: result.classification.raw.errCode,
    rateLimit: result.rateLimit,
    response: result.responseDiagnostics
  }
  write(JSON.stringify(report, null, 2))
  return report
}

if (require.main === module) {
  require("dotenv").config()
  runControlledCenturyVerification().catch(error => {
    process.stderr.write(`${String(error?.message || "Controlled request failed")}\n`)
    process.exitCode = 1
  })
}

module.exports = { runControlledCenturyVerification }
