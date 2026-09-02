const { createHash, createHmac, randomUUID } = require("crypto")

const PROFILES = new Set(["wos", "kingshot"])
const PRODUCTION_WEBSITE_ORIGINS = Object.freeze({
  wos: "https://r-a-c-h-i-e.com",
  kingshot: "https://peggie.r-a-c-h-i-e.com"
})

function canonicalRequest({ method, path, timestamp, nonce, body }) {
  return ["v1", method.toUpperCase(), path, String(timestamp), nonce,
    createHash("sha256").update(body, "utf8").digest("hex")].join("\n")
}

function signRequest({ secret, ...request }) {
  if (typeof secret !== "string" || secret.length < 32) throw new Error("invalid integration secret")
  return `v1=${createHmac("sha256", secret).update(canonicalRequest(request), "utf8").digest("hex")}`
}

function bookingWebsiteConfig(env = process.env) {
  const profile = String(env.GAME_PROFILE || "wos").trim()
  const baseUrl = String(env.BOOKING_WEBSITE_BASE_URL || "").trim().replace(/\/$/, "")
  const secret = String(env.BOOKING_WEBSITE_INTEGRATION_SECRET || "")
  const requested = String(env.BOOKING_WEBSITE_INTEGRATION_ENABLED || "").trim().toLowerCase() === "true"
  const baseUrlConfigured = baseUrl.length > 0
  const secretConfigured = secret.length > 0
  const allowLoopback = String(env.NODE_ENV || "").trim().toLowerCase() !== "production"
  let safeUrl = false
  try {
    const parsed = new URL(baseUrl)
    const loopback = ["localhost", "127.0.0.1", "::1"].includes(parsed.hostname)
    const cleanOrigin = parsed.username === "" && parsed.password === ""
      && parsed.pathname.replace(/\/$/, "") === "" && parsed.search === "" && parsed.hash === ""
    const allowedProtocol = (parsed.protocol === "https:" && (!loopback || allowLoopback))
      || (parsed.protocol === "http:" && loopback && allowLoopback)
    const productionOriginMatches = allowLoopback
      || parsed.origin === PRODUCTION_WEBSITE_ORIGINS[profile]
    safeUrl = cleanOrigin && allowedProtocol && productionOriginMatches
  } catch {}
  const secretValid = secret.length >= 32
  const enabled = requested && PROFILES.has(profile) && safeUrl && secretValid
  const disabledReason = enabled ? null
    : !requested ? "BOOKING_WEBSITE_INTEGRATION_ENABLED is not true"
      : !PROFILES.has(profile) ? "GAME_PROFILE must be wos or kingshot"
        : !baseUrlConfigured ? "BOOKING_WEBSITE_BASE_URL is missing"
          : !safeUrl ? "BOOKING_WEBSITE_BASE_URL is not a valid website origin for this profile"
            : !secretConfigured ? "BOOKING_WEBSITE_INTEGRATION_SECRET is missing"
              : "BOOKING_WEBSITE_INTEGRATION_SECRET must contain at least 32 characters"
  return Object.freeze({ enabled, requested, profile, baseUrl, secret, allowLoopback,
    baseUrlConfigured, secretConfigured, disabledReason,
    pollIntervalMs: Math.max(5000, Number(env.BOOKING_WEBSITE_POLL_INTERVAL_MS) || 10000) })
}

function createBookingWebsiteClient({ config, fetchImplementation = fetch, now = Date.now, createNonce = randomUUID }) {
  if (!config?.enabled) throw new Error("booking website integration is disabled")
  async function post(path, input = {}) {
    const body = JSON.stringify(input)
    const timestamp = String(Math.floor(now() / 1000))
    const nonce = createNonce()
    const signature = signRequest({ secret: config.secret, method: "POST", path, timestamp, nonce, body })
    let response
    try {
      response = await fetchImplementation(`${config.baseUrl}${path}`, {
        method: "POST",
        headers: { "content-type": "application/json", "x-booking-profile": config.profile,
          "x-booking-timestamp": timestamp, "x-booking-nonce": nonce, "x-booking-signature": signature },
        body
      })
    } catch {
      const error = new Error("booking website request failed")
      error.code = "network_error"
      error.operation = path.endsWith("/claim") ? "claim" : "integration_request"
      throw error
    }
    let result = null
    try { result = await response.json() } catch {}
    if (!response.ok) {
      const remoteCode = /^[a-z0-9_]{1,80}$/.test(String(result?.code || ""))
        ? String(result.code) : "http_error"
      const error = new Error("booking website request was refused")
      error.code = remoteCode
      error.status = response.status
      error.operation = path.endsWith("/claim") ? "claim" : "integration_request"
      error.publicCode = remoteCode
      throw error
    }
    return result
  }
  return Object.freeze({
    profile: config.profile,
    baseUrl: config.baseUrl,
    allowLoopback: config.allowLoopback === true,
    claim: limit => post("/api/internal/v1/discord/work/claim", { limit }),
    recipients: (work, recipients) => post(`/api/internal/v1/discord/work/${work.workId}/recipients`, { claimToken: work.claimToken, recipients }),
    outcome: (work, outcome) => post(`/api/internal/v1/discord/work/${work.workId}/outcome`, { claimToken: work.claimToken, ...outcome }),
    approval: (requestId, action, actor) => post(`/api/internal/v1/discord/approval/${requestId}/${action}`, actor),
    communitySetup: input => post("/api/internal/v1/discord/setup/community", input),
    registration: input => post("/api/internal/v1/discord/registration", input)
  })
}

module.exports = { PRODUCTION_WEBSITE_ORIGINS, canonicalRequest, signRequest,
  bookingWebsiteConfig, createBookingWebsiteClient }
