const { createHmac, randomUUID, timingSafeEqual } = require("crypto")

const PROFILES = new Set(["wos", "kingshot"])
const SIGNATURE_PATTERN = /^v1=([0-9a-f]{64})$/
const NONCE_PATTERN = /^[A-Za-z0-9_-]{16,128}$/
const CLOCK_TOLERANCE_SECONDS = 300

function canonicalPublicAllianceEventsRequest({ method, path, profile, timestamp, nonce }) {
  return ["v1", String(method).toUpperCase(), String(path), String(profile),
    String(timestamp), String(nonce)].join("\n")
}

function signPublicAllianceEventsRequest({ secret, ...request }) {
  if (typeof secret !== "string" || secret.length < 32) throw new Error("invalid integration secret")
  return `v1=${createHmac("sha256", secret)
    .update(canonicalPublicAllianceEventsRequest(request), "utf8").digest("hex")}`
}

function verifyPublicAllianceEventsRequest({ secret, method, path, profile, timestamp, nonce,
  signature, now = Date.now, toleranceSeconds = CLOCK_TOLERANCE_SECONDS }) {
  if (!PROFILES.has(profile) || typeof secret !== "string" || secret.length < 32) return false
  if (!/^\d{10}$/.test(String(timestamp)) || !NONCE_PATTERN.test(String(nonce))) return false
  if (Math.abs(Math.floor(now() / 1000) - Number(timestamp)) > toleranceSeconds) return false
  const match = SIGNATURE_PATTERN.exec(String(signature || ""))
  if (!match) return false
  const expected = signPublicAllianceEventsRequest({ secret, method, path, profile, timestamp, nonce })
  const supplied = Buffer.from(match[1], "hex")
  const calculated = Buffer.from(expected.slice(3), "hex")
  return supplied.length === calculated.length && timingSafeEqual(supplied, calculated)
}

function signedPublicAllianceEventsHeaders({ secret, profile, method, path,
  now = Date.now, createNonce = randomUUID }) {
  const timestamp = String(Math.floor(now() / 1000))
  const nonce = createNonce()
  return {
    "x-alliance-events-profile": profile,
    "x-alliance-events-timestamp": timestamp,
    "x-alliance-events-nonce": nonce,
    "x-alliance-events-signature": signPublicAllianceEventsRequest({
      secret, method, path, profile, timestamp, nonce
    })
  }
}

module.exports = {
  CLOCK_TOLERANCE_SECONDS,
  canonicalPublicAllianceEventsRequest,
  signPublicAllianceEventsRequest,
  verifyPublicAllianceEventsRequest,
  signedPublicAllianceEventsHeaders
}
