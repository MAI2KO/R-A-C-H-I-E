const axios = require("axios")
const { centuryAdapter } = require("./adapters")
const { signRequestFields } = require("./signing")
const { classifyCenturyResponse } = require("./responseClassifier")
const { ConservativeRateLimiter, rateLimiterConfig } = require("./rateLimiter")
const { normalizePlayerId, normalizeLocationNumber, normalizeGiftCode } = require("./validation")
const { profileTerminology } = require("./terminology")

function formEncodedPayload(fields) {
  const form = new URLSearchParams()
  for (const [key, value] of Object.entries(fields)) form.set(key, String(value))
  return form.toString()
}

function rateLimitMetadata(headers = {}) {
  const normalized = {}
  for (const [key, value] of Object.entries(headers || {})) {
    normalized[String(key).toLowerCase()] = Array.isArray(value) ? value[0] : value
  }
  return {
    limit: normalized["x-ratelimit-limit"] ?? null,
    remaining: normalized["x-ratelimit-remaining"] ?? null,
    reset: normalized["x-ratelimit-reset"] ?? null,
    retryAfter: normalized["retry-after"] ?? null
  }
}

function boundedResponseData(data, maximumCharacters = 4000) {
  try {
    const serialized = JSON.stringify(data ?? null)
    if (serialized.length > maximumCharacters) {
      return { truncated: true, originalCharacters: serialized.length }
    }
    return JSON.parse(serialized)
  } catch {
    return { unreadable: true }
  }
}

function createCenturyGameClient({
  gameProfile,
  env = process.env,
  adapter = centuryAdapter(gameProfile, env),
  transport = axios,
  limiter = new ConservativeRateLimiter({ gameProfile, ...rateLimiterConfig(env) }),
  now = () => Date.now()
}) {
  const terms = profileTerminology(gameProfile)

  return Object.freeze({
    adapter,
    limiter,

    async redeem({ playerId, code, locationNumber }) {
      const unsigned = {
        fid: normalizePlayerId(playerId, terms.playerLabel),
        cdk: normalizeGiftCode(code),
        kid: normalizeLocationNumber(locationNumber, terms.locationLabel),
        time: String(now())
      }
      const fields = {
        ...unsigned,
        sign: signRequestFields(unsigned, adapter.signingSuffix)
      }
      const body = formEncodedPayload(fields)
      const url = `${adapter.apiBaseUrl}${adapter.redemptionPath}`

      return limiter.schedule(async () => {
        const requestStartedAt = new Date(now())
        try {
          const response = await transport.post(url, body, {
            headers: { "Content-Type": "application/x-www-form-urlencoded" },
            validateStatus: () => true
          })
          const responseReceivedAt = new Date(now())
          const classification = classifyCenturyResponse({
            httpStatus: response.status,
            data: response.data,
            profileMappings: adapter.responseMappings || {}
          })
          return {
            httpStatus: response.status,
            headers: response.headers || {},
            classification,
            retryable: classification.retryable,
            response: classification.raw,
            responseData: boundedResponseData(response.data),
            endpoint: adapter.redemptionPath,
            rateLimit: rateLimitMetadata(response.headers),
            requestStartedAt,
            responseReceivedAt
          }
        } catch (error) {
          const responseReceivedAt = new Date(now())
          return {
            httpStatus: null,
            headers: {},
            classification: {
              state: "temporary_error",
              retryable: true,
              permanent: false,
              raw: { code: null, errCode: null, message: "Network request failed" }
            },
            retryable: true,
            response: { code: null, errCode: null, message: "Network request failed" },
            responseData: null,
            errorCode: String(error?.code || "NETWORK_ERROR").slice(0, 100),
            endpoint: adapter.redemptionPath,
            rateLimit: {},
            requestStartedAt,
            responseReceivedAt
          }
        }
      })
    }
  })
}

module.exports = {
  formEncodedPayload,
  rateLimitMetadata,
  boundedResponseData,
  createCenturyGameClient
}
