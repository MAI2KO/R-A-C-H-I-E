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
        try {
          const response = await transport.post(url, body, {
            headers: { "Content-Type": "application/x-www-form-urlencoded" },
            validateStatus: () => true
          })
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
            response: classification.raw
          }
        } catch (error) {
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
            errorCode: String(error?.code || "NETWORK_ERROR").slice(0, 100)
          }
        }
      })
    }
  })
}

module.exports = { formEncodedPayload, createCenturyGameClient }
