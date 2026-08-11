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
  const normalized = normalizedHeaders(headers)
  const value = name => normalized[name] === undefined
    ? null
    : String(normalized[name]).slice(0, 300)
  return {
    limit: value("x-ratelimit-limit"),
    remaining: value("x-ratelimit-remaining"),
    reset: value("x-ratelimit-reset"),
    retryAfter: value("retry-after")
  }
}

function normalizedHeaders(headers = {}) {
  const normalized = {}
  for (const [key, value] of Object.entries(headers || {})) {
    normalized[String(key).toLowerCase()] = Array.isArray(value) ? value[0] : value
  }
  return normalized
}

const SENSITIVE_KEY = /^(?:authorization|cookie|set-cookie|password|secret|sign|signature|token|fid|kid|cdk)$/i
const SAFE_EDGE_HEADERS = Object.freeze([
  "cf-ray", "cf-cache-status", "via", "x-cache", "x-request-id",
  "x-amz-cf-id", "x-served-by"
])

function sanitizedJson(value) {
  try {
    return JSON.stringify(value ?? null, (key, child) =>
      SENSITIVE_KEY.test(key) ? "[redacted]" : child
    )
  } catch {
    return "[unreadable response]"
  }
}

function sanitizedText(value, redactedValues = []) {
  let text = String(value || "")
    .replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, " ")
    .replace(/<style\b[^>]*>[\s\S]*?<\/style>/gi, " ")
    .replace(/<!--([\s\S]*?)-->/g, " ")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&quot;/gi, '"')
    .replace(/&#(?:39|x27);/gi, "'")
    .replace(/\b(authorization|cookie|password|secret|sign|signature|token|fid|kid|cdk)\b\s*[:=]\s*[^\s,;&]+/gi, "$1=[redacted]")
    .replace(/\s+/g, " ")
    .trim()
  for (const value of redactedValues) {
    const exact = String(value || "")
    if (exact) text = text.split(exact).join("[redacted]")
  }
  return text
}

function responseType(contentType, data) {
  const normalized = String(contentType || "").toLowerCase()
  if (normalized.includes("json") || (data && typeof data === "object")) return "json"
  if (normalized.includes("html") || /<html\b|<!doctype\s+html/i.test(String(data || ""))) return "html"
  if (normalized.startsWith("text/")) return "text"
  return "unknown"
}

function boundedResponseData(data, maximumCharacters = 2048, type = "unknown", redactedValues = []) {
  const raw = data && typeof data === "object"
    ? sanitizedJson(data)
    : String(data ?? "")
  const summary = sanitizedText(raw, redactedValues)
  return {
    summary: summary.slice(0, maximumCharacters),
    truncated: summary.length > maximumCharacters,
    originalCharacters: summary.length
  }
}

function responseDiagnostics(data, headers = {}, maximumCharacters = 2048, redactedValues = []) {
  const normalized = normalizedHeaders(headers)
  const contentType = String(normalized["content-type"] || "").slice(0, 200) || null
  const type = responseType(contentType, data)
  const body = boundedResponseData(data, maximumCharacters, type, redactedValues)
  const edgeHeaders = {}
  for (const name of SAFE_EDGE_HEADERS) {
    if (normalized[name] !== undefined) edgeHeaders[name] = String(normalized[name]).slice(0, 300)
  }
  return {
    responseType: type,
    contentType,
    server: normalized.server ? String(normalized.server).slice(0, 200) : null,
    edgeHeaders,
    bodySummary: body.summary || null,
    bodyTruncated: body.truncated,
    originalBodyCharacters: body.originalCharacters
  }
}

function requestHeaders(adapter) {
  const frontendUrl = String(adapter.frontendUrl || "").replace(/\/$/, "")
  return {
    "Content-Type": "application/x-www-form-urlencoded",
    Accept: "application/json, text/plain, */*",
    Origin: frontendUrl,
    Referer: `${frontendUrl}/`,
    "User-Agent": "R.A.C.H.I.E gift-code client"
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
      const requestTimeMs = Number(now())
      const unsigned = {
        fid: normalizePlayerId(playerId, terms.playerLabel),
        cdk: normalizeGiftCode(code),
        kid: normalizeLocationNumber(locationNumber, terms.locationLabel),
        time: String(Math.floor(requestTimeMs / 1000))
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
            headers: requestHeaders(adapter),
            validateStatus: () => true,
            maxRedirects: 5,
            decompress: true,
            maxContentLength: 64 * 1024
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
            responseDiagnostics: responseDiagnostics(response.data, response.headers, 2048, [
              fields.fid, fields.kid, fields.sign, adapter.signingSuffix
            ]),
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
            responseDiagnostics: null,
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
  normalizedHeaders,
  sanitizedText,
  responseType,
  boundedResponseData,
  responseDiagnostics,
  requestHeaders,
  createCenturyGameClient
}
