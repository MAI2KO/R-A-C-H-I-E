const { verifyPublicAllianceEventsRequest } = require("./publicAllianceEventsAuth")

function header(headers, name) {
  const value = headers?.[name]
  return Array.isArray(value) ? value[0] : String(value || "")
}

async function handleNativeManagerAuthorization(input, {
  config, verifyManager, now = () => new Date()
}) {
  if (input.method !== "GET") {
    return { status: 405, body: { ok: false, code: "method_not_allowed" } }
  }
  const match = /^\/internal\/v1\/manager-authorization\/guild\/(\d{15,22})\/user\/(\d{15,22})$/.exec(input.path)
  if (!match) return { status: 404, body: { ok: false, code: "not_found" } }
  const profile = header(input.headers, "x-alliance-events-profile")
  const authenticated = verifyPublicAllianceEventsRequest({
    secret: config.secret,
    method: input.method,
    path: input.path,
    profile,
    timestamp: header(input.headers, "x-alliance-events-timestamp"),
    nonce: header(input.headers, "x-alliance-events-nonce"),
    signature: header(input.headers, "x-alliance-events-signature"),
    now: () => now().getTime()
  })
  if (!authenticated || profile !== config.profile) {
    return { status: 401, body: { ok: false, code: "authentication_failed" } }
  }
  if (typeof verifyManager !== "function") {
    return { status: 503, body: { ok: false, code: "manager_verification_unavailable" } }
  }
  const result = await verifyManager({ guildId: match[1], discordUserId: match[2] })
  if (result.status === "authorized") {
    return { status: 200, body: { ok: true, canManage: true, via: result.via } }
  }
  if (result.status === "denied") {
    return { status: 200, body: { ok: true, canManage: false } }
  }
  return { status: 503, body: { ok: false, code: "manager_verification_unavailable" } }
}

async function handleNativeGuildOwnership(input, {
  config, verifyOwner, now = () => new Date()
}) {
  if (input.method !== "GET") {
    return { status: 405, body: { ok: false, code: "method_not_allowed" } }
  }
  const match = /^\/internal\/v1\/guild-ownership\/guild\/(\d{15,22})\/user\/(\d{15,22})$/.exec(input.path)
  if (!match) return { status: 404, body: { ok: false, code: "not_found" } }
  const profile = header(input.headers, "x-alliance-events-profile")
  const authenticated = verifyPublicAllianceEventsRequest({
    secret: config.secret, method: input.method, path: input.path, profile,
    timestamp: header(input.headers, "x-alliance-events-timestamp"),
    nonce: header(input.headers, "x-alliance-events-nonce"),
    signature: header(input.headers, "x-alliance-events-signature"),
    now: () => now().getTime()
  })
  if (!authenticated || profile !== config.profile) {
    return { status: 401, body: { ok: false, code: "authentication_failed" } }
  }
  if (typeof verifyOwner !== "function") {
    return { status: 503, body: { ok: false, code: "guild_ownership_unavailable" } }
  }
  const result = await verifyOwner({ guildId: match[1], discordUserId: match[2] })
  if (result.status === "owner") return { status: 200, body: { ok: true, isOwner: true } }
  if (result.status === "not_owner") return { status: 200, body: { ok: true, isOwner: false } }
  return { status: 503, body: { ok: false, code: "guild_ownership_unavailable" } }
}

module.exports = { handleNativeManagerAuthorization, handleNativeGuildOwnership }
