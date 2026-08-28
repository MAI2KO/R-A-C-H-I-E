const http = require("http")

const { verifyPublicAllianceEventsRequest } = require("./publicAllianceEventsAuth")
const { publicAllianceEventsReadModel } = require("./publicAllianceEvents")
const {
  handleNativeManagerAuthorization,
  handleNativeGuildOwnership
} = require("./nativeManagerAuthorizationServer")

function header(headers, name) {
  const value = headers?.[name]
  return Array.isArray(value) ? value[0] : String(value || "")
}

async function handlePublicAllianceEventsRead(input, {
  config, repository, now = () => new Date()
}) {
  if (input.method !== "GET") return { status: 405, body: { ok: false, code: "method_not_allowed" } }
  const guildMatch = /^\/internal\/v1\/public-alliance-events\/guild\/(\d{15,22})$/.exec(input.path)
  const communityMatch = /^\/internal\/v1\/public-alliance-events\/([A-Za-z0-9_-]{1,32})$/.exec(input.path)
  if (!guildMatch && !communityMatch) return { status: 404, body: { ok: false, code: "not_found" } }
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
  try {
    const communityCode = communityMatch ? decodeURIComponent(communityMatch[1]) : undefined
    const events = guildMatch
      ? await repository.listForGuild(guildMatch[1])
      : await repository.listForCommunity(communityCode)
    const model = publicAllianceEventsReadModel({
      profile,
      communityCode,
      events,
      now: now()
    })
    return { status: 200, body: { ok: true, ...model } }
  } catch {
    return { status: 503, body: { ok: false, code: "scheduler_unavailable" } }
  }
}

function createPublicAllianceEventsServer({ config, repository, logger = console,
  verifyManager = null, verifyOwner = null, createServer = http.createServer, now = () => new Date() }) {
  let server = null
  async function requestHandler(request, response) {
    const path = new URL(request.url || "/", "http://internal").pathname
    const input = {
      method: request.method,
      path,
      headers: request.headers
    }
    const result = path.startsWith("/internal/v1/manager-authorization/")
      ? await handleNativeManagerAuthorization(input, { config, verifyManager, now })
      : path.startsWith("/internal/v1/guild-ownership/")
        ? await handleNativeGuildOwnership(input, { config, verifyOwner, now })
        : await handlePublicAllianceEventsRead(input, { config, repository, now })
    response.writeHead(result.status, {
      "content-type": "application/json; charset=utf-8",
      "cache-control": "private, no-store",
      "x-content-type-options": "nosniff"
    })
    response.end(JSON.stringify(result.body))
  }
  async function start() {
    if (server) return { started: false, reason: "already_started" }
    server = createServer((request, response) => {
      void requestHandler(request, response).catch(() => {
        if (!response.headersSent) response.writeHead(503, { "content-type": "application/json" })
        response.end(JSON.stringify({ ok: false, code: "scheduler_unavailable" }))
      })
    })
    await new Promise((resolve, reject) => {
      server.once("error", reject)
      server.listen(config.port, config.host, resolve)
    })
    logger.log(JSON.stringify({ event: "public_alliance_events_read_started",
      game_profile: config.profile, port: config.port }))
    return { started: true }
  }
  async function stop() {
    const current = server
    server = null
    if (!current) return
    await new Promise(resolve => current.close(resolve))
  }
  return Object.freeze({ start, stop, requestHandler })
}

module.exports = { handlePublicAllianceEventsRead, createPublicAllianceEventsServer }
