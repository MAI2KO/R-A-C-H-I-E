const { getPool } = require("./db")
const { createPublicAllianceEventRepository } = require("./publicAllianceEventRepository")
const { createPublicAllianceEventsServer } = require("./publicAllianceEventsServer")

const PROFILES = new Set(["wos", "kingshot"])

function publicAllianceEventsConfig(env = process.env) {
  const requested = String(env.ALLIANCE_EVENTS_READ_ENABLED || "").trim().toLowerCase() === "true"
  const profile = String(env.GAME_PROFILE || "wos").trim()
  const secret = String(env.ALLIANCE_EVENTS_READ_SECRET || "")
  const rawPort = String(env.ALLIANCE_EVENTS_READ_PORT || env.PORT || "3000").trim()
  const port = Number(rawPort)
  const portValid = Number.isInteger(port) && port >= 1 && port <= 65535
  const enabled = requested && PROFILES.has(profile) && secret.length >= 32 && portValid
  const disabledReason = enabled ? null
    : !requested ? "ALLIANCE_EVENTS_READ_ENABLED is not true"
      : !PROFILES.has(profile) ? "GAME_PROFILE must be wos or kingshot"
        : secret.length < 32 ? "ALLIANCE_EVENTS_READ_SECRET must contain at least 32 characters"
          : "ALLIANCE_EVENTS_READ_PORT must be a valid TCP port"
  return Object.freeze({ enabled, requested, profile, secret, port, portValid,
    host: "0.0.0.0", disabledReason, secretConfigured: secret.length > 0 })
}

function createPublicAllianceEventsBootstrap({ initializationPromise, env = process.env,
  logger = console, processRef = process, config = publicAllianceEventsConfig(env),
  getPoolFn = getPool, createRepository = createPublicAllianceEventRepository,
  createServer = createPublicAllianceEventsServer }) {
  logger.log(JSON.stringify({ event: "public_alliance_events_read_configuration",
    game_profile: config.profile, enabled: config.enabled, requested: config.requested,
    secret_configured: config.secretConfigured, port: config.port,
    disabled_reason: config.disabledReason }))
  let runtime = null
  let startPromise = null
  async function start() {
    if (startPromise) return startPromise
    startPromise = (async () => {
      if (!config.enabled) return { started: false, reason: config.disabledReason }
      const health = await initializationPromise
      if (!health?.available || health.gameProfile !== config.profile) {
        return { started: false, reason: "scheduler unavailable" }
      }
      runtime = createServer({
        config,
        repository: createRepository(getPoolFn({ env, logger }), config.profile),
        logger
      })
      return runtime.start()
    })().catch(() => ({ started: false, reason: "startup failed" }))
    return startPromise
  }
  async function stop() { await runtime?.stop() }
  const stopOnSignal = () => { void stop() }
  processRef.once("SIGINT", stopOnSignal)
  processRef.once("SIGTERM", stopOnSignal)
  return Object.freeze({ config, start, stop, getRuntime: () => runtime })
}

module.exports = { publicAllianceEventsConfig, createPublicAllianceEventsBootstrap }
