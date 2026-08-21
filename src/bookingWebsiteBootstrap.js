const {
  bookingWebsiteConfig,
  createBookingWebsiteClient
} = require("./bookingWebsiteClient")
const { createBookingWebsiteRuntime } = require("./bookingDiscordIntegration")

function createBookingWebsiteBootstrap({
  client,
  env = process.env,
  logger = console,
  processRef = process,
  configuration = bookingWebsiteConfig(env),
  createApi = createBookingWebsiteClient,
  createRuntime = createBookingWebsiteRuntime
}) {
  logger.log(JSON.stringify({
    event: "booking_website_integration_configuration",
    game_profile: configuration.profile,
    enabled: configuration.enabled,
    requested: configuration.requested,
    base_url_configured: configuration.baseUrlConfigured,
    secret_configured: configuration.secretConfigured,
    poll_interval_ms: configuration.pollIntervalMs,
    disabled_reason: configuration.disabledReason
  }))

  if (!configuration.enabled) {
    logger.log(JSON.stringify({
      event: "booking_website_integration_disabled",
      game_profile: configuration.profile,
      reason: configuration.disabledReason
    }))
    return Object.freeze({ configuration, api: null, runtime: null, startWorker: () => false })
  }

  let api
  let runtime
  try {
    api = createApi({ config: configuration })
    runtime = createRuntime({
      client,
      api,
      intervalMs: configuration.pollIntervalMs,
      logger
    })
  } catch {
    logger.error(JSON.stringify({
      event: "booking_website_integration_startup_failed",
      game_profile: configuration.profile,
      error_code: "initialization_failed"
    }))
    return Object.freeze({ configuration, api: null, runtime: null, startWorker: () => false })
  }

  let startAttempted = false
  function startWorker() {
    if (startAttempted) return false
    startAttempted = true
    try {
      const result = runtime.start()
      logger.log(JSON.stringify({
        event: "booking_website_integration_worker_started",
        game_profile: configuration.profile,
        worker_started: result?.started === true,
        reason: result?.started === true ? null : String(result?.reason || "not_started").slice(0, 80)
      }))
      return result?.started === true
    } catch {
      logger.error(JSON.stringify({
        event: "booking_website_integration_startup_failed",
        game_profile: configuration.profile,
        error_code: "worker_start_failed"
      }))
      return false
    }
  }

  if (client?.isReady?.()) startWorker()
  else client.once("clientReady", startWorker)

  const stopWorker = () => { void runtime.stop() }
  processRef.once("SIGINT", stopWorker)
  processRef.once("SIGTERM", stopWorker)
  return Object.freeze({ configuration, api, runtime, startWorker })
}

module.exports = { createBookingWebsiteBootstrap }
