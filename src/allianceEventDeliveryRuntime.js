const { getPool } = require("./db")
const {
  shutdownEventSchedulerSubsystem
} = require("./eventSchedulerHealth")
const {
  createEventDeliveryRepository
} = require("./eventDeliveryRepository")
const {
  createEventDeliveryWorker,
  sanitizeDeliveryError
} = require("./eventDeliveryWorker")
const {
  createDiscordEventDeliveryHandler
} = require("./discordEventDelivery")

function createAllianceEventDeliveryRuntime({
  client,
  initializationPromise,
  env = process.env,
  logger = console,
  getPoolFn = getPool,
  createRepositoryFn = createEventDeliveryRepository,
  createHandlerFn = createDiscordEventDeliveryHandler,
  createWorkerFn = createEventDeliveryWorker,
  shutdownFn = shutdownEventSchedulerSubsystem
}) {
  let worker = null
  let startPromise = null
  let shutdownPromise = null
  let shutdownHandlersInstalled = false
  let stopped = false

  async function startOnce() {
    if (!client?.isReady?.()) return { started: false, reason: "Discord client not ready" }
    const health = await initializationPromise
    if (stopped) return { started: false, reason: "stopped" }
    if (!health?.available) {
      return { started: false, reason: health?.reason || "database unavailable" }
    }
    const gameProfile = health.gameProfile
    const botInstanceName = health.botInstanceName
    const repository = createRepositoryFn(
      getPoolFn({ env, logger }),
      gameProfile,
      { targetKind: null }
    )
    const deliveryHandler = createHandlerFn({ client, gameProfile })
    worker = createWorkerFn({
      env,
      health,
      repository,
      gameProfile,
      botInstanceName,
      deliveryHandler,
      logger
    })
    const result = worker.start()
    if (result.started) {
      logger.log(`[Event scheduler] Event delivery worker started for ${gameProfile}`)
    } else {
      logger.error(`[Event scheduler] Event delivery worker not started: ${result.reason}`)
    }
    return result
  }

  function start() {
    if (stopped) return Promise.resolve({ started: false, reason: "stopped" })
    if (!client?.isReady?.()) {
      return Promise.resolve({ started: false, reason: "Discord client not ready" })
    }
    if (startPromise) return startPromise
    startPromise = startOnce().catch(error => {
      logger.error(
        `[Event scheduler] Event delivery startup failed: ${sanitizeDeliveryError(error)}`
      )
      return { started: false, reason: "startup failed" }
    })
    return startPromise
  }

  function stop({ timeoutMs = 5000, destroyClient = false } = {}) {
    if (shutdownPromise) return shutdownPromise
    stopped = true
    shutdownPromise = shutdownFn({ worker, timeoutMs, logger })
      .catch(error => {
        logger.error(
          `[Event scheduler] Event delivery shutdown failed: ${sanitizeDeliveryError(error)}`
        )
        return { workerDrained: false }
      })
      .finally(() => {
        if (destroyClient) {
          try {
            client?.destroy?.()
          } catch {
            logger.error("[Event scheduler] Discord client cleanup failed")
          }
        }
      })
    return shutdownPromise
  }

  function installShutdownHandlers(processRef = process) {
    if (shutdownHandlersInstalled) return false
    shutdownHandlersInstalled = true
    const handleSignal = () => { void stop({ destroyClient: true }) }
    processRef.once("SIGINT", handleSignal)
    processRef.once("SIGTERM", handleSignal)
    return true
  }

  return Object.freeze({
    start,
    stop,
    installShutdownHandlers,
    getWorker: () => worker
  })
}

module.exports = {
  createAllianceEventDeliveryRuntime
}
