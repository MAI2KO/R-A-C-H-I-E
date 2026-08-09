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
const { createWeeklyRoundupRepository } = require("./weeklyRoundupRepository")
const { createDiscordWeeklyRoundupDelivery } = require("./discordWeeklyRoundup")
const { createWeeklyRoundupProcessor } = require("./weeklyRoundupProcessor")
const { createStateEventRepository } = require("./stateEventRepository")
const { createDiscordStateEventDeliveryHandler } = require("./discordStateEventDelivery")
const { createStateEventDeliveryProcessor } = require("./stateEventDeliveryProcessor")

function createAllianceEventDeliveryRuntime({
  client,
  initializationPromise,
  env = process.env,
  logger = console,
  getPoolFn = getPool,
  createRepositoryFn = createEventDeliveryRepository,
  createHandlerFn = createDiscordEventDeliveryHandler,
  createRoundupRepositoryFn = createWeeklyRoundupRepository,
  createRoundupDeliveryFn = createDiscordWeeklyRoundupDelivery,
  createRoundupProcessorFn = createWeeklyRoundupProcessor,
  createStateRepositoryFn = createStateEventRepository,
  createStateHandlerFn = createDiscordStateEventDeliveryHandler,
  createStateProcessorFn = createStateEventDeliveryProcessor,
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
      { targetKind: "alliance" }
    )
    const deliveryHandler = createHandlerFn({ client, gameProfile })
    const roundupRepository = createRoundupRepositoryFn(
      getPoolFn({ env, logger }),
      gameProfile
    )
    const roundupDelivery = createRoundupDeliveryFn({ client, gameProfile })
    const roundupProcessor = createRoundupProcessorFn({
      repository: roundupRepository,
      gameProfile,
      botInstanceName,
      delivery: roundupDelivery,
      config: require("./eventDeliveryConfig").getEventDeliveryConfig(env),
      logger
    })
    let stateProcessor = null
    try {
      const stateRepository = createStateRepositoryFn(
        getPoolFn({ env, logger }),
        gameProfile
      )
      stateProcessor = createStateProcessorFn({
        repository: stateRepository,
        gameProfile,
        botInstanceName,
        deliveryHandler: createStateHandlerFn({ client, gameProfile }),
        config: require("./eventDeliveryConfig").getEventDeliveryConfig(env),
        logger
      })
    } catch (error) {
      logger.error(
        `[Event scheduler] State-event processor not started: ${sanitizeDeliveryError(error)}`
      )
    }
    worker = createWorkerFn({
      env,
      health,
      repository,
      gameProfile,
      botInstanceName,
      deliveryHandler,
      additionalTick: tickNow => Promise.all([
        roundupProcessor.tick(tickNow),
        stateProcessor ? stateProcessor.tick(tickNow) : Promise.resolve(0)
      ]),
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
