const { getPool } = require("../db")
const { createGiftCodeRepository } = require("./repository")
const { createCenturyGameClient } = require("./centuryGameClient")
const { ConservativeRateLimiter, rateLimiterConfig } = require("./rateLimiter")
const {
  verificationWorkerIsEnabled,
  redemptionWorkerIsEnabled,
  getVerifierAccount,
  giftWorkerConfig
} = require("./config")
const {
  createVerificationProcessor,
  createRedemptionProcessor,
  createPollingWorker,
  sanitizeWorkerError
} = require("./workers")
const { createDiscordGiftNotifier } = require("./notifications")
const { createGiftCodeCommunityRepository } = require("./communityRepository")
const { createGiftCodeCommunityService } = require("./communityService")
const { giftAccountConfig } = require("./config")
const { createGiftCodeSourceRepository } = require("./sourceRepository")
const { createGiftCodeSourceIngestionService } = require("./sourceIngestion")
const { createCatalogueAdapter, createCataloguePoller } = require("./catalogueSources")
const { effectiveSourcePollingConfig } = require("./sourceConfig")

let currentRuntime = null

function getGiftCodeRuntime() {
  return currentRuntime
}

function sourceStartupDiagnostic(config, subsystem) {
  return `[Gift code sources] ${config.gameProfile}: ` +
    `polling=${config.pollingEnabled}, profile_source=${config.profileSourceEnabled}, ` +
    `public_catalogue=${config.publicCatalogueEnabled}, subsystem=${subsystem}, ` +
    `interval=${config.intervalMs / 1000}s`
}

function startCatalogueSourceRuntime({
  gameProfile,
  env,
  sourceRepository,
  sourceIngestion,
  logger = console,
  adapterFactory = createCatalogueAdapter,
  pollerFactory = createCataloguePoller,
  workerFactory = createPollingWorker
}) {
  const config = effectiveSourcePollingConfig(gameProfile, env)
  logger.log(sourceStartupDiagnostic(config, true))
  let poller = null
  let worker = null
  try {
    poller = pollerFactory({
      gameProfile,
      sourceRepository,
      ingestion: sourceIngestion,
      adapter: adapterFactory({
        gameProfile,
        timeoutMs: config.timeoutMs,
        maximumBodyBytes: config.maximumBodyBytes
      }),
      enabled: config.publicCatalogueEnabled,
      logger
    })
    worker = workerFactory({
      tick: poller.poll,
      intervalMs: config.intervalMs,
      enabled: config.publicCatalogueEnabled
    })
    const start = worker.start()
    const startError = config.publicCatalogueEnabled && !start.started
      ? String(start.reason || "not_started").slice(0, 100)
      : null
    if (config.publicCatalogueEnabled && !start.started) {
      logger.error(JSON.stringify({
        event: "gift_code_source_poller_start_failed",
        game_profile: gameProfile,
        error_code: startError
      }))
    }
    return { config, poller, worker, start, error: startError }
  } catch (error) {
    const errorCode = sanitizeWorkerError(error)
    logger.error(JSON.stringify({
      event: "gift_code_source_poller_start_failed",
      game_profile: gameProfile,
      error_code: errorCode
    }))
    return {
      config,
      poller,
      worker,
      start: { started: false, reason: errorCode },
      error: errorCode
    }
  }
}

function createGiftCodeWorkflowRuntime({
  client,
  initializationPromise,
  env = process.env,
  logger = console,
  getPoolFn = getPool,
  repositoryFactory = createGiftCodeRepository,
  clientFactory = createCenturyGameClient,
  notifierFactory = createDiscordGiftNotifier
}) {
  let verificationWorker = null
  let redemptionWorker = null
  let verificationProcessor = null
  let limiter = null
  let community = null
  let communityWorker = null
  let sourceIngestion = null
  let sourceWorker = null
  let cataloguePoller = null
  let health = {
    started: false,
    reason: "not started",
    verificationEnabled: verificationWorkerIsEnabled(env),
    redemptionEnabled: redemptionWorkerIsEnabled(env),
    verifierConfigured: false
  }

  const runtime = {
    async start() {
      if (health.started) return { ...health }
      if (!client?.isReady?.()) return { ...health, reason: "Discord client not ready" }
      const subsystem = await initializationPromise
      if (!subsystem?.available) {
        const gameProfile = subsystem?.gameProfile || String(env.GAME_PROFILE || "wos").trim()
        const sourceConfig = effectiveSourcePollingConfig(gameProfile, env)
        logger.log(sourceStartupDiagnostic(sourceConfig, false))
        health = { ...health, reason: subsystem?.reason || "database unavailable" }
        return { ...health }
      }
      try {
        const gameProfile = subsystem.gameProfile
        const repository = repositoryFactory(getPoolFn({ env, logger }), gameProfile)
        const sourceRepository = createGiftCodeSourceRepository(
          getPoolFn({ env, logger }),
          gameProfile
        )
        sourceIngestion = createGiftCodeSourceIngestionService({
          giftRepository: repository,
          sourceRepository,
          gameProfile,
          logger
        })
        const communityRepository = createGiftCodeCommunityRepository(
          getPoolFn({ env, logger }),
          gameProfile
        )
        const config = giftWorkerConfig(env)
        community = createGiftCodeCommunityService({
          repository: communityRepository,
          client,
          gameProfile,
          maximumEnabledAccounts: giftAccountConfig(env).maximumAutoRedeemAccountsPerUser,
          logger
        })
        const verifier = getVerifierAccount(gameProfile, env)
        limiter = new ConservativeRateLimiter({
          gameProfile,
          ...rateLimiterConfig(env),
          maximumRetries: 0
        })
        const centuryClient = clientFactory({ gameProfile, env, limiter })
        verificationProcessor = createVerificationProcessor({
          repository,
          client: centuryClient,
          verifier,
          config,
          botInstanceName: subsystem.botInstanceName,
          community,
          logger
        })
        const redemptionProcessor = createRedemptionProcessor({
          repository,
          client: centuryClient,
          notifier: notifierFactory({ client, gameProfile, logger }),
          community,
          config,
          logger
        })
        verificationWorker = createPollingWorker({
          tick: verificationProcessor.tick,
          intervalMs: config.pollIntervalMs,
          enabled: verificationWorkerIsEnabled(env) && verifier?.configured === true
        })
        redemptionWorker = createPollingWorker({
          tick: redemptionProcessor.tick,
          intervalMs: config.pollIntervalMs,
          enabled: redemptionWorkerIsEnabled(env)
        })
        communityWorker = createPollingWorker({
          tick: community.recoverOne,
          intervalMs: config.pollIntervalMs,
          enabled: true
        })
        const verificationStart = verificationWorker.start()
        const redemptionStart = redemptionWorker.start()
        communityWorker.start()
        const sourceRuntime = startCatalogueSourceRuntime({
          gameProfile,
          env,
          sourceRepository,
          sourceIngestion,
          logger
        })
        cataloguePoller = sourceRuntime.poller
        sourceWorker = sourceRuntime.worker
        health = {
          started: true,
          reason: null,
          gameProfile,
          verificationEnabled: verificationWorkerIsEnabled(env),
          redemptionEnabled: redemptionWorkerIsEnabled(env),
          verifierConfigured: verifier?.configured === true,
          verifierReason: verifier?.configured === false ? verifier.reason : null,
          verificationRunning: verificationStart.started,
          redemptionRunning: redemptionStart.started,
          sourcePollingEnabled: sourceRuntime.config.publicCatalogueEnabled,
          sourcePollingRunning: sourceRuntime.start.started,
          sourcePollingError: sourceRuntime.error
        }
        logger.log(`[Gift codes] Runtime ready for ${gameProfile}; verification=${verificationStart.started}, redemption=${redemptionStart.started}`)
      } catch (error) {
        health = { ...health, started: false, reason: sanitizeWorkerError(error) }
        logger.error(`[Gift codes] Runtime unavailable: ${sanitizeWorkerError(error)}`)
      }
      return { ...health }
    },

    async stop() {
      await Promise.all([
        verificationWorker?.stop(),
        redemptionWorker?.stop(),
        communityWorker?.stop(),
        sourceWorker?.stop()
      ])
      return { stopped: true }
    },

    async verifyCode(code) {
      if (!health.started) return { processed: false, reason: health.reason || "runtime unavailable" }
      if (!health.verificationEnabled) return { processed: false, reason: "verification worker disabled" }
      return verificationProcessor.verifyCode(code)
    },

    async ingestSourceMessage(message) {
      if (!health.started || !sourceIngestion) return null
      return sourceIngestion.ingestDiscordMessage(message)
    },

    status() {
      return {
        ...health,
        verificationRunning: verificationWorker?.isRunning() || false,
        redemptionRunning: redemptionWorker?.isRunning() || false,
        recentRateLimits: limiter?.getObservations().slice(-5) || [],
        sourcePollingRunning: sourceWorker?.isRunning() || false
      }
    },

    community() {
      return community
    }
  }
  currentRuntime = Object.freeze(runtime)
  return currentRuntime
}

module.exports = {
  getGiftCodeRuntime,
  sourceStartupDiagnostic,
  startCatalogueSourceRuntime,
  createGiftCodeWorkflowRuntime
}
