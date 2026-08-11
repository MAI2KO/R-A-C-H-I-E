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

let currentRuntime = null

function getGiftCodeRuntime() {
  return currentRuntime
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
        health = { ...health, reason: subsystem?.reason || "database unavailable" }
        return { ...health }
      }
      try {
        const gameProfile = subsystem.gameProfile
        const repository = repositoryFactory(getPoolFn({ env, logger }), gameProfile)
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
        health = {
          started: true,
          reason: null,
          gameProfile,
          verificationEnabled: verificationWorkerIsEnabled(env),
          redemptionEnabled: redemptionWorkerIsEnabled(env),
          verifierConfigured: verifier?.configured === true,
          verifierReason: verifier?.configured === false ? verifier.reason : null,
          verificationRunning: verificationStart.started,
          redemptionRunning: redemptionStart.started
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
        communityWorker?.stop()
      ])
      return { stopped: true }
    },

    async verifyCode(code) {
      if (!health.started) return { processed: false, reason: health.reason || "runtime unavailable" }
      if (!health.verificationEnabled) return { processed: false, reason: "verification worker disabled" }
      return verificationProcessor.verifyCode(code)
    },

    status() {
      return {
        ...health,
        verificationRunning: verificationWorker?.isRunning() || false,
        redemptionRunning: redemptionWorker?.isRunning() || false,
        recentRateLimits: limiter?.getObservations().slice(-5) || []
      }
    },

    community() {
      return community
    }
  }
  currentRuntime = Object.freeze(runtime)
  return currentRuntime
}

module.exports = { getGiftCodeRuntime, createGiftCodeWorkflowRuntime }
