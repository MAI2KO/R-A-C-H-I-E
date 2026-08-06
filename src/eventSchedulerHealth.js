const {
  schedulerIsEnabled,
  getPool,
  closePool,
  classifyDatabaseError
} = require("./db")
const { runMigrations } = require("./migrate")

const SUPPORTED_GAME_PROFILES = new Set(["wos", "kingshot"])

let health = {
  enabled: false,
  available: false,
  gameProfile: null,
  botInstanceName: null,
  reason: "not initialized"
}

function getEventSchedulerHealth() {
  return { ...health }
}

function eventSchedulerIsAvailable() {
  return health.available
}

function validateOwnershipConfiguration(env) {
  const gameProfile = String(env.GAME_PROFILE || "wos").trim()
  const botInstanceName = String(env.BOT_INSTANCE_NAME || "").trim()

  if (!SUPPORTED_GAME_PROFILES.has(gameProfile)) {
    throw new Error("unsupported GAME_PROFILE")
  }

  if (!botInstanceName) {
    throw new Error("missing BOT_INSTANCE_NAME")
  }

  return { gameProfile, botInstanceName }
}

async function initializeEventSchedulerSubsystem({
  env = process.env,
  logger = console
} = {}) {
  if (!schedulerIsEnabled(env)) {
    health = {
      enabled: false,
      available: false,
      gameProfile: null,
      botInstanceName: null,
      reason: "disabled"
    }
    logger.log("[Event scheduler] Disabled")
    return getEventSchedulerHealth()
  }

  let ownership
  try {
    ownership = validateOwnershipConfiguration(env)
    const pool = getPool({ env, logger })
    await pool.query("SELECT 1")
    await runMigrations({ pool, logger })

    health = {
      enabled: true,
      available: true,
      gameProfile: ownership.gameProfile,
      botInstanceName: ownership.botInstanceName,
      reason: null
    }
    logger.log(
      `[Event scheduler] Available for ${ownership.gameProfile} as ${ownership.botInstanceName}`
    )
  } catch (error) {
    const reason = error.message === "unsupported GAME_PROFILE"
      || error.message === "missing BOT_INSTANCE_NAME"
      ? error.message
      : classifyDatabaseError(error)

    health = {
      enabled: true,
      available: false,
      gameProfile: ownership?.gameProfile || null,
      botInstanceName: ownership?.botInstanceName || null,
      reason
    }
    logger.error(`[Event scheduler] Unavailable: ${reason}`)
    await closePool().catch(() => {})
  }

  return getEventSchedulerHealth()
}

module.exports = {
  SUPPORTED_GAME_PROFILES,
  getEventSchedulerHealth,
  eventSchedulerIsAvailable,
  validateOwnershipConfiguration,
  initializeEventSchedulerSubsystem
}
