const {
  playerGiftCodesIsEnabled,
  schedulerIsEnabled,
  getPool,
  closePool,
  classifyDatabaseError
} = require("../db")
const { runMigrations } = require("../migrate")
const { profileTerminology } = require("./terminology")

let health = {
  enabled: false,
  available: false,
  gameProfile: null,
  botInstanceName: null,
  reason: "not initialized"
}

function getPlayerGiftCodesHealth() {
  return { ...health }
}

async function initializePlayerGiftCodesSubsystem({ env = process.env, logger = console } = {}) {
  if (!playerGiftCodesIsEnabled(env)) {
    health = {
      enabled: false,
      available: false,
      gameProfile: null,
      botInstanceName: null,
      reason: "disabled"
    }
    logger.log("[Player accounts] Disabled")
    return getPlayerGiftCodesHealth()
  }

  const gameProfile = String(env.GAME_PROFILE || "wos").trim()
  const botInstanceName = String(env.BOT_INSTANCE_NAME || "").trim()
  try {
    profileTerminology(gameProfile)
    if (!botInstanceName) throw new Error("missing BOT_INSTANCE_NAME")
    const pool = getPool({ env, logger })
    await pool.query("SELECT 1")
    await runMigrations({ pool, logger })
    health = {
      enabled: true,
      available: true,
      gameProfile,
      botInstanceName,
      reason: null
    }
    logger.log(`[Player accounts] Available for ${gameProfile} as ${botInstanceName}`)
  } catch (error) {
    const reason = error.message === "Unsupported game profile"
      || error.message === "missing BOT_INSTANCE_NAME"
      ? error.message
      : classifyDatabaseError(error)
    health = {
      enabled: true,
      available: false,
      gameProfile: ["wos", "kingshot"].includes(gameProfile) ? gameProfile : null,
      botInstanceName: botInstanceName || null,
      reason
    }
    logger.error(`[Player accounts] Unavailable: ${reason}`)
    if (!schedulerIsEnabled(env)) await closePool().catch(() => {})
  }
  return getPlayerGiftCodesHealth()
}

module.exports = {
  getPlayerGiftCodesHealth,
  initializePlayerGiftCodesSubsystem
}
