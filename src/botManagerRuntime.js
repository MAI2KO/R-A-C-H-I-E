const { getPool, postgresIsEnabled, classifyDatabaseError } = require("./db")
const { runMigrations } = require("./migrate")

let health = Object.freeze({ available: false, reason: "not initialized" })

async function initializeBotManagerSubsystem({ env = process.env, logger = console } = {}) {
  if (!postgresIsEnabled(env)) {
    health = Object.freeze({ available: false, reason: "PostgreSQL is not enabled" })
    return health
  }
  try {
    const pool = getPool({ env, logger })
    await pool.query("SELECT 1")
    await runMigrations({ pool, logger })
    health = Object.freeze({ available: true, reason: null })
  } catch (error) {
    health = Object.freeze({ available: false, reason: classifyDatabaseError(error) })
    logger.error(`[Bot manager] Unavailable: ${health.reason}`)
  }
  return health
}

function getBotManagerHealth() { return health }

module.exports = { initializeBotManagerSubsystem, getBotManagerHealth }
