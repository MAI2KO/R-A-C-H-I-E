const { Pool } = require("pg")

const CONNECTION_TIMEOUT_MS = 5000
const QUERY_TIMEOUT_MS = 15000

let pool = null

class DatabaseConfigurationError extends Error {
  constructor(code, message) {
    super(message)
    this.name = "DatabaseConfigurationError"
    this.code = code
  }
}

function schedulerIsEnabled(env = process.env) {
  return env.EVENT_SCHEDULER_ENABLED === "true"
}

function validateDatabaseUrl(databaseUrl) {
  if (!databaseUrl) {
    throw new DatabaseConfigurationError(
      "DATABASE_URL_MISSING",
      "DATABASE_URL is not configured"
    )
  }

  let parsed
  try {
    parsed = new URL(databaseUrl)
  } catch {
    throw new DatabaseConfigurationError(
      "DATABASE_URL_INVALID",
      "DATABASE_URL is invalid"
    )
  }

  if (!["postgres:", "postgresql:"].includes(parsed.protocol) || !parsed.hostname) {
    throw new DatabaseConfigurationError(
      "DATABASE_URL_INVALID",
      "DATABASE_URL is invalid"
    )
  }
}

function classifyDatabaseError(error) {
  if (error?.code === "DATABASE_URL_MISSING") return "missing DATABASE_URL"
  if (error?.code === "DATABASE_URL_INVALID") return "invalid DATABASE_URL"
  if (error?.code === "ENOTFOUND" || error?.code === "EAI_AGAIN") return "DNS failure"
  if (["ECONNREFUSED", "ECONNRESET", "ENETUNREACH", "EHOSTUNREACH"].includes(error?.code)) {
    return "network failure"
  }
  if (error?.code === "28P01" || error?.code === "28000") return "authentication failure"
  if (error?.code === "ETIMEDOUT" || error?.code === "57014") return "connection timeout"
  if (String(error?.message || "").toLowerCase().includes("timeout")) return "connection timeout"
  return "database failure"
}

function getPool({ env = process.env, logger = console } = {}) {
  if (!schedulerIsEnabled(env)) return null
  if (pool) return pool

  validateDatabaseUrl(env.DATABASE_URL)

  pool = new Pool({
    connectionString: env.DATABASE_URL,
    connectionTimeoutMillis: CONNECTION_TIMEOUT_MS,
    query_timeout: QUERY_TIMEOUT_MS,
    statement_timeout: QUERY_TIMEOUT_MS,
    max: 5,
    idleTimeoutMillis: 30000
  })

  pool.on("error", error => {
    logger.error(
      `[Event scheduler] Postgres pool error: ${classifyDatabaseError(error)}`
    )
  })

  return pool
}

async function closePool() {
  if (!pool) return

  const currentPool = pool
  pool = null
  await currentPool.end()
}

module.exports = {
  CONNECTION_TIMEOUT_MS,
  QUERY_TIMEOUT_MS,
  DatabaseConfigurationError,
  schedulerIsEnabled,
  validateDatabaseUrl,
  classifyDatabaseError,
  getPool,
  closePool
}
