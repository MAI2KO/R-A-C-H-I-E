const fs = require("fs/promises")
const path = require("path")

const {
  postgresIsEnabled,
  getPool,
  closePool,
  classifyDatabaseError
} = require("./db")

const MIGRATION_LOCK_ID = "724194036187"
const DEFAULT_MIGRATIONS_DIR = path.join(__dirname, "..", "migrations")

async function listMigrationFiles(migrationsDir) {
  const entries = await fs.readdir(migrationsDir, { withFileTypes: true })

  return entries
    .filter(entry => entry.isFile() && /^\d+.*\.sql$/.test(entry.name))
    .map(entry => entry.name)
    .sort((left, right) => left.localeCompare(right))
}

function withoutDollarQuotedBodies(sql) {
  let result = ""
  let position = 0
  const openerPattern = /\$(?:[A-Za-z_][A-Za-z0-9_]*)?\$/g
  while (true) {
    openerPattern.lastIndex = position
    const opener = openerPattern.exec(sql)
    if (!opener) return result + sql.slice(position)
    const bodyEnd = sql.indexOf(opener[0], opener.index + opener[0].length)
    if (bodyEnd === -1) return result + sql.slice(position)
    result += sql.slice(position, opener.index)
    position = bodyEnd + opener[0].length
  }
}

function assertMigrationHasNoTransactionControl(fileName, sql) {
  if (/^\s*(BEGIN|COMMIT|ROLLBACK)\b/im.test(withoutDollarQuotedBodies(sql))) {
    throw new Error(`${fileName} must not contain transaction control statements`)
  }
}

async function ensureMigrationsTable(client) {
  await client.query(`
    CREATE TABLE IF NOT EXISTS schema_migrations (
      version text PRIMARY KEY,
      applied_at timestamptz NOT NULL DEFAULT now()
    )
  `)
}

async function runMigrations({
  pool = getPool(),
  migrationsDir = DEFAULT_MIGRATIONS_DIR,
  logger = console
} = {}) {
  if (!pool) {
    return { enabled: false, applied: [] }
  }

  const client = await pool.connect()
  const applied = []
  let lockAcquired = false

  try {
    await client.query("SELECT pg_advisory_lock($1::bigint)", [MIGRATION_LOCK_ID])
    lockAcquired = true

    await ensureMigrationsTable(client)
    const files = await listMigrationFiles(migrationsDir)

    for (const fileName of files) {
      const existing = await client.query(
        "SELECT 1 FROM schema_migrations WHERE version = $1",
        [fileName]
      )

      if (existing.rowCount > 0) continue

      const sql = await fs.readFile(path.join(migrationsDir, fileName), "utf8")
      assertMigrationHasNoTransactionControl(fileName, sql)

      await client.query("BEGIN")
      try {
        await client.query(sql)
        await client.query(
          "INSERT INTO schema_migrations (version) VALUES ($1)",
          [fileName]
        )
        await client.query("COMMIT")
        applied.push(fileName)
        logger.log(`[Event scheduler] Applied migration ${fileName}`)
      } catch (error) {
        await client.query("ROLLBACK")
        throw error
      }
    }

    return { enabled: true, applied }
  } finally {
    if (lockAcquired) {
      try {
        await client.query("SELECT pg_advisory_unlock($1::bigint)", [MIGRATION_LOCK_ID])
      } catch {
        logger.error("[Event scheduler] Could not release the migration lock cleanly")
      }
    }
    client.release()
  }
}

async function runFromCommandLine() {
  if (!postgresIsEnabled()) {
    console.log("[Postgres] Migrations skipped: database-backed subsystems disabled")
    return
  }

  try {
    const result = await runMigrations()
    console.log(
      `[Event scheduler] Migration check complete; applied ${result.applied.length}`
    )
  } catch (error) {
    console.error(
      `[Event scheduler] Migration failed: ${classifyDatabaseError(error)}`
    )
    process.exitCode = 1
  } finally {
    await closePool().catch(() => {})
  }
}

if (require.main === module) {
  runFromCommandLine()
}

module.exports = {
  MIGRATION_LOCK_ID,
  DEFAULT_MIGRATIONS_DIR,
  listMigrationFiles,
  withoutDollarQuotedBodies,
  assertMigrationHasNoTransactionControl,
  runMigrations
}
