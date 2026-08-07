const test = require("node:test")
const assert = require("node:assert/strict")
const fs = require("node:fs/promises")
const os = require("node:os")
const path = require("node:path")
const { Pool } = require("pg")

const { runMigrations } = require("../src/migrate")

const databaseUrl = process.env.TEST_DATABASE_URL

test("concurrent migration runners serialize and apply each migration once", {
  skip: databaseUrl ? false : "TEST_DATABASE_URL is not configured"
}, async () => {
  const adminPool = new Pool({ connectionString: databaseUrl, max: 1 })
  const schema = `migration_concurrency_${process.pid}_${Date.now()}`
  const migrationsDir = await fs.mkdtemp(path.join(os.tmpdir(), "rachie-migrations-"))
  let scopedPool

  try {
    await adminPool.query(`CREATE SCHEMA "${schema}"`)
    await fs.writeFile(
      path.join(migrationsDir, "001_concurrency_probe.sql"),
      `SELECT pg_sleep(0.15);\n` +
      `CREATE TABLE migration_concurrency_probe (id integer PRIMARY KEY);\n` +
      `INSERT INTO migration_concurrency_probe (id) VALUES (1);\n`,
      "utf8"
    )
    scopedPool = new Pool({
      connectionString: databaseUrl,
      max: 4,
      options: `-c search_path=${schema}`
    })
    const logger = { log() {}, error() {} }

    const results = await Promise.all([
      runMigrations({ pool: scopedPool, migrationsDir, logger }),
      runMigrations({ pool: scopedPool, migrationsDir, logger })
    ])

    assert.deepEqual(results.map(result => result.applied.length).sort(), [0, 1])
    assert.equal(results.flatMap(result => result.applied).length, 1)
    assert.deepEqual(results.flatMap(result => result.applied), ["001_concurrency_probe.sql"])

    const migrationCount = await scopedPool.query(
      "SELECT COUNT(*)::integer AS count FROM schema_migrations"
    )
    const probeCount = await scopedPool.query(
      "SELECT COUNT(*)::integer AS count FROM migration_concurrency_probe"
    )
    assert.equal(migrationCount.rows[0].count, 1)
    assert.equal(probeCount.rows[0].count, 1)
  } finally {
    await scopedPool?.end().catch(() => {})
    await adminPool.query(`DROP SCHEMA IF EXISTS "${schema}" CASCADE`).catch(() => {})
    await adminPool.end()
    await fs.rm(migrationsDir, { recursive: true, force: true })
  }
})
