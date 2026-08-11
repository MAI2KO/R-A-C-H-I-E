const test = require("node:test")
const assert = require("node:assert/strict")
const { Pool } = require("pg")

const { closePool, getPool } = require("../src/db")

const databaseUrl = process.env.TEST_DATABASE_URL

test("shared PostgreSQL pool evicts a terminated client and serves later scheduler queries", {
  skip: databaseUrl ? false : "TEST_DATABASE_URL is not configured"
}, async () => {
  const errors = []
  const env = {
    DATABASE_URL: databaseUrl,
    EVENT_SCHEDULER_ENABLED: "true",
    PLAYER_GIFT_CODES_ENABLED: "true"
  }
  const logger = { error: message => errors.push(message) }
  const schedulerPool = getPool({ env, logger })
  const giftCodePool = getPool({ env, logger })
  const admin = new Pool({ connectionString: databaseUrl, max: 1 })
  try {
    assert.equal(schedulerPool, giftCodePool)
    const firstPid = (await schedulerPool.query("SELECT pg_backend_pid() AS pid")).rows[0].pid
    const poolError = new Promise(resolve => schedulerPool.once("error", resolve))
    assert.equal((await admin.query(
      "SELECT pg_terminate_backend($1) AS terminated",
      [firstPid]
    )).rows[0].terminated, true)
    let timeout
    try {
      await Promise.race([
        poolError,
        new Promise((_, reject) => {
          timeout = setTimeout(
            () => reject(new Error("terminated PostgreSQL client was not evicted")),
            2000
          )
        })
      ])
    } finally {
      clearTimeout(timeout)
    }

    const secondPid = (await schedulerPool.query("SELECT pg_backend_pid() AS pid")).rows[0].pid
    assert.notEqual(secondPid, firstPid)
    assert.ok(errors.some(message => message.includes("Postgres pool error")))
  } finally {
    await admin.end()
    await closePool()
  }
})
