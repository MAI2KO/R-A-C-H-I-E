const assert = require("node:assert/strict")
const test = require("node:test")
const { randomUUID } = require("node:crypto")
const fs = require("node:fs/promises")
const path = require("node:path")
const { Pool } = require("pg")

const { runMigrations } = require("../src/migrate")
const { createPlayerRepository } = require("../src/giftCodes/playerRepository")

const databaseUrl = process.env.TEST_DATABASE_URL

test("active identity guard preserves legacy inactive rows and blocks new invalid active rows", {
  skip: databaseUrl ? false : "TEST_DATABASE_URL is not configured"
}, async () => {
  const admin = new Pool({ connectionString: databaseUrl, max: 1 })
  const schema = `player_identity_guard_${process.pid}_${Date.now()}`
  let pool
  try {
    await admin.query(`CREATE SCHEMA "${schema}"`)
    pool = new Pool({ connectionString: databaseUrl, max: 2,
      options: `-c search_path=${schema}` })
    await runMigrations({ pool, logger: { log() {}, error() {} } })
    await pool.query(
      "ALTER TABLE player_accounts DROP CONSTRAINT player_accounts_active_identity_check"
    )
    const legacyId = randomUUID()
    const ownerlessId = randomUUID()
    await pool.query(
      `INSERT INTO player_accounts
         (id,game_profile,discord_user_id,player_id,state_or_kingdom_number,is_active)
       VALUES ($1,'wos','12345','67890','1001',true)`, [legacyId]
    )
    await pool.query(
      `INSERT INTO player_accounts
         (id,game_profile,discord_user_id,player_id,state_or_kingdom_number,
          in_game_name,alliance_abbreviation,is_active)
       VALUES ($1,'wos',NULL,'67894','1001','Ownerless','TAG',true)`, [ownerlessId]
    )
    await pool.query(await fs.readFile(path.join(__dirname, "..", "migrations",
      "022_active_player_account_identity_guard.sql"), "utf8"))
    const constraint = (await pool.query(
      `SELECT convalidated FROM pg_constraint
        WHERE conname='player_accounts_active_identity_check'`
    )).rows[0]
    assert.equal(constraint.convalidated, false)
    assert.equal((await pool.query(
      "SELECT is_active FROM player_accounts WHERE id=$1", [legacyId]
    )).rows[0].is_active, true)
    await assert.rejects(pool.query(
      "UPDATE player_accounts SET updated_at_utc=now() WHERE id=$1", [legacyId]
    ), error => error.code === "23514"
      && error.constraint === "player_accounts_active_identity_check")
    await pool.query("UPDATE player_accounts SET is_active=false WHERE id=$1", [legacyId])
    const ownerless = await createPlayerRepository(pool, "wos").deactivateInvalidUnownedAccount({
      accountId: ownerlessId, playerId: "67894", operatorDiscordUserId: "99999",
      accountRef: "0123456789abcdef", reasons: ["invalid_owner"]
    })
    assert.equal(ownerless.is_active, false)
    assert.equal(ownerless.account_metadata.legacyCleanup.action, "deactivate_invalid_unowned")
    const completed = await createPlayerRepository(pool, "wos")
      .findCompletedPlayerCleanup("0123456789abcdef")
    assert.equal(completed.alreadyCompleted, true)
    assert.equal(completed.previousOwnerDiscordUserId, null)
    await pool.query(
      `UPDATE player_accounts
          SET discord_user_id='12346',in_game_name='Reclaimed',alliance_abbreviation='TAG',
              is_active=true,is_primary=true
        WHERE id=$1`,
      [ownerlessId]
    )
    assert.equal(await createPlayerRepository(pool, "wos")
      .findCompletedPlayerCleanup("0123456789abcdef"), null)
    await assert.rejects(pool.query(
      `INSERT INTO player_accounts
         (id,game_profile,discord_user_id,player_id,state_or_kingdom_number,is_active)
       VALUES ($1,'wos','12345','67891','1001',true)`, [randomUUID()]
    ), error => error.code === "23514"
      && error.constraint === "player_accounts_active_identity_check")
    await assert.rejects(pool.query(
      `INSERT INTO player_accounts
         (id,game_profile,discord_user_id,player_id,state_or_kingdom_number,in_game_name,is_active)
       VALUES ($1,'wos','12345','67893','1001','Valid Player',true)`, [randomUUID()]
    ), error => error.code === "23514"
      && error.constraint === "player_accounts_active_identity_check")
    await pool.query(
      `INSERT INTO player_accounts
         (id,game_profile,discord_user_id,player_id,state_or_kingdom_number,
          in_game_name,alliance_abbreviation,is_active)
       VALUES ($1,'wos','12345','67892','1001','Valid Player','TAG',true)`, [randomUUID()]
    )
  } finally {
    await pool?.end().catch(() => {})
    await admin.query(`DROP SCHEMA IF EXISTS "${schema}" CASCADE`).catch(() => {})
    await admin.end()
  }
})
