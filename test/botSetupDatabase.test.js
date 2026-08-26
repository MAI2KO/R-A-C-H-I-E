const test = require("node:test")
const assert = require("node:assert/strict")
const { Pool } = require("pg")

const { runMigrations } = require("../src/migrate")
const { createBotSetupRepository } = require("../src/botSetupRepository")

const databaseUrl = process.env.TEST_DATABASE_URL

test("bot-managed setup persistence is durable, idempotent and profile scoped", {
  skip: databaseUrl ? false : "TEST_DATABASE_URL is not configured"
}, async () => {
  const admin = new Pool({ connectionString: databaseUrl, max: 1 })
  const schema = `bot_setup_${process.pid}_${Date.now()}`
  let pool
  try {
    await admin.query(`CREATE SCHEMA "${schema}"`)
    pool = new Pool({ connectionString: databaseUrl, max: 8, options: `-c search_path=${schema}` })
    const migrated = await runMigrations({ pool, logger: { log() {}, error() {} } })
    assert.equal(migrated.applied.at(-1), "019_daily_event_recurrence.sql")
    assert.deepEqual((await runMigrations({ pool, logger: { log() {}, error() {} } })).applied, [])

    const guildId = "700000000000000001"
    const values = {
      category_id: "700000000000000002",
      gift_auto_redeem_channel_id: "700000000000000003",
      gift_announcements_channel_id: "700000000000000004",
      minister_sign_up_channel_id: "700000000000000005",
      event_scheduler_channel_id: "700000000000000006",
      event_announcements_channel_id: "700000000000000007",
      gift_auto_redeem_message_id: "700000000000000008",
      minister_sign_up_message_id: "700000000000000009",
      event_scheduler_message_id: "700000000000000010"
    }
    const wos = createBotSetupRepository(pool, "wos")
    const kingshot = createBotSetupRepository(pool, "kingshot")
    await pool.query(
      `INSERT INTO gift_code_guild_settings (
         game_profile, guild_id, gift_code_channel_id
       ) VALUES ('wos', $1, '700000000000000098')`,
      [guildId]
    )
    await pool.query(
      `INSERT INTO gift_code_engagement_events (
         id, game_profile, guild_id, event_type, discord_user_id,
         status, channel_id, message_id
       ) VALUES (
         '11111111-1111-4111-8111-111111111111', 'wos', $1,
         'auto_redeem_join', '700000000000000020', 'completed',
         '700000000000000098', '700000000000000097'
       )`,
      [guildId]
    )
    await wos.save(guildId, values)
    assert.equal(await kingshot.get(guildId), null)
    assert.equal((await wos.get(guildId)).gift_announcements_channel_id, values.gift_announcements_channel_id)

    await pool.query(
      `INSERT INTO event_guild_settings (
         guild_id, game_profile, bot_instance_name, alliance_name, event_channel_id
       ) VALUES ($1, 'wos', 'rachie-wos', 'HwC', '700000000000000099')`,
      [guildId]
    )
    await wos.reconcileDestinations(
      guildId, values.gift_announcements_channel_id, values.event_announcements_channel_id
    )
    assert.equal((await pool.query(
      `SELECT gift_code_channel_id FROM gift_code_guild_settings
        WHERE game_profile = 'wos' AND guild_id = $1`, [guildId]
    )).rows[0].gift_code_channel_id, values.gift_announcements_channel_id)
    const scheduler = (await pool.query(
      `SELECT event_channel_id, weekly_roundup_channel_id FROM event_guild_settings
        WHERE game_profile = 'wos' AND guild_id = $1`, [guildId]
    )).rows[0]
    assert.equal(scheduler.event_channel_id, values.event_announcements_channel_id)
    assert.equal(scheduler.weekly_roundup_channel_id, values.event_announcements_channel_id)
    const untouchedStatusCard = (await pool.query(
      `SELECT channel_id, message_id FROM gift_code_engagement_events
        WHERE id = '11111111-1111-4111-8111-111111111111'`
    )).rows[0]
    assert.deepEqual(untouchedStatusCard, {
      channel_id: "700000000000000098",
      message_id: "700000000000000097"
    }, "/bot-setup must not bulk-migrate existing status cards")

    values.gift_auto_redeem_message_id = "700000000000000011"
    await wos.save(guildId, values)
    assert.equal((await pool.query(
      `SELECT COUNT(*)::integer AS count FROM bot_managed_discord_setups
        WHERE game_profile = 'wos' AND guild_id = $1`, [guildId]
    )).rows[0].count, 1)
    assert.equal((await wos.get(guildId)).gift_auto_redeem_message_id, values.gift_auto_redeem_message_id)
  } finally {
    await pool?.end().catch(() => {})
    await admin.query(`DROP SCHEMA IF EXISTS "${schema}" CASCADE`).catch(() => {})
    await admin.end()
  }
})
