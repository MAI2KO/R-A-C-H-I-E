const test = require("node:test")
const assert = require("node:assert/strict")
const fs = require("node:fs/promises")
const path = require("node:path")
const { Pool } = require("pg")

const { createEventSchedulerRepository } = require("../src/eventSchedulerRepository")

const databaseUrl = process.env.TEST_DATABASE_URL
const migrationsDirectory = path.join(__dirname, "..", "migrations")

test("migration 007 preserves existing channel and state-link IDs", {
  skip: databaseUrl ? false : "TEST_DATABASE_URL is not configured"
}, async () => {
  const pool = new Pool({ connectionString: databaseUrl, max: 1 })
  const client = await pool.connect()
  const schema = `setup_ux_upgrade_${process.pid}`
  try {
    await client.query("SELECT pg_advisory_lock(7000009)")
    await client.query(`CREATE SCHEMA "${schema}"`)
    await client.query(`SET search_path TO "${schema}"`)
    for (const fileName of [
      "001_event_scheduler.sql",
      "002_event_images.sql",
      "003_event_delivery_indexes.sql",
      "004_state_delivery_reconciliation.sql",
      "005_event_management_and_roundups.sql",
      "006_flexible_reminders_and_alliances.sql"
    ]) {
      await client.query(await fs.readFile(path.join(migrationsDirectory, fileName), "utf8"))
    }
    await client.query(
      `INSERT INTO event_guild_settings
         (guild_id, game_profile, bot_instance_name, alliance_name, event_channel_id)
       VALUES ('old-alliance', 'wos', 'rachie-wos', 'YOU', 'old-reminder-channel')`
    )
    await client.query(
      `INSERT INTO event_state_links
         (alliance_guild_id, game_profile, configured_by_bot_instance,
          state_guild_id, state_event_channel_id, sharing_enabled)
       VALUES
         ('old-alliance', 'wos', 'rachie-wos', 'old-state', 'old-state-channel', true)`
    )

    await client.query(await fs.readFile(
      path.join(migrationsDirectory, "007_native_channel_and_state_link_setup.sql"),
      "utf8"
    ))

    assert.deepEqual((await client.query(
      `SELECT event_channel_id FROM event_guild_settings WHERE guild_id = 'old-alliance'`
    )).rows[0], { event_channel_id: "old-reminder-channel" })
    assert.deepEqual((await client.query(
      `SELECT state_guild_id, state_event_channel_id, sharing_enabled
         FROM event_state_links WHERE alliance_guild_id = 'old-alliance'`
    )).rows[0], {
      state_guild_id: "old-state",
      state_event_channel_id: "old-state-channel",
      sharing_enabled: true
    })
  } finally {
    await client.query("SET search_path TO public").catch(() => {})
    await client.query(`DROP SCHEMA IF EXISTS "${schema}" CASCADE`).catch(() => {})
    await client.query("SELECT pg_advisory_unlock(7000009)").catch(() => {})
    client.release()
    await pool.end()
  }
})

test("state destinations and one-time links remain profile scoped", {
  skip: databaseUrl ? false : "TEST_DATABASE_URL is not configured"
}, async () => {
  const pool = new Pool({ connectionString: databaseUrl, max: 4 })
  const allianceGuild = "6888888888888801"
  const otherAllianceGuild = "6888888888888802"
  const stateGuild = "6888888888888803"
  const stateChannel = "6888888888888804"
  const codeHash = "a".repeat(64)
  const expiredHash = "b".repeat(64)

  try {
    await pool.query(
      "DELETE FROM event_state_links WHERE alliance_guild_id IN ($1, $2)",
      [allianceGuild, otherAllianceGuild]
    )
    await pool.query("DELETE FROM event_state_destinations WHERE state_guild_id = $1", [stateGuild])
    await pool.query(
      "DELETE FROM event_guild_settings WHERE guild_id IN ($1, $2)",
      [allianceGuild, otherAllianceGuild]
    )

    const wos = createEventSchedulerRepository(pool, "wos")
    const kingshot = createEventSchedulerRepository(pool, "kingshot")
    await wos.upsertGuildIdentity({
      guildId: allianceGuild,
      botInstanceName: "rachie-wos",
      allianceName: "YOU"
    })
    assert.equal((await wos.getGuildSettings(allianceGuild)).event_channel_id, null)
    await wos.setEventChannel({ guildId: allianceGuild, eventChannelId: "6888888888888805" })
    assert.equal((await wos.getGuildSettings(allianceGuild)).event_channel_id, "6888888888888805")

    await wos.upsertStateDestination({
      stateGuildId: stateGuild,
      configuredByBotInstance: "rachie-wos",
      stateRoundupChannelId: stateChannel
    })
    await kingshot.upsertStateDestination({
      stateGuildId: stateGuild,
      configuredByBotInstance: "peggie-kingshot",
      stateRoundupChannelId: "6888888888888806"
    })
    assert.equal((await wos.getStateDestination(stateGuild)).state_roundup_channel_id, stateChannel)
    assert.equal(
      (await kingshot.getStateDestination(stateGuild)).state_roundup_channel_id,
      "6888888888888806"
    )

    await wos.createStateLinkCode({
      stateGuildId: stateGuild,
      codeHash,
      createdByBotInstance: "rachie-wos",
      createdByUserId: "6888888888888807",
      expiresAt: new Date(Date.now() + 60_000)
    })
    assert.equal(await kingshot.consumeStateLinkCode({
      allianceGuildId: otherAllianceGuild,
      configuredByBotInstance: "peggie-kingshot",
      codeHash
    }), null)

    const linked = await wos.consumeStateLinkCode({
      allianceGuildId: allianceGuild,
      configuredByBotInstance: "rachie-wos",
      codeHash
    })
    assert.equal(linked.state_guild_id, stateGuild)
    assert.equal(linked.state_event_channel_id, stateChannel)
    assert.equal(linked.sharing_enabled, true)
    assert.equal(await wos.consumeStateLinkCode({
      allianceGuildId: allianceGuild,
      configuredByBotInstance: "rachie-wos",
      codeHash
    }), null)

    await wos.createStateLinkCode({
      stateGuildId: stateGuild,
      codeHash: expiredHash,
      createdByBotInstance: "rachie-wos",
      createdByUserId: "6888888888888807",
      expiresAt: new Date(Date.now() + 60_000)
    })
    await pool.query(
      `UPDATE event_state_link_codes
          SET created_at = now() - interval '2 minutes',
              expires_at = now() - interval '1 minute'
        WHERE code_hash = $1`,
      [expiredHash]
    )
    assert.equal(await wos.consumeStateLinkCode({
      allianceGuildId: allianceGuild,
      configuredByBotInstance: "rachie-wos",
      codeHash: expiredHash
    }), null)

    const existing = await wos.getStateLink(allianceGuild)
    assert.equal(existing.state_guild_id, stateGuild)
  } finally {
    await pool.query(
      "DELETE FROM event_state_links WHERE alliance_guild_id IN ($1, $2)",
      [allianceGuild, otherAllianceGuild]
    ).catch(() => {})
    await pool.query("DELETE FROM event_state_destinations WHERE state_guild_id = $1", [stateGuild])
      .catch(() => {})
    await pool.query(
      "DELETE FROM event_guild_settings WHERE guild_id IN ($1, $2)",
      [allianceGuild, otherAllianceGuild]
    ).catch(() => {})
    await pool.end()
  }
})
