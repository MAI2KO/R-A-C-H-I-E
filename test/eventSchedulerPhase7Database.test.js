const test = require("node:test")
const assert = require("node:assert/strict")
const fs = require("node:fs/promises")
const path = require("node:path")
const { Pool } = require("pg")

const { buildDeliveryClaims } = require("../src/eventDeliveryGeneration")
const { createEventDeliveryRepository } = require("../src/eventDeliveryRepository")

const databaseUrl = process.env.TEST_DATABASE_URL
const migrationsDirectory = path.join(__dirname, "..", "migrations")

async function applyMigration(client, fileName) {
  const sql = await fs.readFile(path.join(migrationsDirectory, fileName), "utf8")
  await client.query(sql)
}

test("migration 005 preserves sent state history and terminally invalidates legacy claims", {
  skip: databaseUrl ? false : "TEST_DATABASE_URL is not configured"
}, async () => {
  const pool = new Pool({ connectionString: databaseUrl, max: 1 })
  const client = await pool.connect()
  const schema = `phase7_upgrade_${process.pid}`

  try {
    await client.query("SELECT pg_advisory_lock(7000006)")
    await client.query(`DROP SCHEMA IF EXISTS "${schema}" CASCADE`)
    await client.query(`CREATE SCHEMA "${schema}"`)
    await client.query(`SET search_path TO "${schema}"`)
    for (const fileName of [
      "001_event_scheduler.sql",
      "002_event_images.sql",
      "003_event_delivery_indexes.sql",
      "004_state_delivery_reconciliation.sql"
    ]) {
      await applyMigration(client, fileName)
    }

    await client.query(
      `INSERT INTO event_guild_settings
         (guild_id, game_profile, bot_instance_name, alliance_name, event_channel_id)
       VALUES ('legacy-alliance', 'wos', 'rachie-wos', 'Legacy', 'alliance-channel')`
    )
    const eventId = (await client.query(
      `INSERT INTO scheduled_events (
         guild_id, game_profile, created_by_bot_instance, alliance_name, event_name,
         first_occurrence_date, event_time_utc, recurrence_days,
         advance_reminder_minutes, reminder_at_start, publish_to_alliance,
         publish_to_state, include_in_weekly_roundup, status, created_by_user_id
       ) VALUES (
         'legacy-alliance', 'wos', 'rachie-wos', 'Legacy', 'Legacy Event',
         '2026-08-06', '12:30', 3, 30, true, true, true, true, 'active', 'legacy-user'
       ) RETURNING id`
    )).rows[0].id
    await client.query(
      `INSERT INTO event_delivery_claims (
         event_id, game_profile, occurrence_at, deliver_at, delivery_kind,
         target_kind, target_guild_id, target_channel_id, status,
         claimed_by_bot_instance, claimed_by_worker, claimed_at, claimed_until,
         sent_at, sent_message_id
       ) VALUES
         ($1, 'wos', '2026-08-06T12:30Z', '2026-08-06T12:00Z',
          'advance_reminder', 'state', 'state-guild', 'state-channel', 'sent',
          NULL, NULL, NULL, NULL, '2026-08-06T12:00Z', 'historical-state-message'),
         ($1, 'wos', '2026-08-09T12:30Z', '2026-08-09T12:00Z',
          'advance_reminder', 'state', 'state-guild', 'state-channel', 'pending',
          NULL, NULL, NULL, NULL, NULL, NULL),
         ($1, 'wos', '2026-08-12T12:30Z', '2026-08-12T12:30Z',
          'event_start', 'state', 'state-guild', 'state-channel', 'claimed',
          'rachie-wos', 'legacy-worker', '2026-08-06T11:59Z', '2026-08-06T12:30Z',
          NULL, NULL),
         ($1, 'wos', '2026-08-06T12:30Z', '2026-08-06T12:30Z',
          'event_start', 'alliance', 'legacy-alliance', 'alliance-channel', 'pending',
          NULL, NULL, NULL, NULL, NULL, NULL)`,
      [eventId]
    )

    await applyMigration(client, "005_event_management_and_roundups.sql")

    const rows = (await client.query(
      `SELECT target_kind, delivery_kind, status, claimed_by_bot_instance,
              claimed_by_worker, claimed_at, claimed_until, next_attempt_at,
              sent_message_id, last_error
         FROM event_delivery_claims ORDER BY id`
    )).rows
    assert.equal(rows[0].status, "sent")
    assert.equal(rows[0].sent_message_id, "historical-state-message")
    assert.equal(rows[0].last_error, null)
    assert.ok(rows.slice(1).every(row =>
      row.status === "failed"
      && row.claimed_by_bot_instance === null
      && row.claimed_by_worker === null
      && row.claimed_at === null
      && row.claimed_until === null
      && row.next_attempt_at === null
      && row.last_error
    ))

    const insertLegacyClaim = (deliveryKind, targetKind) => client.query(
      `INSERT INTO event_delivery_claims (
         event_id, game_profile, occurrence_at, deliver_at, delivery_kind,
         target_kind, target_guild_id, target_channel_id
       ) VALUES ($1, 'wos', '2026-08-15T12:30Z', '2026-08-15T12:00Z',
                 $2, $3, 'legacy-alliance', 'alliance-channel')`,
      [eventId, deliveryKind, targetKind]
    )
    await assert.rejects(insertLegacyClaim("advance_reminder", "state"), error =>
      error.code === "23514"
      && error.constraint === "event_delivery_claims_individual_policy_check"
    )
    await assert.rejects(insertLegacyClaim("event_start", "alliance"), error =>
      error.code === "23514"
      && error.constraint === "event_delivery_claims_individual_policy_check"
    )
  } finally {
    await client.query("SET search_path TO public").catch(() => {})
    await client.query(`DROP SCHEMA IF EXISTS "${schema}" CASCADE`).catch(() => {})
    await client.query("SELECT pg_advisory_unlock(7000006)").catch(() => {})
    client.release()
    await pool.end()
  }
})

test("the alliance worker expires final reminders and preserves profile isolation", {
  skip: databaseUrl ? false : "TEST_DATABASE_URL is not configured"
}, async () => {
  const pool = new Pool({ connectionString: databaseUrl, max: 8 })
  const isolationClient = await pool.connect()
  const guildId = "999999999999999921"
  const stateGuildId = "999999999999999922"
  const allianceChannel = "999999999999999923"
  const stateChannel = "999999999999999924"
  const userId = "999999999999999927"
  const now = new Date("2026-08-06T12:00:00Z")

  try {
    await isolationClient.query("SELECT pg_advisory_lock(7000005)")
    await pool.query(
      "DELETE FROM event_guild_settings WHERE guild_id = $1 AND game_profile IN ('wos', 'kingshot')",
      [guildId]
    )
    await pool.query(
      `INSERT INTO event_guild_settings
         (guild_id, game_profile, bot_instance_name, alliance_name, event_channel_id)
       VALUES
         ($1, 'wos', 'rachie-wos', 'WOS North', $2),
         ($1, 'kingshot', 'peggie-kingshot', 'Kingshot North', $2)`,
      [guildId, allianceChannel]
    )
    await pool.query(
      `INSERT INTO event_alliances (
         guild_id, game_profile, alliance_name, is_default, created_by_bot_instance
       ) VALUES
         ($1, 'wos', 'WOS North', true, 'rachie-wos'),
         ($1, 'kingshot', 'Kingshot North', true, 'peggie-kingshot')`,
      [guildId]
    )
    await pool.query(
      `INSERT INTO event_state_links
         (alliance_guild_id, game_profile, configured_by_bot_instance,
          state_guild_id, state_event_channel_id, sharing_enabled)
       VALUES
         ($1, 'wos', 'rachie-wos', $2, $3, true),
         ($1, 'kingshot', 'peggie-kingshot', $2, $3, true)`,
      [guildId, stateGuildId, stateChannel]
    )
    const events = await pool.query(
      `INSERT INTO scheduled_events (
         guild_id, game_profile, created_by_bot_instance, alliance_name, event_name,
         first_occurrence_date, event_time_utc, recurrence_days,
         advance_reminder_minutes, reminder_at_start, publish_to_alliance,
         publish_to_state, include_in_weekly_roundup, status, created_by_user_id
       ) VALUES
         ($1, 'wos', 'rachie-wos', 'WOS North', 'WOS Event',
          '2026-08-06', '12:30', 3, 30, true, true, true, true, 'active', $2),
         ($1, 'kingshot', 'peggie-kingshot', 'Kingshot North', 'Kingshot Event',
          '2026-08-06', '12:30', 3, 30, true, true, true, true, 'active', $2)
       RETURNING id, game_profile`,
      [guildId, userId]
    )
    const wosEventId = events.rows.find(row => row.game_profile === "wos").id
    const kingshotEventId = events.rows.find(row => row.game_profile === "kingshot").id
    const inserted = await pool.query(
      `INSERT INTO event_delivery_claims (
         event_id, game_profile, occurrence_at, deliver_at, delivery_kind,
         target_kind, target_guild_id, target_channel_id, status,
         claimed_by_bot_instance, claimed_by_worker, claimed_at, claimed_until,
         sent_at, sent_message_id
       ) VALUES
         ($1, 'wos', '2026-08-06T12:00Z', '2026-08-06T11:59Z',
          'final_reminder', 'alliance', $2, $3, 'pending', NULL, NULL, NULL, NULL, NULL, NULL),
         ($1, 'wos', '2026-08-06T12:30Z', '2026-08-06T12:00Z',
          'advance_reminder', 'alliance', $2, $3, 'pending', NULL, NULL, NULL, NULL, NULL, NULL),
         ($4, 'kingshot', '2026-08-06T12:30Z', '2026-08-06T12:00Z',
          'advance_reminder', 'alliance', $2, $3, 'pending', NULL, NULL, NULL, NULL, NULL, NULL)
       RETURNING id, game_profile, target_kind, delivery_kind, status`,
      [wosEventId, guildId, allianceChannel, kingshotEventId]
    )
    assert.equal(inserted.rowCount, 3)

    const repository = createEventDeliveryRepository(pool, "wos", { targetKind: "alliance" })
    const claims = await repository.claimDueDeliveries({
      now,
      batchSize: 10,
      leaseSeconds: 60,
      botInstanceName: "rachie-wos",
      workerId: "alliance-only-worker"
    })
    assert.equal(claims.length, 1)
    assert.equal(claims[0].target_kind, "alliance")
    assert.equal(claims[0].delivery_kind, "advance_reminder")

    const reconciled = await pool.query(
      `SELECT target_kind, delivery_kind, status, sent_message_id, next_attempt_at, last_error
         FROM event_delivery_claims
        WHERE event_id = $1
        ORDER BY id`,
      [wosEventId]
    )
    const invalidated = reconciled.rows.filter(row => row.status === "failed")
    assert.equal(invalidated.length, 1)
    assert.ok(invalidated.every(row => row.next_attempt_at === null && row.last_error))
    assert.ok(invalidated.some(row =>
      row.delivery_kind === "final_reminder" &&
      row.target_kind === "alliance" &&
      row.last_error === "Final reminder delivery window has passed."
    ))
    const claimedAlliance = reconciled.rows.find(row => row.status === "claimed")
    assert.equal(claimedAlliance.target_kind, "alliance")
    assert.equal(claimedAlliance.delivery_kind, "advance_reminder")

    const definitions = (await repository.listActiveEventDefinitions({
      rangeEnd: new Date("2026-08-07T00:00:00Z")
    })).filter(row => String(row.id) === String(wosEventId))
    const generated = buildDeliveryClaims(definitions, {
      gameProfile: "wos",
      windowStart: new Date("2026-08-06T11:00:00Z"),
      windowEnd: new Date("2026-08-06T13:00:00Z")
    })
    assert.deepEqual(generated.map(claim => claim.targetKind), ["alliance", "alliance"])
    assert.deepEqual(generated.map(claim => claim.deliveryKind), [
      "advance_reminder",
      "final_reminder"
    ])
    assert.equal(generated[1].deliverAt.toISOString(), "2026-08-06T12:29:00.000Z")

    const kingshotPending = await pool.query(
      `SELECT status FROM event_delivery_claims
        WHERE event_id = $1 AND game_profile = 'kingshot'`,
      [kingshotEventId]
    )
    assert.equal(kingshotPending.rows[0].status, "pending")
    const stateLink = await pool.query(
      `SELECT sharing_enabled, state_guild_id, state_event_channel_id
         FROM event_state_links WHERE alliance_guild_id = $1 AND game_profile = 'wos'`,
      [guildId]
    )
    assert.deepEqual(stateLink.rows[0], {
      sharing_enabled: true,
      state_guild_id: stateGuildId,
      state_event_channel_id: stateChannel
    })
  } finally {
    await pool.query(
      "DELETE FROM event_guild_settings WHERE guild_id = $1 AND game_profile IN ('wos', 'kingshot')",
      [guildId]
    ).catch(() => {})
    await isolationClient.query("SELECT pg_advisory_unlock(7000005)").catch(() => {})
    isolationClient.release()
    await pool.end()
  }
})
