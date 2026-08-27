const test = require("node:test")
const assert = require("node:assert/strict")
const fs = require("node:fs/promises")
const path = require("node:path")
const { Pool } = require("pg")

const { createEventSchedulerRepository } = require("../src/eventSchedulerRepository")
const { runMigrations } = require("../src/migrate")
const { buildStateEventDeliveryClaims } = require("../src/stateEventDeliveryGeneration")
const { createStateEventRepository } = require("../src/stateEventRepository")

const databaseUrl = process.env.TEST_DATABASE_URL
const migrationsDirectory = path.join(__dirname, "..", "migrations")

function scopedPool(schema, max = 8) {
  return new Pool({ connectionString: databaseUrl, max, options: `-c search_path=${schema}` })
}

async function dropSchema(pool, schema) {
  await pool.query(`DROP SCHEMA IF EXISTS "${schema}" CASCADE`).catch(() => {})
}

test("migrations 001-021 apply cleanly, reapply to zero, and serialize concurrent startup", {
  skip: databaseUrl ? false : "TEST_DATABASE_URL is not configured"
}, async () => {
  const admin = new Pool({ connectionString: databaseUrl, max: 1 })
  const cleanSchema = `state_migrate_clean_${process.pid}_${Date.now()}`
  const concurrentSchema = `state_migrate_concurrent_${process.pid}_${Date.now()}`
  let clean
  let concurrent
  const logger = { log() {}, error() {} }
  try {
    await admin.query(`CREATE SCHEMA "${cleanSchema}"`)
    clean = scopedPool(cleanSchema)
    const first = await runMigrations({ pool: clean, logger })
    const second = await runMigrations({ pool: clean, logger })
    assert.equal(first.applied.length, 21)
    assert.equal(first.applied.at(-1), "021_native_bot_manager_role.sql")
    assert.deepEqual(second.applied, [])

    await admin.query(`CREATE SCHEMA "${concurrentSchema}"`)
    concurrent = scopedPool(concurrentSchema)
    const results = await Promise.all([
      runMigrations({ pool: concurrent, logger }),
      runMigrations({ pool: concurrent, logger })
    ])
    assert.deepEqual(results.map(result => result.applied.length).sort((a, b) => a - b), [0, 21])
    const versions = await concurrent.query(
      "SELECT version, COUNT(*)::integer AS count FROM schema_migrations GROUP BY version ORDER BY version"
    )
    assert.equal(versions.rowCount, 21)
    assert.equal(versions.rows.find(row => row.version === "009_state_events.sql").count, 1)
    assert.equal(
      versions.rows.find(row => row.version === "010_cross_midnight_event_groups.sql").count,
      1
    )
    assert.equal(
      versions.rows.find(row => row.version === "011_player_accounts_and_gift_codes.sql").count,
      1
    )
    assert.equal(
      versions.rows.find(row => row.version === "012_gift_code_workers.sql").count,
      1
    )
    assert.equal(
      versions.rows.find(row => row.version === "013_gift_code_community.sql").count,
      1
    )
    assert.equal(
      versions.rows.find(row => row.version === "014_gift_code_community_reconciliation.sql").count,
      1
    )
    assert.equal(
      versions.rows.find(row => row.version === "015_gift_code_guild_enrolment.sql").count,
      1
    )
    assert.equal(
      versions.rows.find(row => row.version === "017_bot_managed_discord_setup.sql").count,
      1
    )
    assert.equal(
      versions.rows.find(row => row.version === "018_player_account_ownership_release.sql").count,
      1
    )
  } finally {
    await clean?.end().catch(() => {})
    await concurrent?.end().catch(() => {})
    await dropSchema(admin, cleanSchema)
    await dropSchema(admin, concurrentSchema)
    await admin.end()
  }
})

test("migrations 009-010 preserve recurrence and legacy group schedules", {
  skip: databaseUrl ? false : "TEST_DATABASE_URL is not configured"
}, async () => {
  const admin = new Pool({ connectionString: databaseUrl, max: 1 })
  const schema = `state_upgrade_${process.pid}_${Date.now()}`
  let pool
  try {
    await admin.query(`CREATE SCHEMA "${schema}"`)
    pool = scopedPool(schema, 1)
    for (const fileName of (await fs.readdir(migrationsDirectory)).filter(name => /^00[1-8]_/.test(name)).sort()) {
      await pool.query(await fs.readFile(path.join(migrationsDirectory, fileName), "utf8"))
    }
    await pool.query(
      `INSERT INTO event_guild_settings
         (guild_id, game_profile, bot_instance_name, alliance_name, event_channel_id)
       VALUES ('upgrade-guild', 'wos', 'rachie-wos', 'YOU', 'upgrade-channel')`
    )
    const alliance = (await pool.query(
      `INSERT INTO event_alliances
         (guild_id, game_profile, alliance_name, is_default, created_by_bot_instance)
       VALUES ('upgrade-guild', 'wos', 'YOU', true, 'rachie-wos') RETURNING id`
    )).rows[0]
    await pool.query(
      `INSERT INTO scheduled_events (
         guild_id, game_profile, created_by_bot_instance, alliance_name, alliance_id,
         event_name, first_occurrence_date, event_time_utc, recurrence_days,
         advance_reminder_minutes, reminder_at_start, created_by_user_id
       ) VALUES ('upgrade-guild', 'wos', 'rachie-wos', 'YOU', $1, 'Three Day',
                 '2026-08-10', '10:00', 3, NULL, false, 'user')`,
      [alliance.id]
    )
    await pool.query(await fs.readFile(path.join(migrationsDirectory, "009_state_events.sql"), "utf8"))
    assert.equal((await pool.query(
      "SELECT recurrence_days FROM scheduled_events WHERE event_name = 'Three Day'"
    )).rows[0].recurrence_days, 3)
    await pool.query("UPDATE scheduled_events SET recurrence_days = 2 WHERE event_name = 'Three Day'")
    await pool.query("UPDATE scheduled_events SET recurrence_days = 42 WHERE event_name = 'Three Day'")
    await pool.query(
      `INSERT INTO scheduled_event_groups
         (event_id, game_profile, group_name, event_time_utc, sort_order)
       SELECT id, game_profile, 'Legacy Group', '00:15', 0
         FROM scheduled_events WHERE event_name = 'Three Day'`
    )
    await pool.query(await fs.readFile(
      path.join(migrationsDirectory, "010_cross_midnight_event_groups.sql"),
      "utf8"
    ))
    await pool.query(await fs.readFile(
      path.join(migrationsDirectory, "019_daily_event_recurrence.sql"),
      "utf8"
    ))
    await pool.query("UPDATE scheduled_events SET recurrence_days = 1 WHERE event_name = 'Three Day'")
    assert.equal((await pool.query(
      "SELECT recurrence_days FROM scheduled_events WHERE event_name = 'Three Day'"
    )).rows[0].recurrence_days, 1)
    const recurrenceConstraints = await pool.query(
      `SELECT conname,pg_get_constraintdef(oid) AS definition
        FROM pg_constraint
        WHERE conname IN ('scheduled_events_recurrence_check','state_events_recurrence_check')
          AND connamespace = current_schema()::regnamespace
        ORDER BY conname`
    )
    assert.equal(recurrenceConstraints.rows.length, 2)
    assert.equal(recurrenceConstraints.rows.every(row => /\b1\b/.test(row.definition)), true)
    const legacyGroup = (await pool.query(
      `SELECT first_occurrence_date::text AS first_occurrence_date
         FROM scheduled_event_groups WHERE group_name = 'Legacy Group'`
    )).rows[0]
    assert.equal(legacyGroup.first_occurrence_date, null)
  } finally {
    await pool?.end().catch(() => {})
    await dropSchema(admin, schema)
    await admin.end()
  }
})

test("state events preserve history, isolate profiles, deduplicate guilds and aggregate roundups", {
  skip: databaseUrl ? false : "TEST_DATABASE_URL is not configured"
}, async () => {
  const admin = new Pool({ connectionString: databaseUrl, max: 1 })
  const schema = `state_runtime_${process.pid}_${Date.now()}`
  let pool
  try {
    await admin.query(`CREATE SCHEMA "${schema}"`)
    pool = scopedPool(schema)
    await runMigrations({ pool, logger: { log() {}, error() {} } })
    const wosScheduler = createEventSchedulerRepository(pool, "wos")
    const ksScheduler = createEventSchedulerRepository(pool, "kingshot")
    const wos = createStateEventRepository(pool, "wos")
    const kingshot = createStateEventRepository(pool, "kingshot")

    await wosScheduler.upsertStateDestination({
      stateGuildId: "state-guild", configuredByBotInstance: "rachie-wos",
      stateRoundupChannelId: "state-channel", stateNumber: "689"
    })
    await ksScheduler.upsertStateDestination({
      stateGuildId: "state-guild", configuredByBotInstance: "peggie-kingshot",
      stateRoundupChannelId: "ks-state-channel", stateNumber: "77"
    })
    await pool.query(
      `INSERT INTO event_guild_settings
         (guild_id, game_profile, bot_instance_name, alliance_name, event_channel_id)
       VALUES ('alliance-guild', 'wos', 'rachie-wos', 'YOU', 'alliance-channel')`
    )
    await pool.query(
      `INSERT INTO event_alliances
         (guild_id, game_profile, alliance_name, is_default, created_by_bot_instance)
       VALUES
         ('alliance-guild', 'wos', 'YOU', true, 'rachie-wos'),
         ('alliance-guild', 'wos', 'Academy', false, 'rachie-wos')`
    )
    await wosScheduler.upsertStateLink({
      allianceGuildId: "alliance-guild", configuredByBotInstance: "rachie-wos",
      stateGuildId: "state-guild", stateEventChannelId: "state-channel", sharingEnabled: true
    })

    const created = await wos.createStateEvent({
      stateGuildId: "state-guild", createdByUserId: "user", createdByBotInstance: "rachie-wos",
      event: {
        eventName: "SvS", firstOccurrenceDate: "2026-08-10", recurrenceDays: 2,
        phases: [
          { phaseName: "Borders", phaseTimeUtc: "10:00", preAlertMinutes: 30,
            announceExact: true, sortOrder: 0 },
          { phaseName: "Battle", phaseTimeUtc: "11:00", preAlertMinutes: 5,
            announceExact: false, sortOrder: 1 }
        ]
      }
    })
    const ksCreated = await kingshot.createStateEvent({
      stateGuildId: "state-guild", createdByUserId: "user", createdByBotInstance: "peggie-kingshot",
      event: {
        eventName: "Kingshot Event", firstOccurrenceDate: "2026-08-10", recurrenceDays: 42,
        phases: [{ phaseName: "Open", phaseTimeUtc: "12:00", announceExact: true }]
      }
    })
    assert.equal(await wos.getStateEvent("state-guild", ksCreated.id), null)
    assert.equal(await kingshot.getStateEvent("state-guild", created.id), null)

    const targets = await wos.listTargetsForStateGuild("state-guild")
    assert.deepEqual(targets.map(target => target.target_guild_id).sort(), [
      "alliance-guild", "state-guild"
    ])
    const definitions = await wos.listActiveStateEventDefinitions({
      rangeEnd: new Date("2026-08-10T12:00:00Z")
    })
    const claims = await buildStateEventDeliveryClaims(definitions, {
      repository: wos, gameProfile: "wos",
      windowStart: new Date("2026-08-10T09:29:00Z"),
      windowEnd: new Date("2026-08-10T11:01:00Z")
    })
    assert.equal(new Set(claims.map(claim => `${claim.phaseId}:${claim.deliveryKind}:${claim.targetGuildId}`)).size,
      claims.length)
    assert.equal(await wos.insertMissingDeliveryClaims([...claims, ...claims]), claims.length)
    const perGuild = await pool.query(
      `SELECT target_guild_id, phase_id, delivery_kind, COUNT(*)::integer AS count
         FROM state_event_delivery_claims GROUP BY target_guild_id, phase_id, delivery_kind`
    )
    assert.ok(perGuild.rows.every(row => row.count === 1))

    const sentClaim = (await pool.query(
      `UPDATE state_event_delivery_claims SET status = 'sent', sent_at = now()
        WHERE id = (SELECT id FROM state_event_delivery_claims ORDER BY id LIMIT 1)
        RETURNING id`
    )).rows[0]
    const loaded = await wos.getStateEvent("state-guild", created.id)
    const originalPhaseIds = loaded.phases.map(phase => String(phase.id))
    await wos.updateStateEvent({
      stateGuildId: "state-guild", eventId: created.id,
      event: {
        eventName: "SvS Updated", firstOccurrenceDate: "2026-08-10", recurrenceDays: 3,
        phases: loaded.phases.map((phase, index) => ({
          ...phaseDraftForUpdate(phase),
          phaseName: index === 0 ? "Borders Updated" : phase.phase_name,
          phaseTimeUtc: index === 0 ? "10:30" : phase.phase_time_utc
        }))
      }
    })
    const updated = await wos.getStateEvent("state-guild", created.id)
    assert.deepEqual(updated.phases.map(phase => String(phase.id)).sort(), originalPhaseIds.sort())
    assert.equal((await pool.query(
      "SELECT COUNT(*)::integer AS count FROM state_event_delivery_claims WHERE id = $1 AND status = 'sent'",
      [sentClaim.id]
    )).rows[0].count, 1)
    assert.equal((await pool.query(
      "SELECT COUNT(*)::integer AS count FROM state_event_delivery_claims WHERE state_event_id = $1 AND status <> 'sent' AND next_attempt_at IS NULL",
      [created.id]
    )).rows[0].count, claims.length - 1)

    await wos.setStateEventStatus({ stateGuildId: "state-guild", eventId: created.id, status: "paused" })
    assert.deepEqual(await wos.stateRoundupOccurrences({
      stateGuildId: "state-guild",
      weekStart: new Date("2026-08-10T00:00:00Z"),
      weekEnd: new Date("2026-08-17T00:00:00Z")
    }), [])
    await wos.setStateEventStatus({ stateGuildId: "state-guild", eventId: created.id, status: "active" })
    const roundup = await wos.stateRoundupOccurrences({
      stateGuildId: "state-guild",
      weekStart: new Date("2026-08-10T00:00:00Z"),
      weekEnd: new Date("2026-08-17T00:00:00Z")
    })
    assert.ok(roundup.some(item => item.phaseName === "Battle" && item.announceExact === false))
    assert.ok(roundup.every(item => item.deliveryKind === undefined))
  } finally {
    await pool?.end().catch(() => {})
    await dropSchema(admin, schema)
    await admin.end()
  }
})

function phaseDraftForUpdate(phase) {
  return {
    id: phase.id,
    phaseName: phase.phase_name,
    phaseTimeUtc: phase.phase_time_utc,
    preAlertMinutes: phase.pre_alert_minutes,
    preAlertMessage: phase.pre_alert_message,
    announceExact: phase.announce_exact,
    exactMessage: phase.exact_message,
    preAlertMedia: phase.pre_alert_media,
    exactMedia: phase.exact_media,
    sortOrder: phase.sort_order
  }
}
