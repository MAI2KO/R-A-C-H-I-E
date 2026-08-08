const test = require("node:test")
const assert = require("node:assert/strict")
const fs = require("node:fs/promises")
const path = require("node:path")
const { Pool } = require("pg")

const { createEventSchedulerRepository } = require("../src/eventSchedulerRepository")
const {
  IDS,
  handleEventSchedulerInteraction
} = require("../src/eventSchedulerInteractions")

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

test("migration 008 adds replay protection and backfills prior state-roundup behavior", {
  skip: databaseUrl ? false : "TEST_DATABASE_URL is not configured"
}, async () => {
  const pool = new Pool({ connectionString: databaseUrl, max: 1 })
  const client = await pool.connect()
  const schema = `roundup_ux_upgrade_${process.pid}`
  try {
    await client.query("SELECT pg_advisory_lock(7000010)")
    await client.query(`CREATE SCHEMA "${schema}"`)
    await client.query(`SET search_path TO "${schema}"`)
    for (const fileName of [
      "001_event_scheduler.sql",
      "002_event_images.sql",
      "003_event_delivery_indexes.sql",
      "004_state_delivery_reconciliation.sql",
      "005_event_management_and_roundups.sql",
      "006_flexible_reminders_and_alliances.sql",
      "007_native_channel_and_state_link_setup.sql"
    ]) {
      await client.query(await fs.readFile(path.join(migrationsDirectory, fileName), "utf8"))
    }
    await client.query(
      `INSERT INTO event_guild_settings
         (guild_id, game_profile, bot_instance_name, alliance_name,
          weekly_roundup_enabled, weekly_roundup_day, weekly_roundup_time_utc,
          weekly_roundup_channel_id)
       VALUES
         ('enabled', 'wos', 'rachie-wos', 'YOU', true, 0, '18:00', 'channel-a'),
         ('disabled', 'kingshot', 'peggie-kingshot', 'YOU', false, 3, '12:30', 'channel-b')`
    )

    await client.query(await fs.readFile(
      path.join(migrationsDirectory, "008_roundup_schedule_controls.sql"),
      "utf8"
    ))
    const settings = (await client.query(
      `SELECT guild_id, weekly_roundup_enabled, state_roundup_enabled,
              weekly_roundup_day, weekly_roundup_time_utc::text,
              weekly_roundup_channel_id, weekly_roundup_not_before::text
         FROM event_guild_settings ORDER BY guild_id`
    )).rows
    assert.deepEqual(settings, [
      {
        guild_id: "disabled",
        weekly_roundup_enabled: false,
        state_roundup_enabled: false,
        weekly_roundup_day: 3,
        weekly_roundup_time_utc: "12:30:00",
        weekly_roundup_channel_id: "channel-b",
        weekly_roundup_not_before: "-infinity"
      },
      {
        guild_id: "enabled",
        weekly_roundup_enabled: true,
        state_roundup_enabled: true,
        weekly_roundup_day: 0,
        weekly_roundup_time_utc: "18:00:00",
        weekly_roundup_channel_id: "channel-a",
        weekly_roundup_not_before: "-infinity"
      }
    ])
  } finally {
    await client.query("SET search_path TO public").catch(() => {})
    await client.query(`DROP SCHEMA IF EXISTS "${schema}" CASCADE`).catch(() => {})
    await client.query("SELECT pg_advisory_unlock(7000010)").catch(() => {})
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

test("roundup configuration changes preserve settings and sent history", {
  skip: databaseUrl ? false : "TEST_DATABASE_URL is not configured"
}, async () => {
  const pool = new Pool({ connectionString: databaseUrl, max: 4 })
  const guildId = "6888888888888811"
  const channelId = "6888888888888812"
  const repository = createEventSchedulerRepository(pool, "wos")
  try {
    await pool.query("DELETE FROM weekly_roundup_claims WHERE source_guild_id = $1", [guildId])
    await pool.query(
      "DELETE FROM event_guild_settings WHERE guild_id = $1 AND game_profile = 'wos'",
      [guildId]
    )
    await repository.upsertGuildSettings({
      guildId,
      botInstanceName: "rachie-wos",
      allianceName: "Roundup Test",
      eventChannelId: "6888888888888813"
    })
    await repository.configureWeeklyRoundup({
      guildId,
      enabled: true,
      stateEnabled: true,
      weekday: 0,
      timeUtc: "18:00",
      channelId,
      postWhenEmpty: false
    })
    await pool.query(
      `INSERT INTO weekly_roundup_claims (
         week_start_date, game_profile, target_kind, target_guild_id,
         target_channel_id, source_guild_id, scheduled_for, status,
         sent_at, sent_message_id
       ) VALUES
         ('2028-02-27', 'wos', 'alliance', $1, $2, $1,
          '2028-02-27T18:00Z', 'sent', '2028-02-27T18:00Z', 'sent-roundup'),
         ('2028-03-05', 'wos', 'alliance', $1, $2, $1,
          '2028-03-05T18:00Z', 'pending', NULL, NULL)`,
      [guildId, channelId]
    )

    await repository.configureWeeklyRoundup({
      guildId,
      enabled: false,
      stateEnabled: true,
      weekday: 0,
      timeUtc: "18:00",
      channelId,
      postWhenEmpty: false
    })
    let settings = await repository.getGuildSettings(guildId)
    assert.equal(settings.weekly_roundup_enabled, false)
    assert.equal(settings.state_roundup_enabled, true)
    assert.equal(settings.weekly_roundup_day, 0)
    assert.equal(String(settings.weekly_roundup_time_utc).slice(0, 5), "18:00")
    assert.equal(settings.weekly_roundup_channel_id, channelId)

    const claims = (await pool.query(
      `SELECT status, sent_message_id, last_error
         FROM weekly_roundup_claims
        WHERE source_guild_id = $1 ORDER BY week_start_date`,
      [guildId]
    )).rows
    assert.deepEqual(claims[0], {
      status: "sent",
      sent_message_id: "sent-roundup",
      last_error: null
    })
    assert.equal(claims[1].status, "failed")
    assert.match(claims[1].last_error, /configuration changed/i)

    await repository.configureWeeklyRoundup({
      guildId,
      enabled: true,
      stateEnabled: false,
      weekday: settings.weekly_roundup_day,
      timeUtc: settings.weekly_roundup_time_utc,
      channelId: settings.weekly_roundup_channel_id,
      postWhenEmpty: settings.roundup_when_empty
    })
    settings = await repository.getGuildSettings(guildId)
    assert.equal(settings.weekly_roundup_enabled, true)
    assert.equal(settings.state_roundup_enabled, false)
    assert.equal(settings.weekly_roundup_day, 0)
    assert.equal(String(settings.weekly_roundup_time_utc).slice(0, 5), "18:00")
    assert.equal(settings.weekly_roundup_channel_id, channelId)
    assert.ok(settings.weekly_roundup_not_before instanceof Date)
  } finally {
    await pool.query("DELETE FROM weekly_roundup_claims WHERE source_guild_id = $1", [guildId])
      .catch(() => {})
    await pool.query(
      "DELETE FROM event_guild_settings WHERE guild_id = $1 AND game_profile = 'wos'",
      [guildId]
    ).catch(() => {})
    await pool.end()
  }
})

test("roundup previews leave production delivery history unchanged", {
  skip: databaseUrl ? false : "TEST_DATABASE_URL is not configured"
}, async () => {
  const pool = new Pool({ connectionString: databaseUrl, max: 4 })
  const guildId = "6888888888888901"
  const channelId = "6888888888888902"
  const repository = createEventSchedulerRepository(pool, "wos")
  try {
    await pool.query("DELETE FROM weekly_roundup_claims WHERE source_guild_id = $1", [guildId])
    await pool.query(
      "DELETE FROM event_guild_settings WHERE guild_id = $1 AND game_profile = 'wos'",
      [guildId]
    )
    await repository.upsertGuildSettings({
      guildId,
      botInstanceName: "rachie-wos",
      allianceName: "Preview Main",
      eventChannelId: "6888888888888903"
    })
    await repository.configureWeeklyRoundup({
      guildId,
      enabled: true,
      stateEnabled: false,
      weekday: 5,
      timeUtc: "15:20",
      channelId,
      postWhenEmpty: false
    })
    const alliance = (await repository.listAlliances(guildId)).alliances[0]
    await repository.createEvent({
      guildId,
      createdByUserId: "6888888888888904",
      createdByBotInstance: "rachie-wos",
      event: {
        allianceId: String(alliance.id),
        allianceName: alliance.alliance_name,
        eventName: "Preview Event",
        firstOccurrenceDate: "2028-03-03",
        eventTimeUtc: "18:00",
        groups: [],
        grouped: false,
        recurrenceDays: 7,
        advanceReminderMinutes: 30,
        advanceReminderMessage: null,
        reminderAtStart: true,
        finalReminderMessage: null,
        publishToAlliance: true,
        publishToState: false,
        includeInWeeklyRoundup: true,
        image: null
      }
    })
    await pool.query(
      `INSERT INTO weekly_roundup_claims (
         week_start_date, game_profile, target_kind, target_guild_id,
         target_channel_id, source_guild_id, scheduled_for, status,
         attempt_count, sent_at, sent_message_id
       ) VALUES
         ('2028-02-25', 'wos', 'alliance', $1, $2, $1,
          '2028-02-25T15:20Z', 'sent', 1, '2028-02-25T15:20Z', 'production-roundup'),
         ('2028-03-03', 'wos', 'alliance', $1, $2, $1,
          '2028-03-03T15:20Z', 'pending', 0, NULL, NULL)`,
      [guildId, channelId]
    )

    const before = (await pool.query(
      `SELECT status, attempt_count, sent_at, sent_message_id
         FROM weekly_roundup_claims WHERE source_guild_id = $1
         ORDER BY week_start_date`,
      [guildId]
    )).rows
    const first = await repository.getWeeklyRoundupPreview(guildId, {
      now: new Date("2028-03-03T16:00:00Z")
    })
    const second = await repository.getWeeklyRoundupPreview(guildId, {
      now: new Date("2028-03-03T16:05:00Z")
    })
    const sentTests = []
    for (let index = 0; index < 2; index += 1) {
      const interaction = {
        commandName: null,
        customId: IDS.roundupAllianceTestConfirm,
        guildId,
        user: { id: "6888888888888904" },
        client: {},
        deferred: false,
        replied: false,
        isChatInputCommand: () => false,
        isButton: () => true,
        isStringSelectMenu: () => false,
        isChannelSelectMenu: () => false,
        isModalSubmit: () => false,
        async deferUpdate() { this.deferred = true },
        async editReply(payload) { this.edited = payload }
      }
      await handleEventSchedulerInteraction(interaction, {
        healthProvider: () => ({
          available: true,
          gameProfile: "wos",
          botInstanceName: "rachie-wos"
        }),
        userCanManageServer: async () => true,
        repositoryProvider: () => repository,
        roundupNow: () => new Date("2028-03-03T16:00:00Z"),
        roundupTargetResolver: async () => ({
          channel: {
            async send(message) {
              sentTests.push(message)
              return { id: `test-${sentTests.length}` }
            }
          }
        }),
        logger: { error() {}, warn() {} }
      })
    }
    const after = (await pool.query(
      `SELECT status, attempt_count, sent_at, sent_message_id
         FROM weekly_roundup_claims WHERE source_guild_id = $1
         ORDER BY week_start_date`,
      [guildId]
    )).rows
    assert.equal(first.occurrences.length, 1)
    assert.equal(second.occurrences.length, 1)
    assert.equal(sentTests.length, 2)
    assert.ok(sentTests.every(message => /— TEST/.test(message.embeds[0].toJSON().title)))
    assert.deepEqual(after, before)
  } finally {
    await pool.query("DELETE FROM weekly_roundup_claims WHERE source_guild_id = $1", [guildId])
      .catch(() => {})
    await pool.query(
      "DELETE FROM event_guild_settings WHERE guild_id = $1 AND game_profile = 'wos'",
      [guildId]
    ).catch(() => {})
    await pool.end()
  }
})

test("ordinary edits and alliance changes retain image and event configuration", {
  skip: databaseUrl ? false : "TEST_DATABASE_URL is not configured"
}, async () => {
  const pool = new Pool({ connectionString: databaseUrl, max: 4 })
  const guildId = "6888888888888821"
  const repository = createEventSchedulerRepository(pool, "wos")
  try {
    await pool.query(
      "DELETE FROM event_guild_settings WHERE guild_id = $1 AND game_profile = 'wos'",
      [guildId]
    )
    await repository.upsertGuildSettings({
      guildId,
      botInstanceName: "rachie-wos",
      allianceName: "Main",
      eventChannelId: "6888888888888822"
    })
    const main = (await repository.listAlliances(guildId)).alliances[0]
    const second = await repository.createAlliance({
      guildId,
      allianceName: "Second",
      createdByBotInstance: "rachie-wos"
    })
    const image = {
      originalFilename: "event.png",
      contentType: "image/png",
      byteSize: 8,
      imageData: Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
    }
    const event = {
      allianceId: String(main.id),
      allianceName: main.alliance_name,
      eventName: "Retained Event",
      firstOccurrenceDate: "2028-02-28",
      eventTimeUtc: "18:00",
      groups: [],
      grouped: false,
      recurrenceDays: 14,
      advanceReminderMinutes: 30,
      advanceReminderMessage: "Prepare",
      reminderAtStart: true,
      finalReminderMessage: "Final call",
      publishToAlliance: true,
      publishToState: true,
      includeInWeeklyRoundup: true,
      image
    }
    const created = await repository.createEvent({
      guildId,
      createdByUserId: "6888888888888823",
      createdByBotInstance: "rachie-wos",
      event
    })

    const ordinary = await repository.updateEvent({
      guildId,
      eventId: created.id,
      event: { ...event, eventName: "Ordinarily Edited" },
      imageAction: "retain"
    })
    assert.equal(String(ordinary.alliance_id), String(main.id))
    assert.equal((await repository.getEvent(guildId, created.id)).image_filename, "event.png")

    await pool.query(
      `INSERT INTO event_delivery_claims (
         event_id, game_profile, schedule_version, occurrence_at, deliver_at,
         delivery_kind, target_kind, target_guild_id, target_channel_id
       ) VALUES ($1, 'wos', 2, '2028-03-13T18:00Z', '2028-03-13T17:30Z',
                 'advance_reminder', 'alliance', $2, '6888888888888822')`,
      [created.id, guildId]
    )
    const changed = await repository.updateEvent({
      guildId,
      eventId: created.id,
      event: {
        ...event,
        eventName: "Ordinarily Edited",
        allianceId: String(second.id),
        allianceName: second.alliance_name
      },
      imageAction: "retain"
    })
    assert.equal(changed.schedule_version, 3)
    assert.equal(String(changed.alliance_id), String(second.id))
    assert.equal((await repository.getEvent(guildId, created.id)).image_filename, "event.png")
    assert.equal(changed.recurrence_days, 14)
    assert.equal(changed.advance_reminder_minutes, 30)
    assert.equal(changed.advance_reminder_message, "Prepare")
    assert.equal(changed.reminder_at_start, true)
    assert.equal(changed.final_reminder_message, "Final call")
    assert.equal(changed.publish_to_alliance, true)
    assert.equal(changed.publish_to_state, true)
    assert.equal(changed.include_in_weekly_roundup, true)
    const unsent = (await pool.query(
      "SELECT status, last_error FROM event_delivery_claims WHERE event_id = $1",
      [created.id]
    )).rows[0]
    assert.equal(unsent.status, "failed")
    assert.match(unsent.last_error, /schedule changed/i)
  } finally {
    await pool.query(
      "DELETE FROM event_guild_settings WHERE guild_id = $1 AND game_profile = 'wos'",
      [guildId]
    ).catch(() => {})
    await pool.end()
  }
})
