const test = require("node:test")
const assert = require("node:assert/strict")
const fs = require("node:fs/promises")
const path = require("node:path")
const { Pool } = require("pg")

const { createEventSchedulerRepository } = require("../src/eventSchedulerRepository")
const { createEventDeliveryRepository } = require("../src/eventDeliveryRepository")
const { buildDeliveryClaims } = require("../src/eventDeliveryGeneration")
const { createWeeklyRoundupRepository } = require("../src/weeklyRoundupRepository")
const { claimsForConfigurations } = require("../src/weeklyRoundupGeneration")
const { formatWeeklyRoundup } = require("../src/weeklyRoundupFormatting")

const databaseUrl = process.env.TEST_DATABASE_URL
const migrationsDirectory = path.join(__dirname, "..", "migrations")

async function applyMigration(client, fileName) {
  const sql = await fs.readFile(path.join(migrationsDirectory, fileName), "utf8")
  await client.query(sql)
}

test("migration 006 backfills alliance ownership and preserves old-writer compatibility", {
  skip: databaseUrl ? false : "TEST_DATABASE_URL is not configured"
}, async () => {
  const pool = new Pool({ connectionString: databaseUrl, max: 1 })
  const client = await pool.connect()
  const schema = `phase9_upgrade_${process.pid}`

  try {
    await client.query("SELECT pg_advisory_lock(7000007)")
    await client.query(`DROP SCHEMA IF EXISTS "${schema}" CASCADE`)
    await client.query(`CREATE SCHEMA "${schema}"`)
    await client.query(`SET search_path TO "${schema}"`)
    for (const fileName of [
      "001_event_scheduler.sql",
      "002_event_images.sql",
      "003_event_delivery_indexes.sql",
      "004_state_delivery_reconciliation.sql",
      "005_event_management_and_roundups.sql"
    ]) {
      await applyMigration(client, fileName)
    }

    await client.query(
      `INSERT INTO event_guild_settings
         (guild_id, game_profile, bot_instance_name, alliance_name, event_channel_id)
       VALUES
         ('shared-guild', 'wos', 'rachie-wos', 'YOU', 'wos-channel'),
         ('shared-guild', 'kingshot', 'peggie-kingshot', 'YOU', 'kingshot-channel')`
    )
    await client.query(
      `INSERT INTO scheduled_events (
         guild_id, game_profile, created_by_bot_instance, alliance_name, event_name,
         first_occurrence_date, event_time_utc, recurrence_days,
         advance_reminder_minutes, reminder_at_start, created_by_user_id
       ) VALUES
         ('shared-guild', 'wos', 'rachie-wos', 'YOU', 'Main Event',
          '2028-02-28', '10:00', 7, 10, true, 'user'),
         ('shared-guild', 'wos', 'rachie-wos', 'YOU2', 'Sub Event',
          '2028-02-28', '11:00', 7, 10, true, 'user')`
    )

    await applyMigration(client, "006_flexible_reminders_and_alliances.sql")

    const alliances = (await client.query(
      `SELECT guild_id, game_profile, alliance_name, is_default
         FROM event_alliances
        ORDER BY game_profile, is_default DESC, alliance_name`
    )).rows
    assert.deepEqual(alliances, [
      { guild_id: "shared-guild", game_profile: "kingshot", alliance_name: "YOU", is_default: true },
      { guild_id: "shared-guild", game_profile: "wos", alliance_name: "YOU", is_default: true },
      { guild_id: "shared-guild", game_profile: "wos", alliance_name: "YOU2", is_default: false }
    ])
    const events = (await client.query(
      `SELECT e.event_name, e.alliance_id, a.alliance_name
         FROM scheduled_events e
         JOIN event_alliances a
           ON a.id = e.alliance_id AND a.guild_id = e.guild_id
          AND a.game_profile = e.game_profile
        ORDER BY e.event_name`
    )).rows
    assert.equal(events[0].alliance_name, "YOU")
    assert.equal(events[1].alliance_name, "YOU2")
    assert.ok(events.every(event => event.alliance_id !== null))

    const compatibleInsert = await client.query(
      `INSERT INTO scheduled_events (
         guild_id, game_profile, created_by_bot_instance, alliance_name, event_name,
         first_occurrence_date, event_time_utc, recurrence_days,
         advance_reminder_minutes, reminder_at_start, created_by_user_id
       ) VALUES (
         'shared-guild', 'wos', 'rachie-wos', 'ignored-old-name', 'Old Writer',
         '2028-02-28', '12:00', 7, 5, true, 'user'
       ) RETURNING alliance_id, alliance_name`
    )
    assert.equal(compatibleInsert.rows[0].alliance_name, "YOU")
    assert.ok(compatibleInsert.rows[0].alliance_id)

    await assert.rejects(client.query(
      `INSERT INTO event_alliances
         (guild_id, game_profile, alliance_name, is_default, created_by_bot_instance)
       VALUES ('shared-guild', 'wos', 'you2', false, 'rachie-wos')`
    ), error => error.code === "23505" && error.constraint === "event_alliances_name_unique_ci")
    await assert.rejects(client.query(
      `UPDATE scheduled_events SET advance_reminder_minutes = 25
        WHERE event_name = 'Old Writer'`
    ), error => error.code === "23514")
    await assert.rejects(client.query(
      `UPDATE scheduled_events SET advance_reminder_message = ' not trimmed '
        WHERE event_name = 'Old Writer'`
    ), error => error.code === "23514")

    const comment = (await client.query(
      `SELECT col_description('scheduled_events'::regclass,
                 (SELECT attnum FROM pg_attribute
                   WHERE attrelid = 'scheduled_events'::regclass
                     AND attname = 'reminder_at_start')) AS value`
    )).rows[0].value
    assert.match(comment, /Legacy column name/)
    assert.match(comment, /one minute before/)
  } finally {
    await client.query("SET search_path TO public").catch(() => {})
    await client.query(`DROP SCHEMA IF EXISTS "${schema}" CASCADE`).catch(() => {})
    await client.query("SELECT pg_advisory_unlock(7000007)").catch(() => {})
    client.release()
    await pool.end()
  }
})

function eventDraft(alliance, name, time, overrides = {}) {
  return {
    allianceId: String(alliance.id),
    allianceName: alliance.alliance_name,
    eventName: name,
    firstOccurrenceDate: "2028-02-28",
    eventTimeUtc: time,
    groups: [],
    grouped: false,
    recurrenceDays: 7,
    advanceReminderMinutes: 15,
    advanceReminderMessage: "Prepare now",
    reminderAtStart: true,
    finalReminderMessage: "Final call",
    publishToAlliance: true,
    publishToState: true,
    includeInWeeklyRoundup: true,
    image: null,
    ...overrides
  }
}

test("one guild safely manages multiple profile-scoped alliances, events and roundups", {
  skip: databaseUrl ? false : "TEST_DATABASE_URL is not configured"
}, async () => {
  const pool = new Pool({ connectionString: databaseUrl, max: 10 })
  const lockClient = await pool.connect()
  const guildId = "777777777777777701"
  const allianceChannel = "777777777777777702"
  const roundupChannel = "777777777777777703"
  const stateGuildId = "777777777777777704"
  const stateChannel = "777777777777777705"
  const userId = "777777777777777706"
  const now = new Date("2028-02-28T09:10:00Z")

  try {
    await lockClient.query("SELECT pg_advisory_lock(7000008)")
    await pool.query(
      "DELETE FROM weekly_roundup_claims WHERE source_guild_id = $1",
      [guildId]
    )
    await pool.query(
      "DELETE FROM event_guild_settings WHERE guild_id = $1 AND game_profile IN ('wos', 'kingshot')",
      [guildId]
    )

    const wos = createEventSchedulerRepository(pool, "wos")
    const kingshot = createEventSchedulerRepository(pool, "kingshot")
    await wos.upsertGuildSettings({
      guildId,
      botInstanceName: "rachie-wos",
      allianceName: "YOU",
      eventChannelId: allianceChannel
    })
    await kingshot.upsertGuildSettings({
      guildId,
      botInstanceName: "peggie-kingshot",
      allianceName: "YOU",
      eventChannelId: allianceChannel
    })
    await wos.configureWeeklyRoundup({
      guildId,
      enabled: true,
      weekday: 1,
      timeUtc: "09:00",
      channelId: roundupChannel,
      postWhenEmpty: false,
      stateEnabled: true
    })
    await wos.upsertStateLink({
      allianceGuildId: guildId,
      configuredByBotInstance: "rachie-wos",
      stateGuildId,
      stateEventChannelId: stateChannel,
      sharingEnabled: true
    })
    assert.equal(await kingshot.getStateLink(guildId), null)
    await kingshot.upsertStateLink({
      allianceGuildId: guildId,
      configuredByBotInstance: "peggie-kingshot",
      stateGuildId,
      stateEventChannelId: "777777777777777715",
      sharingEnabled: true
    })
    assert.equal((await wos.getStateLink(guildId)).state_event_channel_id, stateChannel)
    assert.equal(
      (await kingshot.getStateLink(guildId)).state_event_channel_id,
      "777777777777777715"
    )

    const wosDefault = (await wos.listAlliances(guildId)).alliances[0]
    const kingshotDefault = (await kingshot.listAlliances(guildId)).alliances[0]
    assert.equal(wosDefault.alliance_name, "YOU")
    assert.equal(kingshotDefault.alliance_name, "YOU")
    const you2 = await wos.createAlliance({
      guildId,
      allianceName: "YOU2",
      createdByBotInstance: "rachie-wos"
    })
    const academy = await wos.createAlliance({
      guildId,
      allianceName: "YOU Academy",
      createdByBotInstance: "rachie-wos"
    })
    const kingshotYou2 = await kingshot.createAlliance({
      guildId,
      allianceName: "YOU2",
      createdByBotInstance: "peggie-kingshot"
    })
    await assert.rejects(wos.createAlliance({
      guildId,
      allianceName: "you2",
      createdByBotInstance: "rachie-wos"
    }), error => error.code === "23505")
    assert.equal(await wos.getAlliance(guildId, kingshotYou2.id), null)
    assert.equal(await kingshot.getAlliance(guildId, you2.id), null)

    const mainEvent = await wos.createEvent({
      guildId,
      createdByUserId: userId,
      createdByBotInstance: "rachie-wos",
      event: eventDraft(wosDefault, "Main Foundry", "10:00")
    })
    const subDraft = eventDraft(you2, "Sub Bear", "11:00")
    const subEvent = await wos.createEvent({
      guildId,
      createdByUserId: userId,
      createdByBotInstance: "rachie-wos",
      event: subDraft
    })
    await wos.createEvent({
      guildId,
      createdByUserId: userId,
      createdByBotInstance: "rachie-wos",
      event: eventDraft(academy, "Academy Event", "12:00", { publishToState: false })
    })
    const groupedDraft = eventDraft(academy, "Grouped Drill", null, {
      groups: [{ groupName: "Alpha", eventTimeUtc: "14:00", sortOrder: 0 }],
      grouped: true,
      publishToState: false,
      includeInWeeklyRoundup: false
    })
    const groupedEvent = await wos.createEvent({
      guildId,
      createdByUserId: userId,
      createdByBotInstance: "rachie-wos",
      event: groupedDraft
    })
    await kingshot.createEvent({
      guildId,
      createdByUserId: userId,
      createdByBotInstance: "peggie-kingshot",
      event: eventDraft(kingshotDefault, "Kingshot Event", "13:00")
    })

    const wosEvents = await wos.listEvents(guildId, { limit: 10 })
    assert.deepEqual(
      new Set(wosEvents.events.map(event => event.alliance_name)),
      new Set(["YOU", "YOU2", "YOU Academy"])
    )
    assert.equal((await kingshot.listEvents(guildId, { limit: 10 })).events.length, 1)

    await pool.query(
      `INSERT INTO event_delivery_claims (
         event_id, game_profile, schedule_version, occurrence_at, deliver_at,
         delivery_kind, target_kind, target_guild_id, target_channel_id,
         status, sent_at, sent_message_id
       ) VALUES
         ($1, 'wos', 1, '2028-02-28T11:00Z', '2028-02-28T10:45Z',
          'advance_reminder', 'alliance', $2, $3, 'sent', '2028-02-28T10:45Z', 'sent-history'),
         ($1, 'wos', 1, '2028-03-06T11:00Z', '2028-03-06T10:45Z',
          'advance_reminder', 'alliance', $2, $3, 'pending', NULL, NULL)`,
      [subEvent.id, guildId, allianceChannel]
    )
    const edited = await wos.updateEvent({
      guildId,
      eventId: subEvent.id,
      event: { ...subDraft, advanceReminderMessage: "Changed preparation" },
      imageAction: "retain"
    })
    assert.equal(edited.schedule_version, 2)
    const claimHistory = (await pool.query(
      `SELECT status, sent_message_id, last_error
         FROM event_delivery_claims WHERE event_id = $1 ORDER BY id`,
      [subEvent.id]
    )).rows
    assert.deepEqual(claimHistory[0], {
      status: "sent",
      sent_message_id: "sent-history",
      last_error: null
    })
    assert.equal(claimHistory[1].status, "failed")
    assert.match(claimHistory[1].last_error, /schedule changed/i)

    const originalGroup = (await pool.query(
      `SELECT id FROM scheduled_event_groups
        WHERE event_id = $1 AND game_profile = 'wos'`,
      [groupedEvent.id]
    )).rows[0]
    const groupedClaims = await pool.query(
      `INSERT INTO event_delivery_claims (
         event_id, group_id, game_profile, schedule_version, occurrence_at, deliver_at,
         delivery_kind, target_kind, target_guild_id, target_channel_id,
         status, sent_at, sent_message_id
       ) VALUES
         ($1, $2, 'wos', 1, '2028-02-28T14:00Z', '2028-02-28T13:45Z',
          'advance_reminder', 'alliance', $3, $4, 'sent',
          '2028-02-28T13:45Z', 'group-sent-history'),
         ($1, $2, 'wos', 1, '2028-03-06T14:00Z', '2028-03-06T13:45Z',
          'advance_reminder', 'alliance', $3, $4, 'pending', NULL, NULL)
       RETURNING id, status, sent_message_id, updated_at`,
      [groupedEvent.id, originalGroup.id, guildId, allianceChannel]
    )
    const editedGrouped = await wos.updateEvent({
      guildId,
      eventId: groupedEvent.id,
      event: {
        ...groupedDraft,
        groups: [{ groupName: "Alpha", eventTimeUtc: "14:30", sortOrder: 0 }]
      },
      imageAction: "retain"
    })
    assert.equal(editedGrouped.schedule_version, 2)
    const preservedGroupedClaims = (await pool.query(
      `SELECT id, group_id, group_id_snapshot, group_name_snapshot, status,
              sent_message_id, updated_at
         FROM event_delivery_claims
        WHERE event_id = $1 ORDER BY id`,
      [groupedEvent.id]
    )).rows
    assert.equal(String(preservedGroupedClaims[0].id), String(groupedClaims.rows[0].id))
    assert.equal(preservedGroupedClaims[0].group_id, null)
    assert.equal(String(preservedGroupedClaims[0].group_id_snapshot), String(originalGroup.id))
    assert.equal(preservedGroupedClaims[0].group_name_snapshot, "Alpha")
    assert.equal(preservedGroupedClaims[0].status, "sent")
    assert.equal(preservedGroupedClaims[0].sent_message_id, "group-sent-history")
    assert.equal(
      preservedGroupedClaims[0].updated_at.toISOString(),
      groupedClaims.rows[0].updated_at.toISOString()
    )
    assert.equal(preservedGroupedClaims[1].status, "failed")
    assert.equal(preservedGroupedClaims[1].group_id, null)
    const replacementGroup = (await pool.query(
      `SELECT id, event_time_utc::text AS event_time_utc
         FROM scheduled_event_groups
        WHERE event_id = $1 AND game_profile = 'wos'`,
      [groupedEvent.id]
    )).rows[0]
    assert.notEqual(String(replacementGroup.id), String(originalGroup.id))
    assert.equal(replacementGroup.event_time_utc, "14:30:00")

    const groupedDeliveryRepository = createEventDeliveryRepository(pool, "wos", {
      targetKind: "alliance"
    })
    const versionTwoDefinition = (await groupedDeliveryRepository.listActiveEventDefinitions({
      rangeEnd: new Date("2028-02-28T15:00:00Z")
    })).find(event => String(event.id) === String(groupedEvent.id))
    const versionTwoClaims = buildDeliveryClaims([versionTwoDefinition], {
      gameProfile: "wos",
      windowStart: new Date("2028-02-28T13:00:00Z"),
      windowEnd: new Date("2028-02-28T15:00:00Z")
    })
    assert.deepEqual(versionTwoClaims.map(claim => claim.deliveryKind), [
      "advance_reminder",
      "final_reminder"
    ])
    assert.equal(
      await groupedDeliveryRepository.insertMissingDeliveryClaims(versionTwoClaims),
      2
    )

    const cancelled = await wos.updateEvent({
      guildId,
      eventId: groupedEvent.id,
      event: {
        ...groupedDraft,
        groups: [{ groupName: "Alpha", eventTimeUtc: "14:30", sortOrder: 0 }],
        advanceReminderMinutes: null,
        advanceReminderMessage: null,
        reminderAtStart: false,
        finalReminderMessage: null
      },
      imageAction: "retain"
    })
    assert.equal(cancelled.schedule_version, 3)
    const cancelledStored = await wos.getEvent(guildId, groupedEvent.id)
    assert.equal(cancelledStored.advance_reminder_minutes, null)
    assert.equal(cancelledStored.advance_reminder_message, null)
    assert.equal(cancelledStored.reminder_at_start, false)
    assert.equal(cancelledStored.final_reminder_message, null)
    const cancelledClaims = (await pool.query(
      `SELECT schedule_version, delivery_kind, status, sent_message_id
         FROM event_delivery_claims
        WHERE event_id = $1 ORDER BY id`,
      [groupedEvent.id]
    )).rows
    assert.equal(cancelledClaims[0].status, "sent")
    assert.equal(cancelledClaims[0].sent_message_id, "group-sent-history")
    assert.ok(cancelledClaims.slice(1).every(claim => claim.status === "failed"))
    const cancelledDefinition = (await groupedDeliveryRepository.listActiveEventDefinitions({
      rangeEnd: new Date("2028-02-28T15:00:00Z")
    })).find(event => String(event.id) === String(groupedEvent.id))
    assert.deepEqual(buildDeliveryClaims([cancelledDefinition], {
      gameProfile: "wos",
      windowStart: new Date("2028-02-28T13:00:00Z"),
      windowEnd: new Date("2028-02-28T15:00:00Z")
    }), [])

    const renamed = await wos.renameAlliance({
      guildId,
      allianceId: you2.id,
      allianceName: "YOU Two"
    })
    assert.equal(renamed.alliance_name, "YOU Two")
    const currentSub = await wos.getEvent(guildId, subEvent.id)
    const currentMain = await wos.getEvent(guildId, mainEvent.id)
    assert.equal(currentSub.alliance_name, "YOU Two")
    assert.equal(currentSub.schedule_version, 3)
    assert.equal(currentMain.alliance_name, "YOU")
    assert.equal(currentMain.schedule_version, 1)
    assert.equal((await kingshot.getAlliance(guildId, kingshotYou2.id)).alliance_name, "YOU2")

    const deleteOwned = await wos.deleteAlliance({ guildId, allianceId: you2.id })
    assert.equal(deleteOwned.deleted, false)
    assert.equal(deleteOwned.reason, "events")
    const empty = await wos.createAlliance({
      guildId,
      allianceName: "Temporary",
      createdByBotInstance: "rachie-wos"
    })
    assert.equal((await wos.deleteAlliance({ guildId, allianceId: empty.id })).deleted, true)

    const deliveryRepository = createEventDeliveryRepository(pool, "wos", {
      targetKind: "alliance"
    })
    const definitions = await deliveryRepository.listActiveEventDefinitions({
      rangeEnd: new Date("2028-02-28T14:00:00Z")
    })
    const currentDefinition = definitions.find(event => String(event.id) === String(subEvent.id))
    assert.equal(currentDefinition.schedule_version, 3)
    assert.equal(currentDefinition.advance_reminder_message, "Changed preparation")
    const generated = buildDeliveryClaims([currentDefinition], {
      gameProfile: "wos",
      windowStart: new Date("2028-02-28T09:00:00Z"),
      windowEnd: new Date("2028-02-28T12:00:00Z")
    })
    assert.deepEqual(generated.map(claim => claim.deliveryKind), [
      "advance_reminder",
      "final_reminder"
    ])
    assert.ok(generated.every(claim => claim.scheduleVersion === 3 && claim.targetKind === "alliance"))

    const roundupRepository = createWeeklyRoundupRepository(pool, "wos")
    const configurations = await roundupRepository.listRoundupConfigurations()
    const roundupClaims = claimsForConfigurations(configurations, {
      gameProfile: "wos",
      now,
      graceMinutes: 60
    })
    assert.deepEqual(
      new Set(roundupClaims.map(claim => claim.targetKind)),
      new Set(["alliance", "state"])
    )
    assert.equal(await roundupRepository.insertMissingClaims(roundupClaims), 2)
    const claimed = await roundupRepository.claimDue({
      now,
      batchSize: 10,
      leaseSeconds: 60,
      botInstanceName: "rachie-wos",
      workerId: "phase9-roundup"
    })
    assert.equal(claimed.length, 2)
    const payloads = []
    for (const claim of claimed) {
      payloads.push(await roundupRepository.getClaimPayload({
        claimId: claim.id,
        botInstanceName: "rachie-wos",
        workerId: "phase9-roundup"
      }))
    }
    const alliancePayload = payloads.find(payload => payload.claim.targetKind === "alliance")
    const statePayload = payloads.find(payload => payload.claim.targetKind === "state")
    assert.deepEqual(
      new Set(alliancePayload.occurrences.map(item => item.allianceName)),
      new Set(["YOU", "YOU Two", "YOU Academy"])
    )
    assert.deepEqual(
      new Set(statePayload.occurrences.map(item => item.allianceName)),
      new Set(["YOU", "YOU Two"])
    )
    assert.equal(statePayload.occurrences.some(item => item.eventName === "Academy Event"), false)
    assert.equal(alliancePayload.occurrences.some(item => item.eventName === "Academy Event"), true)
    assert.equal(JSON.stringify(formatWeeklyRoundup(alliancePayload)).includes("YOU Two"), true)
    assert.equal(JSON.stringify(formatWeeklyRoundup(statePayload)).includes("YOU Two"), true)
    assert.equal(JSON.stringify(statePayload).includes("Kingshot Event"), false)
  } finally {
    await pool.query(
      "DELETE FROM weekly_roundup_claims WHERE source_guild_id = $1",
      [guildId]
    ).catch(() => {})
    await pool.query(
      "DELETE FROM event_guild_settings WHERE guild_id = $1 AND game_profile IN ('wos', 'kingshot')",
      [guildId]
    ).catch(() => {})
    await lockClient.query("SELECT pg_advisory_unlock(7000008)").catch(() => {})
    lockClient.release()
    await pool.end()
  }
})
