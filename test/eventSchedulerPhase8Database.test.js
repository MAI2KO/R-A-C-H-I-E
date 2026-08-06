const test = require("node:test")
const assert = require("node:assert/strict")
const { Pool } = require("pg")

const { createEventSchedulerRepository } = require("../src/eventSchedulerRepository")
const { createEventDeliveryRepository } = require("../src/eventDeliveryRepository")
const { createWeeklyRoundupRepository } = require("../src/weeklyRoundupRepository")
const { claimsForConfigurations } = require("../src/weeklyRoundupGeneration")

const databaseUrl = process.env.TEST_DATABASE_URL

function draft(name, allianceName = "North") {
  return {
    allianceId: null,
    allianceName,
    eventName: name,
    firstOccurrenceDate: "2028-02-28",
    eventTimeUtc: "09:00",
    groups: [],
    recurrenceDays: 3,
    advanceReminderMinutes: 10,
    reminderAtStart: true,
    publishToAlliance: true,
    publishToState: true,
    includeInWeeklyRoundup: true,
    image: null
  }
}

test("Phase 8 management and roundups remain atomic, durable and profile isolated", {
  skip: databaseUrl ? false : "TEST_DATABASE_URL is not configured"
}, async () => {
  const pool = new Pool({ connectionString: databaseUrl, max: 10 })
  const isolationClient = await pool.connect()
  const guildA = "888888888888888801"
  const guildB = "888888888888888802"
  const stateGuild = "888888888888888803"
  const channelA = "888888888888888804"
  const channelB = "888888888888888805"
  const stateChannel = "888888888888888806"
  const userId = "888888888888888807"
  const now = new Date("2028-02-28T09:10:00Z")

  try {
    await isolationClient.query("SELECT pg_advisory_lock(7000005)")
    await pool.query(
      "DELETE FROM weekly_roundup_claims WHERE source_guild_id = ANY($1::varchar[])",
      [[guildA, guildB]]
    )
    await pool.query(
      "DELETE FROM event_guild_settings WHERE guild_id = ANY($1::varchar[])",
      [[guildA, guildB]]
    )
    await pool.query(
      `INSERT INTO event_guild_settings (
         guild_id, game_profile, bot_instance_name, alliance_name, event_channel_id,
         weekly_roundup_enabled, weekly_roundup_day, weekly_roundup_time_utc,
         weekly_roundup_channel_id, roundup_when_empty
       ) VALUES
         ($1, 'wos', 'rachie-wos', 'North', $3, true, 1, '09:00', $3, false),
         ($2, 'wos', 'rachie-wos', 'South', $4, true, 1, '09:00', $4, false),
         ($1, 'kingshot', 'peggie-kingshot', 'Kingshot North', $3, true, 1, '09:00', $3, false)`,
      [guildA, guildB, channelA, channelB]
    )
    const alliances = await pool.query(
      `INSERT INTO event_alliances (
         guild_id, game_profile, alliance_name, is_default, created_by_bot_instance
       ) VALUES
         ($1, 'wos', 'North', true, 'rachie-wos'),
         ($2, 'wos', 'South', true, 'rachie-wos'),
         ($1, 'kingshot', 'Kingshot North', true, 'peggie-kingshot')
       RETURNING id, guild_id, game_profile, alliance_name`,
      [guildA, guildB]
    )
    const northAlliance = alliances.rows.find(row =>
      row.guild_id === guildA && row.game_profile === "wos"
    )
    const southAlliance = alliances.rows.find(row =>
      row.guild_id === guildB && row.game_profile === "wos"
    )
    const kingshotAlliance = alliances.rows.find(row => row.game_profile === "kingshot")
    const northDraft = name => ({
      ...draft(name),
      allianceId: String(northAlliance.id),
      allianceName: northAlliance.alliance_name
    })
    await pool.query(
      `INSERT INTO event_state_links (
         alliance_guild_id, game_profile, configured_by_bot_instance,
         state_guild_id, state_event_channel_id, sharing_enabled
       ) VALUES
         ($1, 'wos', 'rachie-wos', $3, $4, true),
         ($2, 'wos', 'rachie-wos', $3, $4, true),
         ($1, 'kingshot', 'peggie-kingshot', $3, $4, true)`,
      [guildA, guildB, stateGuild, stateChannel]
    )

    const wosA = createEventSchedulerRepository(pool, "wos")
    const wosB = createEventSchedulerRepository(pool, "wos")
    const kingshot = createEventSchedulerRepository(pool, "kingshot")
    const image = {
      originalFilename: "original.png",
      contentType: "image/png",
      byteSize: 8,
      imageData: Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
    }
    const managed = await wosA.createEvent({
      guildId: guildA,
      createdByUserId: userId,
      createdByBotInstance: "rachie-wos",
      event: { ...northDraft("Managed"), image }
    })
    const south = await wosB.createEvent({
      guildId: guildB,
      createdByUserId: userId,
      createdByBotInstance: "rachie-wos",
      event: {
        ...draft("South Event", "South"),
        allianceId: String(southAlliance.id)
      }
    })
    const allianceOnly = await wosA.createEvent({
      guildId: guildA,
      createdByUserId: userId,
      createdByBotInstance: "rachie-wos",
      event: { ...northDraft("Alliance Only"), publishToState: false }
    })
    await kingshot.createEvent({
      guildId: guildA,
      createdByUserId: userId,
      createdByBotInstance: "peggie-kingshot",
      event: {
        ...draft("Kingshot Event", "Kingshot North"),
        allianceId: String(kingshotAlliance.id)
      }
    })

    const deliveryRows = await pool.query(
      `INSERT INTO event_delivery_claims (
         event_id, game_profile, schedule_version, occurrence_at, deliver_at,
         delivery_kind, target_kind, target_guild_id, target_channel_id,
         status, claimed_by_bot_instance, claimed_by_worker, claimed_at,
         claimed_until, next_attempt_at, sent_at, sent_message_id
       ) VALUES
         ($1, 'wos', 1, '2028-02-28T09:00Z', '2028-02-28T08:50Z',
          'advance_reminder', 'alliance', $2, $3, 'sent', NULL, NULL, NULL, NULL, NULL,
          '2028-02-28T08:50Z', 'sent-history'),
         ($1, 'wos', 1, '2028-03-02T09:00Z', '2028-03-02T08:50Z',
          'advance_reminder', 'alliance', $2, $3, 'pending', NULL, NULL, NULL, NULL, NULL, NULL, NULL),
         ($1, 'wos', 1, '2028-03-05T09:00Z', '2028-03-05T08:50Z',
          'advance_reminder', 'alliance', $2, $3, 'failed', NULL, NULL, NULL, NULL,
          '2028-03-05T08:40Z', NULL, NULL),
         ($1, 'wos', 1, '2028-03-08T09:00Z', '2028-03-08T08:50Z',
          'advance_reminder', 'alliance', $2, $3, 'claimed', 'rachie-wos', 'active-worker',
          '2028-02-28T09:00Z', '2028-02-28T09:20Z', NULL, NULL, NULL)
       RETURNING id, status`,
      [managed.id, guildA, channelA]
    )
    const activeClaim = deliveryRows.rows.find(row => row.status === "claimed")

    const groupedEdit = {
      ...northDraft("Managed Edited"),
      eventTimeUtc: null,
      groups: [
        { groupName: "Alpha", eventTimeUtc: "10:00", sortOrder: 0 },
        { groupName: "Beta", eventTimeUtc: "11:00", sortOrder: 1 }
      ]
    }
    const edited = await wosA.updateEvent({
      guildId: guildA,
      eventId: managed.id,
      event: groupedEdit,
      imageAction: "retain"
    })
    assert.equal(edited.schedule_version, 2)
    const afterEdit = await wosA.getEvent(guildA, managed.id)
    assert.equal(afterEdit.groups.length, 2)
    assert.equal(afterEdit.image_filename, "original.png")
    const history = await pool.query(
      `SELECT status, sent_message_id, last_error FROM event_delivery_claims
        WHERE event_id = $1 ORDER BY id`,
      [managed.id]
    )
    assert.equal(history.rows[0].status, "sent")
    assert.equal(history.rows[0].sent_message_id, "sent-history")
    assert.ok(history.rows.slice(1).every(row => row.status === "failed" && row.last_error))
    const stalePayload = await createEventDeliveryRepository(pool, "wos").getClaimPayload({
      claimId: activeClaim.id,
      botInstanceName: "rachie-wos",
      workerId: "active-worker"
    })
    assert.equal(stalePayload, null)

    const newClaim = {
      eventId: managed.id,
      groupId: null,
      gameProfile: "wos",
      scheduleVersion: 2,
      occurrenceAt: new Date("2028-03-02T09:00:00Z"),
      deliverAt: new Date("2028-03-02T08:50:00Z"),
      deliveryKind: "advance_reminder",
      targetKind: "alliance",
      targetGuildId: guildA,
      targetChannelId: channelA
    }
    assert.equal(
      await createEventDeliveryRepository(pool, "wos").insertMissingDeliveryClaims([newClaim]),
      1
    )

    const replacement = { ...image, originalFilename: "replacement.png" }
    await wosA.updateEvent({
      guildId: guildA,
      eventId: managed.id,
      event: { ...northDraft("Ungrouped Again"), image: replacement },
      imageAction: "replace"
    })
    let current = await wosA.getEvent(guildA, managed.id)
    assert.equal(current.groups.length, 0)
    assert.equal(current.image_filename, "replacement.png")
    await assert.rejects(wosA.updateEvent({
      guildId: guildA,
      eventId: managed.id,
      event: { ...northDraft("Must Roll Back"), image: { ...replacement, byteSize: 7 } },
      imageAction: "replace"
    }))
    current = await wosA.getEvent(guildA, managed.id)
    assert.equal(current.event_name, "Ungrouped Again")
    assert.equal(current.image_filename, "replacement.png")
    await wosA.updateEvent({
      guildId: guildA,
      eventId: managed.id,
      event: northDraft("No Image"),
      imageAction: "remove"
    })
    current = await wosA.getEvent(guildA, managed.id)
    assert.equal(current.image_filename, null)
    assert.equal(await wosA.updateEvent({
      guildId: guildB,
      eventId: managed.id,
      event: { ...draft("Wrong Guild", "South"), allianceId: String(southAlliance.id) },
      imageAction: "retain"
    }), null)
    assert.equal(await kingshot.updateEvent({
      guildId: guildA,
      eventId: managed.id,
      event: {
        ...draft("Wrong Profile", "Kingshot North"),
        allianceId: String(kingshotAlliance.id)
      },
      imageAction: "retain"
    }), null)

    const anchor = current.first_occurrence_date
    assert.equal((await wosA.setEventStatus({ guildId: guildA, eventId: managed.id, status: "paused" })).status, "paused")
    assert.equal((await createEventDeliveryRepository(pool, "wos").listActiveEventDefinitions({
      rangeEnd: new Date("2028-03-06T00:00:00Z")
    })).some(row => String(row.id) === String(managed.id)), false)
    assert.equal((await wosA.setEventStatus({ guildId: guildA, eventId: managed.id, status: "active" })).status, "active")
    assert.equal((await wosA.getEvent(guildA, managed.id)).first_occurrence_date, anchor)

    const roundupRepository = createWeeklyRoundupRepository(pool, "wos")
    const configs = await roundupRepository.listRoundupConfigurations()
    const roundupClaims = claimsForConfigurations(configs, {
      gameProfile: "wos",
      now,
      graceMinutes: 60
    })
    assert.equal(roundupClaims.filter(claim => claim.targetKind === "alliance").length, 2)
    assert.equal(roundupClaims.filter(claim => claim.targetKind === "state").length, 1)
    assert.equal(await roundupRepository.insertMissingClaims(roundupClaims), 3)
    assert.equal(await roundupRepository.insertMissingClaims(roundupClaims), 0)

    const [workerA, workerB] = await Promise.all([
      roundupRepository.claimDue({
        now, batchSize: 10, leaseSeconds: 60,
        botInstanceName: "rachie-wos", workerId: "roundup-a"
      }),
      roundupRepository.claimDue({
        now, batchSize: 10, leaseSeconds: 60,
        botInstanceName: "rachie-wos", workerId: "roundup-b"
      })
    ])
    assert.equal(workerA.length + workerB.length, 3)
    const claimed = workerA.length ? workerA : workerB
    const owner = workerA.length ? "roundup-a" : "roundup-b"
    const alliancePayload = await roundupRepository.getClaimPayload({
      claimId: claimed.find(row =>
        row.target_kind === "alliance" && row.target_guild_id === guildA
      ).id,
      botInstanceName: "rachie-wos",
      workerId: owner
    })
    assert.ok(alliancePayload.occurrences.some(item =>
      String(item.eventId) === String(allianceOnly.id)
    ))
    assert.equal((await roundupRepository.claimDue({
      now, batchSize: 10, leaseSeconds: 60,
      botInstanceName: "rachie-wos", workerId: "lease-observer"
    })).length, 0)

    const allianceClaims = claimed.filter(row => row.target_kind === "alliance")
    await pool.query(
      "UPDATE weekly_roundup_claims SET claimed_until = $2 WHERE id = $1",
      [allianceClaims[0].id, new Date(now.getTime() - 1000)]
    )
    const reclaimed = await roundupRepository.claimDue({
      now, batchSize: 10, leaseSeconds: 60,
      botInstanceName: "rachie-wos", workerId: "lease-reclaimer"
    })
    assert.equal(reclaimed.length, 1)
    assert.equal(String(reclaimed[0].id), String(allianceClaims[0].id))
    assert.equal(reclaimed[0].attempt_count, 2)

    const retryAt = new Date(now.getTime() + 30000)
    assert.equal(await roundupRepository.markFailed({
      claimId: allianceClaims[1].id,
      botInstanceName: "rachie-wos",
      workerId: owner,
      failedAt: now,
      lastError: "temporary",
      nextAttemptAt: retryAt
    }), true)
    assert.equal((await roundupRepository.claimDue({
      now, batchSize: 10, leaseSeconds: 60,
      botInstanceName: "rachie-wos", workerId: "retry-too-soon"
    })).length, 0)
    const retried = await roundupRepository.claimDue({
      now: retryAt, batchSize: 10, leaseSeconds: 60,
      botInstanceName: "rachie-wos", workerId: "retry-worker"
    })
    assert.equal(retried.length, 1)
    assert.equal(String(retried[0].id), String(allianceClaims[1].id))

    await pool.query(
      `UPDATE weekly_roundup_claims
          SET attempt_count = 5, claimed_until = $2
        WHERE id = $1`,
      [reclaimed[0].id, new Date(now.getTime() - 1000)]
    )
    assert.equal((await roundupRepository.claimDue({
      now, batchSize: 10, leaseSeconds: 60,
      botInstanceName: "rachie-wos", workerId: "exhausted-worker"
    })).length, 0)
    const exhausted = await pool.query(
      "SELECT status, next_attempt_at FROM weekly_roundup_claims WHERE id = $1",
      [reclaimed[0].id]
    )
    assert.deepEqual(exhausted.rows[0], { status: "failed", next_attempt_at: null })

    const stateClaim = claimed.find(row => row.target_kind === "state")
    const statePayload = await roundupRepository.getClaimPayload({
      claimId: stateClaim.id,
      botInstanceName: "rachie-wos",
      workerId: owner
    })
    assert.equal(statePayload.claim.targetIsCurrent, true)
    assert.ok(statePayload.occurrences.some(item => item.allianceName === "North"))
    assert.ok(statePayload.occurrences.some(item => item.allianceName === "South"))
    assert.equal(statePayload.occurrences.some(item => item.allianceName === "Kingshot North"), false)
    assert.equal(statePayload.occurrences.some(item =>
      String(item.eventId) === String(allianceOnly.id)
    ), false)

    assert.equal(await roundupRepository.setPartCount({
      claimId: stateClaim.id, botInstanceName: "rachie-wos", workerId: owner, partCount: 2
    }), true)
    assert.equal(await roundupRepository.recordSentMessage({
      claimId: stateClaim.id, botInstanceName: "rachie-wos", workerId: owner,
      messageIndex: 0, sentMessageId: "roundup-part-0",
      payloadHash: "a".repeat(64)
    }), true)
    assert.equal(await roundupRepository.markSent({
      claimId: stateClaim.id, botInstanceName: "rachie-wos", workerId: owner, sentAt: now
    }), false)
    assert.equal(await roundupRepository.recordSentMessage({
      claimId: stateClaim.id, botInstanceName: "rachie-wos", workerId: owner,
      messageIndex: 1, sentMessageId: "roundup-part-1",
      payloadHash: "b".repeat(64)
    }), true)
    assert.equal(await roundupRepository.markSent({
      claimId: stateClaim.id, botInstanceName: "rachie-wos", workerId: owner, sentAt: now
    }), true)
    const messageRows = await pool.query(
      "SELECT message_index, sent_message_id FROM weekly_roundup_messages WHERE roundup_claim_id = $1 ORDER BY message_index",
      [stateClaim.id]
    )
    assert.deepEqual(messageRows.rows, [
      { message_index: 0, sent_message_id: "roundup-part-0" },
      { message_index: 1, sent_message_id: "roundup-part-1" }
    ])

    const kingshotRoundups = createWeeklyRoundupRepository(pool, "kingshot")
    const kingshotClaims = claimsForConfigurations(
      await kingshotRoundups.listRoundupConfigurations(),
      { gameProfile: "kingshot", now, graceMinutes: 60 }
    )
    assert.equal(await kingshotRoundups.insertMissingClaims(kingshotClaims), 2)
    const kingshotClaimed = await kingshotRoundups.claimDue({
      now, batchSize: 10, leaseSeconds: 60,
      botInstanceName: "peggie-kingshot", workerId: "kingshot-roundup"
    })
    assert.equal(kingshotClaimed.length, 2)
    assert.ok(kingshotClaimed.every(row => row.game_profile === "kingshot"))

    const deleted = await wosA.setEventStatus({
      guildId: guildA,
      eventId: managed.id,
      status: "deleted"
    })
    assert.equal(deleted.status, "deleted")
    assert.equal(await wosA.getEvent(guildA, managed.id), null)
    assert.equal((await wosA.listEvents(guildA)).events.some(row => String(row.id) === String(managed.id)), false)
    const stored = await pool.query(
      "SELECT status FROM scheduled_events WHERE id = $1 AND game_profile = 'wos'",
      [managed.id]
    )
    assert.equal(stored.rows[0].status, "deleted")
    const sentHistory = await pool.query(
      "SELECT status, sent_message_id FROM event_delivery_claims WHERE event_id = $1 AND sent_message_id = 'sent-history'",
      [managed.id]
    )
    assert.deepEqual(sentHistory.rows[0], { status: "sent", sent_message_id: "sent-history" })

    const schema = await pool.query(
      `SELECT to_regclass('weekly_roundup_messages') AS messages,
              to_regclass('weekly_roundup_claims_pending_due_idx') AS due_index`
    )
    assert.equal(schema.rows[0].messages, "weekly_roundup_messages")
    assert.equal(schema.rows[0].due_index, "weekly_roundup_claims_pending_due_idx")
    assert.ok(south.id)
  } finally {
    await pool.query(
      "DELETE FROM weekly_roundup_claims WHERE source_guild_id = ANY($1::varchar[])",
      [[guildA, guildB]]
    ).catch(() => {})
    await pool.query(
      "DELETE FROM event_guild_settings WHERE guild_id = ANY($1::varchar[])",
      [[guildA, guildB]]
    ).catch(() => {})
    await isolationClient.query("SELECT pg_advisory_unlock(7000005)").catch(() => {})
    isolationClient.release()
    await pool.end()
  }
})
