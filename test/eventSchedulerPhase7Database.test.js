const test = require("node:test")
const assert = require("node:assert/strict")
const { Pool } = require("pg")

const { buildDeliveryClaims } = require("../src/eventDeliveryGeneration")
const { createEventDeliveryRepository } = require("../src/eventDeliveryRepository")

const databaseUrl = process.env.TEST_DATABASE_URL

test("Phase 7 state claims remain link-aware, independent and profile isolated", {
  skip: databaseUrl ? false : "TEST_DATABASE_URL is not configured"
}, async () => {
  const pool = new Pool({ connectionString: databaseUrl, max: 8 })
  const isolationClient = await pool.connect()
  const allianceGuildId = "999999999999999921"
  const stateGuildId = "999999999999999922"
  const allianceChannel = "999999999999999923"
  const wosStateChannel = "999999999999999924"
  const kingshotStateChannel = "999999999999999925"
  const changedStateChannel = "999999999999999926"
  const userId = "999999999999999927"
  const generationStart = new Date("2026-08-06T11:00:00Z")
  const generationEnd = new Date("2026-08-06T13:00:00Z")
  const claimNow = new Date("2026-08-06T12:11:00Z")
  let wosEventId
  let kingshotEventId

  function claimInput(botInstanceName, workerId, now = claimNow) {
    return { now, batchSize: 20, leaseSeconds: 60, botInstanceName, workerId }
  }

  function manualStateClaim(eventId, occurrenceAt, channelId = wosStateChannel) {
    return {
      eventId,
      groupId: null,
      gameProfile: "wos",
      occurrenceAt,
      deliverAt: occurrenceAt,
      deliveryKind: "event_start",
      targetKind: "state",
      targetGuildId: stateGuildId,
      targetChannelId: channelId
    }
  }

  try {
    await isolationClient.query("SELECT pg_advisory_lock(7000005)")
    await pool.query(
      "DELETE FROM event_guild_settings WHERE guild_id = $1 AND game_profile IN ('wos', 'kingshot')",
      [allianceGuildId]
    )
    await pool.query(
      `INSERT INTO event_guild_settings (
         guild_id, game_profile, bot_instance_name, alliance_name, event_channel_id
       ) VALUES
         ($1, 'wos', 'rachie-wos', 'WOS North', $2),
         ($1, 'kingshot', 'peggie-kingshot', 'Kingshot North', $2)`,
      [allianceGuildId, allianceChannel]
    )
    await pool.query(
      `INSERT INTO event_state_links (
         alliance_guild_id, game_profile, configured_by_bot_instance,
         state_guild_id, state_event_channel_id, sharing_enabled
       ) VALUES
         ($1, 'wos', 'rachie-wos', $2, $3, true),
         ($1, 'kingshot', 'peggie-kingshot', $2, $4, true)`,
      [allianceGuildId, stateGuildId, wosStateChannel, kingshotStateChannel]
    )
    const events = await pool.query(
      `INSERT INTO scheduled_events (
         guild_id, game_profile, created_by_bot_instance, alliance_name,
         event_name, first_occurrence_date, event_time_utc, recurrence_days,
         advance_reminder_minutes, reminder_at_start, publish_to_alliance,
         publish_to_state, status, created_by_user_id
       ) VALUES
         ($1, 'wos', 'rachie-wos', 'WOS North', 'WOS State Event',
          '2026-08-06', '12:10', 3, 10, true, true, true, 'active', $2),
         ($1, 'kingshot', 'peggie-kingshot', 'Kingshot North', 'Kingshot State Event',
          '2026-08-06', '12:10', 3, 10, true, true, true, 'active', $2)
       RETURNING id, game_profile`,
      [allianceGuildId, userId]
    )
    wosEventId = events.rows.find(row => row.game_profile === "wos").id
    kingshotEventId = events.rows.find(row => row.game_profile === "kingshot").id

    const wos = createEventDeliveryRepository(pool, "wos")
    const kingshot = createEventDeliveryRepository(pool, "kingshot")
    const wosDefinitions = (await wos.listActiveEventDefinitions({ rangeEnd: generationEnd }))
      .filter(event => String(event.id) === String(wosEventId))
    const kingshotDefinitions = (await kingshot.listActiveEventDefinitions({ rangeEnd: generationEnd }))
      .filter(event => String(event.id) === String(kingshotEventId))
    assert.equal(wosDefinitions.length, 1)
    assert.equal(kingshotDefinitions.length, 1)
    assert.equal(wosDefinitions[0].first_occurrence_date, "2026-08-06")
    assert.equal(kingshotDefinitions[0].first_occurrence_date, "2026-08-06")
    assert.equal(wosDefinitions[0].state_event_channel_id, wosStateChannel)
    assert.equal(kingshotDefinitions[0].state_event_channel_id, kingshotStateChannel)

    const wosClaims = buildDeliveryClaims(wosDefinitions, {
      gameProfile: "wos",
      windowStart: generationStart,
      windowEnd: generationEnd
    })
    const kingshotClaims = buildDeliveryClaims(kingshotDefinitions, {
      gameProfile: "kingshot",
      windowStart: generationStart,
      windowEnd: generationEnd
    })
    assert.deepEqual(new Set(wosClaims.map(claim => claim.targetKind)), new Set(["alliance", "state"]))
    assert.equal(await wos.insertMissingDeliveryClaims(wosClaims), 4)
    assert.equal(await wos.insertMissingDeliveryClaims(wosClaims), 0)
    assert.equal(await kingshot.insertMissingDeliveryClaims(kingshotClaims), 4)

    const claimedWos = await wos.claimDueDeliveries(claimInput("rachie-wos", "wos-worker"))
    assert.equal(claimedWos.length, 4)
    assert.ok(claimedWos.every(row => row.game_profile === "wos"))
    const untouchedKingshot = await pool.query(
      `SELECT count(*)::integer AS count FROM event_delivery_claims
        WHERE event_id = $1 AND game_profile = 'kingshot' AND status = 'pending'`,
      [kingshotEventId]
    )
    assert.equal(untouchedKingshot.rows[0].count, 4)
    const claimedKingshot = await kingshot.claimDueDeliveries(
      claimInput("peggie-kingshot", "kingshot-worker")
    )
    assert.equal(claimedKingshot.length, 4)
    assert.ok(claimedKingshot.every(row => row.game_profile === "kingshot"))

    const wosAlliance = claimedWos.find(row => row.target_kind === "alliance")
    const wosState = claimedWos.find(row => row.target_kind === "state")
    assert.equal(await wos.markClaimSent({
      claimId: wosAlliance.id,
      botInstanceName: "rachie-wos",
      workerId: "wos-worker",
      sentAt: claimNow,
      sentMessageId: "alliance-message-id"
    }), true)
    assert.equal(await wos.markClaimSent({
      claimId: wosState.id,
      botInstanceName: "rachie-wos",
      workerId: "wos-worker",
      sentAt: claimNow,
      sentMessageId: "state-message-id"
    }), true)
    const independent = await pool.query(
      `SELECT target_kind, status, sent_message_id
         FROM event_delivery_claims
        WHERE id = ANY($1::bigint[])
        ORDER BY target_kind`,
      [[wosAlliance.id, wosState.id]]
    )
    assert.deepEqual(independent.rows, [
      { target_kind: "alliance", status: "sent", sent_message_id: "alliance-message-id" },
      { target_kind: "state", status: "sent", sent_message_id: "state-message-id" }
    ])

    await pool.query(
      `DELETE FROM event_delivery_claims
        WHERE event_id = $1 AND game_profile = 'wos' AND status <> 'sent'`,
      [wosEventId]
    )
    await wos.insertMissingDeliveryClaims([
      manualStateClaim(wosEventId, new Date("2026-08-09T12:10:00Z"))
    ])
    await pool.query(
      `UPDATE event_state_links SET sharing_enabled = false
        WHERE alliance_guild_id = $1 AND game_profile = 'wos'`,
      [allianceGuildId]
    )
    assert.equal(await wos.reconcileInvalidStateClaims({ now: claimNow }), 1)
    const disabled = await pool.query(
      `SELECT status, next_attempt_at, last_error FROM event_delivery_claims
        WHERE event_id = $1 AND game_profile = 'wos'
          AND target_kind = 'state' AND occurrence_at = '2026-08-09T12:10:00Z'`,
      [wosEventId]
    )
    assert.equal(disabled.rows[0].status, "failed")
    assert.equal(disabled.rows[0].next_attempt_at, null)
    assert.match(disabled.rows[0].last_error, /sharing disabled or target changed/i)
    const sentHistory = await pool.query(
      "SELECT status, sent_message_id FROM event_delivery_claims WHERE id = $1",
      [wosState.id]
    )
    assert.deepEqual(sentHistory.rows[0], {
      status: "sent",
      sent_message_id: "state-message-id"
    })
    const disabledDefinitions = (await wos.listActiveEventDefinitions({ rangeEnd: generationEnd }))
      .filter(event => String(event.id) === String(wosEventId))
    assert.ok(buildDeliveryClaims(disabledDefinitions, {
      gameProfile: "wos",
      windowStart: generationStart,
      windowEnd: generationEnd
    }).every(claim => claim.targetKind === "alliance"))

    await pool.query(
      `DELETE FROM event_delivery_claims
        WHERE event_id = $1 AND game_profile = 'wos' AND status <> 'sent'`,
      [wosEventId]
    )
    await pool.query(
      `UPDATE event_state_links
          SET sharing_enabled = true, state_event_channel_id = $2
        WHERE alliance_guild_id = $1 AND game_profile = 'wos'`,
      [allianceGuildId, wosStateChannel]
    )
    await wos.insertMissingDeliveryClaims([
      manualStateClaim(wosEventId, new Date("2026-08-12T12:10:00Z"))
    ])
    const activeStateClaim = await createEventDeliveryRepository(
      pool,
      "wos",
      { targetKind: "state" }
    ).claimDueDeliveries(claimInput(
      "rachie-wos",
      "active-state-worker",
      new Date("2026-08-12T12:11:00Z")
    ))
    assert.equal(activeStateClaim.length, 1)
    await pool.query(
      `UPDATE event_state_links SET sharing_enabled = false
        WHERE alliance_guild_id = $1 AND game_profile = 'wos'`,
      [allianceGuildId]
    )
    const stalePayload = await wos.getClaimPayload({
      claimId: activeStateClaim[0].id,
      botInstanceName: "rachie-wos",
      workerId: "active-state-worker"
    })
    assert.equal(stalePayload.claim.targetIsCurrent, false)
    assert.equal(await wos.markClaimPermanentlyFailed({
      claimId: activeStateClaim[0].id,
      botInstanceName: "rachie-wos",
      workerId: "active-state-worker",
      failedAt: new Date("2026-08-12T12:11:01Z"),
      lastError: "State sharing disabled or target changed."
    }), true)

    await pool.query(
      `DELETE FROM event_delivery_claims
        WHERE event_id = $1 AND game_profile = 'wos' AND status <> 'sent'`,
      [wosEventId]
    )
    await pool.query(
      `UPDATE event_state_links
          SET sharing_enabled = true, state_event_channel_id = $2
        WHERE alliance_guild_id = $1 AND game_profile = 'wos'`,
      [allianceGuildId, wosStateChannel]
    )
    await wos.insertMissingDeliveryClaims([
      manualStateClaim(wosEventId, new Date("2026-08-09T12:10:00Z"), wosStateChannel)
    ])
    await pool.query(
      `UPDATE event_state_links SET state_event_channel_id = $2
        WHERE alliance_guild_id = $1 AND game_profile = 'wos'`,
      [allianceGuildId, changedStateChannel]
    )
    assert.equal(await wos.reconcileInvalidStateClaims({ now: claimNow }), 1)
    const oldTarget = await pool.query(
      `SELECT status FROM event_delivery_claims
        WHERE event_id = $1 AND game_profile = 'wos'
          AND target_kind = 'state' AND target_channel_id = $2
          AND occurrence_at = '2026-08-09T12:10:00Z'`,
      [wosEventId, wosStateChannel]
    )
    assert.equal(oldTarget.rows[0].status, "failed")

    const changedDefinitions = (await wos.listActiveEventDefinitions({
      rangeEnd: new Date("2026-08-09T13:00:00Z")
    })).filter(event => String(event.id) === String(wosEventId))
    const changedClaims = buildDeliveryClaims(changedDefinitions, {
      gameProfile: "wos",
      windowStart: new Date("2026-08-09T11:00:00Z"),
      windowEnd: new Date("2026-08-09T13:00:00Z")
    })
    assert.ok(changedClaims.some(claim =>
      claim.targetKind === "state" && claim.targetChannelId === changedStateChannel
    ))
    await wos.insertMissingDeliveryClaims(changedClaims)
    const newTarget = await pool.query(
      `SELECT status FROM event_delivery_claims
        WHERE event_id = $1 AND game_profile = 'wos'
          AND target_kind = 'state' AND target_channel_id = $2
          AND occurrence_at = '2026-08-09T12:10:00Z'`,
      [wosEventId, changedStateChannel]
    )
    assert.equal(newTarget.rows[0].status, "pending")

    const indexResult = await pool.query(
      `SELECT 1 FROM pg_indexes
        WHERE tablename = 'event_delivery_claims'
          AND indexname = 'event_delivery_claims_state_unsent_idx'`
    )
    assert.equal(indexResult.rowCount, 1)
  } finally {
    await pool.query(
      "DELETE FROM event_guild_settings WHERE guild_id = $1 AND game_profile IN ('wos', 'kingshot')",
      [allianceGuildId]
    ).catch(() => {})
    await isolationClient.query("SELECT pg_advisory_unlock(7000005)").catch(() => {})
    isolationClient.release()
    await pool.end()
  }
})
