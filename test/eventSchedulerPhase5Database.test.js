const test = require("node:test")
const assert = require("node:assert/strict")
const { Pool } = require("pg")

const {
  createEventDeliveryRepository
} = require("../src/eventDeliveryRepository")

const databaseUrl = process.env.TEST_DATABASE_URL

test("Phase 5 delivery claims are transactional, leased and profile isolated", {
  skip: databaseUrl ? false : "TEST_DATABASE_URL is not configured"
}, async () => {
  const pool = new Pool({ connectionString: databaseUrl, max: 8 })
  const isolationClient = await pool.connect()
  const guildId = "999999999999999951"
  const channelId = "999999999999999952"
  const now = new Date("2026-08-06T12:00:00Z")
  let wosEventId
  let kingshotEventId
  let claimSequence = 0

  async function insertClaim({
    eventId = wosEventId,
    gameProfile = "wos",
    status = "pending",
    deliverAt = new Date(now.getTime() - 60000),
    claimedUntil = null,
    attemptCount = 0,
    nextAttemptAt = null
  } = {}) {
    const result = await pool.query(
      `INSERT INTO event_delivery_claims (
         event_id, game_profile, occurrence_at, deliver_at, delivery_kind,
         target_kind, target_guild_id, target_channel_id, status,
         claimed_by_bot_instance, claimed_by_worker, claimed_at, claimed_until,
         attempt_count, next_attempt_at
       ) VALUES (
         $1, $2, $3::timestamptz, $4, 'event_start', 'alliance', $5, $6, $7::varchar,
         CASE WHEN $7::varchar = 'claimed' THEN 'old-instance' END,
         CASE WHEN $7::varchar = 'claimed' THEN 'old-worker' END,
         CASE WHEN $7::varchar = 'claimed' THEN $3::timestamptz END,
         $8, $9, $10
       ) RETURNING *`,
      [
        eventId,
        gameProfile,
        new Date(deliverAt.getTime() + 60000 + claimSequence++),
        deliverAt,
        guildId,
        channelId,
        status,
        claimedUntil,
        attemptCount,
        nextAttemptAt
      ]
    )
    return result.rows[0]
  }

  function claimInput(workerId, overrides = {}) {
    return {
      now,
      batchSize: 10,
      leaseSeconds: 60,
      botInstanceName: "rachie-wos",
      workerId,
      ...overrides
    }
  }

  try {
    await isolationClient.query("SELECT pg_advisory_lock(7000005)")
    await pool.query(
      "DELETE FROM event_guild_settings WHERE guild_id = $1 AND game_profile IN ('wos', 'kingshot')",
      [guildId]
    )
    await pool.query(
      `INSERT INTO event_guild_settings (
         guild_id, game_profile, bot_instance_name, alliance_name, event_channel_id
       ) VALUES
         ($1, 'wos', 'rachie-wos', 'WOS', $2),
         ($1, 'kingshot', 'peggie-kingshot', 'Kingshot', $2)`,
      [guildId, channelId]
    )
    const events = await pool.query(
      `INSERT INTO scheduled_events (
         guild_id, game_profile, created_by_bot_instance, alliance_name,
         event_name, first_occurrence_date, event_time_utc, recurrence_days,
         advance_reminder_minutes, reminder_at_start, publish_to_alliance,
         status, created_by_user_id
       ) VALUES
         ($1, 'wos', 'rachie-wos', 'WOS', 'WOS Event', '2026-08-06',
          '12:00', 3, 10, true, true, 'active', $2),
         ($1, 'kingshot', 'peggie-kingshot', 'Kingshot', 'Kingshot Event',
          '2026-08-06', '12:00', 3, 10, true, true, 'active', $2)
       RETURNING id, game_profile`,
      [guildId, "999999999999999953"]
    )
    wosEventId = events.rows.find(row => row.game_profile === "wos").id
    kingshotEventId = events.rows.find(row => row.game_profile === "kingshot").id

    const wos = createEventDeliveryRepository(pool, "wos")
    const kingshot = createEventDeliveryRepository(pool, "kingshot")

    const generated = {
      eventId: wosEventId,
      groupId: null,
      gameProfile: "wos",
      occurrenceAt: now,
      deliverAt: new Date(now.getTime() - 60000),
      deliveryKind: "advance_reminder",
      targetKind: "alliance",
      targetGuildId: guildId,
      targetChannelId: channelId
    }
    assert.equal(await wos.insertMissingDeliveryClaims([generated]), 1)
    assert.equal(await wos.insertMissingDeliveryClaims([generated]), 0)
    await pool.query("DELETE FROM event_delivery_claims WHERE event_id = $1", [wosEventId])

    const onlyWos = await insertClaim()
    await insertClaim({ eventId: kingshotEventId, gameProfile: "kingshot" })
    const [workerA, workerB] = await Promise.all([
      wos.claimDueDeliveries(claimInput("worker-a")),
      wos.claimDueDeliveries(claimInput("worker-b"))
    ])
    assert.equal(workerA.length + workerB.length, 1)
    assert.equal(String([...workerA, ...workerB][0].id), String(onlyWos.id))
    assert.equal(
      (await kingshot.claimDueDeliveries({
        ...claimInput("worker-kingshot"),
        botInstanceName: "peggie-kingshot"
      })).length,
      1
    )

    await pool.query("DELETE FROM event_delivery_claims WHERE event_id IN ($1, $2)", [
      wosEventId,
      kingshotEventId
    ])
    const due = await Promise.all([insertClaim(), insertClaim(), insertClaim()])
    await insertClaim({ deliverAt: new Date(now.getTime() + 60000) })
    const bounded = await wos.claimDueDeliveries(claimInput("batch-worker", { batchSize: 2 }))
    assert.equal(bounded.length, 2)
    assert.equal(new Set(bounded.map(row => String(row.id))).size, 2)
    assert.ok(bounded.every(row => due.some(item => String(item.id) === String(row.id))))

    await pool.query("DELETE FROM event_delivery_claims WHERE event_id = $1", [wosEventId])
    const active = await insertClaim({
      status: "claimed",
      claimedUntil: new Date(now.getTime() + 5 * 60000),
      attemptCount: 1
    })
    const expired = await insertClaim({
      status: "claimed",
      claimedUntil: new Date(now.getTime() - 1000),
      attemptCount: 1
    })
    const retry = await insertClaim({
      status: "failed",
      attemptCount: 1,
      nextAttemptAt: new Date(now.getTime() - 1000)
    })
    const reclaimed = await wos.claimDueDeliveries(claimInput("reclaimer"))
    assert.deepEqual(
      new Set(reclaimed.map(row => String(row.id))),
      new Set([String(expired.id), String(retry.id)])
    )
    assert.ok(!reclaimed.some(row => String(row.id) === String(active.id)))
    assert.ok(reclaimed.every(row => row.attempt_count === 2))

    const owned = reclaimed.find(row => String(row.id) === String(expired.id))
    const retryOwned = reclaimed.find(row => String(row.id) === String(retry.id))
    assert.equal(await wos.markClaimSent({
      claimId: owned.id,
      botInstanceName: "rachie-wos",
      workerId: "wrong-worker",
      sentAt: now,
      sentMessageId: "wrong"
    }), false)
    assert.equal(await wos.markClaimFailed({
      claimId: owned.id,
      botInstanceName: "wrong-instance",
      workerId: "reclaimer",
      failedAt: now,
      lastError: "wrong owner",
      nextAttemptAt: new Date(now.getTime() + 60000)
    }), false)
    assert.equal(await wos.markClaimSent({
      claimId: owned.id,
      botInstanceName: "rachie-wos",
      workerId: "reclaimer",
      sentAt: now,
      sentMessageId: "synthetic-integration-id"
    }), true)
    const sent = await pool.query(
      "SELECT status, sent_message_id, claimed_by_worker FROM event_delivery_claims WHERE id = $1",
      [owned.id]
    )
    assert.deepEqual(sent.rows[0], {
      status: "sent",
      sent_message_id: "synthetic-integration-id",
      claimed_by_worker: null
    })

    const retryAt = new Date(now.getTime() + 60000)
    assert.equal(await wos.markClaimFailed({
      claimId: retryOwned.id,
      botInstanceName: "rachie-wos",
      workerId: "reclaimer",
      failedAt: now,
      lastError: "temporary failure",
      nextAttemptAt: retryAt
    }), true)
    const retried = await wos.claimDueDeliveries({
      ...claimInput("retry-worker"),
      now: retryAt
    })
    assert.equal(retried.length, 1)
    assert.equal(String(retried[0].id), String(retry.id))
    assert.equal(retried[0].attempt_count, 3)
    const payload = await wos.getClaimPayload({
      claimId: retry.id,
      botInstanceName: "rachie-wos",
      workerId: "retry-worker"
    })
    assert.equal(payload.event.eventName, "WOS Event")
    assert.equal(payload.claim.gameProfile, "wos")
    assert.equal(await wos.markClaimSent({
      claimId: retry.id,
      botInstanceName: "rachie-wos",
      workerId: "retry-worker",
      sentAt: retryAt
    }), true)

    const exhausted = await insertClaim({
      status: "claimed",
      claimedUntil: new Date(now.getTime() - 1000),
      attemptCount: 5
    })
    assert.equal((await wos.claimDueDeliveries(claimInput("too-late"))).length, 0)
    const exhaustedResult = await pool.query(
      "SELECT status, next_attempt_at, claimed_by_worker FROM event_delivery_claims WHERE id = $1",
      [exhausted.id]
    )
    assert.deepEqual(exhaustedResult.rows[0], {
      status: "failed",
      next_attempt_at: null,
      claimed_by_worker: null
    })

    const indexResult = await pool.query(
      `SELECT indexname FROM pg_indexes
        WHERE tablename = 'event_delivery_claims'
          AND indexname = ANY($1::text[])`,
      [[
        "event_delivery_claims_pending_due_idx",
        "event_delivery_claims_failed_due_idx",
        "event_delivery_claims_expired_lease_idx"
      ]]
    )
    assert.equal(indexResult.rowCount, 3)
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
