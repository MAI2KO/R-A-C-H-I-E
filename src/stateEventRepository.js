const { assertProfile } = require("./eventSchedulerRepository")
const { MAX_DELIVERY_ATTEMPTS } = require("./eventDeliveryRepository")
const { buildStateRoundupOccurrences } = require("./stateEventRoundupCalculation")
const { validateStateEventDraft, normalizeStateNumber } = require("./stateEventValidation")

function requirePool(pool) {
  if (!pool || typeof pool.query !== "function" || typeof pool.connect !== "function") {
    throw new Error("A transactional Postgres pool is required")
  }
}

function requireOwner(value, label) {
  const normalized = String(value || "").trim()
  if (!normalized) throw new Error(`${label} is required`)
  return normalized
}

function cloneDate(value) {
  return value instanceof Date ? new Date(value) : new Date(String(value))
}

function mapMedia(row) {
  if (!row?.image_data) return null
  return Object.freeze({
    originalFilename: row.original_filename,
    contentType: row.content_type,
    byteSize: row.byte_size,
    imageData: Buffer.from(row.image_data)
  })
}

function phaseFromRow(row) {
  return {
    id: row.phase_id || row.id,
    state_event_id: row.state_event_id,
    game_profile: row.game_profile,
    phase_name: row.phase_name,
    phase_time_utc: row.phase_time_utc,
    pre_alert_minutes: row.pre_alert_minutes,
    pre_alert_message: row.pre_alert_message,
    announce_exact: row.announce_exact,
    exact_message: row.exact_message,
    sort_order: row.sort_order,
    pre_alert_media: row.pre_alert_media || null,
    exact_media: row.exact_media || null
  }
}

function attachPhases(events, phaseRows) {
  const phasesByEvent = new Map()
  for (const row of phaseRows) {
    const key = String(row.state_event_id)
    if (!phasesByEvent.has(key)) phasesByEvent.set(key, [])
    phasesByEvent.get(key).push(phaseFromRow(row))
  }
  return events.map(event => ({
    ...event,
    phases: phasesByEvent.get(String(event.id)) || []
  }))
}

async function synchronizeStateEventPhases(client, gameProfile, eventId, phases) {
  const existing = (await client.query(
    `SELECT id FROM state_event_phases
      WHERE state_event_id = $1 AND game_profile = $2
      FOR UPDATE`,
    [eventId, gameProfile]
  )).rows
  const existingIds = new Set(existing.map(row => String(row.id)))
  const retainedIds = []
  await client.query(
    `UPDATE state_event_phases
        SET status = 'deleted', updated_at = now()
      WHERE state_event_id = $1 AND game_profile = $2 AND status = 'active'`,
    [eventId, gameProfile]
  )

  for (const phase of phases) {
    const requestedPhaseId = String(phase.id || "")
    const values = [
      eventId,
      gameProfile,
      phase.phaseName,
      phase.phaseTimeUtc,
      phase.preAlertMinutes,
      phase.preAlertMessage,
      Boolean(phase.announceExact),
      phase.exactMessage,
      phase.sortOrder
    ]
    const phaseResult = existingIds.has(requestedPhaseId)
      ? await client.query(
        `UPDATE state_event_phases
            SET phase_name = $3, phase_time_utc = $4,
                pre_alert_minutes = $5, pre_alert_message = $6,
                announce_exact = $7, exact_message = $8, sort_order = $9,
                status = 'active', updated_at = now()
          WHERE id = $10 AND state_event_id = $1 AND game_profile = $2
          RETURNING id`,
        [...values, requestedPhaseId]
      )
      : await client.query(
        `INSERT INTO state_event_phases (
           state_event_id, game_profile, phase_name, phase_time_utc,
           pre_alert_minutes, pre_alert_message, announce_exact,
           exact_message, sort_order
         ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9)
         RETURNING id`,
        values
      )
    const phaseId = phaseResult.rows[0].id
    retainedIds.push(phaseId)
    for (const [deliveryKind, media] of [
      ["pre_alert", phase.preAlertMedia],
      ["exact", phase.exactMedia]
    ]) {
      if (media === undefined) continue
      await client.query(
        `DELETE FROM state_event_media
          WHERE phase_id = $1 AND game_profile = $2 AND delivery_kind = $3`,
        [phaseId, gameProfile, deliveryKind]
      )
      if (!media) continue
      await client.query(
        `INSERT INTO state_event_media (
           state_event_id, phase_id, game_profile, delivery_kind,
           original_filename, content_type, byte_size, image_data
         ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8)`,
        [
          eventId,
          phaseId,
          gameProfile,
          deliveryKind,
          media.originalFilename,
          media.contentType,
          media.byteSize,
          media.imageData
        ]
      )
    }
  }
  await client.query(
    `UPDATE state_event_phases
        SET status = 'deleted', updated_at = now()
      WHERE state_event_id = $1 AND game_profile = $2
        AND status = 'active' AND NOT (id = ANY($3::uuid[]))`,
    [eventId, gameProfile, retainedIds]
  )
}

async function loadPhaseRows(pool, gameProfile, eventIds) {
  if (!eventIds.length) return []
  const [phaseResult, mediaResult] = await Promise.all([
    pool.query(
      `SELECT * FROM state_event_phases
        WHERE game_profile = $1 AND state_event_id = ANY($2::uuid[])
          AND status = 'active'
        ORDER BY state_event_id, phase_time_utc, sort_order, phase_name, id`,
      [gameProfile, eventIds]
    ),
    pool.query(
      `SELECT m.* FROM state_event_media m
        JOIN state_event_phases p
          ON p.id = m.phase_id AND p.state_event_id = m.state_event_id
         AND p.game_profile = m.game_profile
       WHERE m.game_profile = $1 AND m.state_event_id = ANY($2::uuid[])
         AND p.status = 'active'`,
      [gameProfile, eventIds]
    )
  ])
  const media = new Map()
  for (const row of mediaResult.rows) {
    media.set(`${row.phase_id}:${row.delivery_kind}`, mapMedia(row))
  }
  return phaseResult.rows.map(row => ({
    ...row,
    pre_alert_media: media.get(`${row.id}:pre_alert`) || null,
    exact_media: media.get(`${row.id}:exact`) || null
  }))
}

function createStateEventRepository(pool, gameProfile) {
  requirePool(pool)
  assertProfile(gameProfile)

  const repository = {
    gameProfile,

    async setStateNumber({ stateGuildId, stateNumber }) {
      const result = await pool.query(
        `UPDATE event_state_destinations
            SET state_number = $3, updated_at = now()
          WHERE state_guild_id = $1 AND game_profile = $2 AND enabled = true
          RETURNING *`,
        [stateGuildId, gameProfile, normalizeStateNumber(stateNumber)]
      )
      return result.rows[0] || null
    },

    async createStateEvent({
      stateGuildId,
      createdByUserId,
      createdByBotInstance,
      event
    }) {
      const draft = validateStateEventDraft(event)
      const client = await pool.connect()
      try {
        await client.query("BEGIN")
        const destination = (await client.query(
          `SELECT state_guild_id
             FROM event_state_destinations
            WHERE state_guild_id = $1 AND game_profile = $2 AND enabled = true
            FOR SHARE`,
          [stateGuildId, gameProfile]
        )).rows[0]
        if (!destination) throw new Error("This Discord is not an enabled state destination.")
        const created = (await client.query(
          `INSERT INTO state_events (
             state_guild_id, game_profile, created_by_bot_instance, event_name,
             first_occurrence_date, recurrence_days, created_by_user_id
           ) VALUES ($1, $2, $3, $4, $5, $6, $7)
           RETURNING *`,
          [
            stateGuildId,
            gameProfile,
            createdByBotInstance,
            draft.eventName,
            draft.firstOccurrenceDate,
            draft.recurrenceDays,
            createdByUserId
          ]
        )).rows[0]
        await synchronizeStateEventPhases(client, gameProfile, created.id, draft.phases)
        await client.query("COMMIT")
        return created
      } catch (error) {
        await client.query("ROLLBACK")
        throw error
      } finally {
        client.release()
      }
    },

    async listStateEvents(stateGuildId, { limit = 10, offset = 0 } = {}) {
      const result = await pool.query(
        `SELECT e.*, e.first_occurrence_date::text AS first_occurrence_date,
                d.state_number, COUNT(*) OVER()::integer AS total_count
           FROM state_events e
           JOIN event_state_destinations d
             ON d.state_guild_id = e.state_guild_id AND d.game_profile = e.game_profile
          WHERE e.state_guild_id = $1 AND e.game_profile = $2
            AND e.status IN ('active', 'paused')
          ORDER BY e.first_occurrence_date, e.event_name, e.id
          LIMIT $3 OFFSET $4`,
        [stateGuildId, gameProfile, limit, offset]
      )
      const eventIds = result.rows.map(event => event.id)
      const phases = await loadPhaseRows(pool, gameProfile, eventIds)
      return {
        events: attachPhases(result.rows, phases),
        total: result.rows[0]?.total_count || 0
      }
    },

    async getStateEvent(stateGuildId, eventId) {
      const event = (await pool.query(
        `SELECT e.*, e.first_occurrence_date::text AS first_occurrence_date,
                d.state_number
           FROM state_events e
           JOIN event_state_destinations d
             ON d.state_guild_id = e.state_guild_id AND d.game_profile = e.game_profile
          WHERE e.id = $1 AND e.state_guild_id = $2 AND e.game_profile = $3
            AND e.status IN ('active', 'paused')`,
        [eventId, stateGuildId, gameProfile]
      )).rows[0]
      if (!event) return null
      const phases = await loadPhaseRows(pool, gameProfile, [eventId])
      return { ...event, phases: phases.map(phaseFromRow) }
    },

    async updateStateEvent({ stateGuildId, eventId, event }) {
      const draft = validateStateEventDraft(event)
      const client = await pool.connect()
      try {
        await client.query("BEGIN")
        const updated = (await client.query(
          `UPDATE state_events
              SET event_name = $4, first_occurrence_date = $5,
                  recurrence_days = $6, schedule_version = schedule_version + 1,
                  updated_at = now()
            WHERE id = $1 AND state_guild_id = $2 AND game_profile = $3
              AND status IN ('active', 'paused')
            RETURNING *`,
          [
            eventId,
            stateGuildId,
            gameProfile,
            draft.eventName,
            draft.firstOccurrenceDate,
            draft.recurrenceDays
          ]
        )).rows[0]
        if (!updated) {
          await client.query("ROLLBACK")
          return null
        }
        await synchronizeStateEventPhases(client, gameProfile, eventId, draft.phases)
        await client.query(
          `UPDATE state_event_delivery_claims
              SET status = 'failed', next_attempt_at = NULL,
                  claimed_by_bot_instance = NULL, claimed_by_worker = NULL,
                  claimed_at = NULL, claimed_until = NULL,
                  last_error = 'State event schedule changed.', updated_at = now()
            WHERE state_event_id = $1 AND game_profile = $2 AND status <> 'sent'`,
          [eventId, gameProfile]
        )
        await client.query("COMMIT")
        return updated
      } catch (error) {
        await client.query("ROLLBACK")
        throw error
      } finally {
        client.release()
      }
    },

    async setStateEventStatus({ stateGuildId, eventId, status }) {
      if (!["active", "paused", "deleted"].includes(status)) {
        throw new Error("Unsupported state event status")
      }
      const client = await pool.connect()
      try {
        await client.query("BEGIN")
        const updated = (await client.query(
          `UPDATE state_events
              SET status = $4, schedule_version = schedule_version + 1, updated_at = now()
            WHERE id = $1 AND state_guild_id = $2 AND game_profile = $3
              AND status IN ('active', 'paused')
            RETURNING *`,
          [eventId, stateGuildId, gameProfile, status]
        )).rows[0]
        if (!updated) {
          await client.query("ROLLBACK")
          return null
        }
        await client.query(
          `UPDATE state_event_delivery_claims
              SET status = 'failed', next_attempt_at = NULL,
                  claimed_by_bot_instance = NULL, claimed_by_worker = NULL,
                  claimed_at = NULL, claimed_until = NULL,
                  last_error = $3, updated_at = now()
            WHERE state_event_id = $1 AND game_profile = $2 AND status <> 'sent'`,
          [eventId, gameProfile, status === "deleted" ? "State event deleted." : "State event status changed."]
        )
        await client.query("COMMIT")
        return updated
      } catch (error) {
        await client.query("ROLLBACK")
        throw error
      } finally {
        client.release()
      }
    },

    async listActiveStateEventDefinitions({ rangeEnd }) {
      const events = (await pool.query(
        `SELECT e.*, e.first_occurrence_date::text AS first_occurrence_date,
                d.state_number, d.state_roundup_channel_id
           FROM state_events e
           JOIN event_state_destinations d
             ON d.state_guild_id = e.state_guild_id AND d.game_profile = e.game_profile
          WHERE e.game_profile = $1 AND e.status = 'active' AND d.enabled = true
            AND e.first_occurrence_date <= ($2::timestamptz AT TIME ZONE 'UTC')::date
          ORDER BY e.id`,
        [gameProfile, rangeEnd]
      )).rows
      if (!events.length) return []
      const phases = await loadPhaseRows(pool, gameProfile, events.map(event => event.id))
      return attachPhases(events, phases)
    },

    async listTargetsForStateGuild(stateGuildId) {
      const result = await pool.query(
        `SELECT 'state'::varchar AS target_kind, d.state_guild_id AS target_guild_id,
                d.state_roundup_channel_id AS target_channel_id
           FROM event_state_destinations d
          WHERE d.state_guild_id = $1 AND d.game_profile = $2 AND d.enabled = true
        UNION
         SELECT DISTINCT 'alliance'::varchar AS target_kind,
                s.guild_id AS target_guild_id, s.event_channel_id AS target_channel_id
           FROM event_state_links l
           JOIN event_state_destinations d
             ON d.state_guild_id = l.state_guild_id AND d.game_profile = l.game_profile
           JOIN event_guild_settings s
             ON s.guild_id = l.alliance_guild_id AND s.game_profile = l.game_profile
          WHERE l.state_guild_id = $1 AND l.game_profile = $2
            AND l.sharing_enabled = true AND d.enabled = true
            AND s.event_channel_id IS NOT NULL
          ORDER BY target_kind DESC, target_guild_id`,
        [stateGuildId, gameProfile]
      )
      return result.rows
    },

    async insertMissingDeliveryClaims(claims) {
      if (!Array.isArray(claims) || claims.length === 0) return 0
      const client = await pool.connect()
      let inserted = 0
      try {
        await client.query("BEGIN")
        for (const claim of claims) {
          if (claim.gameProfile !== gameProfile) throw new Error("State event claim profile mismatch")
          const result = await client.query(
            `INSERT INTO state_event_delivery_claims (
               state_event_id, phase_id, game_profile, schedule_version,
               occurrence_at, deliver_at, delivery_kind, target_kind,
               target_guild_id, target_channel_id
             ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10)
             ON CONFLICT DO NOTHING
             RETURNING id`,
            [
              claim.stateEventId,
              claim.phaseId,
              gameProfile,
              claim.scheduleVersion,
              claim.occurrenceAt,
              claim.deliverAt,
              claim.deliveryKind,
              claim.targetKind,
              claim.targetGuildId,
              claim.targetChannelId
            ]
          )
          inserted += result.rowCount
        }
        await client.query("COMMIT")
        return inserted
      } catch (error) {
        await client.query("ROLLBACK")
        throw error
      } finally {
        client.release()
      }
    },

    async reconcileInvalidClaims({ now, client = pool }) {
      const result = await client.query(
        `UPDATE state_event_delivery_claims c
            SET status = 'failed', next_attempt_at = NULL,
                claimed_by_bot_instance = NULL, claimed_by_worker = NULL,
                claimed_at = NULL, claimed_until = NULL,
                last_error = 'State event schedule or target changed.',
                updated_at = $2
           FROM state_events e
           JOIN event_state_destinations d
             ON d.state_guild_id = e.state_guild_id AND d.game_profile = e.game_profile
          WHERE c.state_event_id = e.id AND c.game_profile = e.game_profile
            AND c.game_profile = $1
            AND (
              c.status = 'pending'
              OR (c.status = 'failed' AND c.next_attempt_at IS NOT NULL)
              OR (c.status = 'claimed' AND c.claimed_until <= $2)
            )
            AND NOT COALESCE((
              e.status = 'active' AND e.schedule_version = c.schedule_version
              AND d.enabled = true
              AND (
                (c.target_kind = 'state'
                  AND d.state_guild_id = c.target_guild_id
                  AND d.state_roundup_channel_id = c.target_channel_id)
                OR
                (c.target_kind = 'alliance' AND EXISTS (
                  SELECT 1
                    FROM event_state_links l
                    JOIN event_guild_settings s
                      ON s.guild_id = l.alliance_guild_id AND s.game_profile = l.game_profile
                   WHERE l.game_profile = c.game_profile
                     AND l.state_guild_id = e.state_guild_id
                     AND l.sharing_enabled = true
                     AND s.guild_id = c.target_guild_id
                     AND s.event_channel_id = c.target_channel_id
                ))
              )
            ), false)
          RETURNING c.id`,
        [gameProfile, now]
      )
      return result.rowCount
    },

    async claimDueDeliveries({ now, batchSize, leaseSeconds, botInstanceName, workerId }) {
      const bot = requireOwner(botInstanceName, "Bot instance name")
      const worker = requireOwner(workerId, "Worker ID")
      const client = await pool.connect()
      try {
        await client.query("BEGIN")
        await repository.reconcileInvalidClaims({ now, client })
        await client.query(
          `UPDATE state_event_delivery_claims
              SET status = 'failed', next_attempt_at = NULL,
                  claimed_by_bot_instance = NULL, claimed_by_worker = NULL,
                  claimed_at = NULL, claimed_until = NULL,
                  last_error = 'Claim lease expired after maximum attempts.', updated_at = $2
            WHERE game_profile = $1 AND status = 'claimed'
              AND claimed_until <= $2 AND attempt_count >= $3`,
          [gameProfile, now, MAX_DELIVERY_ATTEMPTS]
        )
        const result = await client.query(
          `WITH claimable AS (
             SELECT id FROM state_event_delivery_claims
              WHERE game_profile = $1 AND attempt_count < $7
                AND (
                  (status = 'pending' AND deliver_at <= $2)
                  OR (status = 'failed' AND next_attempt_at <= $2)
                  OR (status = 'claimed' AND claimed_until <= $2)
                )
              ORDER BY COALESCE(next_attempt_at, claimed_until, deliver_at), id
              FOR UPDATE SKIP LOCKED
              LIMIT $3
           )
           UPDATE state_event_delivery_claims c
              SET status = 'claimed', claimed_by_bot_instance = $4,
                  claimed_by_worker = $5, claimed_at = $2,
                  claimed_until = $2 + ($6 * interval '1 second'),
                  attempt_count = c.attempt_count + 1,
                  next_attempt_at = NULL, last_error = NULL, updated_at = $2
             FROM claimable
            WHERE c.id = claimable.id AND c.game_profile = $1
           RETURNING c.*`,
          [gameProfile, now, batchSize, bot, worker, leaseSeconds, MAX_DELIVERY_ATTEMPTS]
        )
        await client.query("COMMIT")
        return result.rows
      } catch (error) {
        await client.query("ROLLBACK")
        throw error
      } finally {
        client.release()
      }
    },

    async getClaimPayload({ claimId, botInstanceName, workerId }) {
      const row = (await pool.query(
        `SELECT c.id AS claim_id, c.game_profile, c.attempt_count,
                c.schedule_version, c.delivery_kind, c.target_kind,
                c.target_guild_id, c.target_channel_id, c.occurrence_at,
                c.deliver_at,
                e.id AS state_event_id, e.state_guild_id, e.event_name,
                e.recurrence_days, d.state_number,
                p.id AS phase_id, p.phase_name, p.phase_time_utc,
                p.pre_alert_minutes, p.pre_alert_message,
                p.announce_exact, p.exact_message, p.sort_order,
                CASE
                  WHEN e.status = 'active' AND e.schedule_version = c.schedule_version
                   AND d.enabled = true
                   AND (
                     (c.target_kind = 'state'
                       AND d.state_guild_id = c.target_guild_id
                       AND d.state_roundup_channel_id = c.target_channel_id)
                     OR
                     (c.target_kind = 'alliance' AND EXISTS (
                       SELECT 1
                         FROM event_state_links l
                         JOIN event_guild_settings s
                           ON s.guild_id = l.alliance_guild_id AND s.game_profile = l.game_profile
                        WHERE l.game_profile = c.game_profile
                          AND l.state_guild_id = e.state_guild_id
                          AND l.sharing_enabled = true
                          AND s.guild_id = c.target_guild_id
                          AND s.event_channel_id = c.target_channel_id
                     ))
                   )
                  THEN true ELSE false
                END AS target_is_current,
                m.original_filename, m.content_type, m.byte_size, m.image_data
           FROM state_event_delivery_claims c
           JOIN state_events e
             ON e.id = c.state_event_id AND e.game_profile = c.game_profile
           JOIN event_state_destinations d
             ON d.state_guild_id = e.state_guild_id AND d.game_profile = e.game_profile
           JOIN state_event_phases p
             ON p.id = c.phase_id AND p.state_event_id = e.id AND p.game_profile = e.game_profile
           LEFT JOIN state_event_media m
             ON m.phase_id = p.id AND m.game_profile = p.game_profile
            AND m.delivery_kind = c.delivery_kind
          WHERE c.id = $1 AND c.game_profile = $2 AND c.status = 'claimed'
            AND c.claimed_by_bot_instance = $3 AND c.claimed_by_worker = $4`,
        [claimId, gameProfile, botInstanceName, workerId]
      )).rows[0]
      if (!row) return null
      return Object.freeze({
        claim: Object.freeze({
          id: String(row.claim_id),
          gameProfile,
          attemptCount: row.attempt_count,
          scheduleVersion: row.schedule_version,
          deliveryKind: row.delivery_kind,
          targetKind: row.target_kind,
          targetGuildId: row.target_guild_id,
          targetChannelId: row.target_channel_id,
          targetIsCurrent: row.target_is_current === true,
          occurrenceAt: cloneDate(row.occurrence_at),
          deliverAt: cloneDate(row.deliver_at)
        }),
        stateEvent: Object.freeze({
          id: String(row.state_event_id),
          stateGuildId: row.state_guild_id,
          stateNumber: row.state_number,
          eventName: row.event_name,
          recurrenceDays: row.recurrence_days
        }),
        phase: Object.freeze({
          id: String(row.phase_id),
          name: row.phase_name,
          phaseTimeUtc: row.phase_time_utc,
          preAlertMinutes: row.pre_alert_minutes,
          preAlertMessage: row.pre_alert_message,
          announceExact: row.announce_exact,
          exactMessage: row.exact_message,
          sortOrder: row.sort_order
        }),
        image: mapMedia(row)
      })
    },

    async markClaimSent({ claimId, botInstanceName, workerId, sentAt, sentMessageId = null }) {
      const result = await pool.query(
        `UPDATE state_event_delivery_claims
            SET status = 'sent', sent_at = $5, sent_message_id = $6,
                claimed_by_bot_instance = NULL, claimed_by_worker = NULL,
                claimed_at = NULL, claimed_until = NULL,
                last_error = NULL, next_attempt_at = NULL, updated_at = $5
          WHERE id = $1 AND game_profile = $2 AND status = 'claimed'
            AND claimed_by_bot_instance = $3 AND claimed_by_worker = $4`,
        [claimId, gameProfile, botInstanceName, workerId, sentAt, sentMessageId]
      )
      return result.rowCount === 1
    },

    async markClaimFailed({
      claimId,
      botInstanceName,
      workerId,
      failedAt,
      lastError,
      nextAttemptAt
    }) {
      const result = await pool.query(
        `UPDATE state_event_delivery_claims
            SET status = 'failed', last_error = $5, next_attempt_at = $6,
                claimed_by_bot_instance = NULL, claimed_by_worker = NULL,
                claimed_at = NULL, claimed_until = NULL, updated_at = $4
          WHERE id = $1 AND game_profile = $2 AND status = 'claimed'
            AND claimed_by_bot_instance = $3 AND claimed_by_worker = $7`,
        [claimId, gameProfile, botInstanceName, failedAt, lastError, nextAttemptAt, workerId]
      )
      return result.rowCount === 1
    },

    async markClaimPermanentlyFailed(input) {
      return repository.markClaimFailed({ ...input, nextAttemptAt: null })
    },

    async stateRoundupOccurrences({ stateGuildId, weekStart, weekEnd }) {
      const events = (await pool.query(
        `SELECT e.*, e.first_occurrence_date::text AS first_occurrence_date,
                d.state_number
           FROM state_events e
           JOIN event_state_destinations d
             ON d.state_guild_id = e.state_guild_id AND d.game_profile = e.game_profile
          WHERE e.state_guild_id = $1 AND e.game_profile = $2
            AND e.status = 'active' AND d.enabled = true
            AND e.first_occurrence_date <= ($3::timestamptz AT TIME ZONE 'UTC')::date
          ORDER BY e.id`,
        [stateGuildId, gameProfile, weekEnd]
      )).rows
      if (!events.length) return []
      const phases = await loadPhaseRows(pool, gameProfile, events.map(event => event.id))
      return buildStateRoundupOccurrences(attachPhases(events, phases), weekStart, weekEnd)
    }
  }
  return repository
}

module.exports = {
  createStateEventRepository
}
