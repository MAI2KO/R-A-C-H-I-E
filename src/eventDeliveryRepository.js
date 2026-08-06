const { assertProfile } = require("./eventSchedulerRepository")

const MAX_DELIVERY_ATTEMPTS = 5

function requirePool(pool) {
  if (!pool || typeof pool.query !== "function" || typeof pool.connect !== "function") {
    throw new Error("A transactional Postgres pool is required")
  }
}

function requireOwnership(value, label) {
  const normalized = String(value || "").trim()
  if (!normalized) throw new Error(`${label} is required`)
  return normalized
}

function cloneDate(value) {
  return value instanceof Date ? new Date(value) : new Date(String(value))
}

function mapClaimPayload(row) {
  if (!row) return null
  const image = row.image_data
    ? Object.freeze({
        originalFilename: row.image_filename,
        contentType: row.image_content_type,
        byteSize: row.image_byte_size,
        imageData: Buffer.from(row.image_data)
      })
    : null
  const group = row.group_id
    ? Object.freeze({
        id: String(row.group_id),
        name: row.group_name,
        eventTimeUtc: row.group_event_time_utc,
        sortOrder: row.group_sort_order
      })
    : null

  return Object.freeze({
    claim: Object.freeze({
      id: String(row.claim_id),
      gameProfile: row.game_profile,
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
    event: Object.freeze({
      id: String(row.event_id),
      guildId: row.guild_id,
      eventName: row.event_name,
      recurrenceDays: row.recurrence_days,
      advanceReminderMessage: row.advance_reminder_message,
      finalReminderMessage: row.final_reminder_message
    }),
    alliance: Object.freeze({
      id: row.alliance_id ? String(row.alliance_id) : null,
      name: row.alliance_name,
      guildId: row.guild_id
    }),
    group,
    image
  })
}

function createEventDeliveryRepository(pool, gameProfile, { targetKind = null } = {}) {
  requirePool(pool)
  assertProfile(gameProfile)
  if (targetKind !== null && targetKind !== "alliance") {
    throw new Error("Unsupported delivery target kind")
  }

  const repository = {
    gameProfile,
    targetKind,

    async listActiveEventDefinitions({ rangeEnd }) {
      const eventResult = await pool.query(
        `SELECT e.*, e.first_occurrence_date::text AS first_occurrence_date,
                a.alliance_name AS alliance_name,
                s.event_channel_id
           FROM scheduled_events e
           JOIN event_alliances a
             ON a.id = e.alliance_id AND a.guild_id = e.guild_id
            AND a.game_profile = e.game_profile
           JOIN event_guild_settings s
             ON s.guild_id = e.guild_id AND s.game_profile = e.game_profile
          WHERE e.game_profile = $1
            AND e.status = 'active'
            AND e.publish_to_alliance = true
            AND e.first_occurrence_date <= ($2::timestamptz AT TIME ZONE 'UTC')::date
          ORDER BY e.id`,
        [gameProfile, rangeEnd]
      )
      if (eventResult.rows.length === 0) return []

      const eventIds = eventResult.rows.map(event => event.id)
      const groupResult = await pool.query(
        `SELECT id AS group_id, event_id, group_name, event_time_utc, sort_order
           FROM scheduled_event_groups
          WHERE game_profile = $1 AND event_id = ANY($2::bigint[])
          ORDER BY event_id, sort_order, group_name, id`,
        [gameProfile, eventIds]
      )
      const groupsByEvent = new Map()
      for (const group of groupResult.rows) {
        const key = String(group.event_id)
        if (!groupsByEvent.has(key)) groupsByEvent.set(key, [])
        groupsByEvent.get(key).push(group)
      }
      return eventResult.rows.map(event => ({
        ...event,
        groups: groupsByEvent.get(String(event.id)) || []
      }))
    },

    async insertMissingDeliveryClaims(claims) {
      if (!Array.isArray(claims) || claims.length === 0) return 0
      const client = await pool.connect()
      let inserted = 0
      try {
        await client.query("BEGIN")
        for (const claim of claims) {
          if (claim.gameProfile !== gameProfile) {
            throw new Error("Delivery claim game profile does not match repository ownership")
          }
          if (claim.targetKind !== "alliance") {
            throw new Error("Individual state reminder claims are disabled")
          }
          if (targetKind !== null && claim.targetKind !== targetKind) {
            throw new Error("Delivery claim target does not match repository scope")
          }
          const result = await client.query(
            `INSERT INTO event_delivery_claims (
               event_id, group_id, group_id_snapshot, group_name_snapshot,
               game_profile, schedule_version, occurrence_at, deliver_at,
               delivery_kind, target_kind, target_guild_id, target_channel_id
             ) VALUES ($1, $2, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11)
             ON CONFLICT DO NOTHING
             RETURNING id`,
            [
              claim.eventId,
              claim.groupId,
              claim.groupName,
              gameProfile,
              claim.scheduleVersion || 1,
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

    async claimDueDeliveries({
      now,
      batchSize,
      leaseSeconds,
      botInstanceName,
      workerId
    }) {
      const bot = requireOwnership(botInstanceName, "Bot instance name")
      const worker = requireOwnership(workerId, "Worker ID")
      const client = await pool.connect()
      try {
        await client.query("BEGIN")
        await repository.reconcileInvalidEventClaims({ now, client })
        await client.query(
          `UPDATE event_delivery_claims
              SET status = 'failed',
                  claimed_by_bot_instance = NULL,
                  claimed_by_worker = NULL,
                  claimed_at = NULL,
                  claimed_until = NULL,
                  next_attempt_at = NULL,
                  last_error = 'Claim lease expired after maximum attempts.',
                  updated_at = $2
            WHERE game_profile = $1
              AND ($4::varchar IS NULL OR target_kind = $4)
              AND status = 'claimed'
              AND claimed_until <= $2
              AND attempt_count >= $3`,
          [gameProfile, now, MAX_DELIVERY_ATTEMPTS, targetKind]
        )
        const result = await client.query(
          `WITH claimable AS (
             SELECT d.id
               FROM event_delivery_claims d
              WHERE d.game_profile = $1
                AND ($8::varchar IS NULL OR d.target_kind = $8)
                AND d.attempt_count < $7
                AND EXISTS (
                  SELECT 1
                    FROM scheduled_events e
                    JOIN event_guild_settings s
                      ON s.guild_id = e.guild_id AND s.game_profile = e.game_profile
                   WHERE e.id = d.event_id AND e.game_profile = d.game_profile
                     AND e.status = 'active' AND e.schedule_version = d.schedule_version
                     AND d.target_kind = 'alliance'
                     AND e.publish_to_alliance = true
                     AND e.guild_id = d.target_guild_id
                     AND s.event_channel_id = d.target_channel_id
                     AND d.delivery_kind <> 'event_start'
                     AND (d.delivery_kind <> 'final_reminder' OR d.occurrence_at > $2)
                )
                AND (
                  (d.status = 'pending' AND d.deliver_at <= $2)
                  OR (d.status = 'claimed' AND d.claimed_until <= $2)
                  OR (d.status = 'failed' AND d.next_attempt_at <= $2)
                )
              ORDER BY COALESCE(d.next_attempt_at, d.claimed_until, d.deliver_at), d.id
              FOR UPDATE SKIP LOCKED
              LIMIT $3
           )
           UPDATE event_delivery_claims d
              SET status = 'claimed',
                  claimed_by_bot_instance = $4,
                  claimed_by_worker = $5,
                  claimed_at = $2,
                  claimed_until = $2 + ($6 * interval '1 second'),
                  attempt_count = d.attempt_count + 1,
                  next_attempt_at = NULL,
                  last_error = NULL,
                  updated_at = $2
             FROM claimable
            WHERE d.id = claimable.id AND d.game_profile = $1
           RETURNING d.*`,
          [
            gameProfile,
            now,
            batchSize,
            bot,
            worker,
            leaseSeconds,
            MAX_DELIVERY_ATTEMPTS,
            targetKind
          ]
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
      const result = await pool.query(
        `SELECT d.id AS claim_id, d.game_profile, d.attempt_count, d.schedule_version,
                d.delivery_kind, d.target_kind, d.target_guild_id,
                d.target_channel_id, d.occurrence_at, d.deliver_at,
                CASE
                  WHEN e.status = 'active' AND e.schedule_version = d.schedule_version
                   AND d.target_kind = 'alliance' AND e.publish_to_alliance = true
                   AND e.guild_id = d.target_guild_id AND s.event_channel_id = d.target_channel_id
                   AND d.delivery_kind <> 'event_start'
                   AND (d.delivery_kind <> 'final_reminder' OR d.occurrence_at > now())
                  THEN true
                  ELSE false
                END AS target_is_current,
                e.id AS event_id, e.guild_id, e.alliance_id,
                a.alliance_name, e.event_name, e.recurrence_days,
                e.advance_reminder_message, e.final_reminder_message,
                g.id AS group_id,
                COALESCE(g.group_name, d.group_name_snapshot) AS group_name,
                g.event_time_utc AS group_event_time_utc,
                g.sort_order AS group_sort_order,
                i.original_filename AS image_filename,
                i.content_type AS image_content_type,
                i.byte_size AS image_byte_size, i.image_data
           FROM event_delivery_claims d
           JOIN scheduled_events e
             ON e.id = d.event_id AND e.game_profile = d.game_profile
           JOIN event_alliances a
             ON a.id = e.alliance_id AND a.guild_id = e.guild_id
            AND a.game_profile = e.game_profile
           JOIN event_guild_settings s
             ON s.guild_id = e.guild_id AND s.game_profile = e.game_profile
           LEFT JOIN scheduled_event_groups g
             ON g.id = d.group_id
            AND g.event_id = d.event_id
            AND g.game_profile = d.game_profile
           LEFT JOIN scheduled_event_images i
             ON i.event_id = d.event_id AND i.game_profile = d.game_profile
            AND d.delivery_kind = 'advance_reminder'
          WHERE d.id = $1 AND d.game_profile = $2
            AND d.status = 'claimed'
            AND d.claimed_by_bot_instance = $3
            AND d.claimed_by_worker = $4
            AND ($5::varchar IS NULL OR d.target_kind = $5)`,
        [claimId, gameProfile, botInstanceName, workerId, targetKind]
      )
      return mapClaimPayload(result.rows[0])
    },

    async reconcileInvalidEventClaims({ now, client = pool }) {
      const result = await client.query(
        `UPDATE event_delivery_claims d
            SET status = 'failed',
                claimed_by_bot_instance = NULL,
                claimed_by_worker = NULL,
                claimed_at = NULL,
                claimed_until = NULL,
                next_attempt_at = NULL,
                last_error = CASE
                  WHEN d.target_kind = 'state'
                    THEN 'Individual state reminders are disabled.'
                  WHEN d.delivery_kind = 'event_start'
                    THEN 'Exact-start reminders are disabled.'
                  WHEN d.delivery_kind = 'final_reminder' AND d.occurrence_at <= $2
                    THEN 'Final reminder delivery window has passed.'
                  ELSE 'Event schedule or delivery target changed.'
                END,
                updated_at = $2
           FROM scheduled_events e
           JOIN event_guild_settings s
             ON s.guild_id = e.guild_id AND s.game_profile = e.game_profile
          WHERE d.event_id = e.id
            AND d.game_profile = e.game_profile
            AND d.game_profile = $1
            AND (
              d.status = 'pending'
              OR (d.status = 'failed' AND d.next_attempt_at IS NOT NULL)
              OR (d.status = 'claimed' AND d.claimed_until <= $2)
            )
            AND NOT COALESCE((
              e.status = 'active' AND e.schedule_version = d.schedule_version
              AND d.target_kind = 'alliance' AND e.publish_to_alliance = true
              AND e.guild_id = d.target_guild_id AND s.event_channel_id = d.target_channel_id
              AND d.delivery_kind <> 'event_start'
              AND (d.delivery_kind <> 'final_reminder' OR d.occurrence_at > $2)
            ), false)
          RETURNING d.id`,
        [gameProfile, now]
      )
      return result.rowCount
    },

    async reconcileInvalidStateClaims(input) {
      return repository.reconcileInvalidEventClaims(input)
    },

    async markClaimSent({ claimId, botInstanceName, workerId, sentAt, sentMessageId = null }) {
      const result = await pool.query(
        `UPDATE event_delivery_claims
            SET status = 'sent', sent_at = $5, sent_message_id = $6,
                claimed_by_bot_instance = NULL, claimed_by_worker = NULL,
                claimed_at = NULL, claimed_until = NULL,
                last_error = NULL, next_attempt_at = NULL, updated_at = $5
          WHERE id = $1 AND game_profile = $2 AND status = 'claimed'
            AND claimed_by_bot_instance = $3 AND claimed_by_worker = $4
            AND ($7::varchar IS NULL OR target_kind = $7)`,
        [claimId, gameProfile, botInstanceName, workerId, sentAt, sentMessageId, targetKind]
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
        `UPDATE event_delivery_claims
            SET status = 'failed', last_error = $5, next_attempt_at = $6,
                claimed_by_bot_instance = NULL, claimed_by_worker = NULL,
                claimed_at = NULL, claimed_until = NULL, updated_at = $4
          WHERE id = $1 AND game_profile = $2 AND status = 'claimed'
            AND claimed_by_bot_instance = $3 AND claimed_by_worker = $7
            AND ($8::varchar IS NULL OR target_kind = $8)`,
        [
          claimId,
          gameProfile,
          botInstanceName,
          failedAt,
          lastError,
          nextAttemptAt,
          workerId,
          targetKind
        ]
      )
      return result.rowCount === 1
    },

    async markClaimPermanentlyFailed(input) {
      return repository.markClaimFailed({ ...input, nextAttemptAt: null })
    }
  }
  return repository
}

module.exports = {
  MAX_DELIVERY_ATTEMPTS,
  mapClaimPayload,
  createEventDeliveryRepository
}
