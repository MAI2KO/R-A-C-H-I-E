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
      deliveryKind: row.delivery_kind,
      targetKind: row.target_kind,
      targetGuildId: row.target_guild_id,
      targetChannelId: row.target_channel_id,
      occurrenceAt: cloneDate(row.occurrence_at),
      deliverAt: cloneDate(row.deliver_at)
    }),
    event: Object.freeze({
      id: String(row.event_id),
      guildId: row.guild_id,
      eventName: row.event_name,
      recurrenceDays: row.recurrence_days
    }),
    alliance: Object.freeze({
      name: row.alliance_name,
      guildId: row.guild_id
    }),
    group,
    image
  })
}

function createEventDeliveryRepository(pool, gameProfile) {
  requirePool(pool)
  assertProfile(gameProfile)

  const repository = {
    gameProfile,

    async listActiveEventDefinitions({ rangeEnd }) {
      const eventResult = await pool.query(
        `SELECT e.*, s.event_channel_id,
                i.original_filename AS image_filename,
                i.content_type AS image_content_type,
                i.byte_size AS image_byte_size
           FROM scheduled_events e
           JOIN event_guild_settings s
             ON s.guild_id = e.guild_id AND s.game_profile = e.game_profile
           LEFT JOIN scheduled_event_images i
             ON i.event_id = e.id AND i.game_profile = e.game_profile
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
          const result = await client.query(
            `INSERT INTO event_delivery_claims (
               event_id, group_id, game_profile, occurrence_at, deliver_at,
               delivery_kind, target_kind, target_guild_id, target_channel_id
             ) VALUES ($1, $2, $3, $4, $5, $6, 'alliance', $7, $8)
             ON CONFLICT DO NOTHING
             RETURNING id`,
            [
              claim.eventId,
              claim.groupId,
              gameProfile,
              claim.occurrenceAt,
              claim.deliverAt,
              claim.deliveryKind,
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
              AND status = 'claimed'
              AND claimed_until <= $2
              AND attempt_count >= $3`,
          [gameProfile, now, MAX_DELIVERY_ATTEMPTS]
        )
        const result = await client.query(
          `WITH claimable AS (
             SELECT id
               FROM event_delivery_claims
              WHERE game_profile = $1
                AND attempt_count < $7
                AND (
                  (status = 'pending' AND deliver_at <= $2)
                  OR (status = 'claimed' AND claimed_until <= $2)
                  OR (status = 'failed' AND next_attempt_at <= $2)
                )
              ORDER BY COALESCE(next_attempt_at, claimed_until, deliver_at), id
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
            MAX_DELIVERY_ATTEMPTS
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
        `SELECT d.id AS claim_id, d.game_profile, d.attempt_count,
                d.delivery_kind, d.target_kind, d.target_guild_id,
                d.target_channel_id, d.occurrence_at, d.deliver_at,
                e.id AS event_id, e.guild_id, e.alliance_name, e.event_name,
                e.recurrence_days,
                g.id AS group_id, g.group_name,
                g.event_time_utc AS group_event_time_utc,
                g.sort_order AS group_sort_order,
                i.original_filename AS image_filename,
                i.content_type AS image_content_type,
                i.byte_size AS image_byte_size, i.image_data
           FROM event_delivery_claims d
           JOIN scheduled_events e
             ON e.id = d.event_id AND e.game_profile = d.game_profile
           LEFT JOIN scheduled_event_groups g
             ON g.id = d.group_id
            AND g.event_id = d.event_id
            AND g.game_profile = d.game_profile
           LEFT JOIN scheduled_event_images i
             ON i.event_id = d.event_id AND i.game_profile = d.game_profile
          WHERE d.id = $1 AND d.game_profile = $2
            AND d.status = 'claimed'
            AND d.claimed_by_bot_instance = $3
            AND d.claimed_by_worker = $4`,
        [claimId, gameProfile, botInstanceName, workerId]
      )
      return mapClaimPayload(result.rows[0])
    },

    async markClaimSent({ claimId, botInstanceName, workerId, sentAt, sentMessageId = null }) {
      const result = await pool.query(
        `UPDATE event_delivery_claims
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
        `UPDATE event_delivery_claims
            SET status = 'failed', last_error = $5, next_attempt_at = $6,
                claimed_by_bot_instance = NULL, claimed_by_worker = NULL,
                claimed_at = NULL, claimed_until = NULL, updated_at = $4
          WHERE id = $1 AND game_profile = $2 AND status = 'claimed'
            AND claimed_by_bot_instance = $3 AND claimed_by_worker = $7`,
        [
          claimId,
          gameProfile,
          botInstanceName,
          failedAt,
          lastError,
          nextAttemptAt,
          workerId
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
