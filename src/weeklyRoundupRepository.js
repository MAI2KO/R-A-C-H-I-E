const { assertProfile } = require("./eventSchedulerRepository")
const { buildRoundupOccurrences } = require("./weeklyRoundupCalculation")
const { MAX_DELIVERY_ATTEMPTS } = require("./eventDeliveryRepository")

function requirePool(pool) {
  if (!pool || typeof pool.query !== "function" || typeof pool.connect !== "function") {
    throw new Error("A transactional Postgres pool is required")
  }
}

function requireOwner(value, label) {
  const result = String(value || "").trim()
  if (!result) throw new Error(`${label} is required`)
  return result
}

function utcDate(value) {
  return new Date(`${String(value).slice(0, 10)}T00:00:00Z`)
}

function createWeeklyRoundupRepository(pool, gameProfile) {
  requirePool(pool)
  assertProfile(gameProfile)

  const repository = {
    gameProfile,

    async listRoundupConfigurations() {
      const result = await pool.query(
        `SELECT s.guild_id AS source_guild_id, s.game_profile, s.alliance_name,
                s.weekly_roundup_day, s.weekly_roundup_time_utc,
                s.weekly_roundup_channel_id, s.roundup_when_empty,
                s.weekly_roundup_enabled, s.state_roundup_enabled,
                s.weekly_roundup_not_before,
                l.state_guild_id, l.state_event_channel_id, l.sharing_enabled
           FROM event_guild_settings s
           LEFT JOIN event_state_links l
             ON l.alliance_guild_id = s.guild_id AND l.game_profile = s.game_profile
          WHERE s.game_profile = $1
            AND (s.weekly_roundup_enabled = true OR s.state_roundup_enabled = true)
          ORDER BY s.guild_id`,
        [gameProfile]
      )
      return result.rows
    },

    async insertMissingClaims(claims) {
      if (!Array.isArray(claims) || claims.length === 0) return 0
      const client = await pool.connect()
      let inserted = 0
      try {
        await client.query("BEGIN")
        for (const claim of claims) {
          if (claim.gameProfile !== gameProfile) throw new Error("Roundup profile mismatch")
          const result = await client.query(
            `INSERT INTO weekly_roundup_claims (
               week_start_date, game_profile, target_kind, target_guild_id,
               target_channel_id, source_guild_id, scheduled_for, post_when_empty
             ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8)
             ON CONFLICT DO NOTHING
             RETURNING id`,
            [
              claim.weekStartDate,
              gameProfile,
              claim.targetKind,
              claim.targetGuildId,
              claim.targetChannelId,
              claim.sourceGuildId,
              claim.scheduledFor,
              claim.postWhenEmpty
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
        `UPDATE weekly_roundup_claims r
            SET status = 'failed', next_attempt_at = NULL,
                claimed_by_bot_instance = NULL, claimed_by_worker = NULL,
                claimed_at = NULL, claimed_until = NULL,
                last_error = 'Weekly roundup is disabled or target changed.',
                updated_at = $2
          WHERE r.game_profile = $1
            AND (r.status IN ('pending', 'failed')
                 OR (r.status = 'claimed' AND r.claimed_until <= $2))
            AND NOT (
              (r.target_kind = 'alliance' AND EXISTS (
                SELECT 1 FROM event_guild_settings s
                 WHERE s.guild_id = r.source_guild_id
                   AND s.game_profile = r.game_profile
                   AND s.weekly_roundup_enabled = true
                   AND s.guild_id = r.target_guild_id
                   AND s.weekly_roundup_channel_id = r.target_channel_id
              ))
              OR
              (r.target_kind = 'state' AND EXISTS (
                SELECT 1
                  FROM event_guild_settings s
                  JOIN event_state_links l
                    ON l.alliance_guild_id = s.guild_id
                   AND l.game_profile = s.game_profile
                 WHERE s.game_profile = r.game_profile
                   AND s.state_roundup_enabled = true
                   AND l.sharing_enabled = true
                   AND l.state_guild_id = r.target_guild_id
                   AND l.state_event_channel_id = r.target_channel_id
              ))
            )
          RETURNING r.id`,
        [gameProfile, now]
      )
      return result.rowCount
    },

    async claimDue({ now, batchSize, leaseSeconds, botInstanceName, workerId }) {
      const bot = requireOwner(botInstanceName, "Bot instance name")
      const worker = requireOwner(workerId, "Worker ID")
      const client = await pool.connect()
      try {
        await client.query("BEGIN")
        await repository.reconcileInvalidClaims({ now, client })
        await client.query(
          `UPDATE weekly_roundup_claims
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
             SELECT id FROM weekly_roundup_claims
              WHERE game_profile = $1 AND attempt_count < $7
                AND (
                  (status = 'pending' AND scheduled_for <= $2)
                  OR (status = 'failed' AND next_attempt_at <= $2)
                  OR (status = 'claimed' AND claimed_until <= $2)
                )
              ORDER BY COALESCE(next_attempt_at, claimed_until, scheduled_for), id
              FOR UPDATE SKIP LOCKED
              LIMIT $3
           )
           UPDATE weekly_roundup_claims r
              SET status = 'claimed', claimed_by_bot_instance = $4,
                  claimed_by_worker = $5, claimed_at = $2,
                  claimed_until = $2 + ($6 * interval '1 second'),
                  attempt_count = r.attempt_count + 1,
                  next_attempt_at = NULL, last_error = NULL, updated_at = $2
             FROM claimable
            WHERE r.id = claimable.id AND r.game_profile = $1
           RETURNING r.*`,
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
      const claimResult = await pool.query(
        `SELECT r.*,
                r.week_start_date::text AS week_start_text,
                CASE
                  WHEN r.target_kind = 'alliance' THEN EXISTS (
                    SELECT 1 FROM event_guild_settings s
                     WHERE s.guild_id = r.source_guild_id AND s.game_profile = r.game_profile
                       AND s.weekly_roundup_enabled = true
                       AND s.guild_id = r.target_guild_id
                       AND s.weekly_roundup_channel_id = r.target_channel_id
                  )
                  ELSE EXISTS (
                    SELECT 1 FROM event_guild_settings s
                    JOIN event_state_links l
                      ON l.alliance_guild_id = s.guild_id AND l.game_profile = s.game_profile
                     WHERE s.game_profile = r.game_profile
                       AND s.state_roundup_enabled = true AND l.sharing_enabled = true
                       AND l.state_guild_id = r.target_guild_id
                       AND l.state_event_channel_id = r.target_channel_id
                  )
                END AS target_is_current,
                s.alliance_name
           FROM weekly_roundup_claims r
           LEFT JOIN event_guild_settings s
             ON s.guild_id = r.source_guild_id AND s.game_profile = r.game_profile
          WHERE r.id = $1 AND r.game_profile = $2 AND r.status = 'claimed'
            AND r.claimed_by_bot_instance = $3 AND r.claimed_by_worker = $4`,
        [claimId, gameProfile, botInstanceName, workerId]
      )
      const row = claimResult.rows[0]
      if (!row) return null

      const eventResult = await pool.query(
        `SELECT DISTINCT e.*, a.alliance_name AS alliance_name,
                e.first_occurrence_date::text AS first_occurrence_date
           FROM scheduled_events e
           JOIN event_alliances a
             ON a.id = e.alliance_id AND a.guild_id = e.guild_id
            AND a.game_profile = e.game_profile
           ${row.target_kind === "state" ? `JOIN event_state_links l
             ON l.alliance_guild_id = e.guild_id AND l.game_profile = e.game_profile` : ""}
          WHERE e.game_profile = $1 AND e.status = 'active'
            AND e.include_in_weekly_roundup = true
            ${row.target_kind === "alliance"
              ? "AND e.guild_id = $2"
              : `AND e.publish_to_state = true AND l.sharing_enabled = true
                 AND l.state_guild_id = $2 AND l.state_event_channel_id = $3`}
          ORDER BY e.id`,
        row.target_kind === "alliance"
          ? [gameProfile, row.source_guild_id]
          : [gameProfile, row.target_guild_id, row.target_channel_id]
      )
      const eventIds = eventResult.rows.map(event => event.id)
      let groups = []
      if (eventIds.length) {
        groups = (await pool.query(
          `SELECT id AS group_id, event_id, group_name, event_time_utc, sort_order
             FROM scheduled_event_groups
            WHERE game_profile = $1 AND event_id = ANY($2::bigint[])
            ORDER BY event_id, sort_order, group_name, id`,
          [gameProfile, eventIds]
        )).rows
      }
      const groupsByEvent = new Map()
      for (const group of groups) {
        const key = String(group.event_id)
        if (!groupsByEvent.has(key)) groupsByEvent.set(key, [])
        groupsByEvent.get(key).push(group)
      }
      const events = eventResult.rows.map(event => ({
        ...event,
        groups: groupsByEvent.get(String(event.id)) || []
      }))
      const weekStart = utcDate(row.week_start_text)
      const weekEnd = new Date(weekStart.getTime() + 7 * 86400000)
      const messages = (await pool.query(
        `SELECT message_index, sent_message_id, payload_hash FROM weekly_roundup_messages
          WHERE roundup_claim_id = $1 ORDER BY message_index`,
        [row.id]
      )).rows
      return {
        claim: {
          id: String(row.id),
          gameProfile,
          attemptCount: row.attempt_count,
          targetKind: row.target_kind,
          targetGuildId: row.target_guild_id,
          targetChannelId: row.target_channel_id,
          targetIsCurrent: row.target_is_current === true,
          weekStart,
          weekEnd,
          postWhenEmpty: row.post_when_empty
        },
        allianceName: row.alliance_name,
        occurrences: buildRoundupOccurrences(events, weekStart, weekEnd),
        sentMessages: new Map(messages.map(message => [message.message_index, {
          sentMessageId: message.sent_message_id,
          payloadHash: message.payload_hash
        }]))
      }
    },

    async renewLease({ claimId, botInstanceName, workerId, now, leaseSeconds }) {
      const result = await pool.query(
        `UPDATE weekly_roundup_claims
            SET claimed_until = $5 + ($6 * interval '1 second'), updated_at = $5
          WHERE id = $1 AND game_profile = $2 AND status = 'claimed'
            AND claimed_by_bot_instance = $3 AND claimed_by_worker = $4
            AND claimed_until > $5`,
        [claimId, gameProfile, botInstanceName, workerId, now, leaseSeconds]
      )
      return result.rowCount === 1
    },

    async setPartCount({ claimId, botInstanceName, workerId, partCount }) {
      const result = await pool.query(
        `UPDATE weekly_roundup_claims SET part_count = $5, updated_at = now()
          WHERE id = $1 AND game_profile = $2 AND status = 'claimed'
            AND claimed_by_bot_instance = $3 AND claimed_by_worker = $4
            AND (part_count IS NULL OR part_count = $5)`,
        [claimId, gameProfile, botInstanceName, workerId, partCount]
      )
      return result.rowCount === 1
    },

    async recordSentMessage({
      claimId,
      botInstanceName,
      workerId,
      messageIndex,
      sentMessageId,
      payloadHash
    }) {
      const client = await pool.connect()
      try {
        await client.query("BEGIN")
        const owned = await client.query(
          `SELECT 1 FROM weekly_roundup_claims
            WHERE id = $1 AND game_profile = $2 AND status = 'claimed'
              AND claimed_by_bot_instance = $3 AND claimed_by_worker = $4
            FOR UPDATE`,
          [claimId, gameProfile, botInstanceName, workerId]
        )
        if (!owned.rowCount) {
          await client.query("ROLLBACK")
          return false
        }
        await client.query(
          `INSERT INTO weekly_roundup_messages
             (roundup_claim_id, message_index, sent_message_id, payload_hash)
           VALUES ($1, $2, $3, $4)
           ON CONFLICT (roundup_claim_id, message_index) DO NOTHING`,
          [claimId, messageIndex, sentMessageId, payloadHash]
        )
        await client.query("COMMIT")
        return true
      } catch (error) {
        await client.query("ROLLBACK")
        throw error
      } finally {
        client.release()
      }
    },

    async markSent({ claimId, botInstanceName, workerId, sentAt }) {
      const result = await pool.query(
        `UPDATE weekly_roundup_claims r
            SET status = 'sent', sent_at = $5,
                sent_message_id = (
                  SELECT sent_message_id FROM weekly_roundup_messages
                   WHERE roundup_claim_id = r.id ORDER BY message_index LIMIT 1
                ),
                claimed_by_bot_instance = NULL, claimed_by_worker = NULL,
                claimed_at = NULL, claimed_until = NULL,
                next_attempt_at = NULL, last_error = NULL, updated_at = $5
          WHERE r.id = $1 AND r.game_profile = $2 AND r.status = 'claimed'
            AND r.claimed_by_bot_instance = $3 AND r.claimed_by_worker = $4
            AND r.part_count IS NOT NULL
            AND r.part_count = (
              SELECT count(*)::integer FROM weekly_roundup_messages
               WHERE roundup_claim_id = r.id
            )`,
        [claimId, gameProfile, botInstanceName, workerId, sentAt]
      )
      return result.rowCount === 1
    },

    async markFailed({ claimId, botInstanceName, workerId, failedAt, lastError, nextAttemptAt }) {
      const result = await pool.query(
        `UPDATE weekly_roundup_claims
            SET status = 'failed', last_error = $5, next_attempt_at = $6,
                claimed_by_bot_instance = NULL, claimed_by_worker = NULL,
                claimed_at = NULL, claimed_until = NULL, updated_at = $4
          WHERE id = $1 AND game_profile = $2 AND status = 'claimed'
            AND claimed_by_bot_instance = $3 AND claimed_by_worker = $7`,
        [claimId, gameProfile, botInstanceName, failedAt, lastError, nextAttemptAt, workerId]
      )
      return result.rowCount === 1
    },

    async markPermanentlyFailed(input) {
      return repository.markFailed({ ...input, nextAttemptAt: null })
    }
  }
  return repository
}

module.exports = { createWeeklyRoundupRepository }
