const crypto = require("node:crypto")

function createGiftCodeCommunityRepository(pool, gameProfile) {
  if (!pool?.query || !pool?.connect) throw new Error("A transactional PostgreSQL pool is required")
  if (!["wos", "kingshot"].includes(gameProfile)) throw new Error("Unsupported game profile")

  return Object.freeze({
    gameProfile,

    async getSettings(guildId) {
      return (await pool.query(
        `SELECT * FROM gift_code_guild_settings
          WHERE game_profile = $1 AND guild_id = $2`,
        [gameProfile, guildId]
      )).rows[0] || null
    },

    async setChannel(guildId, channelId) {
      return (await pool.query(
        `INSERT INTO gift_code_guild_settings (
           game_profile, guild_id, gift_code_channel_id
         ) VALUES ($1, $2, $3)
         ON CONFLICT (game_profile, guild_id) DO UPDATE
           SET gift_code_channel_id = EXCLUDED.gift_code_channel_id,
               updated_at_utc = now()
         RETURNING *`,
        [gameProfile, guildId, channelId]
      )).rows[0]
    },

    async prepareCodeEngagement(giftCodeId, queuedCount) {
      const client = await pool.connect()
      try {
        await client.query("BEGIN")
        const attribution = (await client.query(
          `SELECT s.submitted_by_discord_user_id,
                  s.raw_metadata->>'guildId' AS guild_id
             FROM gift_code_submissions s
            WHERE s.game_profile = $1
              AND (s.duplicate_of_gift_code_id IS NULL)
              AND s.submitted_by_discord_user_id IS NOT NULL
              AND s.raw_metadata->>'guildId' ~ '^[0-9]{1,32}$'
              AND EXISTS (
                SELECT 1 FROM gift_codes g
                 WHERE g.id = $2 AND g.game_profile = s.game_profile
                   AND g.code = s.submitted_code AND g.status = 'active'
              )
            ORDER BY s.received_at_utc, s.id
            LIMIT 1`,
          [gameProfile, giftCodeId]
        )).rows[0]
        if (!attribution) {
          await client.query("COMMIT")
          return []
        }
        const settings = (await client.query(
          `INSERT INTO gift_code_guild_settings (game_profile, guild_id)
           VALUES ($1, $2)
           ON CONFLICT (game_profile, guild_id) DO UPDATE
             SET updated_at_utc = gift_code_guild_settings.updated_at_utc
           RETURNING *`,
          [gameProfile, attribution.guild_id]
        )).rows[0]
        if (!settings) {
          await client.query("COMMIT")
          return []
        }
        const relevantQueuedCount = (await client.query(
          `SELECT COUNT(DISTINCT r.id)::integer AS count
             FROM gift_code_redemptions r
             JOIN player_accounts a
               ON a.id = r.player_account_id AND a.game_profile = r.game_profile
             JOIN player_account_guilds ag
               ON ag.player_account_id = a.id AND ag.game_profile = a.game_profile
              AND ag.guild_id = $3 AND ag.gift_code_enrolled = true
            WHERE r.game_profile = $1 AND r.gift_code_id = $2
              AND a.is_active = true AND a.gift_redemption_enabled = true
              AND r.status IN ('queued', 'claimed', 'rate_limited', 'temporary_error')`,
          [gameProfile, giftCodeId, attribution.guild_id]
        )).rows[0].count
        const events = []
        if (settings.announcements_enabled && settings.gift_code_channel_id) {
          const event = (await client.query(
             `INSERT INTO gift_code_engagement_events (
               id, game_profile, guild_id, event_type, gift_code_id,
               discord_user_id, progress_remaining, metadata
             ) VALUES ($1, $2, $3, 'code_progress', $4, $5, $6, $7)
             ON CONFLICT DO NOTHING RETURNING *`,
            [
              crypto.randomUUID(), gameProfile, attribution.guild_id,
              giftCodeId, attribution.submitted_by_discord_user_id,
              relevantQueuedCount,
              { queuedCount: relevantQueuedCount }
            ]
          )).rows[0]
          if (event) events.push(event)
        }
        const roleEvent = (await client.query(
          `INSERT INTO gift_code_engagement_events (
             id, game_profile, guild_id, event_type, gift_code_id,
             discord_user_id
           ) VALUES ($1, $2, $3, 'contributor_role', $4, $5)
           ON CONFLICT DO NOTHING RETURNING *`,
          [
            crypto.randomUUID(), gameProfile, attribution.guild_id,
            giftCodeId, attribution.submitted_by_discord_user_id
          ]
        )).rows[0]
        if (roleEvent) events.push(roleEvent)
        await client.query("COMMIT")
        return events
      } catch (error) {
        await client.query("ROLLBACK")
        throw error
      } finally {
        client.release()
      }
    },

    async claimVerificationResultNotification(workerId, now, leaseSeconds = 60, giftCodeId = null) {
      return (await pool.query(
        `WITH due AS (
           SELECT s.id
             FROM gift_code_submissions s
            WHERE s.game_profile = $1
              AND s.duplicate_of_gift_code_id IS NULL
              AND s.submitted_by_discord_user_id IS NOT NULL
              AND ($5::uuid IS NULL OR EXISTS (
                SELECT 1 FROM gift_codes g
                 WHERE g.id = $5 AND g.game_profile = s.game_profile
                   AND g.code = s.submitted_code
              ))
              AND (
                s.raw_metadata #>> '{verificationResultNotification,status}' = 'pending'
                OR (
                  s.raw_metadata #>> '{verificationResultNotification,status}' = 'failed'
                  AND (s.raw_metadata #>> '{verificationResultNotification,nextAttemptAt}')::timestamptz <= $2
                )
                OR (
                  s.raw_metadata #>> '{verificationResultNotification,status}' = 'sending'
                  AND (s.raw_metadata #>> '{verificationResultNotification,claimedUntil}')::timestamptz <= $2
                )
              )
            ORDER BY s.received_at_utc, s.id
            FOR UPDATE SKIP LOCKED
            LIMIT 1
         )
         UPDATE gift_code_submissions s
            SET raw_metadata = jsonb_set(
              s.raw_metadata,
              '{verificationResultNotification}',
              (s.raw_metadata->'verificationResultNotification') || jsonb_build_object(
                'status', 'sending',
                'claimedBy', $3::varchar,
                'claimedAt', $2::timestamptz,
                'claimedUntil', $2::timestamptz + ($4 * interval '1 second'),
                'attemptCount', COALESCE(
                  (s.raw_metadata #>> '{verificationResultNotification,attemptCount}')::integer,
                  0
                ) + 1
              ),
              true
            )
           FROM due
          WHERE s.id = due.id AND s.game_profile = $1
         RETURNING s.id, s.submitted_code, s.submitted_by_discord_user_id,
                   s.raw_metadata #>> '{verificationResultNotification,classification}' AS classification`,
        [gameProfile, now, workerId, leaseSeconds, giftCodeId]
      )).rows[0] || null
    },

    async finishVerificationResultNotification(submissionId, workerId, {
      sent,
      now = new Date(),
      errorCode = null
    }) {
      const status = sent ? "sent" : "failed"
      const retryAt = sent ? null : new Date(now.getTime() + 5 * 60 * 1000)
      return (await pool.query(
        `UPDATE gift_code_submissions s
            SET raw_metadata = jsonb_set(
              s.raw_metadata,
              '{verificationResultNotification}',
              (s.raw_metadata->'verificationResultNotification') || jsonb_build_object(
                'status', $4::varchar,
                'claimedBy', NULL,
                'claimedAt', NULL,
                'claimedUntil', NULL,
                'sentAt', CASE WHEN $5::boolean THEN to_jsonb($6::timestamptz) ELSE 'null'::jsonb END,
                'lastError', $7::varchar,
                'nextAttemptAt', to_jsonb($8::timestamptz)
              ),
              true
            )
          WHERE s.id = $1 AND s.game_profile = $2
            AND s.raw_metadata #>> '{verificationResultNotification,status}' = 'sending'
            AND s.raw_metadata #>> '{verificationResultNotification,claimedBy}' = $3
          RETURNING s.*`,
        [submissionId, gameProfile, workerId, status, Boolean(sent), now, errorCode, retryAt]
      )).rows[0] || null
    },

    async claimContributorRoleProvision(guildId, workerId, now, leaseSeconds = 60) {
      await pool.query(
        `INSERT INTO gift_code_guild_settings (game_profile, guild_id)
         VALUES ($1, $2) ON CONFLICT DO NOTHING`,
        [gameProfile, guildId]
      )
      return (await pool.query(
        `UPDATE gift_code_guild_settings
            SET contributor_role_status = 'claiming',
                contributor_role_claimed_by = $3,
                contributor_role_claimed_until_utc = $4::timestamptz + ($5 * interval '1 second'),
                updated_at_utc = $4::timestamptz
          WHERE game_profile = $1 AND guild_id = $2
            AND (
              contributor_role_claimed_until_utc IS NULL
              OR contributor_role_claimed_until_utc <= $4::timestamptz
              OR contributor_role_claimed_by = $3
            )
          RETURNING *`,
        [gameProfile, guildId, workerId, now, leaseSeconds]
      )).rows[0] || null
    },

    async completeContributorRoleProvision(guildId, workerId, roleId, now) {
      return (await pool.query(
        `UPDATE gift_code_guild_settings
            SET contributor_role_id = $4,
                contributor_role_status = 'ready',
                contributor_role_last_error = NULL,
                contributor_role_claimed_by = NULL,
                contributor_role_claimed_until_utc = NULL,
                updated_at_utc = $5
          WHERE game_profile = $1 AND guild_id = $2
            AND contributor_role_claimed_by = $3
          RETURNING *`,
        [gameProfile, guildId, workerId, roleId, now]
      )).rows[0] || null
    },

    async failContributorRoleProvision(guildId, workerId, errorCode, now) {
      return (await pool.query(
        `UPDATE gift_code_guild_settings
            SET contributor_role_status = 'error',
                contributor_role_last_error = $4,
                contributor_role_claimed_by = NULL,
                contributor_role_claimed_until_utc = $5::timestamptz + interval '5 minutes',
                updated_at_utc = $5
          WHERE game_profile = $1 AND guild_id = $2
            AND contributor_role_claimed_by = $3
          RETURNING *`,
        [gameProfile, guildId, workerId, String(errorCode).slice(0, 200), now]
      )).rows[0] || null
    },

    async markContributorRoleUnavailable(guildId, errorCode, now = new Date()) {
      return (await pool.query(
        `UPDATE gift_code_guild_settings
            SET contributor_role_status = 'error',
                contributor_role_last_error = $3,
                contributor_role_claimed_by = NULL,
                contributor_role_claimed_until_utc = $4::timestamptz + interval '5 minutes',
                updated_at_utc = $4
          WHERE game_profile = $1 AND guild_id = $2
          RETURNING *`,
        [gameProfile, guildId, String(errorCode).slice(0, 200), now]
      )).rows[0] || null
    },

    async claimEvent(eventId, workerId, now, leaseSeconds = 60) {
      return (await pool.query(
        `UPDATE gift_code_engagement_events
            SET status = 'claimed', claimed_by_worker = $3,
                claimed_at_utc = $4,
                claimed_until_utc = $4 + ($5 * interval '1 second'),
                attempt_count = attempt_count + 1, updated_at_utc = $4
          WHERE id = $1 AND game_profile = $2
            AND (
              status = 'pending'
              OR (status = 'claimed' AND claimed_until_utc <= $4)
              OR (status = 'failed' AND next_attempt_at_utc <= $4)
            )
          RETURNING *`,
        [eventId, gameProfile, workerId, now, leaseSeconds]
      )).rows[0] || null
    },

    async claimNextPending(workerId, now, leaseSeconds = 60) {
      return (await pool.query(
        `WITH due AS (
           SELECT id FROM gift_code_engagement_events
            WHERE game_profile = $1
              AND (
                status = 'pending'
                OR (status = 'claimed' AND claimed_until_utc <= $2)
                OR (status = 'failed' AND next_attempt_at_utc <= $2)
              )
            ORDER BY created_at_utc, id
            FOR UPDATE SKIP LOCKED LIMIT 1
         )
         UPDATE gift_code_engagement_events e
            SET status = 'claimed', claimed_by_worker = $3,
                claimed_at_utc = $2,
                claimed_until_utc = $2 + ($4 * interval '1 second'),
                attempt_count = attempt_count + 1, updated_at_utc = $2
           FROM due WHERE e.id = due.id AND e.game_profile = $1
         RETURNING e.*`,
        [gameProfile, now, workerId, leaseSeconds]
      )).rows[0] || null
    },

    async completeEvent(eventId, workerId, {
      channelId = null,
      messageId = null,
      progressCount = 0,
      finalized = false,
      progress = null,
      metadata = {},
      now = new Date()
    } = {}) {
      return (await pool.query(
        `UPDATE gift_code_engagement_events
            SET status = 'completed', channel_id = COALESCE($4, channel_id),
                message_id = COALESCE($5, message_id),
                last_progress_count = GREATEST(last_progress_count, $6),
                progress_successful = GREATEST(progress_successful, COALESCE($10, 0)),
                progress_already_redeemed = GREATEST(progress_already_redeemed, COALESCE($11, 0)),
                progress_account_issues = GREATEST(progress_account_issues, COALESCE($12, 0)),
                progress_restricted = GREATEST(progress_restricted, COALESCE($13, 0)),
                progress_remaining = LEAST(progress_remaining, COALESCE($14, progress_remaining)),
                last_update_at_utc = $8,
                finalized_at_utc = CASE WHEN $7 THEN $8 ELSE finalized_at_utc END,
                metadata = metadata || $9::jsonb,
                claimed_by_worker = NULL, claimed_at_utc = NULL,
                claimed_until_utc = NULL, last_error = NULL,
                next_attempt_at_utc = NULL,
                updated_at_utc = $8
          WHERE id = $1 AND game_profile = $2 AND claimed_by_worker = $3
          RETURNING *`,
        [
          eventId, gameProfile, workerId, channelId, messageId,
          progressCount, Boolean(finalized), now, metadata,
          progress?.successful ?? null,
          progress?.already_redeemed ?? null,
          progress?.account_issues ?? null,
          progress?.restricted ?? null,
          progress?.remaining ?? null
        ]
      )).rows[0] || null
    },

    async failEvent(eventId, workerId, errorCode, now = new Date(), { retryAt = null } = {}) {
      return (await pool.query(
        `UPDATE gift_code_engagement_events
            SET status = 'failed', last_error = $4,
                claimed_by_worker = NULL, claimed_at_utc = NULL,
                claimed_until_utc = NULL, next_attempt_at_utc = $6,
                updated_at_utc = $5
          WHERE id = $1 AND game_profile = $2 AND claimed_by_worker = $3
          RETURNING *`,
        [
          eventId, gameProfile, workerId,
          String(errorCode || "delivery_failed").slice(0, 200), now, retryAt
        ]
      )).rows[0] || null
    },

    async getEventPayload(eventId) {
      return (await pool.query(
        `SELECT e.*, s.gift_code_channel_id, s.contributor_role_id,
                g.code, a.state_or_kingdom_number
           FROM gift_code_engagement_events e
           JOIN gift_code_guild_settings s
             ON s.game_profile = e.game_profile AND s.guild_id = e.guild_id
           LEFT JOIN gift_codes g
             ON g.id = e.gift_code_id AND g.game_profile = e.game_profile
           LEFT JOIN player_accounts a
             ON a.id = e.player_account_id AND a.game_profile = e.game_profile
          WHERE e.id = $1 AND e.game_profile = $2`,
        [eventId, gameProfile]
      )).rows[0] || null
    },

    async codeProgress(giftCodeId, guildId) {
      return (await pool.query(
        `SELECT
           COUNT(*) FILTER (WHERE status = 'success')::integer AS successful,
           COUNT(*) FILTER (WHERE status = 'already_redeemed')::integer AS already_redeemed,
           COUNT(*) FILTER (WHERE status = 'invalid_player')::integer AS account_issues,
           COUNT(*) FILTER (WHERE status = 'restricted')::integer AS restricted,
           COUNT(*) FILTER (WHERE status IN ('queued', 'claimed', 'rate_limited', 'temporary_error'))::integer AS remaining,
           COUNT(*)::integer AS total
         FROM gift_code_redemptions r
         JOIN player_account_guilds ag
           ON ag.player_account_id = r.player_account_id
          AND ag.game_profile = r.game_profile
          AND ag.guild_id = $3
          AND ag.gift_code_enrolled = true
         WHERE r.game_profile = $1 AND r.gift_code_id = $2`,
        [gameProfile, giftCodeId, guildId]
      )).rows[0]
    },

    async accountOwnerStats(discordUserId, guildId) {
      return (await pool.query(
        `SELECT
           COUNT(DISTINCT a.id) FILTER (
             WHERE a.is_active = true AND a.gift_redemption_enabled = true
           )::integer AS "enabledCount",
           COUNT(r.id) FILTER (WHERE r.status = 'success')::integer AS "successfulRedemptions"
         FROM player_accounts a
         JOIN player_account_guilds ag
           ON ag.player_account_id = a.id AND ag.game_profile = a.game_profile
          AND ag.guild_id = $3 AND ag.gift_code_enrolled = true
         LEFT JOIN gift_code_redemptions r
           ON r.player_account_id = a.id AND r.game_profile = a.game_profile
         WHERE a.game_profile = $1 AND a.discord_user_id = $2`,
        [gameProfile, discordUserId, guildId]
      )).rows[0]
    },

    async claimProgressRefresh(giftCodeId, playerAccountId, resultStatus, workerId, now, {
      minimumChange = 5,
      minimumSeconds = 60
    } = {}) {
      const event = (await pool.query(
        `UPDATE gift_code_engagement_events
            SET progress_successful = progress_successful + CASE WHEN $3 = 'success' THEN 1 ELSE 0 END,
                progress_already_redeemed = progress_already_redeemed + CASE WHEN $3 = 'already_redeemed' THEN 1 ELSE 0 END,
                progress_account_issues = progress_account_issues + CASE WHEN $3 = 'invalid_player' THEN 1 ELSE 0 END,
                progress_restricted = progress_restricted + CASE WHEN $3 = 'restricted' THEN 1 ELSE 0 END,
                progress_remaining = GREATEST(0, progress_remaining - 1),
                status = CASE WHEN status = 'completed' AND (
                  progress_remaining <= 1
                  OR (progress_successful + progress_already_redeemed
                      + progress_account_issues + progress_restricted + 1) - last_progress_count >= $6
                  OR last_update_at_utc IS NULL
                  OR last_update_at_utc <= $5::timestamptz - ($7::double precision * interval '1 second')
                ) THEN 'claimed' ELSE status END,
                claimed_by_worker = CASE WHEN status = 'completed' AND (
                  progress_remaining <= 1
                  OR (progress_successful + progress_already_redeemed
                      + progress_account_issues + progress_restricted + 1) - last_progress_count >= $6
                  OR last_update_at_utc IS NULL
                  OR last_update_at_utc <= $5::timestamptz - ($7::double precision * interval '1 second')
                ) THEN $4 ELSE claimed_by_worker END,
                claimed_at_utc = CASE WHEN status = 'completed' AND (
                  progress_remaining <= 1
                  OR (progress_successful + progress_already_redeemed
                      + progress_account_issues + progress_restricted + 1) - last_progress_count >= $6
                  OR last_update_at_utc IS NULL
                  OR last_update_at_utc <= $5::timestamptz - ($7::double precision * interval '1 second')
                ) THEN $5::timestamptz ELSE claimed_at_utc END,
                claimed_until_utc = CASE WHEN status = 'completed' AND (
                  progress_remaining <= 1
                  OR (progress_successful + progress_already_redeemed
                      + progress_account_issues + progress_restricted + 1) - last_progress_count >= $6
                  OR last_update_at_utc IS NULL
                  OR last_update_at_utc <= $5::timestamptz - ($7::double precision * interval '1 second')
                ) THEN $5::timestamptz + interval '60 seconds' ELSE claimed_until_utc END,
                attempt_count = attempt_count + CASE WHEN status = 'completed' AND (
                  progress_remaining <= 1
                  OR (progress_successful + progress_already_redeemed
                      + progress_account_issues + progress_restricted + 1) - last_progress_count >= $6
                  OR last_update_at_utc IS NULL
                  OR last_update_at_utc <= $5::timestamptz - ($7::double precision * interval '1 second')
                ) THEN 1 ELSE 0 END,
                updated_at_utc = $5::timestamptz
          WHERE game_profile = $1 AND gift_code_id = $2
            AND event_type = 'code_progress' AND status IN ('completed', 'claimed')
            AND finalized_at_utc IS NULL
            AND EXISTS (
              SELECT 1 FROM player_account_guilds ag
               WHERE ag.game_profile = gift_code_engagement_events.game_profile
                 AND ag.guild_id = gift_code_engagement_events.guild_id
                 AND ag.player_account_id = $8
                 AND ag.gift_code_enrolled = true
            )
          RETURNING *`,
        [
          gameProfile, giftCodeId, resultStatus, workerId, now,
          minimumChange, minimumSeconds, playerAccountId
        ]
      )).rows[0] || null
      if (!event || event.status !== "claimed" || event.claimed_by_worker !== workerId) return null
      const progress = {
        successful: event.progress_successful,
        already_redeemed: event.progress_already_redeemed,
        account_issues: event.progress_account_issues,
        restricted: event.progress_restricted,
        remaining: event.progress_remaining,
        total: event.progress_successful + event.progress_already_redeemed +
          event.progress_account_issues + event.progress_restricted + event.progress_remaining
      }
      return {
        event,
        progress,
        completed: progress.successful + progress.already_redeemed +
          progress.account_issues + progress.restricted
      }
    },

    async communityStats(guildId = null) {
      const params = [gameProfile, guildId]
      const accountStats = (await pool.query(
        `SELECT
           COUNT(DISTINCT a.discord_user_id)::integer AS registered_users,
           COUNT(DISTINCT a.id)::integer AS registered_accounts,
           COUNT(DISTINCT a.id) FILTER (
             WHERE a.is_active = true AND a.gift_redemption_enabled = true
           )::integer AS enabled_accounts,
           COUNT(DISTINCT a.discord_user_id) FILTER (
             WHERE a.is_active = true AND a.gift_redemption_enabled = true
           )::integer AS auto_redeem_players,
           COUNT(DISTINCT r.id) FILTER (WHERE r.status = 'success')::integer AS successful_redemptions,
           COUNT(DISTINCT r.id) FILTER (WHERE r.status = 'already_redeemed')::integer AS already_redeemed,
           COUNT(DISTINCT r.id) FILTER (
             WHERE r.status = 'success'
               AND r.completed_at_utc >= date_trunc('month', now())
           )::integer AS successful_this_month
         FROM player_accounts a
         LEFT JOIN player_account_guilds ag
           ON ag.player_account_id = a.id AND ag.game_profile = a.game_profile
          AND ag.gift_code_enrolled = true
         LEFT JOIN gift_code_redemptions r
           ON r.player_account_id = a.id AND r.game_profile = a.game_profile
         WHERE a.game_profile = $1
           AND ($2::varchar IS NULL OR ag.guild_id = $2)`,
        params
      )).rows[0]
      const codeStats = (await pool.query(
        `SELECT COUNT(*) FILTER (WHERE status = 'active')::integer AS verified_codes,
                MAX(verified_at_utc) FILTER (WHERE status = 'active') AS latest_verified_at_utc
           FROM gift_codes g WHERE game_profile = $1
            AND ($2::varchar IS NULL OR EXISTS (
              SELECT 1 FROM gift_code_engagement_events e
               WHERE e.game_profile = g.game_profile AND e.gift_code_id = g.id
                 AND e.guild_id = $2 AND e.event_type = 'code_progress'
            ))`,
        params
      )).rows[0]
      const latest = (await pool.query(
        `SELECT code FROM gift_codes g
          WHERE game_profile = $1 AND status = 'active'
            AND ($2::varchar IS NULL OR EXISTS (
              SELECT 1 FROM gift_code_engagement_events e
               WHERE e.game_profile = g.game_profile AND e.gift_code_id = g.id
                 AND e.guild_id = $2 AND e.event_type = 'code_progress'
            ))
          ORDER BY verified_at_utc DESC NULLS LAST, id DESC LIMIT 1`,
        params
      )).rows[0]
      const contributors = (await pool.query(
        `SELECT COUNT(DISTINCT discord_user_id)::integer AS unique_contributors
           FROM gift_code_engagement_events
          WHERE game_profile = $1 AND event_type = 'contributor_role'
            AND status = 'completed'
            AND ($2::varchar IS NULL OR guild_id = $2)`,
        params
      )).rows[0]
      return { ...accountStats, ...codeStats, ...contributors, latest_verified_code: latest?.code || null }
    }
  })
}

module.exports = { createGiftCodeCommunityRepository }
