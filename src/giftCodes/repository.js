const crypto = require("node:crypto")
const { normalizeGiftCode } = require("./validation")

const REDEMPTION_WORK_STATES = Object.freeze(["queued", "rate_limited", "temporary_error"])
const TERMINAL_REDEMPTION_STATES = new Set([
  "success", "already_redeemed", "expired", "invalid_code", "invalid_player",
  "restricted", "retry_exhausted", "unknown", "disabled"
])

function requirePool(pool) {
  if (!pool || typeof pool.query !== "function" || typeof pool.connect !== "function") {
    throw new Error("A transactional PostgreSQL pool is required")
  }
}

function requireProfile(gameProfile) {
  if (!["wos", "kingshot"].includes(gameProfile)) throw new Error("Unsupported game profile")
}

function machineFields(result) {
  const raw = result?.classification?.raw || {}
  return {
    classification: String(result?.classification?.state || "unknown_response"),
    apiCode: raw.code ?? null,
    errCode: raw.errCode ?? null,
    apiMessage: String(raw.message || "").slice(0, 500) || null,
    httpStatus: Number.isInteger(result?.httpStatus) ? result.httpStatus : null,
    metadata: {
      profile: result?.profile || null,
      endpoint: result?.endpoint || "gift_code",
      classification: String(result?.classification?.state || "unknown_response"),
      rateLimit: result?.rateLimit || {},
      transportErrorCode: result?.errorCode || null,
      response: result?.responseDiagnostics || null
    }
  }
}

function createGiftCodeRepository(pool, gameProfile) {
  requirePool(pool)
  requireProfile(gameProfile)

  async function accountStatuses(discordUserId, playerId, guildId, includeInactive) {
    const result = await pool.query(
      `SELECT a.*,
              EXISTS (
                SELECT 1 FROM player_account_guilds ag
                 WHERE ag.game_profile = a.game_profile
                   AND ag.player_account_id = a.id
                   AND ag.guild_id = $4
                   AND ag.gift_code_enrolled = true
              ) AS guild_gift_code_enrolled,
              COALESCE(stats.success_count, 0)::integer AS successful_redemptions,
              COALESCE(stats.already_redeemed_count, 0)::integer AS already_redeemed,
              (
                COALESCE(stats.success_count, 0) +
                COALESCE(stats.already_redeemed_count, 0)
              )::integer AS completed_redemption_checks,
              latest.status AS last_redemption_status,
              latest.api_message AS last_redemption_message
         FROM player_accounts a
         LEFT JOIN LATERAL (
           SELECT
             COUNT(*) FILTER (WHERE status = 'success') AS success_count,
             COUNT(*) FILTER (WHERE status = 'already_redeemed') AS already_redeemed_count
             FROM gift_code_redemptions r
            WHERE r.game_profile = a.game_profile AND r.player_account_id = a.id
              AND r.discord_owner_id_snapshot = a.discord_user_id
         ) stats ON true
         LEFT JOIN LATERAL (
           SELECT status, api_message
             FROM gift_code_redemptions r
            WHERE r.game_profile = a.game_profile AND r.player_account_id = a.id
              AND r.discord_owner_id_snapshot = a.discord_user_id
            ORDER BY COALESCE(r.completed_at_utc, r.attempted_at_utc, r.created_at_utc) DESC, r.id
            LIMIT 1
         ) latest ON true
        WHERE a.game_profile = $1 AND a.discord_user_id = $2
          AND ($3::varchar IS NULL OR a.player_id = $3)
          AND ($5::boolean = true OR a.is_active = true)
        ORDER BY a.is_primary DESC, a.created_at_utc, a.id`,
      [gameProfile, discordUserId, playerId, guildId, includeInactive]
    )
    return result.rows
  }

  const repository = {
    gameProfile,

    async createSource({ sourceType, sourceName, sourceReference = null, trusted = false, metadata = {} }) {
      const result = await pool.query(
        `INSERT INTO gift_code_sources (
           id, game_profile, source_type, source_name, source_reference, trusted, metadata
         ) VALUES ($1, $2, $3, $4, $5, $6, $7)
         RETURNING *`,
        [
          crypto.randomUUID(), gameProfile, String(sourceType || "").trim(),
          String(sourceName || "").trim(), sourceReference, Boolean(trusted), metadata
        ]
      )
      return result.rows[0]
    },

    async recordSubmission({ code, submittedByDiscordUserId = null, sourceId = null, metadata = {} }) {
      const exactCode = normalizeGiftCode(code)
      const client = await pool.connect()
      try {
        await client.query("BEGIN")
        const inserted = await client.query(
          `INSERT INTO gift_codes (
             id, game_profile, code, normalized_code, discovered_by_source_id
           ) VALUES ($1, $2, $3, $4, $5)
           ON CONFLICT (game_profile, code) DO NOTHING
           RETURNING *`,
          [crypto.randomUUID(), gameProfile, exactCode, exactCode, sourceId]
        )
        const duplicate = inserted.rowCount === 0
        const giftCode = inserted.rows[0] || (await client.query(
          `SELECT * FROM gift_codes WHERE game_profile = $1 AND code = $2`,
          [gameProfile, exactCode]
        )).rows[0]
        const submission = (await client.query(
          `INSERT INTO gift_code_submissions (
             id, game_profile, submitted_code, submitted_by_discord_user_id,
             source_id, duplicate_of_gift_code_id, processing_status, raw_metadata
           ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8)
           RETURNING *`,
          [
            crypto.randomUUID(), gameProfile, exactCode, submittedByDiscordUserId,
            sourceId, duplicate ? giftCode.id : null,
            duplicate ? "duplicate" : "pending_verification", metadata
          ]
        )).rows[0]
        await client.query("COMMIT")
        return { giftCode, submission, duplicate }
      } catch (error) {
        await client.query("ROLLBACK")
        throw error
      } finally {
        client.release()
      }
    },

    async getCode(code) {
      const exactCode = normalizeGiftCode(code)
      return (await pool.query(
        `SELECT * FROM gift_codes WHERE game_profile = $1 AND code = $2`,
        [gameProfile, exactCode]
      )).rows[0] || null
    },

    async setAutoRedemption({
      discordUserId,
      playerId,
      enabled,
      guildId = null,
      maximumEnabledAccounts = 2,
      preferenceSource = "user"
    }) {
      const client = await pool.connect()
      try {
        await client.query("BEGIN")
        await client.query(
          "SELECT pg_advisory_xact_lock(hashtext($1))",
          [`gift-auto:${gameProfile}:${discordUserId}`]
        )
        const current = (await client.query(
          `SELECT * FROM player_accounts
            WHERE game_profile = $1 AND discord_user_id = $2
              AND player_id = $3 AND is_active = true
            FOR UPDATE`,
          [gameProfile, discordUserId, playerId]
        )).rows[0] || null
        if (!current) {
          await client.query("COMMIT")
          return { account: null, limitReached: false, enabledCount: 0, engagementEvent: null }
        }
        const beforeCount = (await client.query(
          `SELECT COUNT(*)::integer AS count
             FROM player_accounts
            WHERE game_profile = $1 AND discord_user_id = $2
              AND is_active = true AND gift_redemption_enabled = true`,
          [gameProfile, discordUserId]
        )).rows[0].count
        if (enabled && !current.gift_redemption_enabled && beforeCount >= maximumEnabledAccounts) {
          await client.query("COMMIT")
          return {
            account: current,
            limitReached: true,
            enabledCount: beforeCount,
            engagementEvent: null
          }
        }
        const changed = current.gift_redemption_enabled !== Boolean(enabled)
        const account = (await client.query(
          `UPDATE player_accounts
              SET gift_redemption_enabled = $4,
                  account_metadata = jsonb_set(
                    account_metadata,
                    '{autoRedeemPreference}',
                    jsonb_build_object(
                      'enabled', $4::boolean,
                      'explicit', $5::varchar = 'user'
                    )
                  ),
                  updated_at_utc = now()
            WHERE id = $1 AND game_profile = $2 AND discord_user_id = $3
            RETURNING *`,
          [current.id, gameProfile, discordUserId, Boolean(enabled), preferenceSource]
        )).rows[0]
        let newlyEnrolledInGuild = false
        if (guildId && enabled) {
          const enrolment = (await client.query(
            `INSERT INTO player_account_guilds (
               game_profile, guild_id, player_account_id, gift_code_enrolled,
               gift_code_first_enabled_at_utc, gift_code_updated_at_utc
             ) VALUES ($1, $2, $3, true, now(), now())
             ON CONFLICT (game_profile, guild_id, player_account_id) DO UPDATE
               SET gift_code_enrolled = true,
                   gift_code_first_enabled_at_utc = COALESCE(
                     player_account_guilds.gift_code_first_enabled_at_utc,
                     now()
                   ),
                   gift_code_updated_at_utc = now()
             WHERE player_account_guilds.gift_code_enrolled = false
             RETURNING *`,
            [gameProfile, guildId, account.id]
          )).rows[0] || null
          newlyEnrolledInGuild = Boolean(enrolment)
        }
        let engagementEvent = null
        if (enabled) {
          const activeCodes = (await client.query(
            `SELECT id FROM gift_codes
              WHERE game_profile = $1 AND status = 'active'
              ORDER BY verified_at_utc, id`,
            [gameProfile]
          )).rows
          for (const giftCode of activeCodes) {
            await client.query(
              `INSERT INTO gift_code_redemptions (
                 id, game_profile, gift_code_id, player_account_id,
                 player_id_snapshot, location_number_snapshot,
                 discord_owner_id_snapshot
               ) VALUES ($1, $2, $3, $4, $5, $6, $7)
               ON CONFLICT (game_profile, gift_code_id, player_account_id) DO NOTHING`,
              [
                crypto.randomUUID(), gameProfile, giftCode.id, account.id,
                account.player_id, account.state_or_kingdom_number, discordUserId
              ]
            )
          }
          if (newlyEnrolledInGuild && guildId) {
            engagementEvent = (await client.query(
              `INSERT INTO gift_code_engagement_events (
                 id, game_profile, guild_id, event_type,
                 player_account_id, discord_user_id
               )
               SELECT $1::uuid, $2::varchar, $3::varchar,
                      'auto_redeem_join'::varchar, $4::uuid, $5::varchar
                 FROM gift_code_guild_settings s
                WHERE s.game_profile = $2 AND s.guild_id = $3
                  AND s.announcements_enabled = true
                  AND s.gift_code_channel_id IS NOT NULL
               ON CONFLICT DO NOTHING
               RETURNING *`,
              [crypto.randomUUID(), gameProfile, guildId, account.id, discordUserId]
            )).rows[0] || null
          }
        } else if (account && !enabled) {
          await client.query(
            `UPDATE gift_code_redemptions
                SET status = 'disabled', retryable = false, next_retry_at_utc = NULL,
                    claimed_by_worker = NULL, claimed_at_utc = NULL,
                    claimed_until_utc = NULL, notification_status = 'suppressed',
                    updated_at_utc = now()
              WHERE game_profile = $1 AND player_account_id = $2
                AND status = ANY($3::varchar[])`,
            [gameProfile, account.id, [...REDEMPTION_WORK_STATES, "claimed"]]
          )
        }
        const enabledCount = enabled
          ? beforeCount + (changed ? 1 : 0)
          : Math.max(0, beforeCount - (changed ? 1 : 0))
        await client.query("COMMIT")
        return {
          account,
          limitReached: false,
          enabledCount,
          engagementEvent,
          guildEnrolled: guildId ? enabled || newlyEnrolledInGuild : false
        }
      } catch (error) {
        await client.query("ROLLBACK")
        throw error
      } finally {
        client.release()
      }
    },

    async activeAccountStatuses(discordUserId, playerId = null, guildId = null) {
      return accountStatuses(discordUserId, playerId, guildId, false)
    },

    async historicalAccountStatuses(discordUserId, playerId = null, guildId = null) {
      return accountStatuses(discordUserId, playerId, guildId, true)
    },

    async accountStatuses(discordUserId, playerId = null, guildId = null) {
      return accountStatuses(discordUserId, playerId, guildId, false)
    },

    async redemptionHistory(playerAccountId, discordUserId, limit = 10) {
      const result = await pool.query(
        `SELECT r.status, r.completed_at_utc, r.attempted_at_utc,
                r.location_number_snapshot, g.code
           FROM gift_code_redemptions r
           JOIN gift_codes g
             ON g.id = r.gift_code_id AND g.game_profile = r.game_profile
          WHERE r.game_profile = $1 AND r.player_account_id = $2
            AND r.discord_owner_id_snapshot = $3
          ORDER BY COALESCE(r.completed_at_utc, r.attempted_at_utc, r.created_at_utc) DESC, r.id
          LIMIT $4`,
        [gameProfile, playerAccountId, discordUserId, Math.min(25, Math.max(1, Number(limit) || 10))]
      )
      return result.rows
    },

    async queueRedemption({ giftCodeId, playerAccountId, botInstanceName = null, requireOptIn = false }) {
      const result = await pool.query(
        `INSERT INTO gift_code_redemptions (
           id, game_profile, gift_code_id, player_account_id,
           player_id_snapshot, location_number_snapshot, bot_instance_name,
           discord_owner_id_snapshot
         )
         SELECT $1, $2::varchar, $3, a.id, a.player_id, a.state_or_kingdom_number, $5,
                a.discord_user_id
           FROM player_accounts a
          WHERE a.id = $4 AND a.game_profile = $2 AND a.is_active = true
            AND ($6::boolean = false OR a.gift_redemption_enabled = true)
         ON CONFLICT (game_profile, gift_code_id, player_account_id) DO NOTHING
         RETURNING *`,
        [crypto.randomUUID(), gameProfile, giftCodeId, playerAccountId, botInstanceName, requireOptIn]
      )
      if (result.rowCount) return { created: true, redemption: result.rows[0] }
      const existing = (await pool.query(
        `SELECT * FROM gift_code_redemptions
          WHERE game_profile = $1 AND gift_code_id = $2 AND player_account_id = $3`,
        [gameProfile, giftCodeId, playerAccountId]
      )).rows[0] || null
      return { created: false, redemption: existing }
    },

    async fanOutActiveCode({ giftCodeId, botInstanceName = null, client = pool }) {
      const accounts = (await client.query(
        `SELECT a.id, a.player_id, a.state_or_kingdom_number, a.discord_user_id
           FROM player_accounts a
           JOIN gift_codes c ON c.game_profile = a.game_profile
          WHERE c.id = $2 AND c.game_profile = $1 AND c.status = 'active'
            AND a.is_active = true AND a.gift_redemption_enabled = true
          ORDER BY a.created_at_utc, a.id`,
        [gameProfile, giftCodeId]
      )).rows
      const queued = []
      for (const account of accounts) {
        const row = (await client.query(
          `INSERT INTO gift_code_redemptions (
             id, game_profile, gift_code_id, player_account_id,
             player_id_snapshot, location_number_snapshot, bot_instance_name,
             discord_owner_id_snapshot
           ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8)
           ON CONFLICT (game_profile, gift_code_id, player_account_id) DO NOTHING
           RETURNING *`,
          [
            crypto.randomUUID(), gameProfile, giftCodeId, account.id,
            account.player_id, account.state_or_kingdom_number, botInstanceName,
            account.discord_user_id
          ]
        )).rows[0]
        if (row) queued.push(row)
      }
      return queued
    },

    async claimVerification({ workerId, now, leaseSeconds, code = null, manual = false }) {
      const exactCode = code === null ? null : normalizeGiftCode(code)
      const result = await pool.query(
        `WITH candidate AS (
           SELECT id
             FROM gift_codes
            WHERE game_profile = $1
              AND ($5::varchar IS NULL OR code = $5)
              AND (
                verification_state = 'pending'
                OR (verification_state = 'retry' AND verification_next_retry_at_utc <= $2)
                OR (verification_state = 'claimed' AND verification_claimed_until_utc <= $2)
                OR ($6::boolean = true AND verification_state IN ('blocked', 'review'))
              )
            ORDER BY COALESCE(verification_next_retry_at_utc, first_seen_at_utc), id
            FOR UPDATE SKIP LOCKED
            LIMIT 1
         )
         UPDATE gift_codes g
            SET status = 'verifying', verification_state = 'claimed',
                verification_attempt_count = verification_attempt_count + 1,
                verification_claimed_by_worker = $3,
                verification_claimed_at_utc = $2,
                verification_claimed_until_utc = $2 + ($4 * interval '1 second'),
                verification_next_retry_at_utc = NULL, updated_at_utc = $2
           FROM candidate
          WHERE g.id = candidate.id AND g.game_profile = $1
          RETURNING g.*`,
        [gameProfile, now, workerId, leaseSeconds, exactCode, Boolean(manual)]
      )
      return result.rows[0] || null
    },

    async storedVerificationReview({ errCodes, exactResponses = [], code = null }) {
      const exactCode = code === null ? null : normalizeGiftCode(code)
      const safeErrCodes = [...new Set(errCodes || [])]
        .map(Number)
        .filter(Number.isInteger)
      const exactErrCodes = exactResponses.map(response => Number(response?.errCode))
      const exactMessages = exactResponses.map(response => String(response?.message || "")
        .trim().replace(/[.!?]+$/, "").trim().toUpperCase())
      if (!safeErrCodes.length && !exactErrCodes.length) return null
      return (await pool.query(
        `SELECT g.*, attempt.id AS stored_attempt_id,
                  attempt.http_status AS stored_http_status,
                  attempt.api_code AS stored_api_code,
                  attempt.err_code AS stored_err_code,
                  attempt.api_message AS stored_api_message,
                  attempt.response_metadata AS stored_response_metadata
          FROM gift_codes g
          JOIN LATERAL (
            SELECT a.* FROM gift_code_attempts a
             WHERE a.game_profile = g.game_profile AND a.gift_code_id = g.id
               AND a.attempt_type = 'verification'
             ORDER BY a.attempt_number DESC, a.created_at_utc DESC, a.id DESC
             LIMIT 1
          ) attempt ON true
         WHERE g.game_profile = $1 AND g.status = 'unknown'
           AND g.verification_state = 'review'
           AND attempt.classification = 'unknown_response'
           AND attempt.http_status = 200
           AND (
             attempt.err_code = ANY($2::integer[])
             OR EXISTS (
               SELECT 1
                 FROM unnest($3::integer[], $4::varchar[]) expected(err_code, message)
                WHERE expected.err_code = attempt.err_code
                  AND regexp_replace(upper(btrim(attempt.api_message)), '[.!?]+$', '') = expected.message
             )
           )
           AND ($5::varchar IS NULL OR g.code = $5)
         ORDER BY g.last_verification_at_utc, g.id
         LIMIT 1`,
        [gameProfile, safeErrCodes, exactErrCodes, exactMessages, exactCode]
      )).rows[0] || null
    },

    async finishStoredVerificationRecovery({
      claim,
      classification,
      now,
      codeStatus,
      verificationState,
      nextRetryAt = null,
      botInstanceName = null
    }) {
      const client = await pool.connect()
      try {
        await client.query("BEGIN")
        const updated = (await client.query(
          `UPDATE gift_codes
              SET status = $5::varchar, verification_state = $6::varchar,
                  verification_next_retry_at_utc = $8,
                  verification_claimed_by_worker = NULL,
                  verification_claimed_at_utc = NULL,
                  verification_claimed_until_utc = NULL,
                  verified_at_utc = CASE WHEN $5::varchar = 'active' THEN $4 ELSE verified_at_utc END,
                  last_verification_at_utc = $4,
                  verification_metadata = verification_metadata || jsonb_build_object(
                    'classification', $7::varchar,
                    'storedResponseReclassifiedAt', $4::timestamptz
                  ),
                  updated_at_utc = $4
            WHERE id = $1 AND game_profile = $2
              AND status = 'unknown' AND verification_state = 'review'
              AND EXISTS (
                SELECT 1 FROM gift_code_attempts a
                 WHERE a.id = $3 AND a.game_profile = gift_codes.game_profile
                   AND a.gift_code_id = gift_codes.id
                   AND a.attempt_type = 'verification'
                   AND a.classification = 'unknown_response'
                   AND a.http_status = 200
              )
            RETURNING *`,
          [
            claim.id, gameProfile, claim.stored_attempt_id, now, codeStatus,
            verificationState, classification, nextRetryAt
          ]
        )).rows[0]
        if (!updated) {
          await client.query("COMMIT")
          return null
        }
        const attempt = (await client.query(
          `UPDATE gift_code_attempts
              SET classification = $4,
                  response_metadata = response_metadata || jsonb_build_object(
                    'classification', $4::varchar,
                    'reclassifiedFrom', 'unknown_response',
                    'reclassifiedAt', $5::timestamptz
                  )
            WHERE id = $1 AND game_profile = $2 AND gift_code_id = $3
              AND attempt_type = 'verification'
              AND classification = 'unknown_response'
            RETURNING *`,
          [claim.stored_attempt_id, gameProfile, claim.id, classification, now]
        )).rows[0]
        if (!attempt) throw new Error("Stored verification attempt was not recoverable")
        let queued = []
        if (updated.status === "active") {
          queued = await repository.fanOutActiveCode({
            giftCodeId: updated.id,
            botInstanceName,
            client
          })
        }
        await client.query("COMMIT")
        return { giftCode: updated, queued, attempt }
      } catch (error) {
        await client.query("ROLLBACK")
        throw error
      } finally {
        client.release()
      }
    },

    async finishVerification({ claim, workerId, result, now, codeStatus, verificationState, nextRetryAt = null, botInstanceName = null }) {
      const fields = machineFields({ ...result, profile: gameProfile })
      const client = await pool.connect()
      try {
        await client.query("BEGIN")
        await client.query(
          `INSERT INTO gift_code_attempts (
             id, game_profile, gift_code_id, attempt_type, attempt_number,
             classification, api_code, err_code, api_message, http_status,
             response_metadata, request_started_at_utc, response_received_at_utc
           ) VALUES ($1, $2, $3, 'verification', $4, $5, $6, $7, $8, $9, $10, $11, $12)`,
          [
            crypto.randomUUID(), gameProfile, claim.id, claim.verification_attempt_count,
            fields.classification, fields.apiCode, fields.errCode, fields.apiMessage,
            fields.httpStatus, fields.metadata, result.requestStartedAt || now,
            result.responseReceivedAt || now
          ]
        )
        const updated = (await client.query(
          `UPDATE gift_codes
              SET status = $5::varchar, verification_state = $6::varchar,
                  verification_next_retry_at_utc = $7,
                  verification_claimed_by_worker = NULL,
                  verification_claimed_at_utc = NULL,
                  verification_claimed_until_utc = NULL,
                  verified_at_utc = CASE WHEN $5::varchar = 'active' THEN $4 ELSE verified_at_utc END,
                  last_verification_at_utc = $4,
                  last_api_code = $8, last_err_code = $9,
                  last_api_message = $10, verification_http_status = $11,
                  verification_metadata = $12, updated_at_utc = $4
            WHERE id = $1 AND game_profile = $2
              AND verification_claimed_by_worker = $3
              AND verification_state = 'claimed'
            RETURNING *`,
          [
            claim.id, gameProfile, workerId, now, codeStatus, verificationState,
            nextRetryAt, fields.apiCode, fields.errCode, fields.apiMessage,
            fields.httpStatus, fields.metadata
          ]
        )).rows[0]
        if (!updated) throw new Error("Verification claim ownership was lost")
        let queued = []
        if (updated.status === "active") {
          queued = await repository.fanOutActiveCode({
            giftCodeId: updated.id,
            botInstanceName,
            client
          })
        }
        await client.query(
          `UPDATE gift_code_submissions
              SET processing_status = 'processed'
            WHERE game_profile = $1
              AND (duplicate_of_gift_code_id = $2 OR submitted_code = $3)
              AND processing_status IN ('pending_verification', 'received')`,
          [gameProfile, updated.id, updated.code]
        )
        if (updated.status !== "active" && verificationState !== "retry") {
          await client.query(
            `UPDATE gift_code_submissions
                SET raw_metadata = jsonb_set(
                  raw_metadata,
                  '{verificationResultNotification}',
                  jsonb_build_object(
                    'status', 'pending',
                    'classification', $3::varchar,
                    'createdAt', $4::timestamptz
                  ),
                  true
                )
              WHERE id = (
                SELECT id FROM gift_code_submissions
                 WHERE game_profile = $1
                   AND submitted_code = $2
                   AND duplicate_of_gift_code_id IS NULL
                   AND submitted_by_discord_user_id IS NOT NULL
                   AND raw_metadata->>'submissionKind' = 'user'
                 ORDER BY received_at_utc, id
                 LIMIT 1
              )
                AND game_profile = $1
                AND NOT raw_metadata ? 'verificationResultNotification'`,
            [gameProfile, updated.code, fields.classification, now]
          )
        }
        await client.query("COMMIT")
        return { giftCode: updated, queued }
      } catch (error) {
        await client.query("ROLLBACK")
        throw error
      } finally {
        client.release()
      }
    },

    async claimRedemption({ workerId, now, leaseSeconds }) {
      const client = await pool.connect()
      try {
        await client.query("BEGIN")
        await client.query(
          `UPDATE gift_code_redemptions r
              SET status = 'disabled', retryable = false,
                  next_retry_at_utc = NULL, claimed_by_worker = NULL,
                  claimed_at_utc = NULL, claimed_until_utc = NULL,
                  notification_status = 'suppressed', updated_at_utc = $2
             FROM player_accounts a
            WHERE r.player_account_id = a.id AND r.game_profile = a.game_profile
              AND r.game_profile = $1
              AND r.status = ANY($3::varchar[])
              AND (a.is_active = false OR a.gift_redemption_enabled = false)`,
          [gameProfile, now, [...REDEMPTION_WORK_STATES, "claimed"]]
        )
        const result = await client.query(
          `WITH work AS (
             SELECT r.id, a.player_id, a.state_or_kingdom_number,
                    a.discord_user_id,
                    (
                      SELECT COUNT(*)::integer FROM player_accounts owned
                       WHERE owned.game_profile = a.game_profile
                         AND owned.discord_user_id = a.discord_user_id
                         AND owned.is_active = true
                    ) AS owner_account_count
               FROM gift_code_redemptions r
               JOIN player_accounts a
                 ON a.id = r.player_account_id AND a.game_profile = r.game_profile
               JOIN gift_codes g
                 ON g.id = r.gift_code_id AND g.game_profile = r.game_profile
              WHERE r.game_profile = $1 AND g.status = 'active'
                AND a.is_active = true AND a.gift_redemption_enabled = true
                AND (
                  r.status = 'queued'
                  OR (r.status IN ('rate_limited', 'temporary_error') AND r.next_retry_at_utc <= $2)
                  OR (r.status = 'claimed' AND r.claimed_until_utc <= $2)
                )
              ORDER BY COALESCE(r.next_retry_at_utc, r.claimed_until_utc, r.created_at_utc), r.id
              FOR UPDATE OF r SKIP LOCKED
              LIMIT 1
           )
           UPDATE gift_code_redemptions r
              SET status = 'claimed', attempt_number = attempt_number + 1,
                  player_id_snapshot = work.player_id,
                  location_number_snapshot = work.state_or_kingdom_number,
                  claimed_by_worker = $3, claimed_at_utc = $2,
                  claimed_until_utc = $2 + ($4 * interval '1 second'),
                  next_retry_at_utc = NULL, attempted_at_utc = $2,
                  updated_at_utc = $2
             FROM work
            WHERE r.id = work.id AND r.game_profile = $1
            RETURNING r.*,
              (SELECT code FROM gift_codes WHERE id = r.gift_code_id AND game_profile = r.game_profile) AS code,
              work.discord_user_id, work.owner_account_count`,
          [gameProfile, now, workerId, leaseSeconds]
        )
        await client.query("COMMIT")
        return result.rows[0] || null
      } catch (error) {
        await client.query("ROLLBACK")
        throw error
      } finally {
        client.release()
      }
    },

    async finishRedemption({ claim, workerId, result, now, status, retryable = false, nextRetryAt = null }) {
      if (!REDEMPTION_WORK_STATES.includes(status) && !TERMINAL_REDEMPTION_STATES.has(status)) {
        throw new Error("Unsupported redemption status")
      }
      const fields = machineFields({ ...result, profile: gameProfile })
      const client = await pool.connect()
      try {
        await client.query("BEGIN")
        await client.query(
          `INSERT INTO gift_code_attempts (
             id, game_profile, gift_code_id, redemption_id, attempt_type,
             attempt_number, player_id_snapshot, location_number_snapshot,
             classification, api_code, err_code, api_message, http_status,
             response_metadata, request_started_at_utc, response_received_at_utc
           ) VALUES ($1, $2, $3, $4, 'redemption', $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15)`,
          [
            crypto.randomUUID(), gameProfile, claim.gift_code_id, claim.id,
            claim.attempt_number, claim.player_id_snapshot, claim.location_number_snapshot,
            fields.classification, fields.apiCode, fields.errCode, fields.apiMessage,
            fields.httpStatus, fields.metadata, result.requestStartedAt || now,
            result.responseReceivedAt || now
          ]
        )
        const terminal = TERMINAL_REDEMPTION_STATES.has(status)
        const notificationStatus = ["success", "already_redeemed", "invalid_player", "restricted", "unknown"]
          .includes(status) ? "pending" : "suppressed"
        let updated = (await client.query(
          `UPDATE gift_code_redemptions
              SET status = $5::varchar, api_code = $6, err_code = $7,
                  api_message = $8, http_status = $9,
                  response_metadata = $10, retryable = $11,
                  next_retry_at_utc = $12,
                  completed_at_utc = CASE
                    WHEN $13::boolean THEN $4::timestamptz
                    ELSE NULL::timestamptz
                  END,
                  claimed_by_worker = NULL, claimed_at_utc = NULL,
                  claimed_until_utc = NULL, notification_status = $14,
                  updated_at_utc = $4::timestamptz
            WHERE id = $1 AND game_profile = $2
              AND claimed_by_worker = $3 AND status = 'claimed'
            RETURNING *`,
          [
            claim.id, gameProfile, workerId, now, status, fields.apiCode,
            fields.errCode, fields.apiMessage, fields.httpStatus, fields.metadata,
            Boolean(retryable), nextRetryAt, terminal, notificationStatus
          ]
        )).rows[0]
        let releasedDuringAttempt = false
        if (!updated) {
          const current = (await client.query(
            `SELECT * FROM gift_code_redemptions
              WHERE id = $1 AND game_profile = $2 AND status = 'disabled'`,
            [claim.id, gameProfile]
          )).rows[0] || null
          if (!current) throw new Error("Redemption claim ownership was lost")
          updated = current
          releasedDuringAttempt = true
        }
        if (!releasedDuringAttempt && status === "invalid_player") {
          await client.query(
            `UPDATE player_accounts
                SET verification_status = 'failed', verification_error_code = $3,
                    verification_error_message = $4, last_verified_at_utc = $5,
                    updated_at_utc = $5
              WHERE id = $1 AND game_profile = $2`,
            [claim.player_account_id, gameProfile, fields.errCode?.toString() || null, fields.apiMessage, now]
          )
        } else if (!releasedDuringAttempt && ["success", "already_redeemed"].includes(status)) {
          await client.query(
            `UPDATE player_accounts
                SET verification_status = 'verified', verification_error_code = NULL,
                    verification_error_message = NULL, last_verified_at_utc = $3,
                    updated_at_utc = $3
              WHERE id = $1 AND game_profile = $2`,
            [claim.player_account_id, gameProfile, now]
          )
        }
        await client.query("COMMIT")
        return { ...updated, released_during_attempt: releasedDuringAttempt }
      } catch (error) {
        await client.query("ROLLBACK")
        throw error
      } finally {
        client.release()
      }
    },

    async claimNotification(redemptionId, now = new Date()) {
      return (await pool.query(
        `UPDATE gift_code_redemptions r
            SET notification_status = 'sending', notification_attempted_at_utc = $3,
                updated_at_utc = $3
           FROM player_accounts a
          WHERE r.id = $1 AND r.game_profile = $2 AND r.notification_status = 'pending'
            AND a.id = r.player_account_id AND a.game_profile = r.game_profile
            AND a.is_active = true
            AND a.discord_user_id = r.discord_owner_id_snapshot
          RETURNING r.*, a.discord_user_id`,
        [redemptionId, gameProfile, now]
      )).rows[0] || null
    },

    async finishNotification(redemptionId, { sent, now = new Date(), errorCode = null }) {
      return (await pool.query(
        `UPDATE gift_code_redemptions
            SET notification_status = $3,
                notified_at_utc = CASE
                  WHEN $4::boolean THEN $5::timestamptz
                  ELSE NULL::timestamptz
                END,
                notification_error = $6, updated_at_utc = $5
          WHERE id = $1 AND game_profile = $2 AND notification_status = 'sending'
          RETURNING *`,
        [redemptionId, gameProfile, sent ? "sent" : "failed", Boolean(sent), now, errorCode]
      )).rows[0] || null
    },

    async diagnostics() {
      const counts = (await pool.query(
        `SELECT
           COUNT(*) FILTER (WHERE verification_state IN ('pending', 'retry', 'claimed'))::integer AS pending_candidates,
           COUNT(*) FILTER (WHERE status = 'active')::integer AS active_codes,
           COUNT(*) FILTER (WHERE status = 'expired')::integer AS expired_codes,
           COUNT(*) FILTER (WHERE status = 'invalid')::integer AS invalid_codes,
           COUNT(*) FILTER (
             WHERE status IN ('restricted', 'unknown') OR verification_state = 'review'
           )::integer AS restricted_review_codes
         FROM gift_codes WHERE game_profile = $1`,
        [gameProfile]
      )).rows[0]
      const queue = (await pool.query(
        `SELECT
           COUNT(*) FILTER (WHERE status = 'queued')::integer AS pending_redemptions,
           COUNT(*) FILTER (WHERE status IN ('rate_limited', 'temporary_error'))::integer AS retry_count,
           MIN(created_at_utc) FILTER (WHERE status = 'queued') AS oldest_pending_at_utc,
           MIN(next_retry_at_utc) FILTER (WHERE status IN ('rate_limited', 'temporary_error')) AS next_retry_at_utc
         FROM gift_code_redemptions WHERE game_profile = $1`,
        [gameProfile]
      )).rows[0]
      return { ...counts, ...queue }
    },

    async activeCodeVisibility({ page = 0, pageSize = 15 } = {}) {
      const safePage = Math.max(0, Math.floor(Number(page) || 0))
      const safePageSize = Math.min(25, Math.max(1, Math.floor(Number(pageSize) || 15)))
      const [counts, codes] = await Promise.all([
        pool.query(
          `SELECT
             COUNT(*) FILTER (WHERE status = 'active')::integer AS active_count,
             COUNT(*) FILTER (WHERE status = 'expired')::integer AS expired_count
           FROM gift_codes WHERE game_profile = $1`,
          [gameProfile]
        ),
        pool.query(
          `SELECT code, verified_at_utc
             FROM gift_codes
            WHERE game_profile = $1 AND status = 'active'
            ORDER BY verified_at_utc DESC NULLS LAST, first_seen_at_utc DESC, id DESC
            LIMIT $2 OFFSET $3`,
          [gameProfile, safePageSize, safePage * safePageSize]
        )
      ])
      return {
        codes: codes.rows,
        activeCount: counts.rows[0].active_count,
        expiredCount: counts.rows[0].expired_count,
        page: safePage,
        pageSize: safePageSize
      }
    },

    async codeDiagnostics(code) {
      const exactCode = normalizeGiftCode(code)
      return (await pool.query(
        `SELECT g.*,
           COALESCE(r.pending, 0)::integer AS pending_count,
           COALESCE(r.success, 0)::integer AS success_count,
           COALESCE(r.already, 0)::integer AS already_redeemed_count,
           COALESCE(r.failed, 0)::integer AS failed_count,
           verification.id AS latest_verification_attempt_id,
           verification.http_status AS latest_verification_http_status,
           verification.err_code AS latest_verification_err_code,
           verification.api_message AS latest_verification_api_message,
           verification.classification AS latest_verification_classification,
           verification.response_metadata AS latest_verification_metadata,
           redemption.id AS latest_redemption_attempt_id,
           redemption.http_status AS latest_redemption_http_status,
           redemption.err_code AS latest_redemption_err_code,
           redemption.api_message AS latest_redemption_api_message,
           redemption.classification AS latest_redemption_classification,
           redemption.player_id_snapshot AS latest_redemption_player_id_snapshot,
           s.source_type, s.source_name
         FROM gift_codes g
         LEFT JOIN gift_code_sources s
           ON s.id = g.discovered_by_source_id AND s.game_profile = g.game_profile
         LEFT JOIN LATERAL (
           SELECT
             COUNT(*) FILTER (WHERE status IN ('queued', 'claimed', 'rate_limited', 'temporary_error')) AS pending,
             COUNT(*) FILTER (WHERE status = 'success') AS success,
             COUNT(*) FILTER (WHERE status = 'already_redeemed') AS already,
             COUNT(*) FILTER (WHERE status IN ('invalid_player', 'restricted', 'unknown', 'retry_exhausted')) AS failed
           FROM gift_code_redemptions r
           WHERE r.game_profile = g.game_profile AND r.gift_code_id = g.id
         ) r ON true
         LEFT JOIN LATERAL (
           SELECT id, http_status, err_code, api_message, classification, response_metadata
             FROM gift_code_attempts a
            WHERE a.game_profile = g.game_profile AND a.gift_code_id = g.id
              AND a.attempt_type = 'verification'
            ORDER BY a.attempt_number DESC, a.created_at_utc DESC, a.id DESC
            LIMIT 1
         ) verification ON true
         LEFT JOIN LATERAL (
           SELECT id, http_status, err_code, api_message, classification, player_id_snapshot
             FROM gift_code_attempts a
            WHERE a.game_profile = g.game_profile AND a.gift_code_id = g.id
              AND a.attempt_type = 'redemption'
            ORDER BY a.response_received_at_utc DESC, a.created_at_utc DESC, a.id DESC
            LIMIT 1
         ) redemption ON true
         WHERE g.game_profile = $1 AND g.code = $2`,
        [gameProfile, exactCode]
      )).rows[0] || null
    },

    async withProfileRequestLock(operation, {
      minimumDelayMs = 0,
      sleep = delay => new Promise(resolve => setTimeout(resolve, delay)),
      now = () => new Date()
    } = {}) {
      const client = await pool.connect()
      let locked = false
      try {
        await client.query(
          "SELECT pg_advisory_lock(hashtext($1))",
          [`century-request:${gameProfile}`]
        )
        locked = true
        const state = (await client.query(
          `SELECT last_request_started_at_utc
             FROM gift_code_rate_limit_state
            WHERE game_profile = $1`,
          [gameProfile]
        )).rows[0]
        if (state?.last_request_started_at_utc) {
          const elapsed = now().getTime() - new Date(state.last_request_started_at_utc).getTime()
          const delay = Math.max(0, Number(minimumDelayMs) - elapsed)
          if (delay > 0) await sleep(delay)
        }
        const startedAt = now()
        await client.query(
          `INSERT INTO gift_code_rate_limit_state (
             game_profile, last_request_started_at_utc, updated_at_utc
           ) VALUES ($1, $2, $2)
           ON CONFLICT (game_profile) DO UPDATE
             SET last_request_started_at_utc = EXCLUDED.last_request_started_at_utc,
                 updated_at_utc = EXCLUDED.updated_at_utc`,
          [gameProfile, startedAt]
        )
        return await operation()
      } finally {
        if (locked) {
          await client.query(
            "SELECT pg_advisory_unlock(hashtext($1))",
            [`century-request:${gameProfile}`]
          ).catch(() => {})
        }
        client.release()
      }
    }
  }
  return repository
}

module.exports = {
  REDEMPTION_WORK_STATES,
  TERMINAL_REDEMPTION_STATES,
  machineFields,
  createGiftCodeRepository
}
