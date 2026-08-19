const crypto = require("node:crypto")

const SOURCE_PRIORITIES = Object.freeze({
  discord_mirror: 10,
  manual_user: 20,
  manual_admin: 30,
  public_catalogue: 40
})

function createGiftCodeSourceRepository(pool, gameProfile) {
  if (!pool?.query || !pool?.connect) throw new Error("A transactional PostgreSQL pool is required")
  if (!["wos", "kingshot"].includes(gameProfile)) throw new Error("Unsupported game profile")

  async function ensureSource({ sourceType, sourceName, sourceReference = null, trusted = false }) {
    const client = await pool.connect()
    try {
      await client.query("BEGIN")
      await client.query(
        "SELECT pg_advisory_xact_lock(hashtext($1))",
        [`gift-source:${gameProfile}:${sourceType}:${sourceName}`]
      )
      const existing = (await client.query(
        `SELECT * FROM gift_code_sources
          WHERE game_profile = $1 AND source_type = $2 AND source_name = $3
          ORDER BY created_at_utc, id
          LIMIT 1
          FOR UPDATE`,
        [gameProfile, sourceType, sourceName]
      )).rows[0]
      const source = existing
        ? (await client.query(
          `UPDATE gift_code_sources
              SET source_reference = COALESCE($2, source_reference), enabled = true,
                  updated_at_utc = now()
            WHERE id = $1
            RETURNING *`,
          [existing.id, sourceReference]
        )).rows[0]
        : (await client.query(
          `INSERT INTO gift_code_sources (
             id, game_profile, source_type, source_name, source_reference, trusted, priority
           ) VALUES ($1, $2, $3, $4, $5, $6, $7)
           RETURNING *`,
          [
            crypto.randomUUID(), gameProfile, sourceType, sourceName, sourceReference,
            Boolean(trusted), SOURCE_PRIORITIES[sourceType] || 100
          ]
        )).rows[0]
      await client.query("COMMIT")
      return source
    } catch (error) {
      await client.query("ROLLBACK")
      throw error
    } finally {
      client.release()
    }
  }

  return Object.freeze({
    gameProfile,
    ensureSource,

    async configureDiscordChannel({ guildId, channelId, requireWebhook = false }) {
      const source = await ensureSource({
        sourceType: "discord_mirror",
        sourceName: `Discord mirror ${guildId}/${channelId}`,
        sourceReference: `discord:${guildId}:${channelId}`,
        trusted: true
      })
      return (await pool.query(
        `INSERT INTO gift_code_source_channels (
           game_profile, guild_id, channel_id, source_id, require_webhook
         ) VALUES ($1, $2, $3, $4, $5)
         ON CONFLICT (game_profile, guild_id, channel_id) DO UPDATE
           SET source_id = EXCLUDED.source_id, enabled = true,
               require_webhook = EXCLUDED.require_webhook, updated_at_utc = now()
         RETURNING *`,
        [gameProfile, guildId, channelId, source.id, Boolean(requireWebhook)]
      )).rows[0]
    },

    async discordChannel(channelId, guildId) {
      return (await pool.query(
        `SELECT c.*, s.source_type, s.source_name
           FROM gift_code_source_channels c
           JOIN gift_code_sources s
             ON s.id = c.source_id AND s.game_profile = c.game_profile
          WHERE c.game_profile = $1 AND c.guild_id = $2 AND c.channel_id = $3
            AND c.enabled = true AND s.enabled = true`,
        [gameProfile, guildId, channelId]
      )).rows[0] || null
    },

    async observation(sourceId, observationKey) {
      return (await pool.query(
        `SELECT c.*
           FROM gift_code_source_observations o
           JOIN gift_codes c
             ON c.id = o.gift_code_id AND c.game_profile = o.game_profile
          WHERE o.game_profile = $1 AND o.source_id = $2 AND o.observation_key = $3`,
        [gameProfile, sourceId, String(observationKey).slice(0, 300)]
      )).rows[0] || null
    },

    async recordObservation({
      sourceId,
      giftCodeId,
      code,
      observationKey,
      observedAt,
      sourceReportedExpiryAt = null,
      provenance = {},
      candidateCreated = false
    }) {
      const result = await pool.query(
        `WITH observation AS (
           INSERT INTO gift_code_source_observations (
             id, game_profile, source_id, gift_code_id, observed_code,
             observation_key, observed_at_utc, source_reported_expiry_at_utc, provenance
           ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9)
           ON CONFLICT (game_profile, source_id, observation_key) DO UPDATE
             SET observed_at_utc = LEAST(
                   gift_code_source_observations.observed_at_utc,
                   EXCLUDED.observed_at_utc
                 ),
                 source_reported_expiry_at_utc = COALESCE(
                   gift_code_source_observations.source_reported_expiry_at_utc,
                   EXCLUDED.source_reported_expiry_at_utc
                 ),
                 no_longer_observed_at_utc = NULL,
                 updated_at_utc = now()
           RETURNING (xmax = 0) AS inserted
         )
         UPDATE gift_code_sources s
            SET last_observation_at_utc = GREATEST(last_observation_at_utc, $7),
                last_candidate_at_utc = CASE WHEN $10
                  THEN GREATEST(last_candidate_at_utc, $7)
                  ELSE last_candidate_at_utc
                END,
                observations_count = observations_count + CASE
                  WHEN (SELECT inserted FROM observation) THEN 1 ELSE 0 END,
                candidates_count = candidates_count + CASE WHEN $10 THEN 1 ELSE 0 END,
                last_error = NULL, updated_at_utc = now()
          WHERE s.id = $3 AND s.game_profile = $2
         RETURNING (SELECT inserted FROM observation) AS observation_inserted`,
        [
          crypto.randomUUID(), gameProfile, sourceId, giftCodeId, code,
          String(observationKey).slice(0, 300), observedAt, sourceReportedExpiryAt, provenance,
          Boolean(candidateCreated)
        ]
      )
      return result.rows[0]
    },

    async markCataloguePoll(sourceId, { successful, observedCodes = [], errorCode = null, now = new Date() }) {
      await pool.query(
        `UPDATE gift_code_sources
            SET last_poll_at_utc = $3,
                last_successful_poll_at_utc = CASE WHEN $4 THEN $3 ELSE last_successful_poll_at_utc END,
                last_error = $5, updated_at_utc = $3
          WHERE id = $1 AND game_profile = $2`,
        [sourceId, gameProfile, now, Boolean(successful), errorCode]
      )
      if (!successful) return
      await pool.query(
        `UPDATE gift_code_source_observations
            SET no_longer_observed_at_utc = $3, updated_at_utc = $3
          WHERE game_profile = $1 AND source_id = $2
            AND no_longer_observed_at_utc IS NULL
            AND NOT (observed_code = ANY($4::varchar[]))`,
        [gameProfile, sourceId, now, observedCodes]
      )
    },

    async sourceStatus(guildId = null, { publicCatalogueEnabled = true } = {}) {
      const sources = (await pool.query(
        `SELECT source_type, source_name, enabled, last_observation_at_utc,
                last_candidate_at_utc, last_poll_at_utc, last_successful_poll_at_utc,
                last_error, observations_count, candidates_count
           FROM gift_code_sources s
          WHERE s.game_profile = $1
            AND (
              s.source_type <> 'discord_mirror'
              OR $2::varchar IS NULL
              OR EXISTS (
                SELECT 1 FROM gift_code_source_channels c
                 WHERE c.game_profile = s.game_profile AND c.source_id = s.id
                   AND c.guild_id = $2
              )
            )
          ORDER BY priority, source_name`,
        [gameProfile, guildId]
      )).rows
      const channels = (await pool.query(
        `SELECT guild_id, channel_id, enabled, require_webhook
           FROM gift_code_source_channels
          WHERE game_profile = $1 AND ($2::varchar IS NULL OR guild_id = $2)
          ORDER BY guild_id, channel_id`,
        [gameProfile, guildId]
      )).rows
      const summary = (await pool.query(
        `SELECT
           COUNT(DISTINCT o.gift_code_id)::integer AS codes_observed,
           COUNT(DISTINCT o.gift_code_id) FILTER (
             WHERE g.status = 'candidate'
               AND g.verification_state = 'pending'
               AND g.verification_attempt_count = 0
           )::integer AS new_candidates
           FROM gift_code_source_observations o
           JOIN gift_code_sources s
             ON s.id = o.source_id AND s.game_profile = o.game_profile
           JOIN gift_codes g
             ON g.id = o.gift_code_id AND g.game_profile = o.game_profile
          WHERE o.game_profile = $1
            AND o.no_longer_observed_at_utc IS NULL
            AND s.enabled = true
            AND s.source_type NOT IN ('manual_user', 'manual_admin')
            AND (s.source_type <> 'public_catalogue' OR $3::boolean = true)
            AND (
              s.source_type <> 'discord_mirror'
              OR EXISTS (
                SELECT 1 FROM gift_code_source_channels c
                 WHERE c.game_profile = s.game_profile AND c.source_id = s.id
                   AND c.enabled = true
                   AND ($2::varchar IS NULL OR c.guild_id = $2)
              )
            )`,
        [gameProfile, guildId, Boolean(publicCatalogueEnabled)]
      )).rows[0]
      return { sources, channels, summary }
    }
  })
}

module.exports = { SOURCE_PRIORITIES, createGiftCodeSourceRepository }
