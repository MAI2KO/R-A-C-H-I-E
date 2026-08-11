const crypto = require("node:crypto")
const { normalizeGiftCode } = require("./validation")

function requirePool(pool) {
  if (!pool || typeof pool.query !== "function" || typeof pool.connect !== "function") {
    throw new Error("A transactional PostgreSQL pool is required")
  }
}

function createGiftCodeRepository(pool, gameProfile) {
  requirePool(pool)
  if (!new Set(["wos", "kingshot"]).has(gameProfile)) {
    throw new Error("Unsupported game profile")
  }

  return {
    gameProfile,

    async createSource({
      sourceType,
      sourceName,
      sourceReference = null,
      trusted = false,
      metadata = {}
    }) {
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

    async recordSubmission({
      code,
      submittedByDiscordUserId = null,
      sourceId = null,
      metadata = {}
    }) {
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
            duplicate ? "duplicate" : "pending_verification",
            metadata
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

    async queueRedemption({ giftCodeId, playerAccountId, botInstanceName = null }) {
      const result = await pool.query(
        `INSERT INTO gift_code_redemptions (
           id, game_profile, gift_code_id, player_account_id,
           player_id_snapshot, location_number_snapshot, bot_instance_name
         )
         SELECT $1, $2::varchar, $3, a.id, a.player_id, a.state_or_kingdom_number, $5
           FROM player_accounts a
          WHERE a.id = $4 AND a.game_profile = $2 AND a.is_active = true
         ON CONFLICT (game_profile, gift_code_id, player_account_id) DO NOTHING
         RETURNING *`,
        [crypto.randomUUID(), gameProfile, giftCodeId, playerAccountId, botInstanceName]
      )
      if (result.rowCount) return { created: true, redemption: result.rows[0] }
      const existing = (await pool.query(
        `SELECT * FROM gift_code_redemptions
          WHERE game_profile = $1 AND gift_code_id = $2 AND player_account_id = $3`,
        [gameProfile, giftCodeId, playerAccountId]
      )).rows[0] || null
      return { created: false, redemption: existing }
    }
  }
}

module.exports = { createGiftCodeRepository }
