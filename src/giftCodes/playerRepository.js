const crypto = require("node:crypto")

function requirePool(pool) {
  if (!pool || typeof pool.query !== "function" || typeof pool.connect !== "function") {
    throw new Error("A transactional PostgreSQL pool is required")
  }
}

function createPlayerRepository(pool, gameProfile) {
  requirePool(pool)
  if (!new Set(["wos", "kingshot"]).has(gameProfile)) {
    throw new Error("Unsupported game profile")
  }

  return {
    gameProfile,

    async registerAccount({ discordUserId, playerId, locationNumber }) {
      const client = await pool.connect()
      try {
        await client.query("BEGIN")
        await client.query(
          "SELECT pg_advisory_xact_lock(hashtext($1))",
          [`player-primary:${gameProfile}:${discordUserId}`]
        )
        const hasActive = (await client.query(
          `SELECT 1 FROM player_accounts
            WHERE game_profile = $1 AND discord_user_id = $2 AND is_active = true
            LIMIT 1`,
          [gameProfile, discordUserId]
        )).rowCount > 0
        const id = crypto.randomUUID()
        const account = (await client.query(
          `INSERT INTO player_accounts (
             id, game_profile, discord_user_id, player_id,
             state_or_kingdom_number, is_primary
           ) VALUES ($1, $2, $3, $4, $5, $6)
           RETURNING *`,
          [id, gameProfile, discordUserId, playerId, locationNumber, !hasActive]
        )).rows[0]
        await client.query(
          `INSERT INTO player_location_history (
             id, player_account_id, game_profile, previous_number, new_number,
             changed_by_discord_user_id, change_source
           ) VALUES ($1, $2, $3, NULL, $4, $5, 'user_command')`,
          [crypto.randomUUID(), id, gameProfile, locationNumber, discordUserId]
        )
        await client.query("COMMIT")
        return account
      } catch (error) {
        await client.query("ROLLBACK")
        throw error
      } finally {
        client.release()
      }
    },

    async listOwnedAccounts(discordUserId, { includeInactive = false } = {}) {
      const result = await pool.query(
        `SELECT * FROM player_accounts
          WHERE game_profile = $1 AND discord_user_id = $2
            AND ($3::boolean = true OR is_active = true)
          ORDER BY is_active DESC, is_primary DESC, created_at_utc, id`,
        [gameProfile, discordUserId, includeInactive]
      )
      return result.rows
    },

    async getOwnedAccount(discordUserId, playerId, { includeInactive = false } = {}) {
      const result = await pool.query(
        `SELECT * FROM player_accounts
          WHERE game_profile = $1 AND discord_user_id = $2 AND player_id = $3
            AND ($4::boolean = true OR is_active = true)`,
        [gameProfile, discordUserId, playerId, includeInactive]
      )
      return result.rows[0] || null
    },

    async updateLocation({ discordUserId, playerId, newNumber, changeSource = "user_command" }) {
      const client = await pool.connect()
      try {
        await client.query("BEGIN")
        const current = (await client.query(
          `SELECT * FROM player_accounts
            WHERE game_profile = $1 AND discord_user_id = $2
              AND player_id = $3 AND is_active = true
            FOR UPDATE`,
          [gameProfile, discordUserId, playerId]
        )).rows[0]
        if (!current) {
          await client.query("ROLLBACK")
          return null
        }
        if (current.state_or_kingdom_number === newNumber) {
          await client.query("COMMIT")
          return { account: current, previousNumber: newNumber, changed: false }
        }
        const account = (await client.query(
          `UPDATE player_accounts
              SET state_or_kingdom_number = $4, updated_at_utc = now()
            WHERE id = $1 AND game_profile = $2 AND discord_user_id = $3
            RETURNING *`,
          [current.id, gameProfile, discordUserId, newNumber]
        )).rows[0]
        await client.query(
          `INSERT INTO player_location_history (
             id, player_account_id, game_profile, previous_number, new_number,
             changed_by_discord_user_id, change_source
           ) VALUES ($1, $2, $3, $4, $5, $6, $7)`,
          [
            crypto.randomUUID(), current.id, gameProfile,
            current.state_or_kingdom_number, newNumber, discordUserId, changeSource
          ]
        )
        await client.query("COMMIT")
        return {
          account,
          previousNumber: current.state_or_kingdom_number,
          changed: true
        }
      } catch (error) {
        await client.query("ROLLBACK")
        throw error
      } finally {
        client.release()
      }
    },

    async deactivateAccount({ discordUserId, playerId }) {
      const client = await pool.connect()
      try {
        await client.query("BEGIN")
        await client.query(
          "SELECT pg_advisory_xact_lock(hashtext($1))",
          [`player-primary:${gameProfile}:${discordUserId}`]
        )
        const account = (await client.query(
          `SELECT * FROM player_accounts
            WHERE game_profile = $1 AND discord_user_id = $2
              AND player_id = $3 AND is_active = true
            FOR UPDATE`,
          [gameProfile, discordUserId, playerId]
        )).rows[0]
        if (!account) {
          await client.query("ROLLBACK")
          return null
        }
        await client.query(
          `UPDATE player_accounts
              SET is_active = false, is_primary = false, updated_at_utc = now()
            WHERE id = $1 AND game_profile = $2`,
          [account.id, gameProfile]
        )
        let replacement = null
        if (account.is_primary) {
          replacement = (await client.query(
            `UPDATE player_accounts
                SET is_primary = true, updated_at_utc = now()
              WHERE id = (
                SELECT id FROM player_accounts
                 WHERE game_profile = $1 AND discord_user_id = $2 AND is_active = true
                 ORDER BY created_at_utc, id LIMIT 1
              )
              RETURNING *`,
            [gameProfile, discordUserId]
          )).rows[0] || null
        }
        await client.query("COMMIT")
        return { account: { ...account, is_active: false, is_primary: false }, replacement }
      } catch (error) {
        await client.query("ROLLBACK")
        throw error
      } finally {
        client.release()
      }
    }
  }
}

module.exports = { createPlayerRepository }
