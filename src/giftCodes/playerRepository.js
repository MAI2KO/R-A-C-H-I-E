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

    async registerAccount({ discordUserId, playerId, inGameName, locationNumber,
      allianceAbbreviation, guildId = null }) {
      const client = await pool.connect()
      try {
        await client.query("BEGIN")
        await client.query(
          "SELECT pg_advisory_xact_lock(hashtext($1))",
          [`player-primary:${gameProfile}:${discordUserId}`]
        )
        const existing = (await client.query(
          `SELECT * FROM player_accounts
            WHERE game_profile = $1 AND player_id = $2
            FOR UPDATE`,
          [gameProfile, playerId]
        )).rows[0] || null
        const hasActive = (await client.query(
          `SELECT 1 FROM player_accounts
            WHERE game_profile = $1 AND discord_user_id = $2 AND is_active = true
            LIMIT 1`,
          [gameProfile, discordUserId]
        )).rowCount > 0
        const sameOwner = existing && existing.discord_user_id === discordUserId
        const canReactivate = sameOwner && !existing.is_active
        const canClaim = existing && existing.discord_user_id === null
        const canUpdate = sameOwner || canClaim
        const id = canUpdate ? existing.id : crypto.randomUUID()
        const account = canUpdate
          ? (await client.query(
            `UPDATE player_accounts
                SET discord_user_id = $3, state_or_kingdom_number = $4,
                    in_game_name = $5, alliance_abbreviation = $6, is_active = true,
                    is_primary = $7,
                    gift_redemption_enabled = CASE WHEN is_active THEN gift_redemption_enabled ELSE false END,
                    account_metadata = CASE WHEN discord_user_id IS NULL
                      THEN account_metadata - 'autoRedeemPreference'
                      ELSE account_metadata
                    END,
                    updated_at_utc = now()
              WHERE id = $1 AND game_profile = $2
                AND (discord_user_id = $3 OR discord_user_id IS NULL)
              RETURNING *`,
            [id, gameProfile, discordUserId, locationNumber, inGameName,
              allianceAbbreviation, existing.is_active ? existing.is_primary : !hasActive]
          )).rows[0]
          : (await client.query(
            `INSERT INTO player_accounts (
               id, game_profile, discord_user_id, player_id,
               state_or_kingdom_number, in_game_name, alliance_abbreviation, is_primary
             ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8)
             RETURNING *`,
            [id, gameProfile, discordUserId, playerId, locationNumber,
              inGameName, allianceAbbreviation, !hasActive]
          )).rows[0]
        if (canClaim) {
          await client.query(
            `INSERT INTO player_account_ownership_history (
               id, game_profile, player_account_id, previous_discord_user_id,
               new_discord_user_id, action_type, performed_by_discord_user_id,
               source_metadata
             ) VALUES ($1, $2, $3, NULL, $4, 'claim', $4, $5)`,
            [crypto.randomUUID(), gameProfile, id, discordUserId, { source: "player_register" }]
          )
        }
        if (!canUpdate || existing.state_or_kingdom_number !== locationNumber) {
          await client.query(
            `INSERT INTO player_location_history (
               id, player_account_id, game_profile, previous_number, new_number,
               changed_by_discord_user_id, change_source
             ) VALUES ($1, $2, $3, $4, $5, $6, 'user_command')`,
            [
              crypto.randomUUID(), id, gameProfile,
              canUpdate ? existing.state_or_kingdom_number : null,
              locationNumber, discordUserId
            ]
          )
        }
        if (guildId) {
          await client.query(
            `INSERT INTO player_account_guilds (
               game_profile, guild_id, player_account_id
             ) VALUES ($1, $2, $3)
             ON CONFLICT (game_profile, guild_id, player_account_id) DO UPDATE
               SET gift_code_enrolled = false,
                   gift_code_updated_at_utc = now()`,
            [gameProfile, guildId, id]
          )
        }
        await client.query("COMMIT")
        return {
          ...account,
          registration_status: canReactivate ? "reactivated"
            : canClaim ? "claimed" : sameOwner ? "updated" : "new"
        }
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

    async getAccountByPlayerId(playerId) {
      const result = await pool.query(
        `SELECT a.*,
                COUNT(ag.guild_id)::integer AS guild_enrolment_count
           FROM player_accounts a
           LEFT JOIN player_account_guilds ag
             ON ag.player_account_id = a.id AND ag.game_profile = a.game_profile
          WHERE a.game_profile = $1 AND a.player_id = $2
          GROUP BY a.id`,
        [gameProfile, playerId]
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
              SET is_active = false, is_primary = false,
                  gift_redemption_enabled = false, updated_at_utc = now()
            WHERE id = $1 AND game_profile = $2`,
          [account.id, gameProfile]
        )
        await client.query(
          `UPDATE gift_code_redemptions
              SET status = 'disabled', retryable = false, next_retry_at_utc = NULL,
                  claimed_by_worker = NULL, claimed_at_utc = NULL,
                  claimed_until_utc = NULL, notification_status = 'suppressed',
                  updated_at_utc = now()
            WHERE game_profile = $1 AND player_account_id = $2
              AND status IN ('queued', 'claimed', 'rate_limited', 'temporary_error')`,
          [gameProfile, account.id]
        )
        await client.query(
          `UPDATE player_account_guilds
              SET gift_code_enrolled = false, gift_code_updated_at_utc = now()
            WHERE game_profile = $1 AND player_account_id = $2
              AND gift_code_enrolled = true`,
          [gameProfile, account.id]
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
    },

    async releaseAccount({
      playerId,
      performedByDiscordUserId,
      actionType,
      expectedOwnerDiscordUserId = null,
      expectedAccountId = null
    }) {
      if (!["release", "operator_release"].includes(actionType)) {
        throw new Error("Unsupported ownership release action")
      }
      if (actionType === "operator_release"
        && (!expectedAccountId || !expectedOwnerDiscordUserId)) {
        throw new Error("Operator release requires confirmed account ownership")
      }
      const client = await pool.connect()
      try {
        await client.query("BEGIN")
        const account = (await client.query(
          `SELECT * FROM player_accounts
            WHERE game_profile = $1 AND player_id = $2
            FOR UPDATE`,
          [gameProfile, playerId]
        )).rows[0] || null
        if (
          !account?.discord_user_id
          || (expectedOwnerDiscordUserId !== null
            && account.discord_user_id !== expectedOwnerDiscordUserId)
          || (expectedAccountId !== null && String(account.id) !== String(expectedAccountId))
        ) {
          await client.query("ROLLBACK")
          return null
        }
        if (actionType === "release" && account.discord_user_id !== performedByDiscordUserId) {
          await client.query("ROLLBACK")
          return null
        }

        const previousOwner = account.discord_user_id
        await client.query(
          "SELECT pg_advisory_xact_lock(hashtext($1))",
          [`player-primary:${gameProfile}:${previousOwner}`]
        )
        const guildIds = (await client.query(
          `SELECT guild_id FROM player_account_guilds
            WHERE game_profile = $1 AND player_account_id = $2
            ORDER BY guild_id`,
          [gameProfile, account.id]
        )).rows.map(row => row.guild_id)

        await client.query(
          `UPDATE gift_code_redemptions
              SET status = CASE
                    WHEN status IN ('queued', 'claimed', 'rate_limited', 'temporary_error')
                      THEN 'disabled'
                    ELSE status
                  END,
                  retryable = false, next_retry_at_utc = NULL,
                  claimed_by_worker = NULL, claimed_at_utc = NULL,
                  claimed_until_utc = NULL,
                  notification_status = CASE
                    WHEN notification_status = 'sent' THEN 'sent'
                    ELSE 'suppressed'
                  END,
                  updated_at_utc = now()
            WHERE game_profile = $1 AND player_account_id = $2`,
          [gameProfile, account.id]
        )
        await client.query(
          `UPDATE gift_code_engagement_events
              SET status = 'disabled', claimed_by_worker = NULL,
                  claimed_at_utc = NULL, claimed_until_utc = NULL,
                  next_attempt_at_utc = NULL, updated_at_utc = now()
            WHERE game_profile = $1 AND player_account_id = $2
              AND status IN ('pending', 'claimed', 'failed')`,
          [gameProfile, account.id]
        )
        await client.query(
          `DELETE FROM player_account_guilds
            WHERE game_profile = $1 AND player_account_id = $2`,
          [gameProfile, account.id]
        )
        const released = (await client.query(
          `UPDATE player_accounts
              SET discord_user_id = NULL, is_active = false, is_primary = false,
                  gift_redemption_enabled = false,
                  account_metadata = jsonb_set(
                    account_metadata,
                    '{autoRedeemPreference}',
                    jsonb_build_object(
                      'enabled', false,
                      'explicit', true,
                      'source', $3::varchar
                    ),
                    true
                  ),
                  updated_at_utc = now()
            WHERE id = $1 AND game_profile = $2
            RETURNING *`,
          [account.id, gameProfile, actionType]
        )).rows[0]
        await client.query(
          `INSERT INTO player_account_ownership_history (
             id, game_profile, player_account_id, previous_discord_user_id,
             new_discord_user_id, action_type, performed_by_discord_user_id,
             source_metadata
           ) VALUES ($1, $2, $3, $4, NULL, $5, $6, $7)`,
          [
            crypto.randomUUID(), gameProfile, account.id, previousOwner,
            actionType, performedByDiscordUserId, { source: "discord" }
          ]
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
            [gameProfile, previousOwner]
          )).rows[0] || null
        }
        await client.query("COMMIT")
        return {
          account: released,
          previousOwnerDiscordUserId: previousOwner,
          guildIds,
          replacement
        }
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
