function createBotSetupRepository(pool, gameProfile) {
  if (!pool?.query) throw new Error("A PostgreSQL pool is required")
  if (!["wos", "kingshot"].includes(gameProfile)) throw new Error("Unsupported game profile")

  return Object.freeze({
    async get(guildId) {
      return (await pool.query(
        `SELECT * FROM bot_managed_discord_setups
          WHERE game_profile = $1 AND guild_id = $2`,
        [gameProfile, guildId]
      )).rows[0] || null
    },

    async save(guildId, values) {
      return (await pool.query(
        `INSERT INTO bot_managed_discord_setups (
           game_profile, guild_id, category_id,
           gift_auto_redeem_channel_id, gift_announcements_channel_id,
           minister_sign_up_channel_id, event_scheduler_channel_id,
           event_announcements_channel_id, gift_auto_redeem_message_id,
           minister_sign_up_message_id, event_scheduler_message_id,
           community_number, discord_guild_name, alliance_abbreviation
         ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14)
         ON CONFLICT (game_profile, guild_id) DO UPDATE SET
           category_id = EXCLUDED.category_id,
           gift_auto_redeem_channel_id = EXCLUDED.gift_auto_redeem_channel_id,
           gift_announcements_channel_id = EXCLUDED.gift_announcements_channel_id,
           minister_sign_up_channel_id = EXCLUDED.minister_sign_up_channel_id,
           event_scheduler_channel_id = EXCLUDED.event_scheduler_channel_id,
           event_announcements_channel_id = EXCLUDED.event_announcements_channel_id,
           gift_auto_redeem_message_id = EXCLUDED.gift_auto_redeem_message_id,
           minister_sign_up_message_id = EXCLUDED.minister_sign_up_message_id,
           event_scheduler_message_id = EXCLUDED.event_scheduler_message_id,
           community_number = EXCLUDED.community_number,
           discord_guild_name = EXCLUDED.discord_guild_name,
           alliance_abbreviation = EXCLUDED.alliance_abbreviation,
           updated_at_utc = now()
         RETURNING *`,
        [
          gameProfile, guildId, values.category_id,
          values.gift_auto_redeem_channel_id, values.gift_announcements_channel_id,
          values.minister_sign_up_channel_id, values.event_scheduler_channel_id,
          values.event_announcements_channel_id, values.gift_auto_redeem_message_id,
          values.minister_sign_up_message_id, values.event_scheduler_message_id,
          values.community_number, values.discord_guild_name, values.alliance_abbreviation
        ]
      )).rows[0]
    },

    async getManagerRole(guildId) {
      return (await pool.query(
        `SELECT bot_manager_role_id FROM bot_managed_discord_setups
          WHERE game_profile = $1 AND guild_id = $2`,
        [gameProfile, guildId]
      )).rows[0]?.bot_manager_role_id || null
    },

    async setManagerRole(guildId, roleId) {
      return (await pool.query(
        `INSERT INTO bot_managed_discord_setups (
           game_profile, guild_id, bot_manager_role_id
         ) VALUES ($1, $2, $3)
         ON CONFLICT (game_profile, guild_id) DO UPDATE SET
           bot_manager_role_id = EXCLUDED.bot_manager_role_id,
           updated_at_utc = now()
         RETURNING bot_manager_role_id`,
        [gameProfile, guildId, roleId]
      )).rows[0]?.bot_manager_role_id || null
    },

    async clearManagerRole(guildId) {
      await pool.query(
        `UPDATE bot_managed_discord_setups
            SET bot_manager_role_id = NULL, updated_at_utc = now()
          WHERE game_profile = $1 AND guild_id = $2`,
        [gameProfile, guildId]
      )
    },

    async getDestinations(guildId) {
      const [gift, event] = await Promise.all([
        pool.query(
          `SELECT gift_code_channel_id FROM gift_code_guild_settings
            WHERE game_profile = $1 AND guild_id = $2`,
          [gameProfile, guildId]
        ),
        pool.query(
          `SELECT event_channel_id, weekly_roundup_channel_id FROM event_guild_settings
            WHERE game_profile = $1 AND guild_id = $2`,
          [gameProfile, guildId]
        )
      ])
      return { gift: gift.rows[0] || null, event: event.rows[0] || null }
    },

    async reconcileDestinations({ guildId, giftChannelId, eventChannelId,
      roundupChannelId, botInstanceName, allianceAbbreviation }) {
      await pool.query(
        `INSERT INTO gift_code_guild_settings (
           game_profile, guild_id, gift_code_channel_id
         ) VALUES ($1, $2, $3)
         ON CONFLICT (game_profile, guild_id) DO UPDATE SET
           gift_code_channel_id = EXCLUDED.gift_code_channel_id,
           updated_at_utc = now()`,
        [gameProfile, guildId, giftChannelId]
      )
      await pool.query(
        `INSERT INTO event_guild_settings (
           guild_id, game_profile, bot_instance_name, alliance_name,
           event_channel_id, weekly_roundup_channel_id
         ) VALUES ($1, $2, $3, $4, $5, $6)
         ON CONFLICT (guild_id, game_profile) DO UPDATE SET
           alliance_name = EXCLUDED.alliance_name,
           event_channel_id = EXCLUDED.event_channel_id,
           weekly_roundup_channel_id = EXCLUDED.weekly_roundup_channel_id,
           updated_at = now()`,
        [guildId, gameProfile, botInstanceName, allianceAbbreviation,
          eventChannelId, roundupChannelId]
      )
    }
  })
}

module.exports = { createBotSetupRepository }
