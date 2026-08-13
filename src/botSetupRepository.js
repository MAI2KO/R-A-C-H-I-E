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
           minister_sign_up_message_id, event_scheduler_message_id
         ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11)
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
           updated_at_utc = now()
         RETURNING *`,
        [
          gameProfile, guildId, values.category_id,
          values.gift_auto_redeem_channel_id, values.gift_announcements_channel_id,
          values.minister_sign_up_channel_id, values.event_scheduler_channel_id,
          values.event_announcements_channel_id, values.gift_auto_redeem_message_id,
          values.minister_sign_up_message_id, values.event_scheduler_message_id
        ]
      )).rows[0]
    },

    async reconcileDestinations(guildId, giftChannelId, eventChannelId) {
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
        `UPDATE event_guild_settings
            SET event_channel_id = $3,
                weekly_roundup_channel_id = $3,
                updated_at = now()
          WHERE guild_id = $1 AND game_profile = $2`,
        [guildId, gameProfile, eventChannelId]
      )
    }
  })
}

module.exports = { createBotSetupRepository }
