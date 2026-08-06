const SUPPORTED_PROFILES = new Set(["wos", "kingshot"])

function assertProfile(gameProfile) {
  if (!SUPPORTED_PROFILES.has(gameProfile)) {
    throw new Error("Unsupported scheduler game profile")
  }
}

function createEventSchedulerRepository(pool, gameProfile) {
  if (!pool || typeof pool.query !== "function") {
    throw new Error("A Postgres pool is required")
  }
  assertProfile(gameProfile)

  return {
    async getGuildSettings(guildId) {
      const result = await pool.query(
        `SELECT guild_id, game_profile, bot_instance_name, alliance_name,
                event_channel_id, created_at, updated_at
           FROM event_guild_settings
          WHERE guild_id = $1 AND game_profile = $2`,
        [guildId, gameProfile]
      )
      return result.rows[0] || null
    },

    async upsertGuildSettings({
      guildId,
      botInstanceName,
      allianceName,
      eventChannelId
    }) {
      const result = await pool.query(
        `INSERT INTO event_guild_settings (
           guild_id, game_profile, bot_instance_name, alliance_name,
           event_channel_id
         ) VALUES ($1, $2, $3, $4, $5)
         ON CONFLICT (guild_id, game_profile) DO UPDATE SET
           bot_instance_name = EXCLUDED.bot_instance_name,
           alliance_name = EXCLUDED.alliance_name,
           event_channel_id = EXCLUDED.event_channel_id,
           updated_at = now()
         RETURNING *`,
        [guildId, gameProfile, botInstanceName, allianceName, eventChannelId]
      )
      return result.rows[0]
    },

    async getStateLink(allianceGuildId) {
      const result = await pool.query(
        `SELECT alliance_guild_id, game_profile, configured_by_bot_instance,
                state_guild_id, state_event_channel_id, sharing_enabled,
                created_at, updated_at
           FROM event_state_links
          WHERE alliance_guild_id = $1 AND game_profile = $2`,
        [allianceGuildId, gameProfile]
      )
      return result.rows[0] || null
    },

    async upsertStateLink({
      allianceGuildId,
      configuredByBotInstance,
      stateGuildId,
      stateEventChannelId,
      sharingEnabled = true
    }) {
      const result = await pool.query(
        `INSERT INTO event_state_links (
           alliance_guild_id, game_profile, configured_by_bot_instance,
           state_guild_id, state_event_channel_id, sharing_enabled
         ) VALUES ($1, $2, $3, $4, $5, $6)
         ON CONFLICT (alliance_guild_id, game_profile) DO UPDATE SET
           configured_by_bot_instance = EXCLUDED.configured_by_bot_instance,
           state_guild_id = EXCLUDED.state_guild_id,
           state_event_channel_id = EXCLUDED.state_event_channel_id,
           sharing_enabled = EXCLUDED.sharing_enabled,
           updated_at = now()
         RETURNING *`,
        [
          allianceGuildId,
          gameProfile,
          configuredByBotInstance,
          stateGuildId,
          stateEventChannelId,
          sharingEnabled
        ]
      )
      return result.rows[0]
    },

    async setStateSharing(allianceGuildId, enabled) {
      const result = await pool.query(
        `UPDATE event_state_links
            SET sharing_enabled = $3, updated_at = now()
          WHERE alliance_guild_id = $1 AND game_profile = $2
          RETURNING *`,
        [allianceGuildId, gameProfile, enabled]
      )
      return result.rows[0] || null
    }
  }
}

module.exports = {
  SUPPORTED_PROFILES,
  assertProfile,
  createEventSchedulerRepository
}
