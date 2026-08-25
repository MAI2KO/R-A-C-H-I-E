const SUPPORTED_PROFILES = new Set(["wos", "kingshot"])

function createPublicAllianceEventRepository(pool, gameProfile) {
  if (!pool || typeof pool.query !== "function") throw new Error("A Postgres pool is required")
  if (!SUPPORTED_PROFILES.has(gameProfile)) throw new Error("Unsupported scheduler game profile")

  return Object.freeze({
    gameProfile,
    async listForCommunity(communityCode, limit = 1000) {
      const boundedLimit = Math.max(1, Math.min(Number(limit) || 1000, 1000))
      const events = (await pool.query(
        `SELECT a.id AS alliance_id, a.alliance_name,
                e.id AS event_id, e.event_name,
                e.first_occurrence_date::text AS first_occurrence_date,
                e.event_time_utc, e.recurrence_days
           FROM event_state_destinations destination
           JOIN event_state_links link
             ON link.state_guild_id = destination.state_guild_id
            AND link.game_profile = destination.game_profile
            AND link.sharing_enabled = true
           JOIN scheduled_events e
             ON e.guild_id = link.alliance_guild_id
            AND e.game_profile = link.game_profile
            AND e.status = 'active'
           JOIN event_alliances a
             ON a.id = e.alliance_id AND a.guild_id = e.guild_id
            AND a.game_profile = e.game_profile
          WHERE destination.game_profile = $1
            AND destination.state_number = $2
            AND destination.enabled = true
          ORDER BY lower(a.alliance_name), a.id, e.id
          LIMIT $3`,
        [gameProfile, communityCode, boundedLimit]
      )).rows
      if (!events.length) return []
      const eventIds = events.map(event => event.event_id)
      const groups = (await pool.query(
        `SELECT event_id, group_name,
                first_occurrence_date::text AS first_occurrence_date,
                event_time_utc, sort_order
           FROM scheduled_event_groups
          WHERE game_profile = $1 AND event_id = ANY($2::bigint[])
          ORDER BY event_id, sort_order, lower(group_name), id`,
        [gameProfile, eventIds]
      )).rows
      const groupsByEvent = new Map()
      for (const group of groups) {
        const key = String(group.event_id)
        if (!groupsByEvent.has(key)) groupsByEvent.set(key, [])
        groupsByEvent.get(key).push(group)
      }
      return events.map(event => ({
        ...event,
        groups: groupsByEvent.get(String(event.event_id)) || []
      }))
    }
  })
}

module.exports = { createPublicAllianceEventRepository }
