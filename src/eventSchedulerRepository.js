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
                event_channel_id, weekly_roundup_enabled, weekly_roundup_day,
                weekly_roundup_time_utc, weekly_roundup_channel_id,
                roundup_when_empty, created_at, updated_at
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
    },

    async configureWeeklyRoundup({
      guildId,
      enabled,
      weekday,
      timeUtc,
      channelId,
      postWhenEmpty
    }) {
      const result = await pool.query(
        `UPDATE event_guild_settings
            SET weekly_roundup_enabled = $3,
                weekly_roundup_day = $4,
                weekly_roundup_time_utc = $5,
                weekly_roundup_channel_id = $6,
                roundup_when_empty = $7,
                updated_at = now()
          WHERE guild_id = $1 AND game_profile = $2
          RETURNING *`,
        [guildId, gameProfile, enabled, weekday, timeUtc, channelId, postWhenEmpty]
      )
      return result.rows[0] || null
    },

    async createEvent({ guildId, createdByUserId, createdByBotInstance, event }) {
      if (typeof pool.connect !== "function") {
        throw new Error("The Postgres pool does not support transactions")
      }

      const client = await pool.connect()
      try {
        await client.query("BEGIN")
        const eventResult = await client.query(
          `INSERT INTO scheduled_events (
             guild_id, game_profile, created_by_bot_instance, alliance_name,
             event_name, first_occurrence_date, event_time_utc,
             recurrence_days, image_url, advance_reminder_minutes,
             reminder_at_start, publish_to_alliance, publish_to_state,
             include_in_weekly_roundup, status, created_by_user_id
           ) VALUES (
             $1, $2, $3, $4, $5, $6, $7, $8, NULL, $9, $10, $11, $12,
             $13, 'active', $14
           )
           RETURNING *`,
          [
            guildId,
            gameProfile,
            createdByBotInstance,
            event.allianceName,
            event.eventName,
            event.firstOccurrenceDate,
            event.eventTimeUtc,
            event.recurrenceDays,
            event.advanceReminderMinutes,
            event.reminderAtStart,
            event.publishToAlliance,
            event.publishToState,
            event.includeInWeeklyRoundup,
            createdByUserId
          ]
        )
        const created = eventResult.rows[0]

        for (const group of event.groups) {
          await client.query(
            `INSERT INTO scheduled_event_groups (
               event_id, game_profile, group_name, event_time_utc, sort_order
             ) VALUES ($1, $2, $3, $4, $5)`,
            [created.id, gameProfile, group.groupName, group.eventTimeUtc, group.sortOrder]
          )
        }

        if (event.image) {
          await client.query(
            `INSERT INTO scheduled_event_images (
               event_id, game_profile, original_filename, content_type,
               byte_size, image_data
             ) VALUES ($1, $2, $3, $4, $5, $6)`,
            [
              created.id,
              gameProfile,
              event.image.originalFilename,
              event.image.contentType,
              event.image.byteSize,
              event.image.imageData
            ]
          )
        }

        await client.query("COMMIT")
        return created
      } catch (error) {
        await client.query("ROLLBACK")
        throw error
      } finally {
        client.release()
      }
    },

    async listEvents(guildId, { limit = 10, offset = 0 } = {}) {
      const result = await pool.query(
        `SELECT e.*, e.first_occurrence_date::text AS first_occurrence_date,
                EXISTS (
                  SELECT 1 FROM scheduled_event_images i
                   WHERE i.event_id = e.id AND i.game_profile = e.game_profile
                ) AS has_image,
                COUNT(*) OVER()::integer AS total_count
           FROM scheduled_events e
          WHERE e.guild_id = $1
            AND e.game_profile = $2
            AND e.status IN ('active', 'paused')
          ORDER BY e.first_occurrence_date, e.event_name, e.id
          LIMIT $3 OFFSET $4`,
        [guildId, gameProfile, limit, offset]
      )

      const events = result.rows
      if (events.length === 0) return { events: [], total: 0 }

      const eventIds = events.map(event => event.id)
      const groupsResult = await pool.query(
        `SELECT g.id AS group_id, g.event_id, g.group_name, g.event_time_utc, g.sort_order
           FROM scheduled_event_groups g
           JOIN scheduled_events e
             ON e.id = g.event_id AND e.game_profile = g.game_profile
          WHERE e.guild_id = $1
            AND e.game_profile = $2
            AND g.event_id = ANY($3::bigint[])
          ORDER BY g.event_id, g.sort_order, g.id`,
        [guildId, gameProfile, eventIds]
      )

      const groupsByEvent = new Map()
      for (const group of groupsResult.rows) {
        const key = String(group.event_id)
        if (!groupsByEvent.has(key)) groupsByEvent.set(key, [])
        groupsByEvent.get(key).push(group)
      }

      return {
        events: events.map(event => ({
          ...event,
          groups: groupsByEvent.get(String(event.id)) || []
        })),
        total: events[0].total_count
      }
    },

    async getEvent(guildId, eventId) {
      const result = await pool.query(
        `SELECT e.*, e.first_occurrence_date::text AS first_occurrence_date,
                i.original_filename AS image_filename,
                i.content_type AS image_content_type,
                i.byte_size AS image_byte_size,
                i.image_data
           FROM scheduled_events e
           LEFT JOIN scheduled_event_images i
             ON i.event_id = e.id AND i.game_profile = e.game_profile
          WHERE e.id = $1 AND e.guild_id = $2 AND e.game_profile = $3
            AND e.status IN ('active', 'paused')`,
        [eventId, guildId, gameProfile]
      )
      const event = result.rows[0]
      if (!event) return null

      const groupsResult = await pool.query(
        `SELECT g.id AS group_id, g.group_name, g.event_time_utc, g.sort_order
           FROM scheduled_event_groups g
           JOIN scheduled_events e
             ON e.id = g.event_id AND e.game_profile = g.game_profile
          WHERE g.event_id = $1 AND e.guild_id = $2 AND e.game_profile = $3
          ORDER BY g.sort_order, g.id`,
        [eventId, guildId, gameProfile]
      )
      return { ...event, groups: groupsResult.rows }
    },

    async updateEvent({ guildId, eventId, event, imageAction = "retain" }) {
      if (typeof pool.connect !== "function") throw new Error("Transactions are required")
      if (!["retain", "replace", "remove"].includes(imageAction)) {
        throw new Error("Unsupported image action")
      }
      const client = await pool.connect()
      try {
        await client.query("BEGIN")
        const existing = await client.query(
          `SELECT id, schedule_version FROM scheduled_events
            WHERE id = $1 AND guild_id = $2 AND game_profile = $3
              AND status IN ('active', 'paused')
            FOR UPDATE`,
          [eventId, guildId, gameProfile]
        )
        if (!existing.rowCount) {
          await client.query("ROLLBACK")
          return null
        }
        const updated = await client.query(
          `UPDATE scheduled_events
              SET alliance_name = $4, event_name = $5,
                  first_occurrence_date = $6, event_time_utc = $7,
                  recurrence_days = $8, advance_reminder_minutes = $9,
                  reminder_at_start = $10, publish_to_alliance = $11,
                  publish_to_state = $12, include_in_weekly_roundup = $13,
                  schedule_version = schedule_version + 1, updated_at = now()
            WHERE id = $1 AND guild_id = $2 AND game_profile = $3
            RETURNING *`,
          [
            eventId, guildId, gameProfile, event.allianceName, event.eventName,
            event.firstOccurrenceDate, event.eventTimeUtc, event.recurrenceDays,
            event.advanceReminderMinutes, event.reminderAtStart,
            event.publishToAlliance, event.publishToState,
            event.includeInWeeklyRoundup
          ]
        )
        await client.query(
          "DELETE FROM scheduled_event_groups WHERE event_id = $1 AND game_profile = $2",
          [eventId, gameProfile]
        )
        for (const group of event.groups) {
          await client.query(
            `INSERT INTO scheduled_event_groups
               (event_id, game_profile, group_name, event_time_utc, sort_order)
             VALUES ($1, $2, $3, $4, $5)`,
            [eventId, gameProfile, group.groupName, group.eventTimeUtc, group.sortOrder]
          )
        }
        if (imageAction === "remove" || imageAction === "replace") {
          await client.query(
            "DELETE FROM scheduled_event_images WHERE event_id = $1 AND game_profile = $2",
            [eventId, gameProfile]
          )
        }
        if (imageAction === "replace") {
          if (!event.image) throw new Error("Replacement image is required")
          await client.query(
            `INSERT INTO scheduled_event_images
               (event_id, game_profile, original_filename, content_type, byte_size, image_data)
             VALUES ($1, $2, $3, $4, $5, $6)`,
            [
              eventId, gameProfile, event.image.originalFilename, event.image.contentType,
              event.image.byteSize, event.image.imageData
            ]
          )
        }
        await client.query(
          `UPDATE event_delivery_claims
              SET status = 'failed', next_attempt_at = NULL,
                  claimed_by_bot_instance = NULL, claimed_by_worker = NULL,
                  claimed_at = NULL, claimed_until = NULL,
                  last_error = 'Event schedule changed.', updated_at = now()
            WHERE event_id = $1 AND game_profile = $2 AND status IN ('pending', 'failed')`,
          [eventId, gameProfile]
        )
        await client.query("COMMIT")
        return updated.rows[0]
      } catch (error) {
        await client.query("ROLLBACK")
        throw error
      } finally {
        client.release()
      }
    },

    async setEventStatus({ guildId, eventId, status }) {
      if (!["active", "paused", "deleted"].includes(status)) {
        throw new Error("Unsupported event status")
      }
      const client = await pool.connect()
      try {
        await client.query("BEGIN")
        const result = await client.query(
          `UPDATE scheduled_events
              SET status = $4, schedule_version = schedule_version + 1, updated_at = now()
            WHERE id = $1 AND guild_id = $2 AND game_profile = $3
              AND status IN ('active', 'paused')
            RETURNING *`,
          [eventId, guildId, gameProfile, status]
        )
        if (!result.rowCount) {
          await client.query("ROLLBACK")
          return null
        }
        await client.query(
          `UPDATE event_delivery_claims
              SET status = 'failed', next_attempt_at = NULL,
                  claimed_by_bot_instance = NULL, claimed_by_worker = NULL,
                  claimed_at = NULL, claimed_until = NULL,
                  last_error = $3, updated_at = now()
            WHERE event_id = $1 AND game_profile = $2 AND status IN ('pending', 'failed')`,
          [eventId, gameProfile, status === "deleted" ? "Event deleted." : "Event status changed."]
        )
        await client.query("COMMIT")
        return result.rows[0]
      } catch (error) {
        await client.query("ROLLBACK")
        throw error
      } finally {
        client.release()
      }
    }
  }
}

module.exports = {
  SUPPORTED_PROFILES,
  assertProfile,
  createEventSchedulerRepository
}
