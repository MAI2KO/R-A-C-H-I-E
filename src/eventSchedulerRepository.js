const {
  buildRoundupOccurrences,
  configuredRoundupWindow
} = require("./weeklyRoundupCalculation")

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

  async function withEventGroups(events) {
    const eventIds = events.map(event => event.id)
    let groups = []
    if (eventIds.length) {
      groups = (await pool.query(
        `SELECT id AS group_id, event_id, group_name, event_time_utc, sort_order
           FROM scheduled_event_groups
          WHERE game_profile = $1 AND event_id = ANY($2::bigint[])
          ORDER BY event_id, sort_order, group_name, id`,
        [gameProfile, eventIds]
      )).rows
    }
    const groupsByEvent = new Map()
    for (const group of groups) {
      const key = String(group.event_id)
      if (!groupsByEvent.has(key)) groupsByEvent.set(key, [])
      groupsByEvent.get(key).push(group)
    }
    return events.map(event => ({
      ...event,
      groups: groupsByEvent.get(String(event.id)) || []
    }))
  }

  return {
    async getGuildSettings(guildId) {
      const result = await pool.query(
        `SELECT s.guild_id, s.game_profile, s.bot_instance_name,
                COALESCE(a.alliance_name, s.alliance_name) AS alliance_name,
                a.id AS default_alliance_id,
                s.event_channel_id, s.weekly_roundup_enabled, s.weekly_roundup_day,
                s.weekly_roundup_time_utc, s.weekly_roundup_channel_id,
                s.roundup_when_empty, s.state_roundup_enabled,
                s.weekly_roundup_not_before, s.created_at, s.updated_at,
                (SELECT COUNT(*)::integer FROM event_alliances total
                  WHERE total.guild_id = s.guild_id
                    AND total.game_profile = s.game_profile) AS alliance_count
           FROM event_guild_settings s
           LEFT JOIN event_alliances a
             ON a.guild_id = s.guild_id AND a.game_profile = s.game_profile
            AND a.is_default = true
          WHERE s.guild_id = $1 AND s.game_profile = $2`,
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
        `WITH upserted AS (
           INSERT INTO event_guild_settings (
             guild_id, game_profile, bot_instance_name, alliance_name, event_channel_id
           ) VALUES ($1, $2, $3, $4, $5)
           ON CONFLICT (guild_id, game_profile) DO UPDATE SET
             bot_instance_name = EXCLUDED.bot_instance_name,
             event_channel_id = EXCLUDED.event_channel_id,
             updated_at = now()
           RETURNING *
         ), default_insert AS (
           INSERT INTO event_alliances (
             guild_id, game_profile, alliance_name, is_default, created_by_bot_instance
           )
           SELECT guild_id, game_profile, alliance_name, true, bot_instance_name
             FROM upserted
           ON CONFLICT DO NOTHING
           RETURNING id
         )
         SELECT upserted.*, (SELECT COUNT(*) FROM default_insert) AS defaults_created
           FROM upserted`,
        [guildId, gameProfile, botInstanceName, allianceName, eventChannelId]
      )
      return result.rows[0]
    },

    async upsertGuildIdentity({ guildId, botInstanceName, allianceName }) {
      if (typeof pool.connect !== "function") {
        throw new Error("The Postgres pool does not support transactions")
      }
      const client = await pool.connect()
      try {
        await client.query("BEGIN")
        await client.query(
          `INSERT INTO event_guild_settings (
             guild_id, game_profile, bot_instance_name, alliance_name, event_channel_id
           ) VALUES ($1, $2, $3, $4, NULL)
           ON CONFLICT (guild_id, game_profile) DO UPDATE SET
             bot_instance_name = EXCLUDED.bot_instance_name,
             updated_at = now()`,
          [guildId, gameProfile, botInstanceName, allianceName]
        )
        await client.query(
          `INSERT INTO event_alliances (
             guild_id, game_profile, alliance_name, is_default, created_by_bot_instance
           ) VALUES ($1, $2, $3, true, $4)
           ON CONFLICT DO NOTHING`,
          [guildId, gameProfile, allianceName, botInstanceName]
        )
        const defaultAlliance = (await client.query(
          `UPDATE event_alliances
              SET alliance_name = $3, updated_at = now()
            WHERE guild_id = $1 AND game_profile = $2 AND is_default = true
            RETURNING id`,
          [guildId, gameProfile, allianceName]
        )).rows[0]
        if (!defaultAlliance) throw new Error("The main alliance could not be configured")
        await client.query(
          `UPDATE scheduled_events
              SET alliance_name = $4, updated_at = now()
            WHERE alliance_id = $1 AND guild_id = $2 AND game_profile = $3`,
          [defaultAlliance.id, guildId, gameProfile, allianceName]
        )
        const settings = (await client.query(
          `UPDATE event_guild_settings
              SET alliance_name = $3, updated_at = now()
            WHERE guild_id = $1 AND game_profile = $2
            RETURNING *`,
          [guildId, gameProfile, allianceName]
        )).rows[0]
        await client.query("COMMIT")
        return settings
      } catch (error) {
        await client.query("ROLLBACK")
        throw error
      } finally {
        client.release()
      }
    },

    async setEventChannel({ guildId, eventChannelId }) {
      const result = await pool.query(
        `UPDATE event_guild_settings
            SET event_channel_id = $3, updated_at = now()
          WHERE guild_id = $1 AND game_profile = $2
          RETURNING *`,
        [guildId, gameProfile, eventChannelId]
      )
      return result.rows[0] || null
    },

    async listAlliances(guildId, { limit = 25, offset = 0 } = {}) {
      const result = await pool.query(
        `SELECT a.id, a.guild_id, a.game_profile, a.alliance_name, a.is_default,
                a.created_by_bot_instance, a.created_at, a.updated_at,
                COUNT(*) OVER()::integer AS total_count,
                COUNT(e.id) FILTER (WHERE e.status IN ('active', 'paused'))::integer
                  AS managed_event_count,
                COUNT(e.id)::integer AS total_event_count
           FROM event_alliances a
           LEFT JOIN scheduled_events e
             ON e.alliance_id = a.id AND e.guild_id = a.guild_id
            AND e.game_profile = a.game_profile
          WHERE a.guild_id = $1 AND a.game_profile = $2
          GROUP BY a.id
          ORDER BY a.is_default DESC, lower(a.alliance_name), a.id
          LIMIT $3 OFFSET $4`,
        [guildId, gameProfile, limit, offset]
      )
      return {
        alliances: result.rows,
        total: result.rows[0]?.total_count || 0
      }
    },

    async getAlliance(guildId, allianceId) {
      const result = await pool.query(
        `SELECT id, guild_id, game_profile, alliance_name, is_default,
                created_by_bot_instance, created_at, updated_at
           FROM event_alliances
          WHERE id = $1 AND guild_id = $2 AND game_profile = $3`,
        [allianceId, guildId, gameProfile]
      )
      return result.rows[0] || null
    },

    async createAlliance({ guildId, allianceName, createdByBotInstance }) {
      const result = await pool.query(
        `INSERT INTO event_alliances (
           guild_id, game_profile, alliance_name, is_default, created_by_bot_instance
         ) VALUES ($1, $2, $3, false, $4)
         RETURNING *`,
        [guildId, gameProfile, allianceName, createdByBotInstance]
      )
      return result.rows[0]
    },

    async renameAlliance({ guildId, allianceId, allianceName }) {
      if (typeof pool.connect !== "function") throw new Error("Transactions are required")
      const client = await pool.connect()
      try {
        await client.query("BEGIN")
        const allianceResult = await client.query(
          `UPDATE event_alliances
              SET alliance_name = $4, updated_at = now()
            WHERE id = $1 AND guild_id = $2 AND game_profile = $3
            RETURNING *`,
          [allianceId, guildId, gameProfile, allianceName]
        )
        const alliance = allianceResult.rows[0]
        if (!alliance) {
          await client.query("ROLLBACK")
          return null
        }
        if (alliance.is_default) {
          await client.query(
            `UPDATE event_guild_settings
                SET alliance_name = $3, updated_at = now()
              WHERE guild_id = $1 AND game_profile = $2`,
            [guildId, gameProfile, allianceName]
          )
        }
        const events = await client.query(
          `UPDATE scheduled_events
              SET alliance_name = $4, schedule_version = schedule_version + 1,
                  updated_at = now()
            WHERE alliance_id = $1 AND guild_id = $2 AND game_profile = $3
            RETURNING id`,
          [allianceId, guildId, gameProfile, allianceName]
        )
        if (events.rowCount) {
          await client.query(
            `UPDATE event_delivery_claims d
                SET status = 'failed', next_attempt_at = NULL,
                    claimed_by_bot_instance = NULL, claimed_by_worker = NULL,
                    claimed_at = NULL, claimed_until = NULL,
                    last_error = 'Event alliance changed.', updated_at = now()
              FROM scheduled_events e
             WHERE d.event_id = e.id AND d.game_profile = e.game_profile
               AND e.alliance_id = $1 AND e.guild_id = $2 AND e.game_profile = $3
               AND d.status <> 'sent'`,
            [allianceId, guildId, gameProfile]
          )
        }
        await client.query("COMMIT")
        return alliance
      } catch (error) {
        await client.query("ROLLBACK")
        throw error
      } finally {
        client.release()
      }
    },

    async deleteAlliance({ guildId, allianceId }) {
      if (typeof pool.connect !== "function") throw new Error("Transactions are required")
      const client = await pool.connect()
      try {
        await client.query("BEGIN")
        const alliance = (await client.query(
          `SELECT id, alliance_name, is_default
             FROM event_alliances
            WHERE id = $1 AND guild_id = $2 AND game_profile = $3
            FOR UPDATE`,
          [allianceId, guildId, gameProfile]
        )).rows[0]
        if (!alliance) {
          await client.query("ROLLBACK")
          return { deleted: false, reason: "missing" }
        }
        if (alliance.is_default) {
          await client.query("ROLLBACK")
          return { deleted: false, reason: "default", alliance }
        }
        const eventCount = Number((await client.query(
          `SELECT COUNT(*)::integer AS count FROM scheduled_events
            WHERE alliance_id = $1 AND guild_id = $2 AND game_profile = $3`,
          [allianceId, guildId, gameProfile]
        )).rows[0].count)
        if (eventCount > 0) {
          await client.query("ROLLBACK")
          return { deleted: false, reason: "events", alliance, eventCount }
        }
        await client.query(
          `DELETE FROM event_alliances
            WHERE id = $1 AND guild_id = $2 AND game_profile = $3`,
          [allianceId, guildId, gameProfile]
        )
        await client.query("COMMIT")
        return { deleted: true, alliance }
      } catch (error) {
        await client.query("ROLLBACK")
        throw error
      } finally {
        client.release()
      }
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

    async getStateDestination(stateGuildId) {
      const result = await pool.query(
        `SELECT state_guild_id, game_profile, configured_by_bot_instance,
                state_roundup_channel_id, enabled, created_at, updated_at
           FROM event_state_destinations
          WHERE state_guild_id = $1 AND game_profile = $2`,
        [stateGuildId, gameProfile]
      )
      return result.rows[0] || null
    },

    async upsertStateDestination({
      stateGuildId,
      configuredByBotInstance,
      stateRoundupChannelId
    }) {
      const result = await pool.query(
        `WITH destination AS (
           INSERT INTO event_state_destinations (
             state_guild_id, game_profile, configured_by_bot_instance,
             state_roundup_channel_id, enabled
           ) VALUES ($1, $2, $3, $4, true)
           ON CONFLICT (state_guild_id, game_profile) DO UPDATE SET
             configured_by_bot_instance = EXCLUDED.configured_by_bot_instance,
             state_roundup_channel_id = EXCLUDED.state_roundup_channel_id,
             enabled = true,
             updated_at = now()
           RETURNING *
         ), updated_links AS (
           UPDATE event_state_links link
              SET state_event_channel_id = destination.state_roundup_channel_id,
                  updated_at = now()
             FROM destination
            WHERE link.state_guild_id = destination.state_guild_id
              AND link.game_profile = destination.game_profile
           RETURNING link.alliance_guild_id
         )
         SELECT destination.* FROM destination`,
        [stateGuildId, gameProfile, configuredByBotInstance, stateRoundupChannelId]
      )
      return result.rows[0]
    },

    async createStateLinkCode({
      stateGuildId,
      codeHash,
      createdByBotInstance,
      createdByUserId,
      expiresAt
    }) {
      const result = await pool.query(
         `INSERT INTO event_state_link_codes (
           game_profile, state_guild_id, code_hash, created_by_bot_instance,
           created_by_user_id, expires_at
         )
         SELECT $2::varchar, d.state_guild_id, $3::char(64), $4::varchar, $5::varchar, $6::timestamptz
           FROM event_state_destinations d
          WHERE d.state_guild_id = $1::varchar
            AND d.game_profile = $2::varchar
            AND d.enabled = true
         RETURNING id, game_profile, state_guild_id, expires_at`,
        [stateGuildId, gameProfile, codeHash, createdByBotInstance, createdByUserId, expiresAt]
      )
      return result.rows[0] || null
    },

    async consumeStateLinkCode({ allianceGuildId, configuredByBotInstance, codeHash }) {
      if (typeof pool.connect !== "function") {
        throw new Error("The Postgres pool does not support transactions")
      }
      const client = await pool.connect()
      try {
        await client.query("BEGIN")
        const code = (await client.query(
          `SELECT c.id, c.state_guild_id, d.state_roundup_channel_id
             FROM event_state_link_codes c
             JOIN event_state_destinations d
               ON d.state_guild_id = c.state_guild_id
              AND d.game_profile = c.game_profile
            WHERE c.code_hash = $1
              AND c.game_profile = $2
              AND c.consumed_at IS NULL
              AND c.expires_at > now()
              AND d.enabled = true
            FOR UPDATE OF c`,
          [codeHash, gameProfile]
        )).rows[0]
        if (!code) {
          await client.query("ROLLBACK")
          return null
        }
        const link = (await client.query(
          `INSERT INTO event_state_links (
             alliance_guild_id, game_profile, configured_by_bot_instance,
             state_guild_id, state_event_channel_id, sharing_enabled
           ) VALUES ($1, $2, $3, $4, $5, true)
           ON CONFLICT (alliance_guild_id, game_profile) DO UPDATE SET
             configured_by_bot_instance = EXCLUDED.configured_by_bot_instance,
             state_guild_id = EXCLUDED.state_guild_id,
             state_event_channel_id = EXCLUDED.state_event_channel_id,
             sharing_enabled = true,
             updated_at = now()
           RETURNING *`,
          [
            allianceGuildId,
            gameProfile,
            configuredByBotInstance,
            code.state_guild_id,
            code.state_roundup_channel_id
          ]
        )).rows[0]
        await client.query(
          `UPDATE event_state_link_codes
              SET consumed_at = now(), consumed_by_alliance_guild_id = $2
            WHERE id = $1`,
          [code.id, allianceGuildId]
        )
        await client.query("COMMIT")
        return link
      } catch (error) {
        await client.query("ROLLBACK")
        throw error
      } finally {
        client.release()
      }
    },

    async configureWeeklyRoundup({
      guildId,
      enabled,
      weekday,
      timeUtc,
      channelId,
      postWhenEmpty,
      stateEnabled = null
    }) {
      if (typeof pool.connect !== "function") {
        throw new Error("The Postgres pool does not support transactions")
      }
      const client = await pool.connect()
      try {
        await client.query("BEGIN")
        const result = await client.query(
          `UPDATE event_guild_settings
              SET weekly_roundup_enabled = $3,
                  weekly_roundup_day = $4,
                  weekly_roundup_time_utc = $5,
                  weekly_roundup_channel_id = $6,
                  roundup_when_empty = $7,
                  state_roundup_enabled = COALESCE($8, state_roundup_enabled),
                  weekly_roundup_not_before = now(),
                  updated_at = now()
            WHERE guild_id = $1 AND game_profile = $2
            RETURNING *`,
          [
            guildId,
            gameProfile,
            enabled,
            weekday,
            timeUtc,
            channelId,
            postWhenEmpty,
            stateEnabled
          ]
        )
        if (result.rows[0]) {
          await client.query(
            `UPDATE weekly_roundup_claims
                SET status = 'failed', next_attempt_at = NULL,
                    claimed_by_bot_instance = NULL, claimed_by_worker = NULL,
                    claimed_at = NULL, claimed_until = NULL,
                    last_error = 'Roundup configuration changed.', updated_at = now()
              WHERE source_guild_id = $1 AND game_profile = $2 AND status <> 'sent'`,
            [guildId, gameProfile]
          )
        }
        await client.query("COMMIT")
        return result.rows[0] || null
      } catch (error) {
        await client.query("ROLLBACK")
        throw error
      } finally {
        client.release()
      }
    },

    async getWeeklyRoundupPreview(guildId, { now = new Date() } = {}) {
      const settings = await this.getGuildSettings(guildId)
      if (!settings) return null
      const eventResult = await pool.query(
        `SELECT e.*, a.alliance_name AS alliance_name,
                a.is_default AS is_default_alliance,
                e.first_occurrence_date::text AS first_occurrence_date
           FROM scheduled_events e
           JOIN event_alliances a
             ON a.id = e.alliance_id AND a.guild_id = e.guild_id
            AND a.game_profile = e.game_profile
          WHERE e.guild_id = $1 AND e.game_profile = $2
            AND e.status = 'active' AND e.include_in_weekly_roundup = true
          ORDER BY e.id`,
        [guildId, gameProfile]
      )
      const events = await withEventGroups(eventResult.rows)
      const { weekStart, weekEnd } = configuredRoundupWindow(
        now,
        settings.weekly_roundup_day
      )
      return {
        claim: {
          gameProfile,
          targetKind: "alliance",
          targetGuildId: guildId,
          targetChannelId: settings.weekly_roundup_channel_id,
          targetIsCurrent: true,
          weekStart,
          weekEnd,
          postWhenEmpty: settings.roundup_when_empty === true
        },
        allianceName: settings.alliance_name,
        occurrences: buildRoundupOccurrences(events, weekStart, weekEnd)
      }
    },

    async getStateWeeklyRoundupPreview(guildId, { now = new Date() } = {}) {
      const target = (await pool.query(
        `SELECT s.alliance_name, s.weekly_roundup_day, s.roundup_when_empty,
                l.state_guild_id, l.state_event_channel_id
           FROM event_guild_settings s
           JOIN event_state_links l
             ON l.alliance_guild_id = s.guild_id AND l.game_profile = s.game_profile
           JOIN event_state_destinations d
             ON d.state_guild_id = l.state_guild_id AND d.game_profile = l.game_profile
            AND d.state_roundup_channel_id = l.state_event_channel_id
          WHERE s.guild_id = $1 AND s.game_profile = $2
            AND s.state_roundup_enabled = true
            AND l.sharing_enabled = true AND d.enabled = true`,
        [guildId, gameProfile]
      )).rows[0]
      if (!target) return null

      const eventResult = await pool.query(
        `SELECT DISTINCT e.*, a.alliance_name AS alliance_name,
                a.is_default AS is_default_alliance,
                e.first_occurrence_date::text AS first_occurrence_date
           FROM scheduled_events e
           JOIN event_alliances a
             ON a.id = e.alliance_id AND a.guild_id = e.guild_id
            AND a.game_profile = e.game_profile
           JOIN event_guild_settings s
             ON s.guild_id = e.guild_id AND s.game_profile = e.game_profile
           JOIN event_state_links l
             ON l.alliance_guild_id = e.guild_id AND l.game_profile = e.game_profile
           JOIN event_state_destinations d
             ON d.state_guild_id = l.state_guild_id AND d.game_profile = l.game_profile
            AND d.state_roundup_channel_id = l.state_event_channel_id
          WHERE e.game_profile = $1 AND e.status = 'active'
            AND e.include_in_weekly_roundup = true
            AND s.state_roundup_enabled = true
            AND l.sharing_enabled = true AND d.enabled = true
            AND l.state_guild_id = $2 AND l.state_event_channel_id = $3
          ORDER BY e.id`,
        [gameProfile, target.state_guild_id, target.state_event_channel_id]
      )
      const events = await withEventGroups(eventResult.rows)
      const { weekStart, weekEnd } = configuredRoundupWindow(
        now,
        target.weekly_roundup_day
      )
      return {
        claim: {
          gameProfile,
          targetKind: "state",
          targetGuildId: target.state_guild_id,
          targetChannelId: target.state_event_channel_id,
          targetIsCurrent: true,
          weekStart,
          weekEnd,
          postWhenEmpty: target.roundup_when_empty === true
        },
        allianceName: target.alliance_name,
        occurrences: buildRoundupOccurrences(events, weekStart, weekEnd)
      }
    },

    async createEvent({ guildId, createdByUserId, createdByBotInstance, event }) {
      if (typeof pool.connect !== "function") {
        throw new Error("The Postgres pool does not support transactions")
      }

      const client = await pool.connect()
      try {
        await client.query("BEGIN")
        const alliance = (await client.query(
          `SELECT id, alliance_name
             FROM event_alliances
            WHERE id = $1 AND guild_id = $2 AND game_profile = $3
            FOR SHARE`,
          [event.allianceId, guildId, gameProfile]
        )).rows[0]
        if (!alliance) throw new Error("The selected alliance is no longer available")
        const eventResult = await client.query(
          `INSERT INTO scheduled_events (
             guild_id, game_profile, created_by_bot_instance, alliance_id,
             alliance_name, event_name, first_occurrence_date, event_time_utc,
             recurrence_days, image_url, advance_reminder_minutes,
             reminder_at_start, advance_reminder_message, final_reminder_message,
             publish_to_alliance, publish_to_state, include_in_weekly_roundup,
             status, created_by_user_id
           ) VALUES (
             $1, $2, $3, $4, $5, $6, $7, $8, $9, NULL, $10, $11, $12,
             $13, $14, $15, $16, 'active', $17
           )
           RETURNING *`,
          [
            guildId,
            gameProfile,
            createdByBotInstance,
            alliance.id,
            alliance.alliance_name,
            event.eventName,
            event.firstOccurrenceDate,
            event.eventTimeUtc,
            event.recurrenceDays,
            event.advanceReminderMinutes,
            event.reminderAtStart,
            event.advanceReminderMessage,
            event.finalReminderMessage,
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

    async listEvents(guildId, { limit = 10, offset = 0, allianceId = null } = {}) {
      const result = await pool.query(
        `SELECT e.*, a.alliance_name AS alliance_name,
                e.first_occurrence_date::text AS first_occurrence_date,
                EXISTS (
                  SELECT 1 FROM scheduled_event_images i
                   WHERE i.event_id = e.id AND i.game_profile = e.game_profile
                ) AS has_image,
                COUNT(*) OVER()::integer AS total_count
           FROM scheduled_events e
           JOIN event_alliances a
             ON a.id = e.alliance_id AND a.guild_id = e.guild_id
            AND a.game_profile = e.game_profile
          WHERE e.guild_id = $1
            AND e.game_profile = $2
            AND e.status IN ('active', 'paused')
            AND ($5::bigint IS NULL OR e.alliance_id = $5)
          ORDER BY lower(a.alliance_name), e.first_occurrence_date, e.event_name, e.id
          LIMIT $3 OFFSET $4`,
        [guildId, gameProfile, limit, offset, allianceId]
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
        `SELECT e.*, a.alliance_name AS alliance_name,
                e.first_occurrence_date::text AS first_occurrence_date,
                i.original_filename AS image_filename,
                i.content_type AS image_content_type,
                i.byte_size AS image_byte_size,
                i.image_data
           FROM scheduled_events e
           JOIN event_alliances a
             ON a.id = e.alliance_id AND a.guild_id = e.guild_id
            AND a.game_profile = e.game_profile
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
          `SELECT e.id, e.schedule_version, a.id AS selected_alliance_id,
                  a.alliance_name AS selected_alliance_name
             FROM scheduled_events e
             JOIN event_alliances a
               ON a.id = $4 AND a.guild_id = e.guild_id
              AND a.game_profile = e.game_profile
            WHERE e.id = $1 AND e.guild_id = $2 AND e.game_profile = $3
              AND e.status IN ('active', 'paused')
            FOR UPDATE OF e`,
          [eventId, guildId, gameProfile, event.allianceId]
        )
        if (!existing.rowCount) {
          await client.query("ROLLBACK")
          return null
        }
        const selectedAlliance = existing.rows[0]
        const updated = await client.query(
          `UPDATE scheduled_events
              SET alliance_id = $4, alliance_name = $5, event_name = $6,
                  first_occurrence_date = $7, event_time_utc = $8,
                  recurrence_days = $9, advance_reminder_minutes = $10,
                  reminder_at_start = $11, advance_reminder_message = $12,
                  final_reminder_message = $13, publish_to_alliance = $14,
                  publish_to_state = $15, include_in_weekly_roundup = $16,
                  schedule_version = schedule_version + 1, updated_at = now()
            WHERE id = $1 AND guild_id = $2 AND game_profile = $3
            RETURNING *`,
          [
            eventId, guildId, gameProfile,
            selectedAlliance.selected_alliance_id,
            selectedAlliance.selected_alliance_name,
            event.eventName,
            event.firstOccurrenceDate, event.eventTimeUtc, event.recurrenceDays,
            event.advanceReminderMinutes, event.reminderAtStart,
            event.advanceReminderMessage, event.finalReminderMessage,
            event.publishToAlliance, event.publishToState,
            event.includeInWeeklyRoundup
          ]
        )
        await client.query(
          `UPDATE event_delivery_claims d
              SET group_id_snapshot = COALESCE(d.group_id_snapshot, g.id),
                  group_name_snapshot = COALESCE(d.group_name_snapshot, btrim(g.group_name)),
                  group_id = NULL
             FROM scheduled_event_groups g
            WHERE d.group_id = g.id AND d.event_id = g.event_id
              AND d.game_profile = g.game_profile
              AND d.event_id = $1 AND d.game_profile = $2`,
          [eventId, gameProfile]
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
            WHERE event_id = $1 AND game_profile = $2 AND status <> 'sent'`,
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
            WHERE event_id = $1 AND game_profile = $2 AND status <> 'sent'`,
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
