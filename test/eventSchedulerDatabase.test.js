const test = require("node:test")
const assert = require("node:assert/strict")
const { Pool } = require("pg")

const { createEventSchedulerRepository } = require("../src/eventSchedulerRepository")

const databaseUrl = process.env.TEST_DATABASE_URL

test("Phase 3 Postgres transactions and profile isolation", {
  skip: databaseUrl ? false : "TEST_DATABASE_URL is not configured"
}, async () => {
  const pool = new Pool({ connectionString: databaseUrl, max: 4 })
  const guildId = "999999999999999991"
  const baseEvent = {
    allianceId: null,
    allianceName: "Integration Alliance",
    eventName: "Integration Event",
    firstOccurrenceDate: "2026-08-01",
    eventTimeUtc: "18:30",
    groups: [],
    recurrenceDays: 7,
    advanceReminderMinutes: null,
    reminderAtStart: true,
    publishToAlliance: true,
    publishToState: false,
    includeInWeeklyRoundup: false,
    image: null
  }

  try {
    await pool.query(
      "DELETE FROM event_guild_settings WHERE guild_id = $1 AND game_profile IN ('wos', 'kingshot')",
      [guildId]
    )
    await pool.query(
      `INSERT INTO event_guild_settings (
         guild_id, game_profile, bot_instance_name, alliance_name, event_channel_id
       ) VALUES
         ($1, 'wos', 'rachie-wos', 'WOS Integration', '999999999999999992'),
         ($1, 'kingshot', 'peggie-kingshot', 'Kingshot Integration', '999999999999999993')`,
      [guildId]
    )
    const alliances = await pool.query(
      `INSERT INTO event_alliances (
         guild_id, game_profile, alliance_name, is_default, created_by_bot_instance
       ) VALUES
         ($1, 'wos', 'WOS Integration', true, 'rachie-wos'),
         ($1, 'kingshot', 'Kingshot Integration', true, 'peggie-kingshot')
       RETURNING id, game_profile, alliance_name`,
      [guildId]
    )
    const wosAlliance = alliances.rows.find(row => row.game_profile === "wos")
    const kingshotAlliance = alliances.rows.find(row => row.game_profile === "kingshot")

    const wos = createEventSchedulerRepository(pool, "wos")
    const kingshot = createEventSchedulerRepository(pool, "kingshot")
    const wosEvent = await wos.createEvent({
      guildId,
      createdByUserId: "999999999999999994",
      createdByBotInstance: "rachie-wos",
      event: {
        ...baseEvent,
        allianceId: String(wosAlliance.id),
        allianceName: wosAlliance.alliance_name,
        eventName: "WOS Event"
      }
    })
    const kingshotEvent = await kingshot.createEvent({
      guildId,
      createdByUserId: "999999999999999994",
      createdByBotInstance: "peggie-kingshot",
      event: {
        ...baseEvent,
        allianceId: String(kingshotAlliance.id),
        allianceName: kingshotAlliance.alliance_name,
        eventName: "Kingshot Event"
      }
    })

    const wosList = await wos.listEvents(guildId)
    const kingshotList = await kingshot.listEvents(guildId)
    assert.deepEqual(wosList.events.map(event => event.event_name), ["WOS Event"])
    assert.deepEqual(kingshotList.events.map(event => event.event_name), ["Kingshot Event"])
    assert.equal(await wos.getEvent(guildId, kingshotEvent.id), null)
    assert.equal(await kingshot.getEvent(guildId, wosEvent.id), null)

    await assert.rejects(
      wos.createEvent({
        guildId,
        createdByUserId: "999999999999999994",
        createdByBotInstance: "rachie-wos",
        event: {
          ...baseEvent,
          allianceId: String(wosAlliance.id),
          eventName: "Duplicate Group Rollback",
          eventTimeUtc: null,
          groups: [
            { groupName: "Alpha", eventTimeUtc: "18:00", sortOrder: 0 },
            { groupName: "Alpha", eventTimeUtc: "19:00", sortOrder: 1 }
          ]
        }
      }),
      error => error.code === "23505"
    )
    await assert.rejects(
      wos.createEvent({
        guildId,
        createdByUserId: "999999999999999994",
        createdByBotInstance: "rachie-wos",
        event: {
          ...baseEvent,
          allianceId: String(wosAlliance.id),
          eventName: "Image Rollback",
          image: {
            originalFilename: "not-image.txt",
            contentType: "text/plain",
            byteSize: 3,
            imageData: Buffer.from("bad")
          }
        }
      }),
      error => error.code === "23514"
    )

    const rollbackCheck = await pool.query(
      `SELECT event_name FROM scheduled_events
        WHERE guild_id = $1 AND game_profile = 'wos'
          AND event_name IN ('Duplicate Group Rollback', 'Image Rollback')`,
      [guildId]
    )
    assert.equal(rollbackCheck.rowCount, 0)
  } finally {
    await pool.query(
      "DELETE FROM event_guild_settings WHERE guild_id = $1 AND game_profile IN ('wos', 'kingshot')",
      [guildId]
    ).catch(() => {})
    await pool.end()
  }
})
