const test = require("node:test")
const assert = require("node:assert/strict")
const { Pool } = require("pg")

const { createEventSchedulerRepository } = require("../src/eventSchedulerRepository")
const { createPublicAllianceEventRepository } = require("../src/publicAllianceEventRepository")

const databaseUrl = process.env.TEST_DATABASE_URL

test("Postgres guild schedule is alliance-only, active, profile isolated, and needs no State destination", {
  skip: databaseUrl ? false : "TEST_DATABASE_URL is not configured"
}, async () => {
  const pool = new Pool({ connectionString: databaseUrl, max: 4 })
  const stateGuild = "777777777777777700"
  const guildA = "777777777777777701"
  const guildB = "777777777777777702"
  const user = "777777777777777703"
  const channel = "777777777777777704"
  try {
    await pool.query("SELECT pg_advisory_lock(7000021)")
    await pool.query("DELETE FROM event_state_destinations WHERE state_guild_id = $1", [stateGuild])
    await pool.query("DELETE FROM event_guild_settings WHERE guild_id = ANY($1::varchar[])", [[guildA, guildB]])
    await pool.query(`INSERT INTO event_guild_settings
      (guild_id, game_profile, bot_instance_name, alliance_name, event_channel_id) VALUES
      ($1, 'wos', 'test-wos', 'Zulu', $3), ($2, 'wos', 'test-wos', 'Alpha', $3),
      ($1, 'kingshot', 'test-kingshot', 'Kings', $3)`, [guildA, guildB, channel])
    const alliances = (await pool.query(`INSERT INTO event_alliances
      (guild_id, game_profile, alliance_name, is_default, created_by_bot_instance) VALUES
      ($1, 'wos', 'Zulu', true, 'test-wos'), ($2, 'wos', 'Alpha', true, 'test-wos'),
      ($1, 'kingshot', 'Kings', true, 'test-kingshot') RETURNING id, guild_id, game_profile, alliance_name`,
    [guildA, guildB])).rows
    async function create(profile, guildId, name, status = "active") {
      const identity = alliances.find(row => row.game_profile === profile && row.guild_id === guildId)
      const created = await createEventSchedulerRepository(pool, profile).createEvent({
        guildId, createdByUserId: user, createdByBotInstance: `test-${profile}`,
        event: { allianceId: String(identity.id), allianceName: identity.alliance_name, eventName: name,
          firstOccurrenceDate: "2026-08-24", eventTimeUtc: "19:00", groups: [], recurrenceDays: 2,
          advanceReminderMinutes: null, reminderAtStart: true, publishToAlliance: true,
          publishToState: false, includeInWeeklyRoundup: false, image: null }
      })
      if (status !== "active") await pool.query("UPDATE scheduled_events SET status = $1 WHERE id = $2 AND game_profile = $3", [status, created.id, profile])
      return created
    }
    await create("wos", guildA, "Zulu Active")
    await create("wos", guildA, "Paused", "paused")
    await create("wos", guildA, "Deleted", "deleted")
    await create("wos", guildB, "Alpha Active")
    await create("kingshot", guildA, "Kings Active")
    const wosRepository = createPublicAllianceEventRepository(pool, "wos")
    const wos = [
      ...await wosRepository.listForGuild(guildA),
      ...await wosRepository.listForGuild(guildB)
    ]
    const kingshot = await createPublicAllianceEventRepository(pool, "kingshot").listForGuild(guildA)
    assert.deepEqual(wos.map(row => row.event_name).sort(), ["Alpha Active", "Zulu Active"])
    assert.deepEqual(kingshot.map(row => row.event_name), ["Kings Active"])
    assert.equal(wos.some(row => ["Paused", "Deleted"].includes(row.event_name)), false)
    assert.equal((await pool.query(
      "SELECT count(*)::int AS count FROM event_state_destinations WHERE state_guild_id = $1",
      [stateGuild]
    )).rows[0].count, 0)
  } finally {
    await pool.query("DELETE FROM event_state_destinations WHERE state_guild_id = $1", [stateGuild]).catch(() => {})
    await pool.query("DELETE FROM event_guild_settings WHERE guild_id = ANY($1::varchar[])", [[guildA, guildB]]).catch(() => {})
    await pool.query("SELECT pg_advisory_unlock(7000021)").catch(() => {})
    await pool.end()
  }
})
