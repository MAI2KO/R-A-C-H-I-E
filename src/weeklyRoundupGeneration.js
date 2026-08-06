const { roundupPeriod } = require("./weeklyRoundupCalculation")

function claimsForConfigurations(configurations, { gameProfile, now, graceMinutes }) {
  const claims = []
  const stateClaims = new Map()
  for (const config of configurations || []) {
    if (config.game_profile !== gameProfile) continue
    const period = roundupPeriod(
      now,
      config.weekly_roundup_day,
      config.weekly_roundup_time_utc,
      graceMinutes
    )
    if (!period) continue
    claims.push({
      ...period,
      gameProfile,
      targetKind: "alliance",
      targetGuildId: config.source_guild_id,
      targetChannelId: config.weekly_roundup_channel_id,
      sourceGuildId: config.source_guild_id,
      postWhenEmpty: config.roundup_when_empty === true
    })
    if (
      config.sharing_enabled === true
      && String(config.state_guild_id || "").trim()
      && String(config.state_event_channel_id || "").trim()
    ) {
      const key = [period.weekStartDate, config.state_guild_id, config.state_event_channel_id].join(":")
      if (!stateClaims.has(key)) {
        stateClaims.set(key, {
          ...period,
          gameProfile,
          targetKind: "state",
          targetGuildId: config.state_guild_id,
          targetChannelId: config.state_event_channel_id,
          sourceGuildId: config.source_guild_id,
          postWhenEmpty: config.roundup_when_empty === true
        })
      }
    }
  }
  return [...claims, ...stateClaims.values()]
}

async function generateMissingRoundupClaims({ repository, gameProfile, now, graceMinutes }) {
  const configurations = await repository.listRoundupConfigurations()
  const claims = claimsForConfigurations(configurations, { gameProfile, now, graceMinutes })
  return repository.insertMissingClaims(claims)
}

module.exports = { claimsForConfigurations, generateMissingRoundupClaims }
