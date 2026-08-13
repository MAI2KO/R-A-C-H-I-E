const DISCORD_SNOWFLAKE = /^[1-9][0-9]{14,19}$/

function botOwnerIds(env = process.env) {
  return new Set(String(env.BOT_OWNER_IDS || "")
    .split(",")
    .map(value => value.trim())
    .filter(value => DISCORD_SNOWFLAKE.test(value)))
}

function isBotOperator(discordUserId, env = process.env) {
  const userId = String(discordUserId || "").trim()
  return DISCORD_SNOWFLAKE.test(userId) && botOwnerIds(env).has(userId)
}

module.exports = { DISCORD_SNOWFLAKE, botOwnerIds, isBotOperator }
