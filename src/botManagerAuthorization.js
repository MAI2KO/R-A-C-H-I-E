const { PermissionFlagsBits } = require("discord.js")

function interactionIsGuildOwner(interaction) {
  return Boolean(interaction.guild?.ownerId
    && interaction.user?.id
    && interaction.guild.ownerId === interaction.user.id)
}

function interactionIsDiscordAdministrator(interaction) {
  return Boolean(interaction.memberPermissions?.has(PermissionFlagsBits.Administrator))
}

function createBotManagerDecision({ repositoryProvider, logger = console }) {
  if (typeof repositoryProvider !== "function") throw new Error("repositoryProvider is required")
  return async function decide({ guildId, isOwner, isAdministrator, hasRole }) {
    if (isOwner || isAdministrator) {
      return Object.freeze({ status: "authorized", via: "administrator", guildId })
    }
    if (!guildId) return Object.freeze({ status: "denied", reason: "missing_guild" })
    try {
      const repository = repositoryProvider()
      if (!repository) return Object.freeze({ status: "unavailable", reason: "database_unavailable" })
      const roleId = await repository.getManagerRole(guildId)
      return roleId && hasRole(roleId)
        ? Object.freeze({ status: "authorized", via: "bot_manager_role", guildId })
        : Object.freeze({ status: "denied", reason: "insufficient_permissions" })
    } catch (error) {
      logger.error(JSON.stringify({
        event: "bot_manager_authorization_failed",
        error_code: String(error?.code || "database_unavailable").slice(0, 80)
      }))
      return Object.freeze({ status: "unavailable", reason: "database_unavailable" })
    }
  }
}

function createBotManagerAuthorizer({ repositoryProvider, logger = console }) {
  const decide = createBotManagerDecision({ repositoryProvider, logger })
  return Object.freeze({
    async canManage(interaction) {
      const result = await decide({
        guildId: interaction.guildId,
        isOwner: interactionIsGuildOwner(interaction),
        isAdministrator: interactionIsDiscordAdministrator(interaction),
        hasRole: roleId => Boolean(interaction.member?.roles?.cache?.has(roleId))
      })
      return result.status === "authorized"
    }
  })
}

function createLiveBotManagerVerifier({ client, repositoryProvider, logger = console }) {
  if (!client?.guilds) throw new Error("Discord client is required")
  const decide = createBotManagerDecision({ repositoryProvider, logger })
  return async function verify({ guildId, discordUserId }) {
    try {
      const guild = client.guilds.cache?.get(guildId) || await client.guilds.fetch(guildId)
      const member = await guild.members.fetch(discordUserId)
      return decide({
        guildId,
        isOwner: guild.ownerId === discordUserId,
        isAdministrator: Boolean(member.permissions?.has(PermissionFlagsBits.Administrator)),
        hasRole: roleId => Boolean(member.roles?.cache?.has(roleId))
      })
    } catch (error) {
      if (String(error?.code) === "10007") {
        return Object.freeze({ status: "denied", reason: "not_member" })
      }
      logger.error(JSON.stringify({
        event: "native_manager_member_verification_failed",
        error_code: String(error?.code || "discord_unavailable").slice(0, 80)
      }))
      return Object.freeze({ status: "unavailable", reason: "discord_unavailable" })
    }
  }
}

function createLiveGuildOwnerVerifier({ client, logger = console }) {
  if (!client?.guilds) throw new Error("Discord client is required")
  return async function verify({ guildId, discordUserId }) {
    try {
      const guild = client.guilds.cache?.get(guildId) || await client.guilds.fetch(guildId)
      return guild.ownerId === discordUserId
        ? Object.freeze({ status: "owner" })
        : Object.freeze({ status: "not_owner" })
    } catch (error) {
      logger.error(JSON.stringify({
        event: "native_guild_owner_verification_failed",
        error_code: String(error?.code || "discord_unavailable").slice(0, 80)
      }))
      return Object.freeze({ status: "unavailable" })
    }
  }
}

module.exports = {
  interactionIsGuildOwner,
  interactionIsDiscordAdministrator,
  createBotManagerDecision,
  createBotManagerAuthorizer,
  createLiveBotManagerVerifier,
  createLiveGuildOwnerVerifier
}
