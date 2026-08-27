const { PermissionFlagsBits } = require("discord.js")

function interactionIsGuildOwner(interaction) {
  return Boolean(interaction.guild?.ownerId
    && interaction.user?.id
    && interaction.guild.ownerId === interaction.user.id)
}

function interactionIsDiscordAdministrator(interaction) {
  return Boolean(interaction.memberPermissions?.has(PermissionFlagsBits.Administrator))
}

function createBotManagerAuthorizer({ repositoryProvider, logger = console }) {
  if (typeof repositoryProvider !== "function") throw new Error("repositoryProvider is required")
  return Object.freeze({
    async canManage(interaction) {
      if (interactionIsGuildOwner(interaction) || interactionIsDiscordAdministrator(interaction)) {
        return true
      }
      if (!interaction.guildId) return false
      try {
        const repository = repositoryProvider()
        if (!repository) return false
        const roleId = await repository.getManagerRole(interaction.guildId)
        return Boolean(roleId && interaction.member?.roles?.cache?.has(roleId))
      } catch (error) {
        logger.error(JSON.stringify({
          event: "bot_manager_authorization_failed",
          game_profile: String(repositoryProvider.gameProfile || "unknown").slice(0, 32),
          error_code: String(error?.code || "database_unavailable").slice(0, 80)
        }))
        return false
      }
    }
  })
}

module.exports = {
  interactionIsGuildOwner,
  interactionIsDiscordAdministrator,
  createBotManagerAuthorizer
}
