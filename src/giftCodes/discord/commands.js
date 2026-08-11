const { SlashCommandBuilder } = require("discord.js")
const { playerGiftCodesIsEnabled } = require("../../db")
const { profileTerminology } = require("../terminology")

function buildPlayerCommand(gameProfile) {
  const terms = profileTerminology(gameProfile)
  return new SlashCommandBuilder()
    .setName("player")
    .setDescription(`Manage your ${terms.gameName} player accounts`)
    .addSubcommand(command => command
      .setName("register")
      .setDescription(`Register a ${terms.gameName} character`)
      .addStringOption(option => option
        .setName("player_id")
        .setDescription(terms.playerLabel)
        .setRequired(true))
      .addStringOption(option => option
        .setName(terms.locationLabelLower)
        .setDescription(`${terms.locationLabel} number`)
        .setRequired(true)))
    .addSubcommand(command => command
      .setName("view")
      .setDescription("View your registered player accounts")
      .addStringOption(option => option
        .setName("player_id")
        .setDescription("Optional Player ID to view")
        .setRequired(false)))
    .addSubcommand(command => command
      .setName("location")
      .setDescription(`Change a character's ${terms.locationLabel}`)
      .addStringOption(option => option
        .setName("player_id")
        .setDescription(terms.playerLabel)
        .setRequired(true))
      .addStringOption(option => option
        .setName(terms.locationLabelLower)
        .setDescription(`New ${terms.locationLabel} number`)
        .setRequired(true)))
    .addSubcommand(command => command
      .setName("remove")
      .setDescription("Deactivate one of your player accounts")
      .addStringOption(option => option
        .setName("player_id")
        .setDescription(terms.playerLabel)
        .setRequired(true)))
}

function getPlayerCommandData(env = process.env) {
  if (!playerGiftCodesIsEnabled(env)) return null
  const gameProfile = String(env.GAME_PROFILE || "wos")
  if (!["wos", "kingshot"].includes(gameProfile)) return null
  return buildPlayerCommand(gameProfile).toJSON()
}

module.exports = { buildPlayerCommand, getPlayerCommandData }
