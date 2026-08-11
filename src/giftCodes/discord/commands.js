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

function playerOption(command, terms, required = false) {
  return command.addStringOption(option => option
    .setName("player_id")
    .setDescription(`Optional ${terms.playerLabel} (defaults to your primary)`)
    .setRequired(required))
}

function buildGiftCommand(gameProfile) {
  const terms = profileTerminology(gameProfile)
  return new SlashCommandBuilder()
    .setName("gift")
    .setDescription(`Manage ${terms.gameName} gift codes`)
    .addSubcommand(command => command
      .setName("submit")
      .setDescription("Submit a gift code for controlled verification")
      .addStringOption(option => option
        .setName("code")
        .setDescription("Gift code, preserving its exact case")
        .setRequired(true)))
    .addSubcommand(command => playerOption(command
      .setName("auto-enable")
      .setDescription("Enable automatic redemption for one player account"), terms))
    .addSubcommand(command => playerOption(command
      .setName("auto-disable")
      .setDescription("Disable automatic redemption for one player account"), terms))
    .addSubcommand(command => playerOption(command
      .setName("status")
      .setDescription("View your gift-code settings and recent result"), terms))
}

function buildGiftAdminCommand(gameProfile) {
  const terms = profileTerminology(gameProfile)
  return new SlashCommandBuilder()
    .setName("gift-admin")
    .setDescription(`Inspect ${terms.gameName} gift-code processing`)
    .addSubcommand(command => command
      .setName("status")
      .setDescription("View subsystem and verification status"))
    .addSubcommand(command => command
      .setName("queue")
      .setDescription("View high-level redemption queue health"))
    .addSubcommand(command => command
      .setName("code")
      .setDescription("Inspect one gift code without player details")
      .addStringOption(option => option.setName("code").setDescription("Exact gift code").setRequired(true)))
    .addSubcommand(command => command
      .setName("verify")
      .setDescription("TEST: verify one code using the configured verifier")
      .addStringOption(option => option.setName("code").setDescription("Exact gift code").setRequired(true)))
}

function giftCommands(gameProfile, env = process.env) {
  if (!playerGiftCodesIsEnabled(env)) return []
  if (!["wos", "kingshot"].includes(gameProfile)) return []
  return [buildGiftCommand(gameProfile).toJSON(), buildGiftAdminCommand(gameProfile).toJSON()]
}

function getGiftCommandData(env = process.env) {
  return giftCommands(String(env.GAME_PROFILE || "wos"), env)
}

module.exports = {
  buildPlayerCommand,
  getPlayerCommandData,
  buildGiftCommand,
  buildGiftAdminCommand,
  getGiftCommandData
}
