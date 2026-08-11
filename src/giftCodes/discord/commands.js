const { SlashCommandBuilder } = require("discord.js")
const { playerGiftCodesIsEnabled } = require("../../db")
const { profileTerminology } = require("../terminology")

function buildPlayerRegisterCommand(gameProfile) {
  const terms = profileTerminology(gameProfile)
  return new SlashCommandBuilder()
    .setName("player-register")
    .setDescription(`Register and manage your ${terms.gameName} player accounts`)
}

function buildGiftCodesCommand(gameProfile) {
  const terms = profileTerminology(gameProfile)
  return new SlashCommandBuilder()
    .setName("gift-codes")
    .setDescription(`Gift codes and automatic redemption for ${terms.gameName}`)
}

function buildGiftCodesAdminCommand(gameProfile) {
  const terms = profileTerminology(gameProfile)
  return new SlashCommandBuilder()
    .setName("gift-codes-admin")
    .setDescription(`Manage ${terms.gameName} gift-code automation and community settings`)
}

function profileFromEnvironment(env) {
  const gameProfile = String(env.GAME_PROFILE || "wos")
  return ["wos", "kingshot"].includes(gameProfile) ? gameProfile : null
}

function getPlayerCommandData(env = process.env) {
  if (!playerGiftCodesIsEnabled(env)) return null
  const gameProfile = profileFromEnvironment(env)
  return gameProfile ? buildPlayerRegisterCommand(gameProfile).toJSON() : null
}

function getGiftCommandData(env = process.env) {
  if (!playerGiftCodesIsEnabled(env)) return []
  const gameProfile = profileFromEnvironment(env)
  return gameProfile ? [
    buildGiftCodesCommand(gameProfile).toJSON(),
    buildGiftCodesAdminCommand(gameProfile).toJSON()
  ] : []
}

module.exports = {
  buildPlayerRegisterCommand,
  buildGiftCodesCommand,
  buildGiftCodesAdminCommand,
  getPlayerCommandData,
  getGiftCommandData
}
