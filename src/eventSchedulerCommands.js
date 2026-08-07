const { SlashCommandBuilder } = require("discord.js")
const { schedulerIsEnabled } = require("./db")

function buildEventSchedulerCommand() {
  return new SlashCommandBuilder()
    .setName("event-scheduler")
    .setDescription("Manage alliance events and weekly roundups")
}

function buildEventSchedulerHelpCommand() {
  return new SlashCommandBuilder()
    .setName("event-scheduler-help")
    .setDescription("Show event scheduler setup and management help")
}

function getEventSchedulerCommandData(env = process.env) {
  if (!schedulerIsEnabled(env)) return null
  return buildEventSchedulerCommand().toJSON()
}

function getEventSchedulerHelpCommandData(env = process.env) {
  if (!schedulerIsEnabled(env)) return null
  return buildEventSchedulerHelpCommand().toJSON()
}

module.exports = {
  buildEventSchedulerCommand,
  buildEventSchedulerHelpCommand,
  getEventSchedulerCommandData,
  getEventSchedulerHelpCommandData
}
