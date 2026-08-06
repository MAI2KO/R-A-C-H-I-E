const { SlashCommandBuilder } = require("discord.js")
const { schedulerIsEnabled } = require("./db")

function buildEventSchedulerCommand() {
  return new SlashCommandBuilder()
    .setName("event-scheduler")
    .setDescription("Manage alliance events and weekly roundups")
}

function getEventSchedulerCommandData(env = process.env) {
  if (!schedulerIsEnabled(env)) return null
  return buildEventSchedulerCommand().toJSON()
}

module.exports = {
  buildEventSchedulerCommand,
  getEventSchedulerCommandData
}
