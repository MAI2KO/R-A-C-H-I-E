const {
  ActionRowBuilder,
  ButtonBuilder,
  ButtonStyle,
  StringSelectMenuBuilder
} = require("discord.js")
const { acknowledgeSchedulerInteraction } = require("./interactionResponses")

const HELP_IDS = Object.freeze({
  prefix: "eh:",
  home: "eh:home",
  section: "eh:section"
})

const HELP_SECTIONS = Object.freeze([
  Object.freeze({
    id: "getting-started",
    label: "Getting started",
    description: "Channels, main alliance and sub-alliances",
    content:
      "**Getting started**\n\n" +
      "1. Run `/event-scheduler`, choose **Alliance identity**, and set the main alliance name.\n" +
      "2. Choose **Configure channels** and pick the alliance reminder channel from Discord's channel list.\n" +
      "3. Choose **Weekly roundup settings** to enable or disable the alliance roundup, control state publishing, and change its UTC weekday, time, or channel.\n" +
      "4. Choose **Alliances** to add or rename sub-alliances.\n" +
      "5. Use **State destination** and **State sharing** only when a state Discord should receive the combined weekly roundup.\n\n" +
      "All scheduler administration is private and uses the server's existing management authorisation."
  }),
  Object.freeze({
    id: "creating-event",
    label: "Creating an event",
    description: "Event details, reminders and confirmation",
    content:
      "**Creating an event**\n\n" +
      "Choose **Create event**, select the main alliance or sub-alliance, then enter the event name and first occurrence date. Under **Event timing**, choose **Single time** for one UTC time or **Multiple groups** to add each group name and UTC time through a guided modal. Groups may share a time.\n\n" +
      "Choose recurrence every 3, 7, 14 or 28 days; one optional advance reminder; the optional one-minute final announcement; and optional advance/final custom messages. Use **Manage image** to add, replace or remove one supported image.\n\n" +
      "Publishing controls cover alliance reminders and weekly-roundup inclusion. When Weekly roundup is enabled, state inclusion is automatic while the alliance's state link is enabled. Review the clearly labelled UTC and local-time preview before final confirmation. Nothing is saved before confirmation."
  }),
  Object.freeze({
    id: "accepted-times",
    label: "Accepted time formats",
    description: "Supported UTC input examples",
    content:
      "**Accepted UTC time formats**\n\n" +
      "Supported examples:\n" +
      "- `18:30`\n" +
      "- `1830`\n" +
      "- `1800`\n" +
      "- `6:30pm`\n" +
      "- `6.30 PM`\n" +
      "- `6 PM`\n" +
      "- `18`\n\n" +
      "Dates use `YYYY-MM-DD`. Times are interpreted as UTC. Ambiguous or invalid values are rejected rather than guessed."
  }),
  Object.freeze({
    id: "reminders",
    label: "Reminder behaviour",
    description: "Advance, final and image delivery rules",
    content:
      "**Reminder behaviour**\n\n" +
      "Each event/group occurrence can have one advance reminder: none, 5, 10, 15, 20 or 30 minutes before.\n\n" +
      "The separate final announcement says **About to start** and is delivered one minute before the configured event time. There is no exact-start post.\n\n" +
      "Immediate reminders contain only the alliance, event, optional group, countdown, and optional custom text. The final message uses **About to start** and approximate one-minute wording. Detailed UTC/local schedules belong in previews and weekly roundups. Mentions never ping. An image is attached only to the alliance advance reminder. Individual reminders are alliance-only."
  }),
  Object.freeze({
    id: "managing-events",
    label: "Managing events",
    description: "Edit, cancel, pause, resume and delete",
    content:
      "**Managing events**\n\n" +
      "Use **View events**, select an event, then choose Preview, Edit details, Change alliance, Manage image, Pause/Resume or Delete.\n\n" +
      "**Edit details** keeps the current alliance and image automatically. **Change alliance** changes only ownership and preserves times, recurrence, reminders, messages, publishing settings, groups, and image. **Manage image** explicitly keeps, replaces or removes the stored image. No unrelated edit requires another upload.\n\n" +
      "Select no advance reminder or disable the final announcement to cancel future reminders of that type. Optional messages can be cleared.\n\n" +
      "Pause stops future reminders and roundup appearances while retaining the recurrence anchor. Resume uses that anchor without replaying expired reminders. Delete is soft deletion and preserves sent history."
  }),
  Object.freeze({
    id: "alliance-roundups",
    label: "Alliance roundups",
    description: "Local eligibility and alliance labels",
    content:
      "**Alliance weekly roundups**\n\n" +
      "Open **Weekly roundup settings** at any time, including while disabled. Enable or disable the alliance roundup, independently enable state publishing, change the roundup channel, or preview and save any UTC weekday/time. Disabled settings retain their channel and schedule. Re-enabling starts at the next future schedule and does not replay missed roundups.\n\n" +
      "An active event appears when weekly-roundup inclusion is enabled and an occurrence falls inside the configured weekly window for the same Discord guild and game profile.\n\n" +
      "Main-alliance and sub-alliance entries are labelled separately. Paused or deleted events are excluded. Roundups never include event images."
  }),
  Object.freeze({
    id: "state-setup",
    label: "State setup",
    description: "Link the state Discord and roundup channel",
    content:
      "**State roundup setup**\n\n" +
      "1. Invite the appropriate bot to the state Discord. R.A.C.H.I.E links WOS data; P.E.G.G.I.E links Kingshot data.\n" +
      "2. In the state Discord, run `/event-scheduler`, choose **State destination**, and select the weekly-roundup channel from Discord's channel list.\n" +
      "3. Generate a one-time 15-minute link code. In the alliance Discord choose **State sharing**, select **Link with code**, and enter it.\n\n" +
      "The bot identifies both servers automatically and validates the channel type plus View Channel, Send Messages and Embed Links permissions. The code works only once and only for the same game profile. Normal setup never requires server or channel IDs."
  }),
  Object.freeze({
    id: "state-behaviour",
    label: "State behaviour",
    description: "Combined weekly roundup only",
    content:
      "**State behaviour**\n\n" +
      "A linked state Discord receives one combined weekly roundup containing eligible events from enabled links for the same game profile. WOS and Kingshot data never mix.\n\n" +
      "Active events with Weekly roundup enabled are included automatically while their alliance's state link is enabled. There is no separate per-event state setting, and state sharing never creates an individual state reminder. Advance reminders, final announcements, grouped reminders and images remain alliance-only."
  }),
  Object.freeze({
    id: "troubleshooting",
    label: "Troubleshooting",
    description: "Common scheduler setup and delivery issues",
    content:
      "**Troubleshooting**\n\n" +
      "- Scheduler command absent: confirm `EVENT_SCHEDULER_ENABLED=true` and restart after command registration.\n" +
      "- Scheduler unavailable: check `DATABASE_URL`, `GAME_PROFILE`, `BOT_INSTANCE_NAME`, database access and migration logs. Existing bot features remain available.\n" +
      "- Channel missing: confirm the bot is in that server, then choose an available text channel from the native list.\n" +
      "- State link rejected: generate a fresh code in the state Discord and use the same bot/game profile on both sides.\n" +
      "- Missing permission: grant View Channel, Send Messages and Embed Links; alliance image channels also need Attach Files.\n" +
      "- Invalid date/time: use `YYYY-MM-DD` and a documented UTC format.\n" +
      "- Reminder missing: check event status, publishing, reminder choice, channel and grace window.\n" +
      "- Roundup missing: check active status, roundup inclusion, weekly window/channel, and the state link where applicable. Paused or deleted events do not publish."
  })
])

const HELP_SECTION_BY_ID = new Map(HELP_SECTIONS.map(section => [section.id, section]))

function helpComponents(selectedId) {
  const rows = [new ActionRowBuilder().addComponents(
    new StringSelectMenuBuilder()
      .setCustomId(HELP_IDS.section)
      .setPlaceholder("Choose a scheduler help section")
      .addOptions(HELP_SECTIONS.map(section => ({
        label: section.label,
        description: section.description,
        value: section.id,
        default: section.id === selectedId
      })))
  )]
  if (selectedId !== HELP_SECTIONS[0].id) {
    rows.push(new ActionRowBuilder().addComponents(
      new ButtonBuilder()
        .setCustomId(HELP_IDS.home)
        .setLabel("Getting started")
        .setStyle(ButtonStyle.Secondary)
    ))
  }
  return rows
}

function buildSchedulerHelpView(sectionId = HELP_SECTIONS[0].id) {
  const section = HELP_SECTION_BY_ID.get(sectionId) || HELP_SECTIONS[0]
  return {
    content: section.content,
    components: helpComponents(section.id),
    allowedMentions: { parse: [], repliedUser: false }
  }
}

function isSchedulerHelpInteraction(interaction) {
  return interaction.commandName === "event-scheduler-help"
    || String(interaction.customId || "").startsWith(HELP_IDS.prefix)
}

async function handleEventSchedulerHelpInteraction(interaction) {
  if (!isSchedulerHelpInteraction(interaction)) return false

  if (interaction.isChatInputCommand?.() && interaction.commandName === "event-scheduler-help") {
    await acknowledgeSchedulerInteraction(interaction)
    await interaction.editReply(buildSchedulerHelpView())
    return true
  }

  if (interaction.isStringSelectMenu?.() && interaction.customId === HELP_IDS.section) {
    await acknowledgeSchedulerInteraction(interaction)
    await interaction.editReply(buildSchedulerHelpView(interaction.values[0]))
    return true
  }

  if (interaction.isButton?.() && interaction.customId === HELP_IDS.home) {
    await acknowledgeSchedulerInteraction(interaction)
    await interaction.editReply(buildSchedulerHelpView())
    return true
  }

  return true
}

module.exports = {
  HELP_IDS,
  HELP_SECTIONS,
  buildSchedulerHelpView,
  isSchedulerHelpInteraction,
  handleEventSchedulerHelpInteraction
}
