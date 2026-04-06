require("dotenv").config()

const {
  Client,
  GatewayIntentBits,
  REST,
  Routes,
  SlashCommandBuilder,
  ActionRowBuilder,
  StringSelectMenuBuilder,
  ButtonBuilder,
  ButtonStyle,
  ModalBuilder,
  TextInputBuilder,
  TextInputStyle,
  PermissionFlagsBits
} = require("discord.js")

const axios = require("axios")

console.log("Starting bot...")
console.log("CLIENT_ID present:", !!process.env.CLIENT_ID)
console.log("BOT_TOKEN present:", !!process.env.BOT_TOKEN)
console.log("APPS_SCRIPT_URL present:", !!process.env.APPS_SCRIPT_URL)
console.log("ADMIN_API_KEY present:", !!process.env.ADMIN_API_KEY)

const client = new Client({
  intents: [GatewayIntentBits.Guilds]
})

const pendingUnlinks = new Map()
const pendingBookings = new Map()
const BOOKING_PAGE_SIZE = 25
const BOOKING_TTL_MS = 15 * 60 * 1000

function createBookingToken() {
  return Math.random().toString(36).slice(2, 10) + Date.now().toString(36)
}

function cleanupPendingBookings() {
  const now = Date.now()

  for (const [token, entry] of pendingBookings.entries()) {
    if (now - entry.createdAt > BOOKING_TTL_MS) {
      pendingBookings.delete(token)
    }
  }
}

function getPendingBookingOrThrow(token) {
  cleanupPendingBookings()

  const entry = pendingBookings.get(token)
  if (!entry) {
    throw new Error("This booking session has expired. Run /book again.")
  }

  return entry
}

function buildBookingComponents(token, page = 0) {
  const entry = getPendingBookingOrThrow(token)
  const totalPages = Math.max(1, Math.ceil(entry.times.length / BOOKING_PAGE_SIZE))
  const safePage = Math.max(0, Math.min(page, totalPages - 1))

  const start = safePage * BOOKING_PAGE_SIZE
  const pageTimes = entry.times.slice(start, start + BOOKING_PAGE_SIZE)

  const select = new StringSelectMenuBuilder()
    .setCustomId(`book_select:${token}:${safePage}`)
    .setPlaceholder(pageTimes.length ? "Select a time" : "No times available")
    .setDisabled(pageTimes.length === 0)
    .addOptions(
      pageTimes.length
        ? pageTimes.map(time => ({
            label: time,
            value: time,
            description: `${entry.day} slot`
          }))
        : [{ label: "No times available", value: "none", description: "No slots" }]
    )

  const selectRow = new ActionRowBuilder().addComponents(select)

  const buttons = new ActionRowBuilder().addComponents(
    new ButtonBuilder()
      .setCustomId(`book_page:${token}:${safePage - 1}`)
      .setLabel("Previous")
      .setStyle(ButtonStyle.Secondary)
      .setDisabled(safePage <= 0),
    new ButtonBuilder()
      .setCustomId(`book_page:${token}:${safePage + 1}`)
      .setLabel("Next")
      .setStyle(ButtonStyle.Secondary)
      .setDisabled(safePage >= totalPages - 1)
  )

  const content =
    `Select a time for ${entry.day}\n` +
    `Page ${safePage + 1} of ${totalPages}`

  return {
    content,
    components: [selectRow, buttons]
  }
}

function buildExtrasModal(token) {
  const entry = getPendingBookingOrThrow(token)
  const cfg = entry.config || {}

  const modal = new ModalBuilder()
    .setCustomId(`book_modal:${token}`)
    .setTitle(`${entry.day} booking details`)

  const rows = []

  if (entry.day === "Construction") {
    if (cfg.construction_fc_required) {
      rows.push(
        new ActionRowBuilder().addComponents(
          new TextInputBuilder()
            .setCustomId("fc")
            .setLabel("Fire Crystals (FC)")
            .setStyle(TextInputStyle.Short)
            .setRequired(true)
            .setMaxLength(6)
            .setPlaceholder("Numbers only")
        )
      )
    }

    if (cfg.construction_rfc_required) {
      rows.push(
        new ActionRowBuilder().addComponents(
          new TextInputBuilder()
            .setCustomId("rfc")
            .setLabel("Refined Fire Crystals (RFC)")
            .setStyle(TextInputStyle.Short)
            .setRequired(true)
            .setMaxLength(6)
            .setPlaceholder("Numbers only")
        )
      )
    }

    if (cfg.construction_speedups_required) {
      rows.push(
        new ActionRowBuilder().addComponents(
          new TextInputBuilder()
            .setCustomId("speedups")
            .setLabel("Speed-ups in whole days")
            .setStyle(TextInputStyle.Short)
            .setRequired(true)
            .setMaxLength(6)
            .setPlaceholder("Example: 100")
        )
      )
    }
  }

  if (entry.day === "Research") {
    if (cfg.research_shards_required) {
      rows.push(
        new ActionRowBuilder().addComponents(
          new TextInputBuilder()
            .setCustomId("shards")
            .setLabel("Fire Crystal Shards")
            .setStyle(TextInputStyle.Short)
            .setRequired(true)
            .setMaxLength(6)
            .setPlaceholder("Numbers only")
        )
      )
    }

    if (cfg.research_speedups_required) {
      rows.push(
        new ActionRowBuilder().addComponents(
          new TextInputBuilder()
            .setCustomId("speedups")
            .setLabel("Speed-ups in whole days")
            .setStyle(TextInputStyle.Short)
            .setRequired(true)
            .setMaxLength(6)
            .setPlaceholder("Example: 100")
        )
      )
    }
  }

  if (entry.day === "Troop") {
    if (cfg.troop_speedups_required) {
      rows.push(
        new ActionRowBuilder().addComponents(
          new TextInputBuilder()
            .setCustomId("speedups")
            .setLabel("Speed-ups in whole days")
            .setStyle(TextInputStyle.Short)
            .setRequired(true)
            .setMaxLength(6)
            .setPlaceholder("Example: 100")
        )
      )
    }
  }

  if (rows.length === 0) {
    return null
  }

  modal.addComponents(...rows.slice(0, 5))
  return modal
}

function readModalValue(fields, id) {
  try {
    return fields.getTextInputValue(id)
  } catch {
    return null
  }
}

function ensureNumericOrEmpty(value, fieldName) {
  const raw = String(value || "").trim()
  if (!raw) return ""

  if (!/^[0-9]{1,6}$/.test(raw)) {
    throw new Error(`${fieldName} must be a whole number only, for example 100.`)
  }

  return raw
}

async function submitBookingFromEntry(entry, overrides = {}) {
  return await postToAppsScript({
    action: "book_for_server",
    adminKey: process.env.ADMIN_API_KEY,
    discordServerId: entry.discordServerId,
    discordUserId: entry.discordUserId,
    day: entry.day,
    time: entry.selectedTime,
    fc: overrides.fc ?? null,
    rfc: overrides.rfc ?? null,
    shards: overrides.shards ?? null,
    speedups: overrides.speedups ?? null
  })
}

async function sendBookingDm(user, message) {
  try {
    await user.send(message)
  } catch (error) {
    console.log("Could not send DM:", error?.code, error?.message)
  }
}

function userCanManageServer(interaction) {
  return interaction.memberPermissions?.has(PermissionFlagsBits.ManageGuild)
}

function yesNo(value) {
  return value ? "ON" : "OFF"
}

async function fetchSettingsForServer(interaction) {
  return await postToAppsScript({
    action: "get_settings_for_server",
    adminKey: process.env.ADMIN_API_KEY,
    discordServerId: interaction.guildId
  })
}

function buildSettingsHomeText(result) {
  const s = result.settings

  return (
    `State ${result.state_code}\n` +
    `Sheet: ${result.sheet_name}\n\n` +
    `Max bookings per player per day: ${s.max_bookings_per_player_per_day}\n` +
    `Max linked Discord servers: ${s.max_linked_servers}\n\n` +
    `Choose a settings group below.\n\n` +
    `Construction day\n` +
    `Choose which Construction resources players must enter when booking.\n` +
    `On = the field appears and must be filled in.\n` +
    `Off = the field will not appear during booking.\n\n` +
    `Research day\n` +
    `Choose which Research resources players must enter when booking.\n` +
    `On = the field appears and must be filled in.\n` +
    `Off = the field will not appear during booking.\n\n` +
    `Troop training day\n` +
    `Choose whether players must enter troop speed-ups when booking.\n` +
    `On = the field appears and must be filled in.\n` +
    `Off = the field will not appear during booking.`
  )
}

function buildSettingsHomeComponents(result) {
  const s = result.settings

  return [
    new ActionRowBuilder().addComponents(
      new ButtonBuilder()
        .setCustomId("settings_group:construction")
        .setLabel("Construction day")
        .setStyle(ButtonStyle.Secondary),
      new ButtonBuilder()
        .setCustomId("settings_group:research")
        .setLabel("Research day")
        .setStyle(ButtonStyle.Secondary),
      new ButtonBuilder()
        .setCustomId("settings_group:troop")
        .setLabel("Troop training day")
        .setStyle(ButtonStyle.Secondary)
    ),
    new ActionRowBuilder().addComponents(
      new StringSelectMenuBuilder()
        .setCustomId("settings_max_bookings_select")
        .setPlaceholder(`Max bookings per player: ${s.max_bookings_per_player_per_day}`)
        .addOptions([
          { label: "6 bookings per day", value: "6", description: "Original plus one rebook for each minister day" },
          { label: "9 bookings per day", value: "9", description: "3 extra bookings above default" },
          { label: "12 bookings per day", value: "12", description: "6 extra bookings above default" },
          { label: "15 bookings per day", value: "15", description: "9 extra bookings above default" },
          { label: "18 bookings per day", value: "18", description: "12 extra bookings above default" }
        ])
    ),
    new ActionRowBuilder().addComponents(
      new StringSelectMenuBuilder()
        .setCustomId("settings_max_links_select")
        .setPlaceholder(`Max linked Discord servers: ${s.max_linked_servers}`)
        .addOptions([
          { label: "5 linked servers", value: "5", description: "Default recommendation" },
          { label: "10 linked servers", value: "10", description: "For larger states" },
          { label: "15 linked servers", value: "15", description: "For many alliance servers" },
          { label: "20 linked servers", value: "20", description: "High capacity" }
        ])
    )
  ]
}

function buildConstructionSettingsText(result) {
  const s = result.settings

  return (
    `Construction day settings\n\n` +
    `FC required: ${yesNo(s.construction_fc_required)}\n` +
    `RFC required: ${yesNo(s.construction_rfc_required)}\n` +
    `Construction speed-ups required: ${yesNo(s.construction_speedups_required)}\n\n` +
    `On = the field appears and must be filled in.\n` +
    `Off = the field will not appear during booking.`
  )
}

function buildConstructionSettingsComponents(result) {
  const s = result.settings

  return [
    new ActionRowBuilder().addComponents(
      new ButtonBuilder()
        .setCustomId("settings_toggle:construction_fc_required:construction")
        .setLabel(`FC: ${yesNo(s.construction_fc_required)}`)
        .setStyle(s.construction_fc_required ? ButtonStyle.Success : ButtonStyle.Secondary),
      new ButtonBuilder()
        .setCustomId("settings_toggle:construction_rfc_required:construction")
        .setLabel(`RFC: ${yesNo(s.construction_rfc_required)}`)
        .setStyle(s.construction_rfc_required ? ButtonStyle.Success : ButtonStyle.Secondary),
      new ButtonBuilder()
        .setCustomId("settings_toggle:construction_speedups_required:construction")
        .setLabel(`Speed-ups: ${yesNo(s.construction_speedups_required)}`)
        .setStyle(s.construction_speedups_required ? ButtonStyle.Success : ButtonStyle.Secondary)
    ),
    new ActionRowBuilder().addComponents(
      new ButtonBuilder()
        .setCustomId("settings_back:home")
        .setLabel("Back")
        .setStyle(ButtonStyle.Primary)
    )
  ]
}

function buildResearchSettingsText(result) {
  const s = result.settings

  return (
    `Research day settings\n\n` +
    `Shards required: ${yesNo(s.research_shards_required)}\n` +
    `Research speed-ups required: ${yesNo(s.research_speedups_required)}\n\n` +
    `On = the field appears and must be filled in.\n` +
    `Off = the field will not appear during booking.`
  )
}

function buildResearchSettingsComponents(result) {
  const s = result.settings

  return [
    new ActionRowBuilder().addComponents(
      new ButtonBuilder()
        .setCustomId("settings_toggle:research_shards_required:research")
        .setLabel(`Shards: ${yesNo(s.research_shards_required)}`)
        .setStyle(s.research_shards_required ? ButtonStyle.Success : ButtonStyle.Secondary),
      new ButtonBuilder()
        .setCustomId("settings_toggle:research_speedups_required:research")
        .setLabel(`Speed-ups: ${yesNo(s.research_speedups_required)}`)
        .setStyle(s.research_speedups_required ? ButtonStyle.Success : ButtonStyle.Secondary)
    ),
    new ActionRowBuilder().addComponents(
      new ButtonBuilder()
        .setCustomId("settings_back:home")
        .setLabel("Back")
        .setStyle(ButtonStyle.Primary)
    )
  ]
}

function buildTroopSettingsText(result) {
  const s = result.settings

  return (
    `Troop training day settings\n\n` +
    `Troop speed-ups required: ${yesNo(s.troop_speedups_required)}\n\n` +
    `On = the field appears and must be filled in.\n` +
    `Off = the field will not appear during booking.`
  )
}

function buildTroopSettingsComponents(result) {
  const s = result.settings

  return [
    new ActionRowBuilder().addComponents(
      new ButtonBuilder()
        .setCustomId("settings_toggle:troop_speedups_required:troop")
        .setLabel(`Speed-ups: ${yesNo(s.troop_speedups_required)}`)
        .setStyle(s.troop_speedups_required ? ButtonStyle.Success : ButtonStyle.Secondary)
    ),
    new ActionRowBuilder().addComponents(
      new ButtonBuilder()
        .setCustomId("settings_back:home")
        .setLabel("Back")
        .setStyle(ButtonStyle.Primary)
    )
  ]
}

function renderSettingsView(result, view = "home") {
  if (view === "construction") {
    return {
      content: buildConstructionSettingsText(result),
      components: buildConstructionSettingsComponents(result)
    }
  }

  if (view === "research") {
    return {
      content: buildResearchSettingsText(result),
      components: buildResearchSettingsComponents(result)
    }
  }

  if (view === "troop") {
    return {
      content: buildTroopSettingsText(result),
      components: buildTroopSettingsComponents(result)
    }
  }

  return {
    content: buildSettingsHomeText(result),
    components: buildSettingsHomeComponents(result)
  }
}

/* -------------------- COMMAND DEFINITIONS -------------------- */

const commands = [

  new SlashCommandBuilder()
   .setName("linked-servers")
   .setDescription("Show all linked Discord servers for this state")
   .setDefaultMemberPermissions(PermissionFlagsBits.ManageGuild),

  new SlashCommandBuilder()
    .setName("help")
    .setDescription("Show help for using the booking system"),

  new SlashCommandBuilder()
    .setName("unlink-state")
    .setDescription("Remove a linked Discord server from this state")
    .setDefaultMemberPermissions(PermissionFlagsBits.ManageGuild),

  new SlashCommandBuilder()
    .setName("setup")
    .setDescription("Create a new state and link this Discord")
    .setDefaultMemberPermissions(PermissionFlagsBits.ManageGuild),

  new SlashCommandBuilder()
    .setName("link-state")
    .setDescription("Link this Discord to an existing state")
    .setDefaultMemberPermissions(PermissionFlagsBits.ManageGuild),

  new SlashCommandBuilder()
    .setName("settings")
    .setDescription("Open the admin settings panel")
    .setDefaultMemberPermissions(PermissionFlagsBits.ManageGuild),

  new SlashCommandBuilder()
    .setName("sheet-link")
    .setDescription("Get the booking sheet link for this server")
    .setDefaultMemberPermissions(PermissionFlagsBits.ManageGuild),

  new SlashCommandBuilder()
    .setName("open-bookings")
    .setDescription("Open bookings for this server")
    .setDefaultMemberPermissions(PermissionFlagsBits.ManageGuild),

  new SlashCommandBuilder()
    .setName("close-bookings")
    .setDescription("Close bookings for this server")
    .setDefaultMemberPermissions(PermissionFlagsBits.ManageGuild),

  new SlashCommandBuilder()
    .setName("my-bookings")
    .setDescription("Show your current bookings for this state"),

  new SlashCommandBuilder()
    .setName("booking-link")
    .setDescription("Get the booking link for this server"),

  new SlashCommandBuilder()
    .setName("times")
    .setDescription("Check available times")
    .addStringOption(option =>
      option
        .setName("day")
        .setDescription("Which day")
        .setRequired(true)
        .addChoices(
          { name: "Construction", value: "Construction" },
          { name: "Research", value: "Research" },
          { name: "Troop", value: "Troop" }
        )
    ),

  new SlashCommandBuilder()
    .setName("register")
    .setDescription("Register your player details")
    .addStringOption(option =>
      option
        .setName("alliance")
        .setDescription("Alliance tag")
        .setRequired(true)
    )
    .addStringOption(option =>
      option
        .setName("name")
        .setDescription("In game name")
        .setRequired(true)
    )
    .addStringOption(option =>
      option
        .setName("id")
        .setDescription("Player ID numbers only")
        .setRequired(true)
    ),

  new SlashCommandBuilder()
    .setName("my-info")
    .setDescription("Show your registered player info"),

  new SlashCommandBuilder()
    .setName("unregister")
    .setDescription("Delete your registration"),

  new SlashCommandBuilder()
    .setName("book")
    .setDescription("Book a minister slot")
    .addStringOption(option =>
      option
        .setName("day")
        .setDescription("Which day")
        .setRequired(true)
        .addChoices(
          { name: "Construction", value: "Construction" },
          { name: "Research", value: "Research" },
          { name: "Troop", value: "Troop" }
        )
    ),

  new SlashCommandBuilder()
    .setName("reset-state-password")
    .setDescription("Generate a new join password for a state")
    .setDefaultMemberPermissions(PermissionFlagsBits.ManageGuild),

  new SlashCommandBuilder()
    .setName("remove-booking")
    .setDescription("Remove your booking")
    .addStringOption(option =>
      option
        .setName("day")
        .setDescription("Which day")
        .setRequired(true)
        .addChoices(
          { name: "Construction", value: "Construction" },
          { name: "Research", value: "Research" },
          { name: "Troop", value: "Troop" }
        )
    )
].map(command => command.toJSON())

/* -------------------- COMMAND REGISTRATION -------------------- */

const rest = new REST({ version: "10" }).setToken(process.env.BOT_TOKEN)

async function registerCommands() {
  await rest.put(
    Routes.applicationCommands(process.env.CLIENT_ID),
    { body: commands }
  )

  console.log("Slash commands registered")
}

/* -------------------- API HELPER -------------------- */

async function postToAppsScript(payload) {
  const response = await axios.post(
    process.env.APPS_SCRIPT_URL,
    payload,
    { headers: { "Content-Type": "application/json" } }
  )

  return response.data
}

/* -------------------- BOT READY -------------------- */

client.once("clientReady", () => {
  console.log(`Bot ready: ${client.user.tag}`)
})

function buildSetupModal() {
  return new ModalBuilder()
    .setCustomId("setup_modal")
    .setTitle("Setup new state")
    .addComponents(
      new ActionRowBuilder().addComponents(
        new TextInputBuilder()
          .setCustomId("state")
          .setLabel("State number")
          .setStyle(TextInputStyle.Short)
          .setRequired(true)
          .setMaxLength(10)
          .setPlaceholder("Example: 9999")
      ),
      new ActionRowBuilder().addComponents(
        new TextInputBuilder()
          .setCustomId("server")
          .setLabel("Alliance/Discord server name")
          .setStyle(TextInputStyle.Short)
          .setRequired(true)
          .setMaxLength(50)
          .setPlaceholder("Example: myalliancediscord or 9999statediscord")
      )
    )
}

function buildLinkStateModal() {
  return new ModalBuilder()
    .setCustomId("link_state_modal")
    .setTitle("Link to existing state")
    .addComponents(
      new ActionRowBuilder().addComponents(
        new TextInputBuilder()
          .setCustomId("state")
          .setLabel("State number")
          .setStyle(TextInputStyle.Short)
          .setRequired(true)
          .setMaxLength(10)
          .setPlaceholder("Example: 9999")
      ),
      new ActionRowBuilder().addComponents(
        new TextInputBuilder()
          .setCustomId("password")
          .setLabel("State join password")
          .setStyle(TextInputStyle.Short)
          .setRequired(true)
          .setMaxLength(50)
          .setPlaceholder("Paste the join password exactly")
      ),
      new ActionRowBuilder().addComponents(
        new TextInputBuilder()
          .setCustomId("server")
          .setLabel("Alliance/Discord server name")
          .setStyle(TextInputStyle.Short)
          .setRequired(true)
          .setMaxLength(50)
          .setPlaceholder("Example: myalliancediscord or 9999statediscord")
      )
    )
}

/* -------------------- COMMAND HANDLER -------------------- */

client.on("interactionCreate", async interaction => {
  try {

    if (interaction.isStringSelectMenu()) {

      if (interaction.customId === "settings_max_bookings_select") {
        if (!userCanManageServer(interaction)) {
          await interaction.reply({
            content: "❌ You do not have permission to use this control.",
            flags: 64
          })
          return
        }

        await interaction.deferUpdate()

        const value = interaction.values[0]

        const updateResult = await postToAppsScript({
          action: "update_setting_for_server",
          adminKey: process.env.ADMIN_API_KEY,
          discordServerId: interaction.guildId,
          key: "max_bookings_per_player_per_day",
          value: value
        })

        if (!updateResult.ok) {
          await interaction.editReply(`❌ ${updateResult.error || "Could not update max bookings."}`)
          return
        }

        const refreshed = await fetchSettingsForServer(interaction)

        if (!refreshed.ok) {
          await interaction.editReply("❌ Could not refresh settings.")
          return
        }

        await interaction.editReply(renderSettingsView(refreshed, "home"))
        return
      }

      if (interaction.customId === "settings_max_links_select") {
        if (!userCanManageServer(interaction)) {
          await interaction.reply({
            content: "❌ You do not have permission to use this control.",
            flags: 64
          })
          return
        }

        await interaction.deferUpdate()

        const value = interaction.values[0]

        const updateResult = await postToAppsScript({
          action: "update_setting_for_server",
          adminKey: process.env.ADMIN_API_KEY,
          discordServerId: interaction.guildId,
          key: "max_linked_servers",
          value: value
        })

        if (!updateResult.ok) {
          await interaction.editReply(`❌ ${updateResult.error || "Could not update max linked servers."}`)
          return
        }

        const refreshed = await fetchSettingsForServer(interaction)

        if (!refreshed.ok) {
          await interaction.editReply("❌ Could not refresh settings.")
          return
        }

        await interaction.editReply(renderSettingsView(refreshed, "home"))
        return
      }

      if (interaction.customId.startsWith("book_select:")) {
        const [, token] = interaction.customId.split(":")
        const entry = getPendingBookingOrThrow(token)

        if (interaction.user.id !== entry.discordUserId) {
          await interaction.reply({
            content: "❌ This booking menu belongs to someone else.",
            flags: 64
          })
          return
        }

        const time = interaction.values[0]
        if (!time || time === "none") {
          await interaction.reply({
            content: "❌ No valid time selected.",
            flags: 64
          })
          return
        }

        entry.selectedTime = time
        pendingBookings.set(token, entry)

        const modal = buildExtrasModal(token)

        if (modal) {
          await interaction.showModal(modal)
          return
        }

        await interaction.deferUpdate()

        const result = await submitBookingFromEntry(entry)

        if (!result.ok) {
          await interaction.editReply({
            content: `❌ ${result.error || "Could not book slot."}`,
            components: []
          })
          pendingBookings.delete(token)
          return
        }

        let message = `✅ ${result.day} booked for ${result.playerName}\nTime: ${result.time}`

        if (result.moved && result.oldTime) {
          message = `✅ ${result.day} changed for ${result.playerName}\nFrom: ${result.oldTime}\nTo: ${result.time}`
        }

        await interaction.editReply({
          content: message,
          components: []
        })

        let dmMessage =
          `📅 Booking confirmed\n\n` +
          `State: ${result.state_code}\n` +
          `Day: ${result.day}\n` +
          `Time: ${result.time}`

        if (result.moved && result.oldTime) {
          dmMessage =
            `🔁 Booking changed\n\n` +
            `State: ${result.state_code}\n` +
            `Day: ${result.day}\n` +
            `Old time: ${result.oldTime}\n` +
            `New time: ${result.time}`
        }

        await sendBookingDm(interaction.user, dmMessage)

        pendingBookings.delete(token)
        return
      }

      if (interaction.customId.startsWith("unlink_select:")) {
        const token = interaction.customId.split(":")[1]
        const pending = pendingUnlinks.get(token)

        if (!pending) {
          await interaction.reply({
            content: "❌ This unlink menu has expired. Run /unlink-state again.",
            flags: 64
          })
          return
        }

        if (interaction.user.id !== pending.requestedBy) {
          await interaction.reply({
            content: "❌ This unlink menu belongs to someone else.",
            flags: 64
          })
          return
        }

        await interaction.deferUpdate()

        const targetDiscordServerId = interaction.values[0]

        const result = await postToAppsScript({
          action: "unlink_state_server_by_id",
          adminKey: process.env.ADMIN_API_KEY,
          targetDiscordServerId: targetDiscordServerId
        })

        pendingUnlinks.delete(token)

        if (!result.ok) {
          await interaction.editReply({
            content: `❌ ${result.error || "Could not unlink server."}`,
            components: []
          })
          return
        }

        if (!result.removed) {
          await interaction.editReply({
            content: "❌ That linked server could not be found.",
            components: []
          })
          return
        }

        if (result.state_deleted) {
          await interaction.editReply({
            content:
              `✅ Removed "${result.discord_server_name}" from state ${result.state_code}.\n\n` +
              `That was the last linked Discord server, so the state record was deleted and the sheet was moved to trash.`,
            components: []
          })
          return
        }

        await interaction.editReply({
          content: `✅ Removed "${result.discord_server_name}" from state ${result.state_code}.`,
          components: []
        })
        return
      }

      return
    }

    if (interaction.isButton()) {
      if (interaction.customId.startsWith("settings_group:")) {
        if (!userCanManageServer(interaction)) {
          await interaction.reply({
            content: "❌ You do not have permission to use this control.",
            flags: 64
          })
          return
        }

        await interaction.deferUpdate()

        const view = interaction.customId.split(":")[1]
        const refreshed = await fetchSettingsForServer(interaction)

        if (!refreshed.ok) {
          await interaction.editReply("❌ Could not load settings.")
          return
        }

        await interaction.editReply(renderSettingsView(refreshed, view))
        return
      }

      if (interaction.customId.startsWith("settings_toggle:")) {
        if (!userCanManageServer(interaction)) {
          await interaction.reply({
            content: "❌ You do not have permission to use this control.",
            flags: 64
          })
          return
        }

        await interaction.deferUpdate()

        const [, key, view] = interaction.customId.split(":")

        const settingsResult = await fetchSettingsForServer(interaction)

        if (!settingsResult.ok) {
          await interaction.editReply("❌ Could not load settings.")
          return
        }

        const currentValue = Boolean(settingsResult.settings[key])

        const updateResult = await postToAppsScript({
          action: "update_setting_for_server",
          adminKey: process.env.ADMIN_API_KEY,
          discordServerId: interaction.guildId,
          key: key,
          value: String(!currentValue)
        })

        if (!updateResult.ok) {
          await interaction.editReply(`❌ ${updateResult.error || "Could not update setting."}`)
          return
        }

        const refreshed = await fetchSettingsForServer(interaction)

        if (!refreshed.ok) {
          await interaction.editReply("❌ Could not refresh settings.")
          return
        }

        await interaction.editReply(renderSettingsView(refreshed, view || "home"))
        return
      }

      if (interaction.customId === "settings_back:home") {
        if (!userCanManageServer(interaction)) {
          await interaction.reply({
            content: "❌ You do not have permission to use this control.",
            flags: 64
          })
          return
        }

        await interaction.deferUpdate()

        const refreshed = await fetchSettingsForServer(interaction)

        if (!refreshed.ok) {
          await interaction.editReply("❌ Could not load settings.")
          return
        }

        await interaction.editReply(renderSettingsView(refreshed, "home"))
        return
      }

      if (interaction.customId.startsWith("book_page:")) {
        const [, token, pageRaw] = interaction.customId.split(":")
        const entry = getPendingBookingOrThrow(token)

        if (interaction.user.id !== entry.discordUserId) {
          await interaction.reply({
            content: "❌ This booking menu belongs to someone else.",
            flags: 64
          })
          return
        }

        const page = parseInt(pageRaw, 10) || 0
        const ui = buildBookingComponents(token, page)

        await interaction.update(ui)
        return
      }

      return
    }

    if (interaction.isModalSubmit()) {

      if (interaction.customId === "setup_modal") {
        if (!userCanManageServer(interaction)) {
          await interaction.reply({
            content: "❌ You do not have permission to use this command.",
            flags: 64
          })
          return
        }

        await interaction.deferReply({ flags: 64 })

        const state = String(interaction.fields.getTextInputValue("state") || "").trim()
        const serverName = String(interaction.fields.getTextInputValue("server") || "").trim()

        const result = await postToAppsScript({
          action: "setup_state",
          adminKey: process.env.ADMIN_API_KEY,
          stateCode: state,
          discordServerId: interaction.guildId,
          discordServerName: serverName,
          createdBy: interaction.user.username
        })

        if (!result.ok) {
          await interaction.editReply(`❌ ${result.error || "Could not create state."}`)
          return
        }

        await interaction.editReply(
          `✅ State ${result.state_code} created\n\n` +
          `Sheet:\n${result.sheet_url}\n\n` +
          `Booking page:\n${result.booking_url}\n\n` +
          `Join password:\n${result.join_password}\n\n` +
          `Share that password only with trusted server admins.`
        )
        return
      }

      if (interaction.customId === "link_state_modal") {
        if (!userCanManageServer(interaction)) {
          await interaction.reply({
            content: "❌ You do not have permission to use this command.",
            flags: 64
          })
          return
        }

        await interaction.deferReply({ flags: 64 })

        const state = String(interaction.fields.getTextInputValue("state") || "").trim()
        const password = String(interaction.fields.getTextInputValue("password") || "").trim()
        const serverName = String(interaction.fields.getTextInputValue("server") || "").trim()

        const result = await postToAppsScript({
          action: "link_state",
          adminKey: process.env.ADMIN_API_KEY,
          stateCode: state,
          joinPassword: password,
          discordServerId: interaction.guildId,
          discordServerName: serverName,
          createdBy: interaction.user.username
        })

        if (!result.ok) {
          await interaction.editReply(`❌ ${result.error || "Could not link state."}`)
          return
        }

        if (result.already_linked) {
          await interaction.editReply(
            `✅ This server is already linked to state ${result.state_code}\n\n` +
            `Sheet:\n${result.sheet_url}\n\n` +
            `Booking page:\n${result.booking_url}`
          )
          return
        }

        await interaction.editReply(
          `✅ This server is now linked to state ${result.state_code}\n\n` +
          `Sheet:\n${result.sheet_url}\n\n` +
          `Booking page:\n${result.booking_url}`
        )
        return
      }

      if (!interaction.customId.startsWith("book_modal:")) return

      const [, token] = interaction.customId.split(":")
      const entry = getPendingBookingOrThrow(token)

      if (interaction.user.id !== entry.discordUserId) {
        await interaction.reply({
          content: "❌ This booking modal belongs to someone else.",
          flags: 64
        })
        return
      }

      await interaction.deferReply({ flags: 64 })

      let fc, rfc, shards, speedups

      try {
        fc = ensureNumericOrEmpty(readModalValue(interaction.fields, "fc"), "FC")
        rfc = ensureNumericOrEmpty(readModalValue(interaction.fields, "rfc"), "RFC")
        shards = ensureNumericOrEmpty(readModalValue(interaction.fields, "shards"), "Shards")
        speedups = ensureNumericOrEmpty(readModalValue(interaction.fields, "speedups"), "Speed-ups")
      } catch (err) {
        await interaction.editReply(`❌ ${err.message}`)
        return
      }

      const result = await submitBookingFromEntry(entry, {
        fc,
        rfc,
        shards,
        speedups
      })

      if (!result.ok) {
        await interaction.editReply(`❌ ${result.error || "Could not book slot."}`)
        pendingBookings.delete(token)
        return
      }

      let message = `✅ ${result.day} booked for ${result.playerName}\nTime: ${result.time}`

      if (result.moved && result.oldTime) {
        message = `✅ ${result.day} changed for ${result.playerName}\nFrom: ${result.oldTime}\nTo: ${result.time}`
      }

      await interaction.editReply(message)

      let dmMessage =
        `📅 Booking confirmed\n\n` +
        `State: ${result.state_code}\n` +
        `Day: ${result.day}\n` +
        `Time: ${result.time}`

      if (result.moved && result.oldTime) {
        dmMessage =
          `🔁 Booking changed\n\n` +
          `State: ${result.state_code}\n` +
          `Day: ${result.day}\n` +
          `Old time: ${result.oldTime}\n` +
          `New time: ${result.time}`
      }

      await sendBookingDm(interaction.user, dmMessage)

      pendingBookings.delete(token)
      return
    }

    if (!interaction.isChatInputCommand()) return

    if (interaction.commandName === "help") {
      await interaction.reply({
        flags: 64,
        content:
`R.A.C.H.I.E Minister Booking Bot

USER COMMANDS

/register
Register your in-game name, alliance and player ID.

/book
Book a minister slot.

/remove-booking
Cancel your booking.

/my-bookings
See your current bookings.

/times
Check available times for each minister day.


ADMIN COMMANDS

/setup
Create a new state booking sheet.

/link-state
Link this Discord to an existing state.

/unlink-state
Remove a linked Discord server.

/settings
Configure booking requirements and limits.

/open-bookings
Allow players to start booking.

/close-bookings
Stop new bookings.

/reset-state-password
Generate a new join password for linking servers.


TIP

If a slot is taken while you are booking,
simply run /book again and choose another time.`
      })

      return
    }

    if (interaction.commandName === "reset-state-password") {
      await interaction.deferReply({ flags: 64 })

      if (!userCanManageServer(interaction)) {
        await interaction.editReply("❌ You do not have permission to use this command.")
        return
      }

      const result = await postToAppsScript({
        action: "reset_state_password",
        adminKey: process.env.ADMIN_API_KEY,
        discordServerId: interaction.guildId
      })

      if (!result.ok) {
        await interaction.editReply(`❌ ${result.error || "Could not reset state password."}`)
        return
      }

      await interaction.editReply(
        `✅ Join password reset for state ${result.state_code}\n\n` +
        `New password:\n${result.join_password}\n\n` +
        `Share it only with trusted server admins.`
      )
      return
    }

    if (interaction.commandName === "linked-servers") {
  await interaction.deferReply({ flags: 64 })

  if (!userCanManageServer(interaction)) {
    await interaction.editReply("❌ You do not have permission to use this command.")
    return
  }

  const result = await postToAppsScript({
    action: "get_linked_servers_for_current_state",
    adminKey: process.env.ADMIN_API_KEY,
    discordServerId: interaction.guildId
  })

  if (!result.ok) {
    await interaction.editReply(`❌ ${result.error || "Could not load linked servers."}`)
    return
  }

  const links = Array.isArray(result.links) ? result.links : []

  if (!links.length) {
    await interaction.editReply(`State ${result.state_code}\nNo linked Discord servers found.`)
    return
  }

  const lines = links.map((link, index) =>
    `${index + 1}. ${link.discord_server_name || "[unnamed server]"}`
  )

  await interaction.editReply(
    `State ${result.state_code}\n\nLinked Discord servers\n${lines.join("\n")}`
  )
  return
}

    /* -------------------- SHEET LINK -------------------- */

    if (interaction.commandName === "sheet-link") {
      await interaction.deferReply({ flags: 64 })

      if (!userCanManageServer(interaction)) {
        await interaction.editReply("❌ You do not have permission to use this command.")
        return
      }

      const result = await postToAppsScript({
        action: "get_sheet_link_for_server",
        adminKey: process.env.ADMIN_API_KEY,
        discordServerId: interaction.guildId
      })

      if (!result.ok) {
        await interaction.editReply(`❌ ${result.error || "Could not get sheet link."}`)
        return
      }

      await interaction.editReply(
        `State ${result.state_code}\n` +
        `Sheet:\n${result.sheet_url}\n\n` +
        `Booking page:\n${result.booking_url}`
      )
      return
    }

    /* -------------------- SETTINGS -------------------- */

    if (interaction.commandName === "settings") {
      await interaction.deferReply({ flags: 64 })

      if (!userCanManageServer(interaction)) {
        await interaction.editReply("❌ You do not have permission to use this command.")
        return
      }

      const result = await fetchSettingsForServer(interaction)

      if (!result.ok) {
        await interaction.editReply(`❌ ${result.error || "Could not load settings."}`)
        return
      }

      await interaction.editReply(renderSettingsView(result, "home"))
      return
    }

    /* -------------------- OPEN BOOKINGS -------------------- */

if (interaction.commandName === "open-bookings") {
  await interaction.deferReply({ flags: 64 })

  if (!userCanManageServer(interaction)) {
    await interaction.editReply("❌ You do not have permission to use this command.")
    return
  }

  const result = await postToAppsScript({
    action: "open_bookings_for_server",
    adminKey: process.env.ADMIN_API_KEY,
    discordServerId: interaction.guildId
  })

  if (!result.ok) {
    await interaction.editReply(`❌ ${result.error || "Could not open bookings."}`)
    return
  }

  const announcement =
`📢 Bookings are now OPEN for state ${result.state_code}

Use /book to reserve a minister slot.

Available minister days:
Construction
Research
Troop Training

You can check available times with:
/times

Need to update your details first?
Use:
/register

Check your bookings:
/my-bookings

Need help?
Use:
/help`

  try {
    await interaction.channel.send(announcement)
    await interaction.editReply("✅ Bookings opened and announcement posted in this channel.")
  } catch (error) {
    console.error("Could not post open-bookings announcement:", error)
    await interaction.editReply(
      "✅ Bookings opened, but I could not post the announcement in this channel. Check my channel permissions."
    )
  }

  return
}

    /* -------------------- CLOSE BOOKINGS -------------------- */

if (interaction.commandName === "close-bookings") {
  await interaction.deferReply({ flags: 64 })

  if (!userCanManageServer(interaction)) {
    await interaction.editReply("❌ You do not have permission to use this command.")
    return
  }

  const result = await postToAppsScript({
    action: "close_bookings_for_server",
    adminKey: process.env.ADMIN_API_KEY,
    discordServerId: interaction.guildId
  })

  if (!result.ok) {
    await interaction.editReply(`❌ ${result.error || "Could not close bookings."}`)
    return
  }

  const announcement =
`📢 Bookings are now CLOSED for state ${result.state_code}

New bookings are currently disabled.

If you already have a booking you can still check it with:
/my-bookings

Need help?
Use:
/help`

  try {
    await interaction.channel.send(announcement)
    await interaction.editReply("✅ Bookings closed and announcement posted in this channel.")
  } catch (error) {
    console.error("Could not post close-bookings announcement:", error)
    await interaction.editReply(
      "✅ Bookings closed, but I could not post the announcement in this channel. Check my channel permissions."
    )
  }

  return
}

    /* -------------------- SETUP -------------------- */

    if (interaction.commandName === "setup") {
      if (!userCanManageServer(interaction)) {
        await interaction.reply({
          content: "❌ You do not have permission to use this command.",
          flags: 64
        })
        return
      }

      await interaction.showModal(buildSetupModal())
      return
    }

    /* -------------------- LINK STATE -------------------- */

    if (interaction.commandName === "link-state") {
      if (!userCanManageServer(interaction)) {
        await interaction.reply({
          content: "❌ You do not have permission to use this command.",
          flags: 64
        })
        return
      }

      await interaction.showModal(buildLinkStateModal())
      return
    }

    /* -------------------- BOOKING LINK -------------------- */

    if (interaction.commandName === "booking-link") {
      await interaction.deferReply({ flags: 64 })

      const result = await postToAppsScript({
        action: "get_booking_link_for_server",
        adminKey: process.env.ADMIN_API_KEY,
        discordServerId: interaction.guildId
      })

      if (!result.ok) {
        await interaction.editReply(`❌ ${result.error}`)
        return
      }

      await interaction.editReply(`State ${result.state_code}\n${result.booking_url}`)
      return
    }

    /* -------------------- TIMES -------------------- */

    if (interaction.commandName === "times") {
      await interaction.deferReply({ flags: 64 })

      const day = interaction.options.getString("day")

      const result = await postToAppsScript({
        action: "get_times_for_server",
        adminKey: process.env.ADMIN_API_KEY,
        discordServerId: interaction.guildId,
        day: day
      })

      if (!result.ok) {
        await interaction.editReply(`❌ ${result.error}`)
        return
      }

      const timesText = result.times.length
        ? result.times.join(", ")
        : "No times available"

      await interaction.editReply(`${day} times:\n${timesText}`)
      return
    }

    /* -------------------- MY BOOKINGS -------------------- */

    if (interaction.commandName === "my-bookings") {
      await interaction.deferReply({ flags: 64 })

      const result = await postToAppsScript({
        action: "get_my_bookings_for_server",
        adminKey: process.env.ADMIN_API_KEY,
        discordServerId: interaction.guildId,
        discordUserId: interaction.user.id
      })

      if (!result.ok) {
        await interaction.editReply(`❌ ${result.error || "Could not get your bookings."}`)
        return
      }

      const construction = result.bookings?.Construction || "Not booked"
      const research = result.bookings?.Research || "Not booked"
      const troop = result.bookings?.Troop || "Not booked"

      await interaction.editReply(
        `Bookings for ${result.playerName}\n` +
        `Construction: ${construction}\n` +
        `Research: ${research}\n` +
        `Troop: ${troop}`
      )
      return
    }

    /* -------------------- REGISTER -------------------- */

    if (interaction.commandName === "register") {
      await interaction.deferReply({ flags: 64 })

      const alliance = interaction.options.getString("alliance")
      const name = interaction.options.getString("name")
      const id = interaction.options.getString("id")

      const result = await postToAppsScript({
        action: "register_player_for_server",
        adminKey: process.env.ADMIN_API_KEY,
        discordServerId: interaction.guildId,
        discordUserId: interaction.user.id,
        discordTag: interaction.user.tag,
        inGameName: name,
        playerId: id,
        alliance: alliance
      })

      if (!result.ok) {
        await interaction.editReply(`❌ ${result.error}`)
        return
      }

      await interaction.editReply(
        `Registered\nAlliance: ${result.alliance}\nName: ${result.inGameName}\nPlayer ID: ${result.playerId}`
      )
      return
    }

    /* -------------------- MY INFO -------------------- */

    if (interaction.commandName === "my-info") {
      await interaction.deferReply({ flags: 64 })

      const result = await postToAppsScript({
        action: "get_registered_player_for_server",
        adminKey: process.env.ADMIN_API_KEY,
        discordServerId: interaction.guildId,
        discordUserId: interaction.user.id
      })

      if (!result.found) {
        await interaction.editReply("You are not registered")
        return
      }

      await interaction.editReply(
        `Alliance: ${result.alliance}\nName: ${result.inGameName}\nPlayer ID: ${result.playerId}`
      )
      return
    }

    /* -------------------- UNREGISTER -------------------- */

    if (interaction.commandName === "unregister") {
      await interaction.deferReply({ flags: 64 })

      const result = await postToAppsScript({
        action: "delete_registered_player_for_server",
        adminKey: process.env.ADMIN_API_KEY,
        discordServerId: interaction.guildId,
        discordUserId: interaction.user.id
      })

      if (!result.deleted) {
        await interaction.editReply("No registration found")
        return
      }

      await interaction.editReply("Registration deleted")
      return
    }

    /* -------------------- BOOK -------------------- */

    if (interaction.commandName === "book") {
      await interaction.deferReply({ flags: 64 })

      const day = interaction.options.getString("day")

      const [timesResult, configResult] = await Promise.all([
        postToAppsScript({
          action: "get_times_for_server",
          adminKey: process.env.ADMIN_API_KEY,
          discordServerId: interaction.guildId,
          day: day
        }),
        postToAppsScript({
          action: "get_booking_config_for_server",
          adminKey: process.env.ADMIN_API_KEY,
          discordServerId: interaction.guildId
        })
      ])

      if (!timesResult.ok) {
        await interaction.editReply(`❌ ${timesResult.error || "Could not load available times."}`)
        return
      }

      if (!configResult.ok) {
        await interaction.editReply(`❌ ${configResult.error || "Could not load booking settings."}`)
        return
      }

      if (!Array.isArray(timesResult.times) || timesResult.times.length === 0) {
        await interaction.editReply(`❌ No available times found for ${day}.`)
        return
      }

      const token = createBookingToken()

      pendingBookings.set(token, {
        createdAt: Date.now(),
        discordServerId: interaction.guildId,
        discordUserId: interaction.user.id,
        day: day,
        times: timesResult.times,
        config: configResult
      })

      const ui = buildBookingComponents(token, 0)
      await interaction.editReply(ui)
      return
    }

    /* -------------------- UNLINK STATE -------------------- */

    if (interaction.commandName === "unlink-state") {
      await interaction.deferReply({ flags: 64 })

      if (!userCanManageServer(interaction)) {
        await interaction.editReply("❌ You do not have permission to use this command.")
        return
      }

      const result = await postToAppsScript({
        action: "get_linked_servers_for_current_state",
        adminKey: process.env.ADMIN_API_KEY,
        discordServerId: interaction.guildId
      })

      if (!result.ok) {
        await interaction.editReply(`❌ ${result.error || "Could not load linked servers."}`)
        return
      }

      const links = Array.isArray(result.links) ? result.links : []

      if (!links.length) {
        await interaction.editReply("❌ No linked Discord servers were found for this state.")
        return
      }

      const token = Math.random().toString(36).slice(2, 10) + Date.now().toString(36)
      pendingUnlinks.set(token, {
        stateCode: result.state_code,
        requestedBy: interaction.user.id
      })

      const options = links.slice(0, 25).map(link => ({
        label: String(link.discord_server_name || link.discord_server_id).slice(0, 100),
        value: String(link.discord_server_id),
        description: `Server ID: ${String(link.discord_server_id).slice(0, 80)}`
      }))

      const row = new ActionRowBuilder().addComponents(
        new StringSelectMenuBuilder()
          .setCustomId(`unlink_select:${token}`)
          .setPlaceholder("Select a linked Discord server to remove")
          .addOptions(options)
      )

      await interaction.editReply({
        content: `State ${result.state_code}\nSelect a linked Discord server to unlink.`,
        components: [row]
      })
      return
    }

    /* -------------------- REMOVE BOOKING -------------------- */

    if (interaction.commandName === "remove-booking") {
      await interaction.deferReply({ flags: 64 })

      const day = interaction.options.getString("day")

      const result = await postToAppsScript({
        action: "remove_booking_for_server",
        adminKey: process.env.ADMIN_API_KEY,
        discordServerId: interaction.guildId,
        discordUserId: interaction.user.id,
        day: day
      })

      if (!result.ok) {
        await interaction.editReply(`❌ ${result.error}`)
        return
      }

      if (!result.removed) {
        await interaction.editReply(`You do not have a ${day} booking`)
        return
      }

      await interaction.editReply(
        `${day} booking removed\nPrevious time: ${result.oldTime}`
      )

      await sendBookingDm(
        interaction.user,
        `❌ Booking removed\n\n` +
        `State: ${result.state_code}\n` +
        `Day: ${day}\n` +
        `Previous time: ${result.oldTime}`
      )

      return
    }

  } catch (error) {
    console.error(error)

    try {
      if (interaction.deferred || interaction.replied) {
        await interaction.editReply("Something went wrong")
      } else {
        await interaction.reply({ content: "Something went wrong", flags: 64 })
      }
    } catch (replyError) {
      console.error(replyError)
    }
  }
})

/* -------------------- START BOT -------------------- */

async function main() {
  try {
    console.log("Registering slash commands...")
    await registerCommands()

    console.log("Logging in...")
    await client.login(process.env.BOT_TOKEN)
  } catch (error) {
    console.error("Bot startup failed")
    console.error(error)
  }
}

main()