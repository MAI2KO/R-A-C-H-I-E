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
const OpenAI = require("openai")

console.log("Starting bot...")
console.log("CLIENT_ID present:", !!process.env.CLIENT_ID)
console.log("BOT_TOKEN present:", !!process.env.BOT_TOKEN)
console.log("APPS_SCRIPT_URL present:", !!process.env.APPS_SCRIPT_URL)
console.log("ADMIN_API_KEY present:", !!process.env.ADMIN_API_KEY)
console.log("OPENAI_API_KEY present:", !!process.env.OPENAI_API_KEY)

const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY
})
const client = new Client({
  intents: [
    GatewayIntentBits.Guilds,
    GatewayIntentBits.GuildMessages,
    GatewayIntentBits.MessageContent
  ]
})

const pendingUnlinks = new Map()
const pendingBookings = new Map()
const pendingAdminBookings = new Map()
const pendingAdminReserves = new Map()
const pendingAdminRemoveReserved = new Map()

const messageBuffers = new Map()
const channelCooldowns = new Map()

const MESSAGE_LIMIT = 10
const COOLDOWN_MS = 30 * 60 * 1000

const BOOKING_PAGE_SIZE = 25
const BOOKING_TTL_MS = 15 * 60 * 1000

const banterConfigCache = new Map()
const BANTER_CONFIG_TTL_MS = 5 * 60 * 1000

function createBookingToken() {
  return Math.random().toString(36).slice(2, 10) + Date.now().toString(36)
}

function cleanupPendingSessions() {
  const now = Date.now()

  const maps = [
   pendingBookings,
   pendingAdminBookings,
   pendingAdminReserves,
   pendingAdminRemoveReserved
  ]

  for (const map of maps) {
    for (const [token, entry] of map.entries()) {
      if (!entry?.createdAt || now - entry.createdAt > BOOKING_TTL_MS) {
        map.delete(token)
      }
    }
  }
}

function getPendingBookingOrThrow(token) {
  cleanupPendingSessions()

  const entry = pendingBookings.get(token)
  if (!entry) {
    throw new Error("This booking session has expired. Run /book again.")
  }

  return entry
}

function getPendingAdminBookingOrThrow(token) {
  cleanupPendingSessions()

  const entry = pendingAdminBookings.get(token)
  if (!entry) {
    throw new Error("This admin booking session has expired. Run /admin-add-booking again.")
  }

  return entry
}

function getPendingAdminReserveOrThrow(token) {
  cleanupPendingSessions()

  const entry = pendingAdminReserves.get(token)
  if (!entry) {
    throw new Error("This reserve session has expired. Run /admin-reserve-slots again.")
  }

  return entry
}

function getPendingAdminRemoveReservedOrThrow(token) {
  cleanupPendingSessions()

  const entry = pendingAdminRemoveReserved.get(token)
  if (!entry) {
    throw new Error("This reserved-slot removal session has expired. Run /admin-remove-reserved again.")
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

  return {
    content:
      `Select a time for ${entry.day}\n` +
      `Page ${safePage + 1} of ${totalPages}`,
    components: [selectRow, buttons]
  }
}

function buildAdminBookingComponents(token, page = 0) {
  const entry = getPendingAdminBookingOrThrow(token)
  const totalPages = Math.max(1, Math.ceil(entry.times.length / BOOKING_PAGE_SIZE))
  const safePage = Math.max(0, Math.min(page, totalPages - 1))

  const start = safePage * BOOKING_PAGE_SIZE
  const pageTimes = entry.times.slice(start, start + BOOKING_PAGE_SIZE)

  const select = new StringSelectMenuBuilder()
    .setCustomId(`admin_book_select:${token}:${safePage}`)
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
      .setCustomId(`admin_book_page:${token}:${safePage - 1}`)
      .setLabel("Previous")
      .setStyle(ButtonStyle.Secondary)
      .setDisabled(safePage <= 0),
    new ButtonBuilder()
      .setCustomId(`admin_book_page:${token}:${safePage + 1}`)
      .setLabel("Next")
      .setStyle(ButtonStyle.Secondary)
      .setDisabled(safePage >= totalPages - 1)
  )

  return {
    content:
      `Admin booking for ${entry.day}\n` +
      `Select a time\n` +
      `Page ${safePage + 1} of ${totalPages}`,
    components: [selectRow, buttons]
  }
}

function buildAdminReserveComponents(token, page = 0) {
  const entry = getPendingAdminReserveOrThrow(token)
  const totalPages = Math.max(1, Math.ceil(entry.times.length / BOOKING_PAGE_SIZE))
  const safePage = Math.max(0, Math.min(page, totalPages - 1))

  const start = safePage * BOOKING_PAGE_SIZE
  const pageTimes = entry.times.slice(start, start + BOOKING_PAGE_SIZE)

  const select = new StringSelectMenuBuilder()
    .setCustomId(`admin_reserve_select:${token}:${safePage}`)
    .setPlaceholder(pageTimes.length ? "Select up to 5 times" : "No times available")
    .setDisabled(pageTimes.length === 0)
    .setMinValues(1)
    .setMaxValues(Math.min(5, pageTimes.length))
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
      .setCustomId(`admin_reserve_page:${token}:${safePage - 1}`)
      .setLabel("Previous")
      .setStyle(ButtonStyle.Secondary)
      .setDisabled(safePage <= 0),
    new ButtonBuilder()
      .setCustomId(`admin_reserve_page:${token}:${safePage + 1}`)
      .setLabel("Next")
      .setStyle(ButtonStyle.Secondary)
      .setDisabled(safePage >= totalPages - 1)
  )

  return {
    content:
      `Admin reserve for ${entry.day}\n` +
      `Select up to 5 time slots\n` +
      `Page ${safePage + 1} of ${totalPages}`,
    components: [selectRow, buttons]
  }
}

function buildAdminRemoveReservedComponents(token, page = 0) {
  const entry = getPendingAdminRemoveReservedOrThrow(token)
  const totalPages = Math.max(1, Math.ceil(entry.times.length / BOOKING_PAGE_SIZE))
  const safePage = Math.max(0, Math.min(page, totalPages - 1))

  const start = safePage * BOOKING_PAGE_SIZE
  const pageTimes = entry.times.slice(start, start + BOOKING_PAGE_SIZE)

  const select = new StringSelectMenuBuilder()
    .setCustomId(`admin_remove_reserved_select:${token}:${safePage}`)
    .setPlaceholder(pageTimes.length ? "Select reserved slots to remove" : "No reserved slots found")
    .setDisabled(pageTimes.length === 0)
    .setMinValues(pageTimes.length ? 1 : 0)
    .setMaxValues(pageTimes.length ? Math.min(5, pageTimes.length) : 1)
    .addOptions(
      pageTimes.length
        ? pageTimes.map(time => ({
            label: time,
            value: time,
            description: `${entry.day} reserved slot`
          }))
        : [{ label: "No reserved slots", value: "none", description: "No reserved slots" }]
    )

  const selectRow = new ActionRowBuilder().addComponents(select)

  const buttons = new ActionRowBuilder().addComponents(
    new ButtonBuilder()
      .setCustomId(`admin_remove_reserved_page:${token}:${safePage - 1}`)
      .setLabel("Previous")
      .setStyle(ButtonStyle.Secondary)
      .setDisabled(safePage <= 0),
    new ButtonBuilder()
      .setCustomId(`admin_remove_reserved_page:${token}:${safePage + 1}`)
      .setLabel("Next")
      .setStyle(ButtonStyle.Secondary)
      .setDisabled(safePage >= totalPages - 1)
  )

  return {
    content:
      `Admin remove reserved slots for ${entry.day}\n` +
      `Select up to 5 reserved time slots\n` +
      `Page ${safePage + 1} of ${totalPages}`,
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

function buildAdminAddBookingModal(token) {
  return new ModalBuilder()
    .setCustomId(`admin_add_booking_modal:${token}`)
    .setTitle("Add player booking")
    .addComponents(
      new ActionRowBuilder().addComponents(
        new TextInputBuilder()
          .setCustomId("alliance")
          .setLabel("Alliance tag")
          .setStyle(TextInputStyle.Short)
          .setRequired(true)
          .setMaxLength(3)
          .setPlaceholder("Example: ABC")
      ),
      new ActionRowBuilder().addComponents(
        new TextInputBuilder()
          .setCustomId("name")
          .setLabel("Player name")
          .setStyle(TextInputStyle.Short)
          .setRequired(true)
          .setMaxLength(30)
          .setPlaceholder("Example: PlayerName")
      ),
      new ActionRowBuilder().addComponents(
        new TextInputBuilder()
          .setCustomId("player_id")
          .setLabel("Player ID")
          .setStyle(TextInputStyle.Short)
          .setRequired(true)
          .setMaxLength(20)
          .setPlaceholder("Numbers only")
      )
    )
}

function buildAdminRemoveBookingModal(day) {
  return new ModalBuilder()
    .setCustomId(`admin_remove_booking_modal:${day}`)
    .setTitle("Remove player booking")
    .addComponents(
      new ActionRowBuilder().addComponents(
        new TextInputBuilder()
          .setCustomId("player_id")
          .setLabel("Player ID")
          .setStyle(TextInputStyle.Short)
          .setRequired(true)
          .setMaxLength(20)
          .setPlaceholder("Enter the player's ID")
      )
    )
}

function buildGrantAccessModal() {
  return new ModalBuilder()
    .setCustomId("grant_access_modal")
    .setTitle("Grant sheet access")
    .addComponents(
      new ActionRowBuilder().addComponents(
        new TextInputBuilder()
          .setCustomId("email")
          .setLabel("Email address")
          .setStyle(TextInputStyle.Short)
          .setRequired(true)
          .setMaxLength(100)
          .setPlaceholder("example@gmail.com")
      )
    )
}

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

function buildClearBookingsModal() {
  return new ModalBuilder()
    .setCustomId("clear_bookings_modal")
    .setTitle("Clear all bookings")
    .addComponents(
      new ActionRowBuilder().addComponents(
        new TextInputBuilder()
          .setCustomId("confirm")
          .setLabel("Type CLEAR to confirm")
          .setStyle(TextInputStyle.Short)
          .setRequired(true)
          .setMaxLength(5)
          .setPlaceholder("CLEAR")
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

function buildRegisterModal() {
  return new ModalBuilder()
    .setCustomId("register_modal")
    .setTitle("Register player details")
    .addComponents(
      new ActionRowBuilder().addComponents(
        new TextInputBuilder()
          .setCustomId("alliance")
          .setLabel("Alliance tag")
          .setStyle(TextInputStyle.Short)
          .setRequired(true)
          .setMaxLength(3)
          .setPlaceholder("Example: YOU")
      ),
      new ActionRowBuilder().addComponents(
        new TextInputBuilder()
          .setCustomId("name")
          .setLabel("In game name")
          .setStyle(TextInputStyle.Short)
          .setRequired(true)
          .setMaxLength(30)
          .setPlaceholder("Example: NO BEND")
      ),
      new ActionRowBuilder().addComponents(
        new TextInputBuilder()
          .setCustomId("id")
          .setLabel("Player ID numbers only")
          .setStyle(TextInputStyle.Short)
          .setRequired(true)
          .setMaxLength(20)
          .setPlaceholder("Example: 8008135")
      )
    )
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

async function userCanManageServer(interaction) {
  if (interaction.memberPermissions?.has(PermissionFlagsBits.ManageGuild)) {
    return true
  }

  try {
    const result = await postToAppsScript({
      action: "get_bot_admin_role_for_server",
      adminKey: process.env.ADMIN_API_KEY,
      discordServerId: interaction.guildId
    })

    const roleId = String(result.bot_admin_role_id || "").trim()
    if (!roleId) return false

    return interaction.member?.roles?.cache?.has(roleId) || false
  } catch (error) {
    console.error("Permission check failed:", error)
    return false
  }
}

function yesNo(value) {
  return value ? "ON" : "OFF"
}

function isTooAggressive(text) {
  const banned = ["idiot", "moron", "retard"]
  return banned.some(word => text.toLowerCase().includes(word))
}

function soundsTooForcedBritish(text) {
  const bannedPhrases = [
    "blimey",
    "old cobblers",
    "good heavens",
    "crikey",
    "haven't we",
    "what's left for us",
    "talk about"
  ]

  const lower = text.toLowerCase()
  return bannedPhrases.some(phrase => lower.includes(phrase))
}

async function sendBookingDm(user, message) {
  try {
    await user.send(message)
  } catch (error) {
    console.log("Could not send DM:", error?.code, error?.message)
  }
}

async function fetchSettingsForServer(interaction) {
  return await postToAppsScript({
    action: "get_settings_for_server",
    adminKey: process.env.ADMIN_API_KEY,
    discordServerId: interaction.guildId
  })
}

async function getBanterConfigForGuild(guildId) {
  const cached = banterConfigCache.get(guildId)

  if (cached && Date.now() - cached.timestamp < BANTER_CONFIG_TTL_MS) {
    return cached.data
  }

  const result = await postToAppsScript({
    action: "get_banter_config_for_server",
    adminKey: process.env.ADMIN_API_KEY,
    discordServerId: guildId
  })

  const data = {
    banterChannelId: String(result.banter_channel_id || "").trim(),
    spiceLevel: result.spice_level || "standard"
  }

  banterConfigCache.set(guildId, {
    data,
    timestamp: Date.now()
  })

  return data
}

async function triggerBanter(channel, messages, spiceLevel = "standard") {
  if (!openai) return

  try {
    const combinedLength = messages.reduce((total, m) => total + m.content.length, 0)
    if (combinedLength < 120) return

    const textBlock = messages
      .map(m => `${m.author}: ${m.content}`)
      .join("\n")

    let spiceInstruction = `
Keep it playful, dry and sharp.
Use light mockery.
Avoid strong vulgarity.
`

if (spiceLevel === "mild") {
  spiceInstruction = `
Keep it light, teasing and observational.
Prioritise bemused or deadpan reactions over insults.
Avoid vulgarity.
Avoid calling people embarrassing, weird, pathetic, or similar.
Sound more amused than savage.
`
} else if (spiceLevel === "spicy") {
  spiceInstruction = `
Be noticeably sharper, cheekier and more cutting.
Mild vulgarity is allowed if it fits naturally.
You can sound more dismissive, more fed up, and more personally mocking.
Lean more towards a roast than a light tease.
Do not become abusive, hateful, or bullying.
`
}

    let prompt = `
You are R.A.C.H.I.E, a witty Manchester woman in a Discord server.

Read the last messages and find ONE specific message, opinion, claim, or short exchange that is the easiest to mock playfully.

If nothing clearly stands out, return exactly:
NO_REPLY

If something does stand out, reply with one short, natural, context-aware line reacting to that specific part of the chat.

Voice:
- dry, sharp, playful
- natural northern English, lightly Manc
- sounds like a real person in chat
- mildly vulgar is fine
- not theatrical, not exaggerated

Rules:
- under 14 words
- 1 sentence only
- react to one specific moment, not the whole chat in general
- do not explain what you are doing
- no slurs
- no direct harassment
- no generic filler
- no forced British phrases
- do not sound like a stereotype
- only use words like muppet, sausage, or absolute salad occasionally and only if they fit naturally
- avoid repeating stock phrases like "not you", "state of this", or "you lot" too often
- make the tone clearly match the requested spice level

Good examples for mild:
- that is a proper odd thing to say
- bold claim for this time of day
- that’s a strange hill to stand on
- fair enough, but that is still nonsense

Good examples for standard:
- that logic’s in the bin
- proper weird thing to admit out loud
- bold thing to say with your chest
- you’ve fully embarrassed yourself there

Good examples for spicy:
- that is a shocking amount of confidence for such a daft point
- you are chatting complete shite there
- that might be the dumbest thing said all hour
- proper clown behaviour, that

Bad examples:
- generic insults that could fit any chat
- comments about the whole conversation unless one clear pattern stands out
- random British caricature phrases
- repeating the same stock opener every time

Spice:
${spiceInstruction}

Conversation:
${textBlock}
`

    if (Math.random() < 0.2) {
      const slangOptions = [
        "muppet",
        "sausage",
        "you salad",
        "weapon",
        "donut",
        "proper clown behaviour"
      ]

      const slang = slangOptions[Math.floor(Math.random() * slangOptions.length)]
      prompt += `\nIf it genuinely fits, you may naturally use a playful insult like "${slang}", but do not force it.`
    }

    const response = await openai.chat.completions.create({
      model: "gpt-4o-mini",
      messages: [{ role: "user", content: prompt }],
      temperature: 0.75,
      max_tokens: 40
    })

    const reply = response.choices[0]?.message?.content?.trim()

    if (!reply || reply === "NO_REPLY") return
    if (isTooAggressive(reply)) return
    if (soundsTooForcedBritish(reply)) return

    await channel.send(reply)
  } catch (err) {
    console.error("Banter error:", err)
  }
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
  .setName("set-banter-spice")
  .setDescription("Set how spicy R.A.C.H.I.E banter is")
  .addStringOption(option =>
    option
      .setName("level")
      .setDescription("Spice level")
      .setRequired(true)
      .addChoices(
        { name: "Mild", value: "mild" },
        { name: "Standard", value: "standard" },
        { name: "Spicy", value: "spicy" }
      )
  )
  .setDefaultMemberPermissions(PermissionFlagsBits.ManageGuild),

  new SlashCommandBuilder()
    .setName("set-banter-channel")
    .setDescription("Choose the channel where R.A.C.H.I.E banter is allowed")
    .addChannelOption(option =>
      option
       .setName("channel")
       .setDescription("Channel for banter")
       .addChannelTypes(0)
       .setRequired(true)
      )
    .setDefaultMemberPermissions(PermissionFlagsBits.ManageGuild),

  new SlashCommandBuilder()
    .setName("clear-banter-channel")
    .setDescription("Disable the assigned R.A.C.H.I.E banter channel")
    .setDefaultMemberPermissions(PermissionFlagsBits.ManageGuild),

  new SlashCommandBuilder()
    .setName("set-bot-admin-role")
    .setDescription("Set the role allowed to manage R.A.C.H.I.E")
    .addRoleOption(option =>
      option
       .setName("role")
       .setDescription("Role to allow")
       .setRequired(true)
      )
    .setDefaultMemberPermissions(PermissionFlagsBits.ManageGuild),

  new SlashCommandBuilder()
    .setName("clear-bot-admin-role")
    .setDescription("Clear the custom R.A.C.H.I.E admin role")
    .setDefaultMemberPermissions(PermissionFlagsBits.ManageGuild),

  new SlashCommandBuilder()
    .setName("banter-test")
    .setDescription("Test the AI banter reply in this channel")
    .setDefaultMemberPermissions(PermissionFlagsBits.ManageGuild),

  new SlashCommandBuilder()
    .setName("admin-remove-booking")
    .setDescription("Remove a booking for a player")
    .addStringOption(option =>
      option
        .setName("day")
        .setDescription("Which booking to remove")
        .setRequired(true)
        .addChoices(
          { name: "Construction", value: "Construction" },
          { name: "Research", value: "Research" },
          { name: "Troop", value: "Troop" },
          { name: "All days", value: "ALL" }
        )
    )
    .setDefaultMemberPermissions(PermissionFlagsBits.ManageGuild),
  
  new SlashCommandBuilder()
    .setName("setup-help")
    .setDescription("Show setup instructions for admins")
    .setDefaultMemberPermissions(PermissionFlagsBits.ManageGuild),  

  new SlashCommandBuilder()
    .setName("set-announcements")
    .setDescription("Set the channel for booking announcements")
    .addChannelOption(option =>
      option
        .setName("channel")
        .setDescription("Channel to send announcements in")
        .addChannelTypes(0)
        .setRequired(true)
    )
    .setDefaultMemberPermissions(PermissionFlagsBits.ManageGuild),

  new SlashCommandBuilder()
    .setName("admin-help")
    .setDescription("Show admin help for managing the booking system")
    .setDefaultMemberPermissions(PermissionFlagsBits.ManageGuild),

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
    .setName("clear-bookings")
    .setDescription("Clear all booking entries from the booking sheet")
    .setDefaultMemberPermissions(PermissionFlagsBits.ManageGuild),

  new SlashCommandBuilder()
    .setName("my-bookings")
    .setDescription("Show your current bookings for this state"),

  new SlashCommandBuilder()
    .setName("booking-link")
    .setDescription("Get the booking link for this server"),

  new SlashCommandBuilder()
    .setName("times")
    .setDescription("Check available times and the current booking date")
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
    .setDescription("Register your player details"),

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
    .setName("admin-reserve-slots")
    .setDescription("Reserve up to 5 booking slots")
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
    .setDefaultMemberPermissions(PermissionFlagsBits.ManageGuild),

  new SlashCommandBuilder()
    .setName("admin-remove-reserved")
    .setDescription("Remove reserved booking slots")
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
    .setDefaultMemberPermissions(PermissionFlagsBits.ManageGuild),

  new SlashCommandBuilder()
    .setName("reset-state-password")
    .setDescription("Generate a new join password for a state")
    .setDefaultMemberPermissions(PermissionFlagsBits.ManageGuild),

  new SlashCommandBuilder()
    .setName("grant-access")
    .setDescription("Give a user edit access to the booking sheet")
    .setDefaultMemberPermissions(PermissionFlagsBits.ManageGuild),

  new SlashCommandBuilder()
    .setName("admin-add-booking")
    .setDescription("Add a booking for a player")
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
    .setDefaultMemberPermissions(PermissionFlagsBits.ManageGuild),

    new SlashCommandBuilder()
     .setName("set-booking-date")
     .setDescription("Set the booking date for a minister day")
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
     .addIntegerOption(option =>
        option
        .setName("year")
        .setDescription("Year in UTC")
        .setRequired(true)
        .setMinValue(2024)
        .setMaxValue(2100)
      )
      .addIntegerOption(option =>
        option
         .setName("month")
         .setDescription("Month number")
         .setRequired(true)
         .setMinValue(1)
         .setMaxValue(12)
      )
      .addIntegerOption(option =>
        option
         .setName("day_of_month")
         .setDescription("Day of month")
         .setRequired(true)
         .setMinValue(1)
         .setMaxValue(31)
      )
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

/* -------------------- COMMAND HANDLER -------------------- */

client.on("interactionCreate", async interaction => {
  try {
    if (interaction.isStringSelectMenu()) {
      if (interaction.customId === "settings_max_bookings_select") {
        if (!(await userCanManageServer(interaction))) {
          await interaction.reply({
            content: "❌ You do not have permission to use this control.",
            flags: 64
          })
          return
        }

        await interaction.deferUpdate()

        await interaction.editReply({
          content: "⏳ Updating settings...",
          components: []
        })

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
        if (!(await userCanManageServer(interaction))) {
          await interaction.reply({
            content: "❌ You do not have permission to use this control.",
            flags: 64
          })
          return
        }

        await interaction.deferUpdate()

         await interaction.editReply({
          content: "⏳ Updating settings...",
          components: []
        })

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

         await interaction.editReply({
          content: `⏳ Booking ${time} for ${entry.day}...`,
          components: []
        })

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
         `Date: ${result.booking_date_display || "No date set"}\n` +
         `Time: ${result.time} UTC`

        if (result.moved && result.oldTime) {
         dmMessage =
          `🔁 Booking changed\n\n` +
          `State: ${result.state_code}\n` +
          `Day: ${result.day}\n` +
          `Date: ${result.booking_date_display || "No date set"}\n` +
          `Old time: ${result.oldTime} UTC\n` +
          `New time: ${result.time} UTC`
        }

        await sendBookingDm(interaction.user, dmMessage)

        pendingBookings.delete(token)
        return
      }

      if (interaction.customId.startsWith("admin_book_select:")) {
        const [, token] = interaction.customId.split(":")
        const entry = getPendingAdminBookingOrThrow(token)

        if (interaction.user.id !== entry.requestedBy) {
          await interaction.reply({
            content: "❌ This admin booking menu belongs to someone else.",
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
        pendingAdminBookings.set(token, entry)

        await interaction.showModal(buildAdminAddBookingModal(token))
        return
      }

      if (interaction.customId.startsWith("admin_reserve_select:")) {
        const [, token] = interaction.customId.split(":")
        const entry = getPendingAdminReserveOrThrow(token)

        if (interaction.user.id !== entry.requestedBy) {
          await interaction.reply({
            content: "❌ This reserve menu belongs to someone else.",
            flags: 64
          })
          return
        }

        const times = interaction.values.filter(v => v && v !== "none")

        if (!times.length) {
          await interaction.reply({
            content: "❌ No valid times selected.",
            flags: 64
          })
          return
        }

        await interaction.deferUpdate()

         await interaction.editReply({
          content: `⏳ Reserving selected slot(s) for ${entry.day}...`,
          components: []
        })

        const result = await postToAppsScript({
          action: "admin_reserve_slots_for_server",
          adminKey: process.env.ADMIN_API_KEY,
          discordServerId: interaction.guildId,
          day: entry.day,
          times: times
        })

        pendingAdminReserves.delete(token)

        if (!result.ok) {
          await interaction.editReply({
            content: `❌ ${result.error || "Could not reserve slots."}`,
            components: []
          })
          return
        }

        let message = `✅ Reserved ${result.count} slot(s) for ${result.day}`

        if (Array.isArray(result.times) && result.times.length) {
          message += `\nSuccess: ${result.times.join(", ")}`
        }

        if (Array.isArray(result.failed) && result.failed.length) {
          message += `\nFailed: ${result.failed.map(x => `${x.time} (${x.error})`).join(", ")}`
        }

        await interaction.editReply({
          content: message,
          components: []
        })
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

        await interaction.editReply({
          content: "⏳ Unlinking selected server...",
          components: []
        })

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

      if (interaction.customId.startsWith("admin_remove_reserved_select:")) {
  const [, token] = interaction.customId.split(":")
  const entry = getPendingAdminRemoveReservedOrThrow(token)

  if (interaction.user.id !== entry.requestedBy) {
    await interaction.reply({
      content: "❌ This reserved-slot removal menu belongs to someone else.",
      flags: 64
    })
    return
  }

  const times = interaction.values.filter(v => v && v !== "none")

  if (!times.length) {
    await interaction.reply({
      content: "❌ No valid reserved slots selected.",
      flags: 64
    })
    return
  }

  await interaction.deferUpdate()

  await interaction.editReply({
    content: `⏳ Removing reserved slot(s) for ${entry.day}...`,
    components: []
  })

  const result = await postToAppsScript({
    action: "admin_remove_reserved_slots_for_server",
    adminKey: process.env.ADMIN_API_KEY,
    discordServerId: interaction.guildId,
    day: entry.day,
    times: times
  })

  pendingAdminRemoveReserved.delete(token)

  if (!result.ok) {
    await interaction.editReply({
      content: `❌ ${result.error || "Could not remove reserved slots."}`,
      components: []
    })
    return
  }

  let message = `✅ Removed ${result.count} reserved slot(s) for ${result.day}`

  if (Array.isArray(result.times) && result.times.length) {
    message += `\nRemoved: ${result.times.join(", ")}`
  }

  if (Array.isArray(result.failed) && result.failed.length) {
    message += `\nFailed: ${result.failed.map(x => `${x.time} (${x.error})`).join(", ")}`
  }

  await interaction.editReply({
    content: message,
    components: []
  })
  return
}

      return
    }

    if (interaction.isButton()) {
      if (interaction.customId.startsWith("settings_group:")) {
        if (!(await userCanManageServer(interaction))) {
          await interaction.reply({
            content: "❌ You do not have permission to use this control.",
            flags: 64
          })
          return
        }

        await interaction.deferUpdate()

         await interaction.editReply({
          content: "⏳ Loading settings...",
          components: []
        })

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
        if (!(await userCanManageServer(interaction))) {
          await interaction.reply({
            content: "❌ You do not have permission to use this control.",
            flags: 64
          })
          return
        }

        await interaction.deferUpdate()

         await interaction.editReply({
          content: "⏳ Updating settings...",
          components: []
        })

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
        if (!(await userCanManageServer(interaction))) {
          await interaction.reply({
            content: "❌ You do not have permission to use this control.",
            flags: 64
          })
          return
        }

        await interaction.deferUpdate()

         await interaction.editReply({
          content: "⏳ Loading settings...",
          components: []
        })

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

      if (interaction.customId.startsWith("admin_book_page:")) {
        const [, token, pageRaw] = interaction.customId.split(":")
        const entry = getPendingAdminBookingOrThrow(token)

        if (interaction.user.id !== entry.requestedBy) {
          await interaction.reply({
            content: "❌ This admin booking menu belongs to someone else.",
            flags: 64
          })
          return
        }

        const page = parseInt(pageRaw, 10) || 0
        const ui = buildAdminBookingComponents(token, page)

        await interaction.update(ui)
        return
      }

      if (interaction.customId.startsWith("admin_reserve_page:")) {
        const [, token, pageRaw] = interaction.customId.split(":")
        const entry = getPendingAdminReserveOrThrow(token)

        if (interaction.user.id !== entry.requestedBy) {
          await interaction.reply({
            content: "❌ This reserve menu belongs to someone else.",
            flags: 64
          })
          return
        }

        const page = parseInt(pageRaw, 10) || 0
        const ui = buildAdminReserveComponents(token, page)

        await interaction.update(ui)
        return
      }

      if (interaction.customId.startsWith("admin_remove_reserved_page:")) {
        const [, token, pageRaw] = interaction.customId.split(":")
        const entry = getPendingAdminRemoveReservedOrThrow(token)

        if (interaction.user.id !== entry.requestedBy) {
         await interaction.reply({
           content: "❌ This reserved-slot removal menu belongs to someone else.",
           flags: 64
          })
           return
        }

          const page = parseInt(pageRaw, 10) || 0
          const ui = buildAdminRemoveReservedComponents(token, page)

          await interaction.update(ui)
          return
        }

      return
    }

    if (interaction.isModalSubmit()) {
      if (interaction.customId === "setup_modal") {
        if (!(await userCanManageServer(interaction))) {
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
          `Share that password only with trusted server admins.\n\n` +
          `Next step:\n` +
          `Run /set-announcements and choose the channel where booking updates should be posted.`
        )
        return
      }

      if (interaction.customId === "link_state_modal") {
        if (!(await userCanManageServer(interaction))) {
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
            `Booking page:\n${result.booking_url}\n\n` +
            `You can run /set-announcements at any time to change where booking updates are posted.`
          )
          return
        }

        await interaction.editReply(
          `✅ This server is now linked to state ${result.state_code}\n\n` +
          `Sheet:\n${result.sheet_url}\n\n` +
          `Booking page:\n${result.booking_url}\n\n` +
          `Next step:\n` +
          `Run /set-announcements and choose the channel where booking updates should be posted.`
        )
        return
      }

      if (interaction.customId === "register_modal") {
        await interaction.deferReply({ flags: 64 })

        const allianceRaw = String(interaction.fields.getTextInputValue("alliance") || "").trim()
        const name = String(interaction.fields.getTextInputValue("name") || "").trim()
        const id = String(interaction.fields.getTextInputValue("id") || "").trim()

        const alliance = allianceRaw.toUpperCase()

        if (!/^[A-Z0-9]{3}$/.test(alliance)) {
          await interaction.editReply("❌ Alliance tag must be 3 letters or numbers only.")
          return
        }

        if (!/^[0-9]+$/.test(id)) {
          await interaction.editReply("❌ Player ID must contain numbers only.")
          return
        }

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

      if (interaction.customId === "clear_bookings_modal") {
        if (!(await userCanManageServer(interaction))) {
          await interaction.reply({
            content: "❌ You do not have permission to use this command.",
            flags: 64
          })
          return
        }

        await interaction.deferReply({ flags: 64 })

        const confirm = String(interaction.fields.getTextInputValue("confirm") || "").trim()

        if (confirm !== "CLEAR") {
          await interaction.editReply(
            "❌ Clear cancelled. You must type CLEAR exactly to wipe all booking entries from the sheet."
          )
          return
        }

        const result = await postToAppsScript({
          action: "clear_bookings_for_server",
          adminKey: process.env.ADMIN_API_KEY,
          discordServerId: interaction.guildId
        })

        if (!result.ok) {
          await interaction.editReply(`❌ ${result.error || "Could not clear bookings."}`)
          return
        }

        await interaction.editReply(
          `✅ All booking entries have been cleared for state ${result.state_code}.`
        )
        return
      }

      if (interaction.customId === "grant_access_modal") {
        if (!(await userCanManageServer(interaction))) {
          await interaction.reply({
            content: "❌ You do not have permission to use this command.",
            flags: 64
          })
          return
        }

        await interaction.deferReply({ flags: 64 })

        const email = String(interaction.fields.getTextInputValue("email") || "").trim()

        if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
          await interaction.editReply("❌ Please enter a valid email address.")
          return
        }

        const result = await postToAppsScript({
          action: "grant_sheet_access_for_server",
          adminKey: process.env.ADMIN_API_KEY,
          discordServerId: interaction.guildId,
          email: email
        })

        if (!result.ok) {
          await interaction.editReply(`❌ ${result.error || "Could not grant access."}`)
          return
        }

        await interaction.editReply(
          `✅ ${result.email} now has edit access to the sheet for state ${result.state_code}.`
        )
        return
      }

      if (interaction.customId.startsWith("admin_remove_booking_modal:")) {
        if (!(await userCanManageServer(interaction))) {
          await interaction.reply({
            content: "❌ You do not have permission to use this command.",
            flags: 64
          })
          return
        }

        await interaction.deferReply({ flags: 64 })

        const day = interaction.customId.split(":")[1]
        const playerId = String(interaction.fields.getTextInputValue("player_id") || "").trim()

        if (!/^[0-9]+$/.test(playerId)) {
          await interaction.editReply("❌ Player ID must contain numbers only.")
          return
        }

        const result = await postToAppsScript({
          action: "admin_remove_booking_for_server",
          adminKey: process.env.ADMIN_API_KEY,
          discordServerId: interaction.guildId,
          playerId: playerId,
          day: day
        })

        if (!result.ok) {
          await interaction.editReply(`❌ ${result.error || "Could not remove booking."}`)
          return
        }

        if (!result.removed) {
          await interaction.editReply(`❌ No booking found for player ID ${playerId} on ${day}.`)
          return
        }

        if (day === "ALL") {
          await interaction.editReply(
            `✅ Removed ${result.removed_count} booking(s) for player ID ${playerId}.`
          )
          return
        }

        await interaction.editReply(
          `✅ Removed ${day} booking for player ID ${playerId}.`
        )
        return
      }

      if (interaction.customId.startsWith("admin_add_booking_modal:")) {
        if (!(await userCanManageServer(interaction))) {
          await interaction.reply({
            content: "❌ You do not have permission to use this command.",
            flags: 64
          })
          return
        }

        await interaction.deferReply({ flags: 64 })

        const [, token] = interaction.customId.split(":")
        const entry = getPendingAdminBookingOrThrow(token)

        if (interaction.user.id !== entry.requestedBy) {
          await interaction.editReply("❌ This admin booking modal belongs to someone else.")
          return
        }

        const allianceRaw = String(interaction.fields.getTextInputValue("alliance") || "").trim()
        const name = String(interaction.fields.getTextInputValue("name") || "").trim()
        const playerId = String(interaction.fields.getTextInputValue("player_id") || "").trim()

        const alliance = allianceRaw.toUpperCase()

        if (!/^[A-Z0-9]{3}$/.test(alliance)) {
          await interaction.editReply("❌ Alliance tag must be exactly 3 letters or numbers.")
          return
        }

        if (!name) {
          await interaction.editReply("❌ Player name is required.")
          return
        }

        if (!/^[0-9]+$/.test(playerId)) {
          await interaction.editReply("❌ Player ID must contain numbers only.")
          return
        }

        const result = await postToAppsScript({
          action: "admin_add_booking_for_server",
          adminKey: process.env.ADMIN_API_KEY,
          discordServerId: interaction.guildId,
          day: entry.day,
          time: entry.selectedTime,
          alliance: alliance,
          inGameName: name,
          playerId: playerId
        })

        pendingAdminBookings.delete(token)

        if (!result.ok) {
          await interaction.editReply(`❌ ${result.error || "Could not add booking."}`)
          return
        }

        await interaction.editReply(
          `✅ Booking added for state ${result.state_code}\n` +
          `Day: ${result.day}\n` +
          `Time: ${result.time}\n` +
          `Alliance: ${result.alliance}\n` +
          `Name: ${result.playerName}\n` +
          `Player ID: ${result.playerId}`
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
       `Date: ${result.booking_date_display || "No date set"}\n` +
       `Time: ${result.time} UTC`

      if (result.moved && result.oldTime) {
         dmMessage =
          `🔁 Booking changed\n\n` +
          `State: ${result.state_code}\n` +
          `Day: ${result.day}\n` +
          `Date: ${result.booking_date_display || "No date set"}\n` +
          `Old time: ${result.oldTime} UTC\n` +
          `New time: ${result.time} UTC`
        }

      await sendBookingDm(interaction.user, dmMessage)

      pendingBookings.delete(token)
      return
    }

    if (!interaction.isChatInputCommand()) return

  if (interaction.commandName === "set-banter-channel") {
  await interaction.deferReply({ flags: 64 })

  if (!(await userCanManageServer(interaction))) {
    await interaction.editReply("❌ You do not have permission to use this command.")
    return
  }

  const channel = interaction.options.getChannel("channel")

  if (!channel || typeof channel.send !== "function") {
    await interaction.editReply("❌ Please select a valid text channel.")
    return
  }

  const result = await postToAppsScript({
    action: "set_banter_channel_for_server",
    adminKey: process.env.ADMIN_API_KEY,
    discordServerId: interaction.guildId,
    channelId: channel.id
  })

  if (!result.ok) {
    await interaction.editReply(`❌ ${result.error || "Could not set banter channel."}`)
    return
  }

  await interaction.editReply(`✅ R.A.C.H.I.E banter channel set to #${channel.name}.`)
  return
}

if (interaction.commandName === "clear-banter-channel") {
  await interaction.deferReply({ flags: 64 })

  if (!(await userCanManageServer(interaction))) {
    await interaction.editReply("❌ You do not have permission to use this command.")
    return
  }

  const result = await postToAppsScript({
    action: "clear_banter_channel_for_server",
    adminKey: process.env.ADMIN_API_KEY,
    discordServerId: interaction.guildId
  })

  if (!result.ok) {
    await interaction.editReply(`❌ ${result.error || "Could not clear banter channel."}`)
    return
  }

  await interaction.editReply("✅ R.A.C.H.I.E banter channel cleared.")
  return
}

if (interaction.commandName === "set-banter-spice") {
  await interaction.deferReply({ flags: 64 })

  if (!(await userCanManageServer(interaction))) {
    await interaction.editReply("❌ You do not have permission to use this command.")
    return
  }

  const level = interaction.options.getString("level")

  const result = await postToAppsScript({
    action: "set_banter_spice_for_server",
    adminKey: process.env.ADMIN_API_KEY,
    discordServerId: interaction.guildId,
    spiceLevel: level
  })

  if (!result.ok) {
    await interaction.editReply(`❌ ${result.error || "Could not update spice level."}`)
    return
  }

  await interaction.editReply(`🔥 Banter spice level set to ${level}.`)
  return
}

  if (interaction.commandName === "set-bot-admin-role") {
  await interaction.deferReply({ flags: 64 })

  if (!interaction.memberPermissions?.has(PermissionFlagsBits.ManageGuild)) {
    await interaction.editReply("❌ Only a server admin can set the bot admin role.")
    return
  }

  const role = interaction.options.getRole("role")

  const result = await postToAppsScript({
    action: "set_bot_admin_role_for_server",
    adminKey: process.env.ADMIN_API_KEY,
    discordServerId: interaction.guildId,
    roleId: role.id
  })

  if (!result.ok) {
    await interaction.editReply(`❌ ${result.error || "Could not save bot admin role."}`)
    return
  }

  await interaction.editReply(`✅ Bot admin role set to ${role}.`)
  return
}

if (interaction.commandName === "clear-bot-admin-role") {
  await interaction.deferReply({ flags: 64 })

  if (!interaction.memberPermissions?.has(PermissionFlagsBits.ManageGuild)) {
    await interaction.editReply("❌ Only a server admin can clear the bot admin role.")
    return
  }

  const result = await postToAppsScript({
    action: "clear_bot_admin_role_for_server",
    adminKey: process.env.ADMIN_API_KEY,
    discordServerId: interaction.guildId
  })

  if (!result.ok) {
    await interaction.editReply(`❌ ${result.error || "Could not clear bot admin role."}`)
    return
  }

  await interaction.editReply("✅ Bot admin role cleared.")
  return
}

if (interaction.commandName === "set-booking-date") {
  await interaction.deferReply({ flags: 64 })

  if (!(await userCanManageServer(interaction))) {
    await interaction.editReply("❌ You do not have permission to use this command.")
    return
  }

  const day = interaction.options.getString("day")
  const year = interaction.options.getInteger("year")
  const month = interaction.options.getInteger("month")
  const dayOfMonth = interaction.options.getInteger("day_of_month")

  const isoDate =
    String(year).padStart(4, "0") + "-" +
    String(month).padStart(2, "0") + "-" +
    String(dayOfMonth).padStart(2, "0")

  const result = await postToAppsScript({
    action: "set_booking_date_for_server",
    adminKey: process.env.ADMIN_API_KEY,
    discordServerId: interaction.guildId,
    day: day,
    date: isoDate
  })

  if (!result.ok) {
    await interaction.editReply(`❌ ${result.error || "Could not update booking date."}`)
    return
  }

  await interaction.editReply(
    `✅ ${result.day} date updated.\n` +
    `Date: ${result.display_date}\n` +
    `Stored as: ${result.iso_date}\n` +
    `Time zone: UTC`
  )
  return
}

    if (interaction.commandName === "grant-access") {
      if (!(await userCanManageServer(interaction))) {
        await interaction.reply({
          content: "❌ You do not have permission to use this command.",
          flags: 64
        })
        return
      }

      await interaction.showModal(buildGrantAccessModal())
      return
    }

    if (interaction.commandName === "clear-bookings") {
      if (!(await userCanManageServer(interaction))) {
        await interaction.reply({
          content: "❌ You do not have permission to use this command.",
          flags: 64
        })
        return
      }

      await interaction.showModal(buildClearBookingsModal())
      return
    }

    if (interaction.commandName === "admin-remove-booking") {
      if (!(await userCanManageServer(interaction))) {
        await interaction.reply({
          content: "❌ You do not have permission to use this command.",
          flags: 64
        })
        return
      }

      const day = interaction.options.getString("day")
      await interaction.showModal(buildAdminRemoveBookingModal(day))
      return
    }

    if (interaction.commandName === "admin-reserve-slots") {
      await interaction.deferReply({ flags: 64 })

      if (!(await userCanManageServer(interaction))) {
        await interaction.editReply("❌ You do not have permission to use this command.")
        return
      }

      const day = interaction.options.getString("day")

      const result = await postToAppsScript({
        action: "get_times_for_server",
        adminKey: process.env.ADMIN_API_KEY,
        discordServerId: interaction.guildId,
        day: day
      })

      if (!result.ok) {
        await interaction.editReply(`❌ ${result.error || "Could not load available times."}`)
        return
      }

      if (!Array.isArray(result.times) || result.times.length === 0) {
        await interaction.editReply(`❌ No available times found for ${day}.`)
        return
      }

      const token = createBookingToken()

      pendingAdminReserves.set(token, {
        createdAt: Date.now(),
        discordServerId: interaction.guildId,
        requestedBy: interaction.user.id,
        day: day,
        times: result.times
      })

      const ui = buildAdminReserveComponents(token, 0)
      await interaction.editReply(ui)
      return
    }

    if (interaction.commandName === "admin-add-booking") {
      await interaction.deferReply({ flags: 64 })

      if (!(await userCanManageServer(interaction))) {
        await interaction.editReply("❌ You do not have permission to use this command.")
        return
      }

      const day = interaction.options.getString("day")

      const result = await postToAppsScript({
        action: "get_times_for_server",
        adminKey: process.env.ADMIN_API_KEY,
        discordServerId: interaction.guildId,
        day: day
      })

      if (!result.ok) {
        await interaction.editReply(`❌ ${result.error || "Could not load available times."}`)
        return
      }

      if (!Array.isArray(result.times) || result.times.length === 0) {
        await interaction.editReply(`❌ No available times found for ${day}.`)
        return
      }

      const token = createBookingToken()

      pendingAdminBookings.set(token, {
        createdAt: Date.now(),
        discordServerId: interaction.guildId,
        requestedBy: interaction.user.id,
        day: day,
        times: result.times
      })

      const ui = buildAdminBookingComponents(token, 0)
      await interaction.editReply(ui)
      return
    }

    if (interaction.commandName === "help") {
  await interaction.reply(
`R.A.C.H.I.E User Guide

/register
Set your alliance, name and player ID.

/book
Book a minister slot.

/remove-booking
Cancel one of your bookings.

/my-bookings
See your current bookings.

/my-info
See your saved player details.

/times
Check available times and the current UTC booking date for Construction, Research and Troop Training.

/booking-link
Get the booking page link for this state.

Tip

Register first, then book.
All booking times are in UTC.
If a slot is taken, run /book again and choose another time.`
  )
  return
}

    if (interaction.commandName === "reset-state-password") {
      await interaction.deferReply({ flags: 64 })

      if (!(await userCanManageServer(interaction))) {
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

    if (interaction.commandName === "banter-test") {
  await interaction.deferReply({ flags: 64 })

  if (!(await userCanManageServer(interaction))) {
    await interaction.editReply("❌ You do not have permission to use this command.")
    return
  }

  try {
    const fetched = await interaction.channel.messages.fetch({ limit: 10 })

    const messages = fetched
      .filter(message => !message.author.bot && message.content && message.content.trim())
      .map(message => ({
        author: message.member?.displayName || message.author.username,
        content: message.content.trim()
      }))
      .reverse()

    if (!messages.length) {
      await interaction.editReply("❌ No recent user messages found in this channel.")
      return
    }

    await interaction.editReply("⏳ Generating banter...")

    const banterConfig = await getBanterConfigForGuild(interaction.guildId)
    await triggerBanter(interaction.channel, messages, banterConfig.spiceLevel)

    await interaction.editReply("✅ Banter test sent.")
    return
  } catch (error) {
    console.error("banter-test failed:", error)
    await interaction.editReply("❌ Could not generate a banter reply.")
    return
  }
}

   if (interaction.commandName === "admin-help") {
  await interaction.deferReply({ flags: 64 })

  if (!(await userCanManageServer(interaction))) {
    await interaction.editReply("❌ You do not have permission to use this command.")
    return
  }

  await interaction.editReply(
`R.A.C.H.I.E Admin Guide

/setup
Create a new state sheet and link this server.

/setup-help
Show full setup instructions and settings help.

/link-state
Link this server to an existing state.

/unlink-state
Remove a linked Discord server from this state.

/linked-servers
Show all linked Discord servers for this state.

/set-announcements
Choose which channel receives booking announcements.

/settings
Manage booking requirements and limits.

/set-booking-date
Set the UTC booking date for Construction, Research or Troop.

/sheet-link
Get the state sheet link and booking page link.

/booking-link
Get the booking page link only.

/open-bookings
Open bookings and post announcements to linked servers.

/close-bookings
Close bookings and post announcements to linked servers.

/admin-add-booking
Add a booking for a player.

/admin-remove-booking
Remove a booking for a player.

/admin-reserve-slots
Reserve up to 5 slots at once.

/admin-remove-reserved
Remove reserved slots.

/clear-bookings
Clear all booking entries from the sheet.

/grant-access
Give a trusted user edit access to the sheet.

/reset-state-password
Generate a new join password for future server linking.

Tip

The state join password can be reset at any time using /reset-state-password.
This will not affect servers that are already linked.
Only new servers will need the updated password.`
  )
  return
}
    if (interaction.commandName === "setup-help") {
  await interaction.deferReply({ flags: 64 })

  if (!(await userCanManageServer(interaction))) {
    await interaction.editReply("❌ You do not have permission to use this command.")
    return
  }

  await interaction.editReply(
`R.A.C.H.I.E Setup Guide

1. Create or link a state
Use /setup for a new state.
Use /link-state if the state already exists.

2. Set announcements
Use /set-announcements to choose where booking updates are posted.

3. Optional sheet edit access
Use /grant-access to give a trusted admin edit access to the sheet.

4. Open or close bookings
Use /open-bookings to allow booking.
Use /close-bookings to stop new bookings.

5. Link more Discord servers
Use /link-state on other alliance servers.
Use /linked-servers to view them.
Use /unlink-state to remove one.

6. Admin booking tools
/admin-add-booking
/admin-remove-booking
/admin-reserve-slots
/admin-remove-reserved

7. Booking dates
Use /set-booking-date to set the UTC date for each minister day.

8. Settings
Use /settings to control required booking fields, booking limits and linked server limits.

9. Useful links
Use /sheet-link for the sheet and booking page.
Use /booking-link for the booking page only.

Recommended order
/setup or /link-state
/set-announcements
/settings
/set-booking-date
/open-bookings`
  )
  return
}

    if (interaction.commandName === "linked-servers") {
      await interaction.deferReply({ flags: 64 })

      if (!(await userCanManageServer(interaction))) {
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

    if (interaction.commandName === "set-announcements") {
      await interaction.deferReply({ flags: 64 })

      if (!(await userCanManageServer(interaction))) {
        await interaction.editReply("❌ You do not have permission to use this command.")
        return
      }

      const channel = interaction.options.getChannel("channel")

      if (!channel || typeof channel.send !== "function") {
        await interaction.editReply("❌ Please select a valid text channel.")
        return
      }

      const result = await postToAppsScript({
        action: "set_announcement_channel",
        adminKey: process.env.ADMIN_API_KEY,
        discordServerId: interaction.guildId,
        channelId: channel.id
      })

      if (!result.ok) {
        await interaction.editReply(`❌ ${result.error || "Could not save channel."}`)
        return
      }

      await interaction.editReply(`✅ Announcement channel set to #${channel.name}`)
      return
    }

    if (interaction.commandName === "sheet-link") {
      await interaction.deferReply({ flags: 64 })

      if (!(await userCanManageServer(interaction))) {
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

    if (interaction.commandName === "settings") {
      await interaction.deferReply({ flags: 64 })

      if (!(await userCanManageServer(interaction))) {
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

    if (interaction.commandName === "open-bookings") {
      await interaction.deferReply({ flags: 64 })

      if (!(await userCanManageServer(interaction))) {
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
        const sent = await sendAnnouncementToLinkedServers(interaction, announcement)

        await interaction.editReply(
          `✅ Bookings opened for state ${sent.state_code}.\n` +
          `Announcements sent to ${sent.sent_count} linked server channel(s).`
        )
      } catch (error) {
        console.error("Could not send open-bookings announcements:", error)
        await interaction.editReply(
          `✅ Bookings opened for state ${result.state_code}, but announcements could not be sent.`
        )
      }

      return
    }

    if (interaction.commandName === "close-bookings") {
  await interaction.deferReply({ flags: 64 })

  if (!(await userCanManageServer(interaction))) {
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
    const sent = await sendAnnouncementToLinkedServers(interaction, announcement)

    await interaction.editReply(
      `✅ Bookings closed for state ${sent.state_code}.\n` +
      `Announcements sent to ${sent.sent_count} linked server channel(s).`
    )
  } catch (error) {
    console.error("Could not send close-bookings announcements:", error)
    await interaction.editReply(
      `✅ Bookings closed for state ${result.state_code}, but announcements could not be sent.`
    )
  }

  return
}

if (interaction.commandName === "admin-remove-reserved") {
  await interaction.deferReply({ flags: 64 })

  if (!(await userCanManageServer(interaction))) {
    await interaction.editReply("❌ You do not have permission to use this command.")
    return
  }

  const day = interaction.options.getString("day")

  const result = await postToAppsScript({
    action: "get_reserved_times_for_server",
    adminKey: process.env.ADMIN_API_KEY,
    discordServerId: interaction.guildId,
    day: day
  })

  if (!result.ok) {
    await interaction.editReply(`❌ ${result.error || "Could not load reserved slots."}`)
    return
  }

  if (!Array.isArray(result.times) || result.times.length === 0) {
    await interaction.editReply(`❌ No reserved slots found for ${day}.`)
    return
  }

  const token = createBookingToken()

  pendingAdminRemoveReserved.set(token, {
    createdAt: Date.now(),
    discordServerId: interaction.guildId,
    requestedBy: interaction.user.id,
    day: day,
    times: result.times
  })

  const ui = buildAdminRemoveReservedComponents(token, 0)
  await interaction.editReply(ui)
  return
}

    if (interaction.commandName === "setup") {
      if (!(await userCanManageServer(interaction))) {
        await interaction.reply({
          content: "❌ You do not have permission to use this command.",
          flags: 64
        })
        return
      }

      await interaction.showModal(buildSetupModal())
      return
    }

    if (interaction.commandName === "link-state") {
      if (!(await userCanManageServer(interaction))) {
        await interaction.reply({
          content: "❌ You do not have permission to use this command.",
          flags: 64
        })
        return
      }

      await interaction.showModal(buildLinkStateModal())
      return
    }

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

    if (interaction.commandName === "times") {
  await interaction.deferReply({ flags: 64 })

  const day = interaction.options.getString("day")

  const [timesResult, dateResult] = await Promise.all([
    postToAppsScript({
      action: "get_times_for_server",
      adminKey: process.env.ADMIN_API_KEY,
      discordServerId: interaction.guildId,
      day: day
    }),
    postToAppsScript({
      action: "get_booking_date_for_server",
      adminKey: process.env.ADMIN_API_KEY,
      discordServerId: interaction.guildId,
      day: day
    })
  ])

  if (!timesResult.ok) {
    await interaction.editReply(`❌ ${timesResult.error}`)
    return
  }

  if (!dateResult.ok) {
    await interaction.editReply(`❌ ${dateResult.error || "Could not load booking date."}`)
    return
  }

  const timesText = timesResult.times.length
    ? timesResult.times.map(time => `${time} UTC`).join(", ")
    : "No times available"

  const dateText = dateResult.display_date || "No date set"

  await interaction.editReply(
    `${day} date: ${dateText}\n` +
    `Times (UTC):\n${timesText}`
  )
  return
}

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
       `Construction: ${construction} ${construction !== "Not booked" ? `on ${result.dates?.Construction || "No date set"}` : ""}\n` +
       `Research: ${research} ${research !== "Not booked" ? `on ${result.dates?.Research || "No date set"}` : ""}\n` +
       `Troop: ${troop} ${troop !== "Not booked" ? `on ${result.dates?.Troop || "No date set"}` : ""}`
      )
      return
    }

    if (interaction.commandName === "register") {
      await interaction.showModal(buildRegisterModal())
      return
    }

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

    if (interaction.commandName === "unlink-state") {
      await interaction.deferReply({ flags: 64 })

      if (!(await userCanManageServer(interaction))) {
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

      const token = createBookingToken()
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

client.on("messageCreate", async message => {
  try {
    if (message.author.bot) return
    if (!message.guild) return
    if (!message.content || message.content.trim().length < 4) return

    const banterConfig = await getBanterConfigForGuild(message.guildId)

    if (!banterConfig.banterChannelId) {
      return
    }

    if (message.channel.id !== banterConfig.banterChannelId) {
      return
    }

    const channelId = message.channel.id

    const lastTime = channelCooldowns.get(channelId) || 0
    if (Date.now() - lastTime < COOLDOWN_MS) {
      return
    }

    if (!messageBuffers.has(channelId)) {
      messageBuffers.set(channelId, [])
    }

    const buffer = messageBuffers.get(channelId)

    buffer.push({
      author: message.member?.displayName || message.author.username,
      content: message.content.trim(),
      sourceMessage: message
    })

    if (buffer.length > MESSAGE_LIMIT) {
      buffer.shift()
    }

    if (buffer.length < MESSAGE_LIMIT) return

    await triggerBanter(message.channel, [...buffer], banterConfig.spiceLevel)

    messageBuffers.set(channelId, [])
    channelCooldowns.set(channelId, Date.now())
  } catch (err) {
    console.error("Message handler error:", err)
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