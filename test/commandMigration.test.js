const assert = require("node:assert/strict")
const fs = require("node:fs")
const path = require("node:path")
const test = require("node:test")

const { buildBotSetupCommand } = require("../src/botSetupInteractions")
const { getPlayerCommandData } = require("../src/giftCodes/discord/commands")
const {
  RETIRED_LEGACY_BOOKING_COMMANDS,
  RETIRED_LEGACY_BOOKING_COMPONENTS,
  handleRetiredLegacyBookingInteraction
} = require("../src/legacyBookingRetirement")

const source = fs.readFileSync(path.join(__dirname, "..", "index.js"), "utf8")

test("canonical setup and registration commands are profile-aware native entry points", () => {
  assert.equal(buildBotSetupCommand().toJSON().name, "setup")
  for (const gameProfile of ["wos", "kingshot"]) {
    const command = getPlayerCommandData({ PLAYER_GIFT_CODES_ENABLED: "true", GAME_PROFILE: gameProfile })
    assert.equal(command.name, "register")
    assert.match(command.description, gameProfile === "wos" ? /Whiteout Survival/ : /Kingshot/)
  }
})

test("every legacy Sheet booking write surface is retired and stale controls fail closed", async () => {
  for (const command of [
    "admin-add-booking", "admin-remove-booking", "admin-remove-reserved",
    "admin-reserve-slots", "book", "clear-bookings", "close-bookings", "grant-access",
    "link-state", "open-bookings", "remove-booking", "reset-state-password",
    "set-announcements", "set-booking-date", "settings", "setup", "unlink-state", "unregister",
  ]) {
    assert.equal(RETIRED_LEGACY_BOOKING_COMMANDS.has(command), true)
  }
  for (const component of ["book_modal:", "register_modal", "settings_", "setup_modal"]) {
    assert.equal(RETIRED_LEGACY_BOOKING_COMPONENTS.includes(component), true)
  }
  assert.ok(source.indexOf("handleRetiredLegacyBookingInteraction(interaction)")
    < source.indexOf("if (interaction.isStringSelectMenu())"))

  const staleCommand = {
    commandName: "book", isChatInputCommand: () => true,
    async reply(payload) { this.payload = payload }
  }
  assert.equal(await handleRetiredLegacyBookingInteraction(staleCommand), true)
  assert.match(staleCommand.payload.content, /retired/)

  for (const customId of ["register_modal", "book_select:old-token:0", "settings_toggle:fc"]) {
    const staleControl = {
      customId, isChatInputCommand: () => false,
      async reply(payload) { this.payload = payload }
    }
    assert.equal(await handleRetiredLegacyBookingInteraction(staleControl), true)
    assert.match(staleControl.payload.content, /retired/)
  }
})

test("sheet-link stays registered as a strictly read-only legacy navigation command", () => {
  assert.equal(RETIRED_LEGACY_BOOKING_COMMANDS.has("sheet-link"), false)
  assert.match(source, /\.setName\("sheet-link"\)\s*\.setDescription\("Open the emergency read-only legacy booking sheet"\)/)

  const start = source.indexOf('if (interaction.commandName === "sheet-link")')
  const end = source.indexOf('if (interaction.commandName === "settings")', start)
  assert.ok(start > -1 && end > start)
  const handler = source.slice(start, end)
  assert.match(handler, /stateAppsScriptClient\.post\(\{\s*action: "get_sheet_link_for_server"/)
  assert.match(handler, /discordServerId: interaction\.guildId/)
  assert.match(handler, /Emergency legacy booking fallback — read-only/)
  assert.match(handler, /scheduled for removal after one successful native live booking cycle/)
  assert.match(handler, /result\.sheet_url/)
  assert.doesNotMatch(handler, /result\.booking_url/)
  assert.doesNotMatch(handler,
    /setup_state|open_bookings|close_bookings|book_for_server|remove_booking|update_setting|grant_sheet_access|set_announcement/)
})
