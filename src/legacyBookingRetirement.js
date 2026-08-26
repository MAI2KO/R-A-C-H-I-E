const { MessageFlags } = require("discord.js")

const RETIRED_LEGACY_BOOKING_COMMANDS = new Set([
  "admin-add-booking", "admin-help", "admin-remove-booking", "admin-remove-reserved",
  "admin-reserve-slots", "book", "booking-link", "clear-bookings", "close-bookings",
  "grant-access", "help", "link-state", "linked-servers", "my-bookings", "my-info",
  "open-bookings", "register", "remove-booking", "reset-state-password", "set-announcements",
  "set-booking-date", "settings", "setup", "setup-help", "times",
  "unlink-state", "unregister"
])

const RETIRED_LEGACY_BOOKING_COMPONENTS = [
  "admin_add_booking_modal:", "admin_book_page:", "admin_book_select:",
  "admin_remove_booking_modal:", "admin_remove_reserved_page:",
  "admin_remove_reserved_select:", "admin_reserve_page:", "admin_reserve_select:",
  "book_modal:", "book_page:", "book_select:", "clear_bookings_modal",
  "grant_access_modal", "link_state_modal", "register_modal", "settings_",
  "setup_modal", "unlink_select:"
]

function isRetiredLegacyBookingComponent(interaction) {
  const customId = String(interaction.customId || "")
  return RETIRED_LEGACY_BOOKING_COMPONENTS.some(prefix => customId.startsWith(prefix))
}

async function handleRetiredLegacyBookingInteraction(interaction) {
  if (interaction.isChatInputCommand?.()
      && RETIRED_LEGACY_BOOKING_COMMANDS.has(interaction.commandName)) {
    await interaction.reply({
      content: "This legacy Sheet-backed command has been retired. Use the authenticated community website for minister bookings, `/setup` for bot-managed channels, or `/register` for your game character.",
      flags: MessageFlags.Ephemeral
    })
    return true
  }
  if (isRetiredLegacyBookingComponent(interaction)) {
    await interaction.reply({
      content: "This legacy Sheet-backed control has been retired. Continue on the authenticated community website.",
      flags: MessageFlags.Ephemeral
    })
    return true
  }
  return false
}

module.exports = {
  RETIRED_LEGACY_BOOKING_COMMANDS,
  RETIRED_LEGACY_BOOKING_COMPONENTS,
  handleRetiredLegacyBookingInteraction,
  isRetiredLegacyBookingComponent
}
