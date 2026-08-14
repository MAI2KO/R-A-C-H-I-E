const { createAppsScriptClient } = require("./transport")

const BOOKING_APPS_SCRIPT_URL_ENV = "BOOKING_APPS_SCRIPT_URL"

function createBookingAppsScriptClient(options = {}) {
  return createAppsScriptClient({
    ...options,
    urlEnvName: BOOKING_APPS_SCRIPT_URL_ENV
  })
}

module.exports = {
  BOOKING_APPS_SCRIPT_URL_ENV,
  createBookingAppsScriptClient
}
