const { createAppsScriptClient } = require("./transport")

const BANTER_APPS_SCRIPT_URL_ENV = "BANTER_APPS_SCRIPT_URL"

function createBanterAppsScriptClient(options = {}) {
  return createAppsScriptClient({
    ...options,
    urlEnvName: BANTER_APPS_SCRIPT_URL_ENV
  })
}

module.exports = {
  BANTER_APPS_SCRIPT_URL_ENV,
  createBanterAppsScriptClient
}
