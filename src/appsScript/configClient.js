const { createAppsScriptClient } = require("./transport")

const CONFIG_APPS_SCRIPT_URL_ENV = "CONFIG_APPS_SCRIPT_URL"

function createConfigAppsScriptClient(options = {}) {
  return createAppsScriptClient({
    ...options,
    urlEnvName: CONFIG_APPS_SCRIPT_URL_ENV
  })
}

module.exports = {
  CONFIG_APPS_SCRIPT_URL_ENV,
  createConfigAppsScriptClient
}
