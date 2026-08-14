const { createAppsScriptClient } = require("./transport")

const STATE_APPS_SCRIPT_URL_ENV = "STATE_APPS_SCRIPT_URL"

function createStateAppsScriptClient(options = {}) {
  return createAppsScriptClient({
    ...options,
    urlEnvName: STATE_APPS_SCRIPT_URL_ENV
  })
}

module.exports = {
  STATE_APPS_SCRIPT_URL_ENV,
  createStateAppsScriptClient
}
