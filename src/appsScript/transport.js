const axios = require("axios")

const JSON_HEADERS = Object.freeze({
  "Content-Type": "application/json"
})

function resolveAppsScriptUrl(env, overrideName) {
  return env[overrideName] || env.APPS_SCRIPT_URL
}

function createAppsScriptTransport({ httpClient = axios } = {}) {
  return Object.freeze({
    async post(url, payload) {
      const response = await httpClient.post(url, payload, {
        headers: JSON_HEADERS
      })
      return response.data
    }
  })
}

function createAppsScriptClient({
  urlEnvName,
  env = process.env,
  transport = createAppsScriptTransport()
}) {
  return Object.freeze({
    async post(payload) {
      return transport.post(resolveAppsScriptUrl(env, urlEnvName), payload)
    }
  })
}

module.exports = {
  JSON_HEADERS,
  createAppsScriptClient,
  createAppsScriptTransport,
  resolveAppsScriptUrl
}
