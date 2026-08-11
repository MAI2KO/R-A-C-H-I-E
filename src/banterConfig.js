const DEFAULT_BANTER_CONFIG = Object.freeze({
  banterChannelId: "",
  spiceLevel: "standard"
})

function safeErrorDetails(error) {
  return {
    error_name: String(error?.name || "Error").slice(0, 100),
    error_code: String(error?.code || error?.response?.status || "unknown").slice(0, 100),
    error_message: String(error?.message || "Banter configuration lookup failed")
      .replace(/https?:\/\/\S+/gi, "[url]")
      .replace(/(admin_api_key|authorization|token)\s*[=:]\s*\S+/gi, "$1=[redacted]")
      .slice(0, 300)
  }
}

function createBanterConfigLookup({
  fetchAction,
  ttlMs = 5 * 60 * 1000,
  now = () => Date.now(),
  logger = console
}) {
  if (typeof fetchAction !== "function") throw new Error("fetchAction is required")
  const cache = new Map()
  const pending = new Map()

  async function load(guildId) {
    try {
      const [channelResult, spiceResult] = await Promise.all([
        fetchAction("get_banter_channel_for_server", guildId),
        fetchAction("get_banter_spice_for_server", guildId)
      ])
      return {
        banterChannelId: String(channelResult?.banter_channel_id || "").trim(),
        spiceLevel: String(spiceResult?.banter_spice_level || "standard").trim().toLowerCase()
      }
    } catch (error) {
      logger.warn(JSON.stringify({
        event: "banter_config_lookup_failed",
        guild_id: /^[0-9]{1,32}$/.test(String(guildId)) ? String(guildId) : undefined,
        ...safeErrorDetails(error)
      }))
      return { ...DEFAULT_BANTER_CONFIG }
    }
  }

  return Object.freeze({
    async get(guildId) {
      const key = String(guildId || "")
      const cached = cache.get(key)
      if (cached && now() - cached.timestamp < ttlMs) return cached.data
      if (pending.has(key)) return pending.get(key)

      const request = load(key).then(data => {
        cache.set(key, { data, timestamp: now() })
        pending.delete(key)
        return data
      }, error => {
        pending.delete(key)
        throw error
      })
      pending.set(key, request)
      return request
    },

    invalidate(guildId) {
      cache.delete(String(guildId || ""))
    }
  })
}

module.exports = { DEFAULT_BANTER_CONFIG, safeErrorDetails, createBanterConfigLookup }
