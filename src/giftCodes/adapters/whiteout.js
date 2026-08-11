const DEFAULT_SIGNING_SUFFIX = "tB87#kPtkxqOS2"

function whiteoutAdapter(env = process.env) {
  return Object.freeze({
    gameProfile: "wos",
    gameName: "Whiteout Survival",
    frontendUrl: "https://wos-giftcode.centurygame.com",
    apiBaseUrl: "https://wos-giftcode-api.centurygame.com/api",
    redemptionPath: "/gift_code",
    signingSuffix: String(env.CENTURY_WOS_SIGNING_SUFFIX || DEFAULT_SIGNING_SUFFIX)
  })
}

module.exports = { DEFAULT_SIGNING_SUFFIX, whiteoutAdapter }
