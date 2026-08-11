const DEFAULT_SIGNING_SUFFIX = "mN4!pQs6JrYwV9"

function kingshotAdapter(env = process.env) {
  return Object.freeze({
    gameProfile: "kingshot",
    gameName: "Kingshot",
    frontendUrl: "https://ks-giftcode.centurygame.com",
    apiBaseUrl: "https://kingshot-giftcode.centurygame.com/api",
    redemptionPath: "/gift_code",
    signingSuffix: String(env.CENTURY_KINGSHOT_SIGNING_SUFFIX || DEFAULT_SIGNING_SUFFIX)
  })
}

module.exports = { DEFAULT_SIGNING_SUFFIX, kingshotAdapter }
