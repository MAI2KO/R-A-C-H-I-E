const DEFAULT_SIGNING_SUFFIX = "mN4!pQs6JrYwV9"

function kingshotAdapter(env = process.env) {
  return Object.freeze({
    gameProfile: "kingshot",
    gameName: "Kingshot",
    frontendUrl: "https://ks-giftcode.centurygame.com",
    apiBaseUrl: "https://kingshot-giftcode.centurygame.com/api",
    redemptionPath: "/gift_code",
    signingSuffix: String(env.CENTURY_KINGSHOT_SIGNING_SUFFIX || DEFAULT_SIGNING_SUFFIX),
    responseMappings: Object.freeze({
      errCodes: Object.freeze({
        20000: "success",
        40008: "already_redeemed",
        40007: "expired",
        40014: "invalid_code"
      })
    })
  })
}

module.exports = { DEFAULT_SIGNING_SUFFIX, kingshotAdapter }
