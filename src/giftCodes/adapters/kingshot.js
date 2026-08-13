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
      semanticResponses: Object.freeze([
        Object.freeze({
          httpStatus: 200,
          errCode: 40004,
          message: "TIMEOUT RETRY",
          state: "server_busy_timeout"
        })
      ]),
      errCodes: Object.freeze({
        20000: "success",
        40008: "already_redeemed",
        40011: "redemption_limit",
        40007: "expired",
        40005: "claim_limit",
        40006: "level_restriction",
        40012: "account_age_restriction",
        40014: "invalid_code",
        40017: "account_restriction",
        40019: "simultaneous_action_throttle",
        40020: "invalid_player",
        40100: "verification_throttle",
        40102: "verification_error",
        40103: "verification_error"
      })
    })
  })
}

module.exports = { DEFAULT_SIGNING_SUFFIX, kingshotAdapter }
