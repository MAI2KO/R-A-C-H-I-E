const DEFAULT_SIGNING_SUFFIX = "tB87#kPtkxqOS2"

function whiteoutAdapter(env = process.env) {
  return Object.freeze({
    gameProfile: "wos",
    gameName: "Whiteout Survival",
    frontendUrl: "https://wos-giftcode.centurygame.com",
    apiBaseUrl: "https://wos-giftcode-api.centurygame.com/api",
    redemptionPath: "/gift_code",
    signingSuffix: String(env.CENTURY_WOS_SIGNING_SUFFIX || DEFAULT_SIGNING_SUFFIX),
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

module.exports = { DEFAULT_SIGNING_SUFFIX, whiteoutAdapter }
