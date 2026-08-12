const KNOWN_ERROR_STATES = Object.freeze({ 20000: "success" })

function integerOrNull(value) {
  const number = Number(value)
  return Number.isInteger(number) ? number : null
}

function classifyCenturyResponse({ httpStatus, data, profileMappings = {} }) {
  const raw = Object.freeze({
    code: integerOrNull(data?.code),
    errCode: integerOrNull(data?.err_code),
    message: String(data?.msg || "").slice(0, 500)
  })
  let state
  if (Number(httpStatus) === 429) state = "rate_limited"
  else if (Number(httpStatus) >= 500 || Number(httpStatus) === 408) state = "temporary_error"
  else {
    const errCodes = profileMappings.errCodes || profileMappings
    const codes = profileMappings.codes || {}
    const messages = profileMappings.messages || {}
    state = errCodes[raw.errCode]
      || KNOWN_ERROR_STATES[raw.errCode]
      || codes[raw.code]
      || messages[raw.message]
      || ([401, 403].includes(Number(httpStatus)) ? "upstream_rejection" : "unknown_response")
  }

  return Object.freeze({
    state,
    retryable: [
      "rate_limited", "temporary_error", "verification_throttle",
      "simultaneous_action_throttle"
    ].includes(state),
    permanent: [
      "success", "already_redeemed", "expired", "invalid_code", "invalid_player",
      "eligibility_restriction", "redemption_limit", "claim_limit", "level_restriction",
      "account_restriction", "account_age_restriction", "verification_error"
    ].includes(state),
    raw
  })
}

module.exports = { KNOWN_ERROR_STATES, classifyCenturyResponse }
