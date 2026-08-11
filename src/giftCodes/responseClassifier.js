const KNOWN_ERROR_STATES = Object.freeze({
  40008: "already_redeemed",
  40007: "expired",
  40014: "invalid_code",
  40020: "invalid_player"
})

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
  else if (raw.code === 0 && raw.errCode === 20000) state = "success"
  else state = profileMappings[raw.errCode] || KNOWN_ERROR_STATES[raw.errCode] || "unknown_response"

  return Object.freeze({
    state,
    retryable: ["rate_limited", "temporary_error"].includes(state),
    permanent: [
      "already_redeemed", "expired", "invalid_code", "invalid_player",
      "eligibility_restriction", "redemption_limit"
    ].includes(state),
    raw
  })
}

module.exports = { KNOWN_ERROR_STATES, classifyCenturyResponse }
