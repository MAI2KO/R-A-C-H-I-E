const { sanitizedText } = require("../centuryGameClient")

function displayValue(value) {
  return value === null || value === undefined || value === "" ? "None" : String(value)
}

function centuryMessage(value, redactedValues = []) {
  if (value === null || value === undefined || value === "") return "Not recorded"
  const sanitized = sanitizedText(value, redactedValues)
    .replace(/@everyone|@here/gi, "[mention redacted]")
    .replace(/<@!?\d+>/g, "[mention redacted]")
    .replace(/\b\d{6,32}\b/g, "[identifier redacted]")
    .replace(/[`*_~|]/g, "")
    .slice(0, 300)
  return sanitized || "Not recorded"
}

function titleCase(value) {
  if (["html", "json", "http"].includes(String(value).toLowerCase())) {
    return String(value).toUpperCase()
  }
  return String(value || "unknown_response")
    .split("_")
    .map(word => word.charAt(0).toUpperCase() + word.slice(1))
    .join(" ")
}

function classificationMeaning(value) {
  return {
    redemption_limit: "Another gift code of this type was already redeemed on this character",
    claim_limit: "This character has reached the claim limit",
    level_restriction: "This character's City or Town Center level is too low",
    account_restriction: "This account does not meet the redemption requirements",
    account_age_restriction: "This account does not meet the required account age",
    verification_throttle: "Verification requests were made too frequently",
    verification_error: "The frontend verification code was incorrect or expired",
    simultaneous_action_throttle: "Too many simultaneous actions were in progress"
  }[value] || null
}

function formatCodeDiagnostics(code) {
  if (!code) return "No matching gift code was found."
  const metadata = code.verification_metadata || {}
  const response = metadata.response || {}
  const legacyResponse = metadata.rawResponse
  const legacyResponseType = legacyResponse && typeof legacyResponse === "object"
    ? "json"
    : /<html\b|<!doctype\s+html/i.test(String(legacyResponse || "")) ? "html" : null
  const classification = metadata.classification
    || ([401, 403].includes(Number(code.verification_http_status))
      ? "upstream_rejection"
      : "unknown_response")
  const hasVerificationAttempt = Boolean(code.latest_verification_attempt_id)
  const verificationHttpStatus = hasVerificationAttempt
    ? code.latest_verification_http_status
    : code.verification_http_status
  const verificationErrCode = hasVerificationAttempt
    ? code.latest_verification_err_code
    : code.last_err_code
  const verificationMessage = hasVerificationAttempt
    ? code.latest_verification_api_message
    : code.last_api_message
  const verificationMetadata = code.latest_verification_metadata || metadata
  const verificationResponse = verificationMetadata.response || response
  const verificationClassification = code.latest_verification_classification || classification
  const lines = [
    `**Gift code ${code.code}**`,
    `Current status: ${code.status}`,
    `Current verification state: ${code.verification_state}`
  ]
  if (code.first_seen_at_utc) {
    lines.push(`First seen: <t:${Math.floor(new Date(code.first_seen_at_utc).getTime() / 1000)}:f>`)
  }
  lines.push(
    `Source: ${code.source_name || code.source_type || "Discord submission"}`,
    "",
    "**Latest recorded verification attempt**",
    `HTTP status: ${displayValue(verificationHttpStatus)}`,
    `Century err_code: ${displayValue(verificationErrCode)}`,
    `Century message: ${centuryMessage(verificationMessage)}`,
    `Response type: ${titleCase(verificationResponse.responseType || legacyResponseType || "unknown")}`,
    `Classification: ${titleCase(verificationClassification)}`
  )
  const verificationMeaning = classificationMeaning(verificationClassification)
  if (verificationMeaning) lines.push(`Meaning: ${verificationMeaning}`)
  if (code.latest_redemption_attempt_id) {
    lines.push(
      "",
      "**Latest player redemption**",
      `HTTP status: ${displayValue(code.latest_redemption_http_status)}`,
      `Century err_code: ${displayValue(code.latest_redemption_err_code)}`,
      `Century message: ${centuryMessage(code.latest_redemption_api_message, [
        code.latest_redemption_player_id_snapshot
      ])}`,
      `Classification: ${titleCase(code.latest_redemption_classification)}`
    )
    const redemptionMeaning = classificationMeaning(code.latest_redemption_classification)
    if (redemptionMeaning) lines.push(`Meaning: ${redemptionMeaning}`)
  }
  lines.push(
    `Queue counts: pending ${code.pending_count}, success ${code.success_count}, ` +
      `already redeemed ${code.already_redeemed_count}, review/failed ${code.failed_count}`
  )
  return lines.join("\n")
}

module.exports = {
  displayValue,
  titleCase,
  classificationMeaning,
  centuryMessage,
  formatCodeDiagnostics
}
