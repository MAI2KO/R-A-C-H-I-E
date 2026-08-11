function displayValue(value) {
  return value === null || value === undefined || value === "" ? "None" : String(value)
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
    `HTTP status: ${displayValue(code.verification_http_status)}`,
    `Century err_code: ${displayValue(code.last_err_code)}`,
    `Response type: ${titleCase(response.responseType || legacyResponseType || "unknown")}`,
    `Classification: ${titleCase(classification)}`,
    `Queue counts: pending ${code.pending_count}, success ${code.success_count}, ` +
      `already redeemed ${code.already_redeemed_count}, review/failed ${code.failed_count}`
  )
  return lines.join("\n")
}

module.exports = { displayValue, titleCase, formatCodeDiagnostics }
