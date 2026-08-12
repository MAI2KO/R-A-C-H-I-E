const MONTHS = Object.freeze({
  january: 0, february: 1, march: 2, april: 3, may: 4, june: 5,
  july: 6, august: 7, september: 8, october: 9, november: 10, december: 11
})

function cleanCode(value) {
  const code = String(value || "").trim().replace(/^[`'"([{<]+|[`'"\])}>.,;:!?]+$/g, "")
  return /^[A-Za-z0-9]{3,64}$/.test(code) ? code : null
}

function parseSourceExpiry(text, observedAt = new Date()) {
  const match = String(text || "").match(
    /(?:valid\s+until|expires?(?:\s+at)?)\s*:\s*([A-Za-z]+)\s+(\d{1,2})(?:,?\s+(\d{4}))?\s*,?\s*(\d{1,2}):(\d{2})\s*\((?:UTC|GMT)(?:\+?0(?::?00)?)?\)/i
  )
  if (!match) return null
  const month = MONTHS[match[1].toLowerCase()]
  if (month === undefined) return null
  const reference = new Date(observedAt)
  let year = match[3] ? Number(match[3]) : reference.getUTCFullYear()
  const day = Number(match[2])
  const hour = Number(match[4])
  const minute = Number(match[5])
  function dateFor(candidateYear) {
    const value = new Date(Date.UTC(candidateYear, month, day, hour, minute))
    return value.getUTCFullYear() === candidateYear && value.getUTCMonth() === month &&
      value.getUTCDate() === day && value.getUTCHours() === hour &&
      value.getUTCMinutes() === minute
      ? value
      : null
  }
  let result = dateFor(year)
  if (!result) return null
  if (!match[3] && result.getTime() < reference.getTime() - 31 * 86400000) {
    year += 1
    result = dateFor(year)
  }
  return result
}

function parseDiscordMirrorMessage(content, observedAt = new Date()) {
  const match = String(content || "").match(
    /(?:^|\n)\s*(?:[^A-Za-z0-9\n]{0,4}\s*)?(?:gift\s+code|redeem\s+code|code)\s*:\s*([^\s\n]+)/im
  )
  const code = cleanCode(match?.[1])
  if (!code) return null
  return { code, sourceReportedExpiryAt: parseSourceExpiry(content, observedAt) }
}

function decodeHtml(value) {
  return String(value || "")
    .replace(/&amp;/g, "&").replace(/&quot;/g, '"').replace(/&#39;/g, "'")
    .replace(/&lt;/g, "<").replace(/&gt;/g, ">")
}

function parseActiveCatalogueHtml(html) {
  const body = String(html || "")
  const codes = new Set()
  const patterns = [
    /<[^>]+data-(?:status|state)=["']active["'][^>]*data-code=["']([^"']+)["'][^>]*>/gi,
    /<[^>]+data-code=["']([^"']+)["'][^>]*data-(?:status|state)=["']active["'][^>]*>/gi,
    /<[^>]+class=["'][^"']*active-code[^"']*["'][^>]*>([\s\S]*?)<\/[^>]+>/gi
  ]
  for (const pattern of patterns) {
    for (const match of body.matchAll(pattern)) {
      const withoutTags = decodeHtml(match[1]).replace(/<[^>]*>/g, " ").trim()
      const code = cleanCode(withoutTags)
      if (code) codes.add(code)
    }
  }
  return [...codes]
}

module.exports = {
  cleanCode,
  parseSourceExpiry,
  parseDiscordMirrorMessage,
  parseActiveCatalogueHtml
}
