const crypto = require("node:crypto")

const CODE_ALPHABET = "23456789ABCDEFGHJKLMNPQRSTUVWXYZ"
const CODE_LENGTH = 12
const CODE_TTL_MINUTES = 15

function generateStateLinkCode() {
  const compact = Array.from(
    { length: CODE_LENGTH },
    () => CODE_ALPHABET[crypto.randomInt(CODE_ALPHABET.length)]
  ).join("")
  return compact.match(/.{1,4}/g).join("-")
}

function normalizeStateLinkCode(value) {
  const compact = String(value || "").trim().toUpperCase().replace(/[\s-]+/g, "")
  if (compact.length !== CODE_LENGTH || [...compact].some(char => !CODE_ALPHABET.includes(char))) {
    return null
  }
  return compact
}

function hashStateLinkCode(value) {
  const normalized = normalizeStateLinkCode(value)
  if (!normalized) return null
  return crypto.createHash("sha256").update(normalized).digest("hex")
}

module.exports = {
  CODE_TTL_MINUTES,
  generateStateLinkCode,
  normalizeStateLinkCode,
  hashStateLinkCode
}
