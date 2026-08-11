const crypto = require("node:crypto")

function encodeSigningValue(value) {
  return encodeURIComponent(String(value))
    .replace(/[!'()*]/g, character =>
      `%${character.charCodeAt(0).toString(16).toUpperCase()}`)
}

function signingMaterial(fields) {
  return Object.keys(fields)
    .sort()
    .map(key => `${encodeSigningValue(key)}=${encodeSigningValue(fields[key])}`)
    .join("&")
}

function signRequestFields(fields, signingSuffix) {
  const suffix = String(signingSuffix || "")
  if (!suffix) throw new Error("Century signing suffix is not configured")
  return crypto.createHash("md5").update(`${signingMaterial(fields)}${suffix}`, "utf8").digest("hex")
}

module.exports = { encodeSigningValue, signingMaterial, signRequestFields }
