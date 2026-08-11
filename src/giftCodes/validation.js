class PlayerValidationError extends Error {
  constructor(message, code = "PLAYER_VALIDATION_ERROR") {
    super(message)
    this.name = "PlayerValidationError"
    this.code = code
  }
}

function normalizeNumericIdentifier(value, label, maximumLength) {
  const normalized = String(value || "").trim()
  if (!new RegExp(`^\\d{1,${maximumLength}}$`).test(normalized)) {
    throw new PlayerValidationError(
      `${label} must contain 1 to ${maximumLength} digits.`,
      "INVALID_NUMERIC_IDENTIFIER"
    )
  }
  return normalized
}

function normalizeDiscordUserId(value) {
  return normalizeNumericIdentifier(value, "Discord user ID", 32)
}

function normalizeGuildId(value) {
  return normalizeNumericIdentifier(value, "Discord server ID", 32)
}

function normalizePlayerId(value, playerLabel = "Player ID") {
  return normalizeNumericIdentifier(value, playerLabel, 32)
}

function normalizeLocationNumber(value, locationLabel) {
  return normalizeNumericIdentifier(value, `${locationLabel} number`, 10)
}

function normalizeGiftCode(value) {
  const code = String(value || "").trim()
  if (!code || code.length > 128 || /[\u0000-\u001f\u007f]/.test(code)) {
    throw new PlayerValidationError(
      "Gift code must be 1 to 128 printable characters.",
      "INVALID_GIFT_CODE"
    )
  }
  return code
}

module.exports = {
  PlayerValidationError,
  normalizeNumericIdentifier,
  normalizeDiscordUserId,
  normalizeGuildId,
  normalizePlayerId,
  normalizeLocationNumber,
  normalizeGiftCode
}
