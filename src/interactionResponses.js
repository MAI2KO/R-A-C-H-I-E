const { MessageFlags } = require("discord.js")

const EXPIRED_INTERACTION = 10062
const ALREADY_ACKNOWLEDGED = 40060

function interactionErrorCode(error) {
  return Number(error?.code ?? error?.rawError?.code ?? error?.data?.code) || null
}

function isExpectedInteractionResponseError(error) {
  return [EXPIRED_INTERACTION, ALREADY_ACKNOWLEDGED].includes(interactionErrorCode(error))
}

function interactionLabel(interaction) {
  return interaction?.commandName || interaction?.customId || "unknown interaction"
}

function logExpectedInteractionResponseError(error, interaction, logger = console) {
  const code = interactionErrorCode(error)
  const label = interactionLabel(interaction)
  if (code === EXPIRED_INTERACTION) {
    logger.warn(`[Discord] Interaction expired before acknowledgement: ${label}`)
  } else if (code === ALREADY_ACKNOWLEDGED) {
    logger.warn(`[Discord] Interaction was already acknowledged: ${label}`)
  }
}

function interactionType(interaction) {
  if (interaction?.isChatInputCommand?.()) return "slash command"
  if (interaction?.isModalSubmit?.()) return "modal submission"
  if (interaction?.isChannelSelectMenu?.()) return "channel selector"
  if (interaction?.isStringSelectMenu?.()) return "string selector"
  if (interaction?.isButton?.()) return "button"
  return "interaction"
}

function sanitizeValidationMessage(message) {
  return String(message || "Invalid value")
    .replace(/https?:\/\/\S+/gi, "[redacted URL]")
    .replace(/\b(token|authorization|password|secret)\b\s*[:=]?\s*\S+/gi, "$1 [redacted]")
    .replace(/[\r\n]+/g, " ")
    .slice(0, 180)
}

function sanitizedValidationIssues(errors, maximumIssues = 6) {
  const issues = []
  function visit(value, path = []) {
    if (!value || issues.length >= maximumIssues) return
    if (Array.isArray(value)) {
      for (let index = 0; index < value.length && issues.length < maximumIssues; index += 1) {
        visit(value[index], [...path, String(index)])
      }
      return
    }
    if (typeof value !== "object") return
    if (Array.isArray(value._errors)) {
      for (const issue of value._errors.slice(0, maximumIssues - issues.length)) {
        const field = path.join(".").replace(/[^A-Za-z0-9_.-]/g, "?") || "body"
        const code = String(issue?.code || "INVALID_VALUE")
          .replace(/[^A-Za-z0-9_-]/g, "?").slice(0, 80)
        const message = sanitizeValidationMessage(issue?.message)
        issues.push(`${field}: ${code} (${message})`)
      }
    }
    for (const [key, child] of Object.entries(value)) {
      if (key !== "_errors") visit(child, [...path, key])
    }
  }
  visit(errors)
  return issues
}

function logDiscordApiError(error, interaction, logger = console) {
  const code = interactionErrorCode(error)
  if (!code) return false

  const issues = sanitizedValidationIssues(error?.rawError?.errors || error?.data?.errors)
  const duplicateId = issues.some(issue => issue.includes("COMPONENT_CUSTOM_ID_DUPLICATED"))
  const message = code === 50035 && duplicateId
    ? `duplicate custom_id in rendered components; ${issues.join("; ")}`
    : code === 50035 && issues.length
      ? `invalid form body; ${issues.join("; ")}`
      : "Discord rejected the interaction response"
  logger.error(
    `[Event scheduler] Discord API error ${code} during ${interactionType(interaction)} ` +
    `${interactionLabel(interaction)}: ${message}`
  )
  return true
}

async function acknowledgeSchedulerInteraction(interaction) {
  if (interaction.deferred || interaction.replied) return false
  if (interaction.isChatInputCommand?.() || interaction.isModalSubmit?.()) {
    await interaction.deferReply({ flags: MessageFlags.Ephemeral })
  } else {
    await interaction.deferUpdate()
  }
  return true
}

async function editOrReply(interaction, payload) {
  if (interaction.deferred) {
    const { flags, ephemeral, ...editablePayload } = payload
    return interaction.editReply(editablePayload)
  }
  if (interaction.replied) return interaction.followUp({
    ...payload,
    flags: payload.flags ?? MessageFlags.Ephemeral
  })
  return interaction.reply(payload)
}

async function safelyRespondToInteraction(
  interaction,
  payload,
  { logger = console } = {}
) {
  try {
    return await editOrReply(interaction, payload)
  } catch (error) {
    if (isExpectedInteractionResponseError(error)) {
      logExpectedInteractionResponseError(error, interaction, logger)
      return null
    }
    throw error
  }
}

async function handleInteractionFailure(
  interaction,
  error,
  { logger = console, message = "Something went wrong" } = {}
) {
  if (isExpectedInteractionResponseError(error)) {
    logExpectedInteractionResponseError(error, interaction, logger)
    return
  }

  if (!logDiscordApiError(error, interaction, logger)) {
    logger.error("Interaction handler error:", error)
  }
  try {
    await safelyRespondToInteraction(interaction, {
      content: message,
      flags: MessageFlags.Ephemeral,
      components: []
    }, { logger })
  } catch (responseError) {
    logger.error("Interaction error response failed:", responseError)
  }
}

async function schedulerInteractionWasHandled(interaction, handler, options) {
  return Boolean(await handler(interaction, options))
}

module.exports = {
  EXPIRED_INTERACTION,
  ALREADY_ACKNOWLEDGED,
  interactionErrorCode,
  isExpectedInteractionResponseError,
  interactionLabel,
  logExpectedInteractionResponseError,
  interactionType,
  sanitizedValidationIssues,
  logDiscordApiError,
  acknowledgeSchedulerInteraction,
  editOrReply,
  safelyRespondToInteraction,
  handleInteractionFailure,
  schedulerInteractionWasHandled
}
