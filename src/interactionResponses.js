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

async function acknowledgeSchedulerInteraction(interaction) {
  if (interaction.deferred || interaction.replied) return false
  if (interaction.isChatInputCommand?.() || interaction.isModalSubmit?.()) {
    await interaction.deferReply({ ephemeral: true })
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

  logger.error("Interaction handler error:", error)
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
  acknowledgeSchedulerInteraction,
  editOrReply,
  safelyRespondToInteraction,
  handleInteractionFailure,
  schedulerInteractionWasHandled
}
