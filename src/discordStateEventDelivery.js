const {
  normalizeDiscordDeliveryError,
  prepareStoredEventImage,
  resolveDeliveryTarget
} = require("./discordEventDelivery")
const { PermanentDeliveryError, RetryableDeliveryError } = require("./eventDeliveryWorker")
const { formatStateEventDelivery } = require("./stateEventDeliveryFormatting")

function permanent(message) {
  return new PermanentDeliveryError(message)
}

function retryable(message) {
  return new RetryableDeliveryError(message)
}

function createDiscordStateEventDeliveryHandler({ client, gameProfile }) {
  if (!client) throw new Error("Discord client is required")
  const expectedProfile = String(gameProfile || "").trim()
  if (!expectedProfile) throw new Error("Game profile is required")

  return async function deliverStateEventClaim(payload) {
    try {
      if (!client.isReady?.()) throw retryable("Discord client is not ready.")
      if (payload?.claim?.gameProfile !== expectedProfile) {
        throw permanent("Delivery game profile does not match this bot.")
      }
      if (!["state", "alliance"].includes(payload?.claim?.targetKind)) {
        throw permanent("State event target type is unsupported.")
      }
      if (payload?.claim?.targetIsCurrent !== true) {
        throw permanent("State event target changed or sharing was disabled.")
      }
      const image = payload.image ? prepareStoredEventImage(payload.image) : null
      const channel = await resolveDeliveryTarget(client, payload, { hasImage: Boolean(image) })
      const message = formatStateEventDelivery(payload, {
        imageFilename: image?.filename || null
      })
      if (image) message.files = [image.file]
      const sent = await channel.send(message)
      if (!String(sent?.id || "").trim()) {
        throw retryable("Discord did not return a message identifier.")
      }
      return Object.freeze({ sentMessageId: String(sent.id) })
    } catch (error) {
      throw normalizeDiscordDeliveryError(error)
    }
  }
}

module.exports = {
  createDiscordStateEventDeliveryHandler
}
