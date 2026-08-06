const { resolveDeliveryTarget, normalizeDiscordDeliveryError } = require("./discordEventDelivery")
const { formatWeeklyRoundup } = require("./weeklyRoundupFormatting")
const { PermanentDeliveryError } = require("./eventDeliveryWorker")

function createDiscordWeeklyRoundupDelivery({ client, gameProfile }) {
  if (!client) throw new Error("Discord client is required")
  const expectedProfile = String(gameProfile || "").trim()
  if (!expectedProfile) throw new Error("Game profile is required")

  return async function prepareRoundup(payload) {
    try {
      if (!client.isReady?.()) throw new Error("Discord client is not ready")
      if (payload?.claim?.gameProfile !== expectedProfile) {
        throw new PermanentDeliveryError("Roundup game profile does not match this bot.")
      }
      const targetPayload = {
        claim: {
          targetKind: payload.claim.targetKind,
          targetGuildId: payload.claim.targetGuildId,
          targetChannelId: payload.claim.targetChannelId,
          targetIsCurrent: payload.claim.targetIsCurrent
        },
        alliance: { guildId: payload.claim.targetGuildId },
        event: { guildId: payload.claim.targetGuildId }
      }
      const channel = await resolveDeliveryTarget(client, targetPayload, { hasImage: false })
      const messages = formatWeeklyRoundup(payload)
      return {
        messages,
        async sendPart(index) {
          try {
            const sent = await channel.send(messages[index])
            if (!String(sent?.id || "").trim()) throw new Error("Discord returned no message ID")
            return String(sent.id)
          } catch (error) {
            throw normalizeDiscordDeliveryError(error)
          }
        }
      }
    } catch (error) {
      throw normalizeDiscordDeliveryError(error)
    }
  }
}

module.exports = { createDiscordWeeklyRoundupDelivery }
