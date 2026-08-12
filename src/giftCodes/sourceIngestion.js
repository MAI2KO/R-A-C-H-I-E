const { MessageReferenceType } = require("discord.js")
const { normalizeGiftCode } = require("./validation")
const { parseDiscordMirrorMessage } = require("./sourceParsers")

function safeSourceError(error) {
  return String(error?.code || error?.name || "source_error").slice(0, 100)
}

function snapshotText(snapshot) {
  const parts = []
  if (snapshot?.content) parts.push(String(snapshot.content))
  for (const embed of Array.isArray(snapshot?.embeds) ? snapshot.embeds : []) {
    if (embed?.title) parts.push(String(embed.title))
    if (embed?.description) parts.push(String(embed.description))
    for (const field of Array.isArray(embed?.fields) ? embed.fields : []) {
      if (field?.name && field?.value) parts.push(`${field.name}: ${field.value}`)
    }
  }
  return parts.join("\n")
}

function forwardedSnapshotTexts(message) {
  if (message?.reference?.type !== MessageReferenceType.Forward) return []
  const snapshots = message.messageSnapshots
  if (!snapshots || typeof snapshots.values !== "function") return []
  return [...snapshots.values()].map(snapshotText).filter(text => text.trim())
}

function parseDiscordSourceMessage(message, observedAt) {
  if (message?.reference?.type !== MessageReferenceType.Forward) {
    const parsed = parseDiscordMirrorMessage(message?.content, observedAt)
    return parsed ? { ...parsed, sourceMessageKind: "message" } : null
  }
  for (const text of forwardedSnapshotTexts(message)) {
    const parsed = parseDiscordMirrorMessage(text, observedAt)
    if (parsed) return { ...parsed, sourceMessageKind: "forwarded_snapshot" }
  }
  return null
}

function createGiftCodeSourceIngestionService({ giftRepository, sourceRepository, gameProfile, logger = console }) {
  async function ingest({
    source,
    code,
    observationKey,
    observedAt = new Date(),
    sourceReportedExpiryAt = null,
    provenance = {},
    submittedByDiscordUserId = null,
    submissionKind = source.sourceType
  }) {
    const exactCode = normalizeGiftCode(code)
    const sourceRow = await sourceRepository.ensureSource(source)
    const priorObservation = await sourceRepository.observation(sourceRow.id, observationKey)
    if (priorObservation) {
      await sourceRepository.recordObservation({
        sourceId: sourceRow.id,
        giftCodeId: priorObservation.id,
        code: exactCode,
        observationKey,
        observedAt,
        sourceReportedExpiryAt,
        provenance,
        candidateCreated: false
      })
      return {
        giftCode: priorObservation,
        submission: null,
        duplicate: true,
        duplicateObservation: true,
        source: sourceRow
      }
    }
    const recorded = await giftRepository.recordSubmission({
      code: exactCode,
      submittedByDiscordUserId,
      sourceId: sourceRow.id,
      metadata: {
        submissionKind,
        transport: provenance.transport || source.sourceType,
        guildId: provenance.guildId || null,
        sourceReportedExpiryAt: sourceReportedExpiryAt?.toISOString?.() || null
      }
    })
    await sourceRepository.recordObservation({
      sourceId: sourceRow.id,
      giftCodeId: recorded.giftCode.id,
      code: exactCode,
      observationKey,
      observedAt,
      sourceReportedExpiryAt,
      provenance,
      candidateCreated: !recorded.duplicate
    })
    return { ...recorded, source: sourceRow }
  }

  return Object.freeze({
    ingest,

    async ingestDiscordMessage(message) {
      try {
        if (!message?.guildId || !message?.channelId || !message?.id) return null
        const configured = await sourceRepository.discordChannel(message.channelId, message.guildId)
        if (!configured || (configured.require_webhook && !message.webhookId)) return null
        const observedAt = new Date(message.createdTimestamp || Date.now())
        const parsed = parseDiscordSourceMessage(message, observedAt)
        if (!parsed) return null
        return ingest({
          source: {
            sourceType: "discord_mirror",
            sourceName: configured.source_name,
            sourceReference: `discord:${message.guildId}:${message.channelId}`,
            trusted: true
          },
          code: parsed.code,
          observationKey: `discord-message:${message.id}`,
          observedAt,
          sourceReportedExpiryAt: parsed.sourceReportedExpiryAt,
          provenance: {
            transport: "discord",
            guildId: message.guildId,
            channelId: message.channelId,
            messageId: message.id,
            webhookId: message.webhookId || null,
            sourceDisplayName: String(message.author?.username || message.member?.displayName || "").slice(0, 100) || null,
            ...(parsed.sourceMessageKind === "forwarded_snapshot" ? {
              sourceMessageKind: "forwarded_snapshot",
              forwardedSourceGuildId: message.reference?.guildId || null,
              forwardedSourceChannelId: message.reference?.channelId || null,
              forwardedSourceMessageId: message.reference?.messageId || null
            } : {})
          }
        })
      } catch (error) {
        logger.warn(JSON.stringify({
          event: "gift_code_source_message_failed",
          game_profile: gameProfile,
          guild_id: message?.guildId,
          channel_id: message?.channelId,
          error_code: safeSourceError(error)
        }))
        return null
      }
    }
  })
}

module.exports = {
  safeSourceError,
  snapshotText,
  forwardedSnapshotTexts,
  parseDiscordSourceMessage,
  createGiftCodeSourceIngestionService
}
