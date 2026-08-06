const { getOccurrencesInRange } = require("./occurrenceCalculation")

const MINUTE_MS = 60 * 1000
const MAX_ADVANCE_MINUTES = 30

function requireInstant(value, label) {
  if (!(value instanceof Date) || !Number.isFinite(value.getTime())) {
    throw new Error(`${label} must be a valid Date`)
  }
  return value.getTime()
}

function deliveryWindow(now, { lookaheadMinutes, graceMinutes }) {
  const nowMs = requireInstant(now, "Generation time")
  return {
    start: new Date(nowMs - graceMinutes * MINUTE_MS),
    end: new Date(nowMs + lookaheadMinutes * MINUTE_MS)
  }
}

function claimFor(event, occurrence, deliveryKind, deliverAt) {
  return {
    eventId: event.id,
    groupId: occurrence.groupId,
    gameProfile: event.game_profile,
    occurrenceAt: occurrence.occurrenceAt,
    deliverAt,
    deliveryKind,
    targetKind: "alliance",
    targetGuildId: event.guild_id,
    targetChannelId: event.event_channel_id
  }
}

function buildDeliveryClaims(events, { gameProfile, windowStart, windowEnd }) {
  const startMs = requireInstant(windowStart, "Delivery window start")
  const endMs = requireInstant(windowEnd, "Delivery window end")
  if (endMs < startMs) throw new Error("Delivery window end must not be before its start")
  if (endMs === startMs) return []

  const claims = []
  const occurrenceEnd = new Date(endMs + MAX_ADVANCE_MINUTES * MINUTE_MS)
  for (const event of events) {
    if (
      event.game_profile !== gameProfile
      || event.status !== "active"
      || event.publish_to_alliance !== true
      || !String(event.event_channel_id || "").trim()
    ) continue

    const occurrences = getOccurrencesInRange(event, windowStart, occurrenceEnd)
    for (const occurrence of occurrences) {
      const reminderMinutes = Number(event.advance_reminder_minutes)
      if ([10, 30].includes(reminderMinutes)) {
        const deliverAt = new Date(
          occurrence.occurrenceAt.getTime() - reminderMinutes * MINUTE_MS
        )
        if (deliverAt.getTime() >= startMs && deliverAt.getTime() < endMs) {
          claims.push(claimFor(event, occurrence, "advance_reminder", deliverAt))
        }
      }
      if (
        event.reminder_at_start === true
        && occurrence.occurrenceAt.getTime() >= startMs
        && occurrence.occurrenceAt.getTime() < endMs
      ) {
        claims.push(claimFor(
          event,
          occurrence,
          "event_start",
          new Date(occurrence.occurrenceAt)
        ))
      }
    }
  }
  return claims
}

async function generateMissingDeliveryClaims({
  repository,
  gameProfile,
  now = new Date(),
  config
}) {
  const window = deliveryWindow(now, config)
  const events = await repository.listActiveEventDefinitions({
    rangeStart: window.start,
    rangeEnd: new Date(window.end.getTime() + MAX_ADVANCE_MINUTES * MINUTE_MS)
  })
  const claims = buildDeliveryClaims(events, {
    gameProfile,
    windowStart: window.start,
    windowEnd: window.end
  })
  const inserted = await repository.insertMissingDeliveryClaims(claims)
  return { window, consideredEvents: events.length, generated: claims.length, inserted }
}

module.exports = {
  MINUTE_MS,
  MAX_ADVANCE_MINUTES,
  deliveryWindow,
  buildDeliveryClaims,
  generateMissingDeliveryClaims
}
