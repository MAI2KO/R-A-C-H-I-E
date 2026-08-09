const { MINUTE_MS, deliveryWindow } = require("./eventDeliveryGeneration")
const { getStateEventOccurrencesInRange } = require("./stateEventOccurrenceCalculation")

const MAX_STATE_PRE_ALERT_MINUTES = 30

function claimFor(event, occurrence, deliveryKind, deliverAt, target) {
  return {
    stateEventId: event.id,
    phaseId: occurrence.phaseId,
    gameProfile: event.game_profile,
    scheduleVersion: event.schedule_version || 1,
    occurrenceAt: occurrence.occurrenceAt,
    deliverAt,
    deliveryKind,
    targetKind: target.target_kind || target.kind,
    targetGuildId: target.target_guild_id || target.guildId,
    targetChannelId: target.target_channel_id || target.channelId
  }
}

function uniqueTargetsByGuild(targets) {
  const unique = new Map()
  for (const target of targets || []) {
    const guildId = target.target_guild_id || target.guildId
    if (!guildId || unique.has(String(guildId))) continue
    unique.set(String(guildId), target)
  }
  return [...unique.values()]
}

async function buildStateEventDeliveryClaims(events, {
  repository,
  gameProfile,
  windowStart,
  windowEnd
}) {
  const startMs = windowStart.getTime()
  const endMs = windowEnd.getTime()
  if (endMs < startMs) throw new Error("Delivery window end must not be before its start")
  if (endMs === startMs) return []

  const claims = []
  const targetsByStateGuild = new Map()
  const occurrenceEnd = new Date(endMs + MAX_STATE_PRE_ALERT_MINUTES * MINUTE_MS)
  for (const event of events) {
    if (event.game_profile !== gameProfile || event.status !== "active") continue
    if (!targetsByStateGuild.has(event.state_guild_id)) {
      targetsByStateGuild.set(
        event.state_guild_id,
        uniqueTargetsByGuild(await repository.listTargetsForStateGuild(event.state_guild_id))
      )
    }
    const targets = targetsByStateGuild.get(event.state_guild_id)
    if (!targets.length) continue
    for (const occurrence of getStateEventOccurrencesInRange(event, windowStart, occurrenceEnd)) {
      for (const target of targets) {
        const preAlertMinutes = Number(occurrence.preAlertMinutes)
        if ([5, 10, 15, 20, 30].includes(preAlertMinutes)) {
          const deliverAt = new Date(occurrence.occurrenceAt.getTime() - preAlertMinutes * MINUTE_MS)
          if (deliverAt.getTime() >= startMs && deliverAt.getTime() < endMs) {
            claims.push(claimFor(event, occurrence, "pre_alert", deliverAt, target))
          }
        }
        if (occurrence.announceExact === true) {
          const deliverAt = new Date(occurrence.occurrenceAt)
          if (deliverAt.getTime() >= startMs && deliverAt.getTime() < endMs) {
            claims.push(claimFor(event, occurrence, "exact", deliverAt, target))
          }
        }
      }
    }
  }
  return claims
}

async function generateMissingStateEventDeliveryClaims({
  repository,
  gameProfile,
  now = new Date(),
  config
}) {
  const window = deliveryWindow(now, config)
  const events = await repository.listActiveStateEventDefinitions({
    rangeEnd: new Date(window.end.getTime() + MAX_STATE_PRE_ALERT_MINUTES * MINUTE_MS)
  })
  const claims = await buildStateEventDeliveryClaims(events, {
    repository,
    gameProfile,
    windowStart: window.start,
    windowEnd: window.end
  })
  const inserted = await repository.insertMissingDeliveryClaims(claims)
  return { window, consideredEvents: events.length, generated: claims.length, inserted }
}

module.exports = {
  MAX_STATE_PRE_ALERT_MINUTES,
  claimFor,
  uniqueTargetsByGuild,
  buildStateEventDeliveryClaims,
  generateMissingStateEventDeliveryClaims
}
