const {
  compareStateOccurrences,
  getStateEventOccurrencesInRange
} = require("./stateEventOccurrenceCalculation")

function stateRoundupKey(item) {
  return [
    item.stateEventId,
    item.phaseId,
    item.occurrenceAt.toISOString()
  ].join(":")
}

function compareStateRoundupOccurrences(left, right) {
  return left.occurrenceAt.getTime() - right.occurrenceAt.getTime()
    || String(left.stateNumber || "").localeCompare(String(right.stateNumber || ""))
    || String(left.eventName || "").localeCompare(String(right.eventName || ""))
    || compareStateOccurrences(left, right)
}

function buildStateRoundupOccurrences(events, weekStart, weekEnd) {
  const unique = new Map()
  for (const event of events || []) {
    if (event.status !== "active") continue
    for (const occurrence of getStateEventOccurrencesInRange(event, weekStart, weekEnd)) {
      const item = {
        ...occurrence,
        stateGuildId: event.state_guild_id,
        stateNumber: event.state_number,
        eventName: event.event_name
      }
      unique.set(stateRoundupKey(item), item)
    }
  }
  return [...unique.values()].sort(compareStateRoundupOccurrences)
}

module.exports = {
  stateRoundupKey,
  compareStateRoundupOccurrences,
  buildStateRoundupOccurrences
}
