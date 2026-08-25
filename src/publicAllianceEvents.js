const { getNextOccurrences } = require("./occurrenceCalculation")
const { recurrenceLabel } = require("./eventSchedulerFormatting")

function compareText(left, right) {
  return String(left).localeCompare(String(right), "en", { sensitivity: "base" })
}

function publicAllianceEventsReadModel({ profile, communityCode, events, now = new Date(), count = 3 }) {
  const alliances = new Map()
  for (const event of events) {
    const allianceKey = String(event.alliance_id)
    if (!alliances.has(allianceKey)) {
      alliances.set(allianceKey, {
        sortKey: allianceKey,
        name: String(event.alliance_name),
        abbreviation: null,
        events: []
      })
    }
    const upcoming = getNextOccurrences({
      id: event.event_id,
      alliance_name: event.alliance_name,
      event_name: event.event_name,
      first_occurrence_date: event.first_occurrence_date,
      event_time_utc: event.event_time_utc,
      recurrence_days: event.recurrence_days,
      status: "active",
      groups: event.groups || []
    }, now, count).map(occurrence => ({
      at: occurrence.occurrenceAt.toISOString(),
      group: occurrence.groupName || null
    }))
    alliances.get(allianceKey).events.push({
      sortKey: String(event.event_id),
      name: String(event.event_name),
      recurrence: {
        days: Number(event.recurrence_days),
        summary: recurrenceLabel(Number(event.recurrence_days))
      },
      upcoming
    })
  }

  const publicAlliances = [...alliances.values()]
    .sort((left, right) => compareText(left.name, right.name)
      || compareText(left.sortKey, right.sortKey))
    .map(alliance => ({
      name: alliance.name,
      abbreviation: alliance.abbreviation,
      events: alliance.events
        .sort((left, right) => String(left.upcoming[0]?.at || "").localeCompare(
          String(right.upcoming[0]?.at || "")
        ) || compareText(left.name, right.name) || compareText(left.sortKey, right.sortKey))
        .map(({ sortKey, ...event }) => event)
    }))
  return communityCode === undefined
    ? { profile, alliances: publicAlliances }
    : { profile, communityCode, alliances: publicAlliances }
}

module.exports = { publicAllianceEventsReadModel }
