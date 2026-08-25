# Website Alliance Events read API

The scheduler PostgreSQL tables remain authoritative. This opt-in internal listener turns active alliance schedules into a narrow public-safe model for the website; it does not write schedules or call Discord.

Set all of the following on each bot service to enable it:

- `ALLIANCE_EVENTS_READ_ENABLED=true`
- `ALLIANCE_EVENTS_READ_SECRET` to a random value of at least 32 characters (different per profile)
- `ALLIANCE_EVENTS_READ_PORT` to the private listener port

When configuration is absent or invalid the listener is dormant and existing bot behavior is unchanged. It starts only after the scheduler database health check succeeds. Requests are signed with HMAC-SHA256 over method, path, profile, timestamp, and nonce; timestamps have a five-minute tolerance. Put the listener behind private service networking/TLS and do not expose its secret in browser code.

`GET /internal/v1/public-alliance-events/guild/{guildId}` reads active `scheduled_events` and their `event_alliances` and `scheduled_event_groups` for that Discord guild and the process `GAME_PROFILE`. Discord guild IDs must be 15-22 digits. It does not read `event_state_destinations`, `event_state_links`, or `state_events`: those tables describe Discord State delivery, not website community membership. Paused/deleted events are filtered by `scheduled_events.status = 'active'`; publish and roundup settings are unrelated.

The older `GET /internal/v1/public-alliance-events/{stateNumber}` route remains available for compatibility with existing internal consumers, but the website does not use it.

The guild response contains only profile, alliance name/optional abbreviation, event name, recurrence summary, group name, and the next three UTC instants. Alliance abbreviations are currently not stored by the scheduler, so the field is `null`. Discord/database IDs, delivery history, claims, images, settings, and secrets never cross this boundary. Occurrences come directly from `occurrenceCalculation.js`; the website does not implement recurrence arithmetic.

Failures return bounded codes without database or secret details. The website handles them as a page-local degraded state. Images and transfer/chat features are deliberately outside this integration.
