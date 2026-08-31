# Native booking website integration

## Player-local appointment display

All player-facing booking messages—automatic confirmations, manually approved confirmations, reschedules, cancellations, and 30-minute reminders—show both the canonical UTC appointment and a Discord-native `<t:UNIX:F>` timestamp. Discord renders the native timestamp in the recipient's configured locale and timezone, so daylight-saving rules are handled by Discord and the bot does not persist user timezones or fixed UTC offsets.

The website supplies the exact PostgreSQL appointment instant from the booked slot. Reschedules supply separate old and new instants. This is presentation-only: UTC storage, reminder timing, reminder deduplication, reschedule supersession, and cancellation suppression are unchanged. The formatting uses the same safe Discord timestamp helper as the event scheduler. Missing or malformed instants never produce invalid Discord markup.

This subsystem connects one bot deployment to the matching native booking website profile. It is disabled unless `BOOKING_WEBSITE_INTEGRATION_ENABLED=true` and a valid profile, HTTPS base URL, and 32-character-or-longer secret are all present. Plain HTTP is accepted only for loopback local development.

The website owns bookings, approval transitions, notification decisions, persistence, retries, and reminder scheduling. The bot has no website database credentials. It polls signed internal HTTPS endpoints, discovers current Discord managers, sends or edits DMs, receives approval buttons, and reports delivery outcomes. The same signed client also previews/applies `/setup` community linkage and projects `/register` identity into native booking prefill. The signed host and `GAME_PROFILE` determine WOS or Kingshot; neither command exposes a profile selector.

Configure independently per deployment:

```text
R.A.C.H.I.E: GAME_PROFILE=wos
P.E.G.G.I.E: GAME_PROFILE=kingshot
BOOKING_WEBSITE_INTEGRATION_ENABLED=true
BOOKING_WEBSITE_BASE_URL=https://<matching-profile-hostname>
BOOKING_WEBSITE_INTEGRATION_SECRET=<matching-profile-only-random-secret>
BOOKING_WEBSITE_POLL_INTERVAL_MS=10000
```

Never share the WOS and Kingshot secrets. The HMAC covers method, path, timestamp, nonce, and exact body. The website validates its hostname profile, clock tolerance, and one-use nonce before doing work.

## Discord setup

`/setup` requires this integration because it reconciles the native community
link. The preview is read-only. Apply reuses an active matching community, or
creates a WOS community with the platform's deterministic cycle defaults when
none exists, then links the current guild. Creation and linking are profile
scoped, transactionally audited, advisory-lock serialized, and safe to rerun.
An existing link or a community already claimed by another active guild fails
closed instead of being remapped. Kingshot creation currently fails with an
explicit unsupported-defaults response; no cycle is invented. `/register` stores the complete canonical
identity in the bot database, then upserts the matching website participant by
the signed guild/profile scope. If the website is temporarily unavailable the
player is told to retry; no legacy Sheet registration is attempted.

Manager discovery must enumerate current guild members so it can include guild owners, current Administrators, and current holders of `bot_manager_role_id`, including uncached members. In each Discord Developer Portal application, enable **Bot → Privileged Gateway Intents → Server Members Intent** before turning on the integration. When enabled, this code conditionally requests `GuildMembers`; when disabled it does not alter existing intents.

`bot_manager_role_id` is stored only in the profile-scoped
`bot_managed_discord_setups` PostgreSQL row. Migration 021 is additive. There is
no automatic Apps Script import because the legacy deployments are separate
external authorities and bulk import cannot be made transactional or safely
profile-scoped. After migration, an owner or Discord Administrator must run
`/set-bot-admin-role` once in each guild that used a custom role.

## Automatic booking-open delivery

For every open WOS window, the website transaction creates at most one
`booking_window_open` work row and one window-bound guest link. The token is a
profile/community/window-specific HMAC derived from the integration secret, so
retries reproduce the same opaque URL while PostgreSQL stores only its SHA-256
hash and six-character hint. Opening a later window revokes the previous link;
Sunday 12:00 UTC reconciliation revokes the current link and copied URLs stop
authorising guest booking.

The bot posts the public message to each linked guild's native managed
`minister-sign-up` channel using stable Discord nonces. It sends the
copy-friendly link block by DM to the deduplicated guild owner,
Administrators, and configured native manager-role holders. This is the
existing safe manager-only delivery path; no new public admin channel is
invented. The member button checks the canonical PostgreSQL `/register`
identity: registered members receive the native booking URL, while others are
directed through `/register` first. The guest button is the fresh window URL.

The bot must remain a member of every guild linked to the profile's State/Kingdom. It needs ordinary permission to access those guilds and send DMs; no broad website or PostgreSQL permission is added.

## Startup and poll diagnostics

`npm start` executes `node index.js`. `main()` calls the booking bootstrap before Discord login; the bootstrap parses configuration, creates the API client/worker when enabled, and registers the worker start on Discord's `clientReady` event. The first poll is immediate after that event. The interval then continues after empty responses, successful delivery, and transient failures.

Startup emits structured, secret-free records showing profile, enabled/requested state, whether URL and secret are configured, poll interval, disabled reason, and whether the worker started. Polling logs only the first successful website connection, recovery after a failure, non-zero claimed counts, and one record per consecutive failure category. It does not log successful empty polls, URLs, signatures, headers, nonces, bodies, or credentials.

The exact flag value is case-insensitive `true` with surrounding whitespace ignored. No profile-prefixed bot variable is expected. `GAME_PROFILE` is `wos` for R.A.C.H.I.E and `kingshot` for P.E.G.G.I.E; blank defaults to `wos`. Missing/unsafe URL, missing/short secret, unsupported profile, or any other flag value keeps the subsystem disabled and produces a bounded reason.

## Operation and failures

Work is claimed in short leases. DM sends use a stable Discord nonce with nonce enforcement, limiting duplicates after a send succeeds but acknowledgement fails without displaying an internal identifier. Discord `50007` (cannot send messages to this user) and other known permanent access/not-found responses are terminal; temporary errors are returned for persisted retry. Booking state is never rolled back because a DM failed.

Approval buttons contain only a request UUID and action. The bot supplies the clicking user's canonical Discord ID, but the website performs live authorization and calls its existing atomic approval domain. Other manager copies are edited by later durable jobs. All interaction replies are ephemeral.

Logs contain only static event names, game profile, work type, and bounded error codes. They never include message bodies, integration secrets, bot tokens, OAuth values, cookies, session tokens, or raw website responses.

To stop delivery immediately, set `BOOKING_WEBSITE_INTEGRATION_ENABLED=false` and restart the bot. The website keeps queued work for later. Do not delete queue rows or reverse the additive website migration.

Local tests mock both Discord and HTTP. Never put real credentials in `.env` for the test suite:

```bash
npm run check
npm test
git diff --check
```
