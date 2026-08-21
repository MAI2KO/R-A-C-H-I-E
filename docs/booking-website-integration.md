# Native booking website integration

This optional subsystem connects one bot deployment to the matching native booking website profile. It is disabled unless `BOOKING_WEBSITE_INTEGRATION_ENABLED=true` and a valid profile, HTTPS base URL, and 32-character-or-longer secret are all present. Plain HTTP is accepted only for loopback local development. Existing Apps Script, event, banter, setup, and gift-code behavior continues when it is disabled or misconfigured.

The website owns bookings, approval transitions, notification decisions, persistence, retries, and reminder scheduling. The bot has no website database credentials. It polls signed internal HTTPS endpoints, discovers current Discord managers, sends or edits DMs, receives approval buttons, and reports delivery outcomes.

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

Manager discovery must enumerate current guild members so it can include guild owners, current Administrators, and current holders of `bot_manager_role_id`, including uncached members. In each Discord Developer Portal application, enable **Bot → Privileged Gateway Intents → Server Members Intent** before turning on the integration. When enabled, this code conditionally requests `GuildMembers`; when disabled it does not alter existing intents.

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
