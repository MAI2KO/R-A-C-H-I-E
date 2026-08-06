# R.A.C.H.I.E / P.E.G.G.I.E

This repository runs two Discord bots from the same codebase:

- **R.A.C.H.I.E** uses the `wos` game profile and normally `BANTER_PROFILE=rachie`.
- **P.E.G.G.I.E** uses the `kingshot` game profile and normally `BANTER_PROFILE=peggie`.

They are separate Discord applications and separate Railway services. Each has its own `BOT_TOKEN`, `CLIENT_ID`, `APPS_SCRIPT_URL`, Apps Script deployment, and booking sheet. Those Apps Script systems remain independent. The optional Postgres event scheduler may share one `DATABASE_URL`; every scheduler query and claim remains isolated by `game_profile`.

## Existing And Scheduler Boundaries

Apps Script continues to own existing booking, state-linking, booking announcements, administration, and related sheet-backed behavior. The scheduler does not change those actions, URLs, sheets, or announcement channels.

Postgres owns only scheduler guild settings, alliance identities, events, images, delivery claims, scheduler state links, and weekly roundups. If Postgres is disabled or unavailable, the Discord bot still starts and existing Apps Script-backed features continue to operate. Scheduler commands and polling remain unavailable until scheduler health recovers on a later restart.

See [Architecture](docs/architecture.md) for the complete boundary and data model.

## Event Scheduler

Administrators use `/event-scheduler`; authorization uses the existing `userCanManageServer(interaction)` behavior.

A Discord guild and game profile can contain one main alliance and multiple sub-alliances, such as `YOU`, `YOU2`, and `YOU Academy`. Events select an alliance through an ephemeral opaque control. Alliance channels, weekly-roundup channels, and state links remain guild/profile settings and may be shared by every alliance in that guild.

Events support:

- Custom event names and a required alliance selection.
- One UTC time, or named groups with a separate UTC time per group.
- First occurrence dates and recurrence every 3, 7, 14, or 28 days.
- One optional advance reminder: 5, 10, 15, 20, or 30 minutes before.
- One optional final announcement one minute before, worded **About to start**.
- Optional advance and final custom messages, each trimmed and limited to 500 characters.
- An optional image attached only to the alliance advance reminder.
- Alliance-only individual reminders.
- Alliance weekly roundups and combined state weekly roundups.
- Editing, image retain/replace/remove, pause, resume, and soft deletion.

Accepted UTC time examples include `18:30`, `1830`, `1800`, `6:30pm`, and `6.30 PM`. Ambiguous values are rejected. The legacy database column `reminder_at_start` now means the one-minute final announcement; no exact-start delivery is created.

Custom messages supplement the standard alliance, event, group, UTC time, and Discord timestamp details. Discord mention parsing is disabled. Images are never attached to final announcements, state Discord messages, roundups, or management previews. If an image is stored while the advance reminder is disabled, it remains available for a later edit but is not posted.

Each alliance roundup contains eligible events for its guild/profile and lists alliance names. A linked state Discord receives one combined profile-scoped roundup containing only active, roundup-enabled, `publish_to_state=true` events from valid enabled state links. State Discords never receive individual event reminders.

See [Event Scheduler](docs/event-scheduler.md) for setup and workflow details.

## Installation

Requirements: Node.js 18 or later, npm, a Discord application, and the existing Apps Script deployment. Postgres is optional unless the event scheduler is enabled.

```bash
npm install
cp .env.example .env
npm run check
npm test
npm start
```

`.env` is ignored by Git. Keep real credentials in local environment configuration or Railway variables; never commit them.

## Environment

All variables read by the code are represented in [.env.example](.env.example):

- Required existing bot variables: `BOT_TOKEN`, `CLIENT_ID`, `APPS_SCRIPT_URL`, `ADMIN_API_KEY`, and `OPENAI_API_KEY` where those existing features are used.
- Profile variables: `GAME_PROFILE` (`wos` default or `kingshot`) and `BANTER_PROFILE` (`rachie` default).
- Scheduler variables: `EVENT_SCHEDULER_ENABLED`, `DATABASE_URL`, and `BOT_INSTANCE_NAME`.
- Optional scheduler tuning: `EVENT_SCHEDULER_LOOKAHEAD_MINUTES`, `EVENT_SCHEDULER_GRACE_MINUTES`, `EVENT_SCHEDULER_POLL_INTERVAL_MS`, `EVENT_SCHEDULER_BATCH_SIZE`, `EVENT_SCHEDULER_CLAIM_LEASE_SECONDS`, and `EVENT_SCHEDULER_HANDLER_TIMEOUT_MS`.
- Test-only database variable: `TEST_DATABASE_URL`, which must reference a disposable Postgres database.

Invalid scheduler tuning values safely fall back to code defaults. `EVENT_SCHEDULER_ENABLED` must be exactly `true` to enable the subsystem.

## Migrations

Migrations run during scheduler initialization under a Postgres advisory lock. They can also be checked explicitly:

```bash
EVENT_SCHEDULER_ENABLED=true npm run migrate
```

Migration files are append-only. Migrations 001 through 005 preserve the initial scheduler history; migration 006 adds flexible reminders, custom messages, and profile-scoped alliance entities while backfilling existing events. Run migrations again to confirm that zero files reapply.

## Testing

```bash
npm run check
npm test
```

Database tests use `TEST_DATABASE_URL` and must point only to a disposable PostgreSQL database:

```bash
TEST_DATABASE_URL=postgresql://user:password@127.0.0.1:5432/disposable npm test
```

Tests use mocked Discord clients and must not target live Discord, Apps Script, Railway, or production Postgres. See [Testing](docs/testing.md).

## Railway Deployment

R.A.C.H.I.E and P.E.G.G.I.E remain separate services. They may share Postgres but must use distinct identities:

```text
R.A.C.H.I.E: GAME_PROFILE=wos,      BOT_INSTANCE_NAME=rachie-wos
P.E.G.G.I.E: GAME_PROFILE=kingshot, BOT_INSTANCE_NAME=peggie-kingshot
Both:         EVENT_SCHEDULER_ENABLED=true
```

Each service retains its own Discord and Apps Script variables. Follow [Deployment](docs/deployment.md) for clean migration validation, staged rollout, health checks, and rollback.

## Safe Rollout And Rollback

Back up Postgres, validate migrations on a disposable database, deploy one service at a time, and verify scheduler health before enabling public use. Database idempotency, leases, profile scope, and a compatibility trigger protect ordinary restarts and overlapping deployments.

The immediate rollback is `EVENT_SCHEDULER_ENABLED=false`. This disables scheduler commands and polling without affecting existing bot features. Leave additive migrations applied; reverting to old scheduler code after migration 006 is supported for ordinary default-alliance writes, but old code does not understand sub-alliance management or custom messages.

## Troubleshooting

- Scheduler says unavailable: inspect startup logs for missing `DATABASE_URL`, invalid `GAME_PROFILE`, missing `BOT_INSTANCE_NAME`, authentication, network, or migration errors.
- Commands are absent: confirm `EVENT_SCHEDULER_ENABLED=true` and that global command registration completed for that service's `CLIENT_ID`.
- Messages are not delivered: verify the configured guild/channel, bot membership, View Channel, Send Messages, Embed Links, and Attach Files permissions.
- State roundup is empty: verify the profile-scoped scheduler state link, sharing status, event `publish_to_state`, roundup inclusion, active status, and weekly window.
- Duplicate alliance rejected: names are unique case-insensitively within one guild/profile.

## Security And Limitations

- Credentials stay in environment variables; no secrets belong in Git.
- Custom messages are plain text only, bounded to 500 characters, mention-neutralized, and cannot provide raw embeds or Discord payloads.
- Delivery is database-idempotent and transactionally claimed, but Discord itself does not offer a transactional send. A crash after Discord accepts a message and before its ID is stored can cause an at-least-once duplicate on retry.
- All scheduler dates and times are UTC; there is no per-user timezone conversion during creation.
- Alliance deletion is intentionally conservative: the main alliance cannot be deleted, and any alliance with event history must be retained.
- State links in this subsystem are independent from the existing Apps Script state-linking contract.
