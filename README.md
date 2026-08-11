# R.A.C.H.I.E / P.E.G.G.I.E

This repository runs two Discord bots from the same codebase:

- **R.A.C.H.I.E** uses the `wos` game profile and normally `BANTER_PROFILE=rachie`.
- **P.E.G.G.I.E** uses the `kingshot` game profile and normally `BANTER_PROFILE=peggie`.

They are separate Discord applications and separate Railway services. Each has its own `BOT_TOKEN`, `CLIENT_ID`, `APPS_SCRIPT_URL`, Apps Script deployment, and booking sheet. Those Apps Script systems remain independent. The optional Postgres event scheduler may share one `DATABASE_URL`; every scheduler query and claim remains isolated by `game_profile`.

## Existing And Scheduler Boundaries

Apps Script continues to own existing booking, state-linking, booking announcements, administration, and related sheet-backed behavior. The scheduler does not change those actions, URLs, sheets, or announcement channels.

Postgres owns scheduler data and, when enabled, canonical player accounts and gift-code records. If Postgres is disabled or unavailable, the Discord bot still starts and existing Apps Script-backed features continue to operate. Database-backed commands report temporary unavailability without taking down booking or banter.

See [Architecture](docs/architecture.md) for the complete boundary and data model.

## Event Scheduler

Administrators use `/event-scheduler`; authorization uses the existing `userCanManageServer(interaction)` behavior. `/event-scheduler-help` provides private, database-independent setup, creation, reminder, management, roundup, state-link and troubleshooting pages. The management home also includes a **Help** control.

Normal channel setup uses Discord's native channel selector. Set the main alliance through **Alliance identity**, use **Configure channels** for the alliance reminder channel, and use **Weekly roundup settings** for roundup enablement, state publishing, UTC weekday/time, and channel. No channel ID is typed. To link a state Discord, configure **State destination** inside that Discord, select its roundup channel, generate a one-time 15-minute code, and enter that code under **State sharing** in the alliance Discord. Guild and channel IDs are resolved internally.

An enabled state destination also exposes **State events**. After setting the state number/name, state admins can create canonical one-time or recurring events with one or more named phases. Each phase uses the shared free-form UTC parser, one of the standard pre-alert choices, an independent exact-time toggle, separate custom messages, and separate PNG/JPEG/GIF/WebP media for pre-alert and exact-time posts. Linked alliance Discords cannot create or manage these events.

A Discord guild and game profile can contain one main alliance and multiple sub-alliances, such as `YOU`, `YOU2`, and `YOU Academy`. Events select an alliance through an ephemeral opaque control. Alliance channels, weekly-roundup channels, and state links remain guild/profile settings and may be shared by every alliance in that guild.

Events support:

- Custom event names and a required alliance selection.
- A guided choice between one UTC time and named groups managed one at a time with separate UTC times; groups may share a time.
- First occurrence dates and recurrence every 2, 3, 7, 14, 21, 28, 35, or 42 days.
- One optional advance reminder: 5, 10, 15, 20, or 30 minutes before.
- One optional final announcement one minute before, worded **About to start**.
- Optional advance and final custom messages, each trimmed and limited to 500 characters.
- An optional image attached only to the alliance advance reminder.
- Alliance-only individual reminders.
- Alliance weekly roundups and combined state weekly roundups.
- Ordinary editing that retains alliance and image, separate alliance changes, explicit image retain/replace/remove, pause, resume, and soft deletion.
- Reminder cancellation by selecting no advance reminder or disabling the final announcement.

Accepted UTC time examples include `18:30`, `1830`, `1800`, `6:30pm`, and `6.30 PM`. Ambiguous values are rejected. The legacy database column `reminder_at_start` now means the one-minute final announcement; no exact-start delivery is created.

Immediate reminders contain the alliance, event, optional group, countdown, and optional custom text without date, recurrence, or timestamp fields. Detailed UTC and clearly labelled Discord-local times appear in previews and weekly roundups. Discord mention parsing is disabled. Images are never attached to final announcements, state Discord messages, roundups, or management previews. Images persist through ordinary edits and alliance changes unless explicitly replaced or removed.

Each alliance roundup contains eligible events for its guild/profile and lists alliance names. A linked state Discord receives one combined profile-scoped roundup containing active, roundup-enabled alliance events from valid enabled state links. Alliance events never produce individual state reminders.

Canonical state-event phase alerts publish once to the configured state Discord and once to each uniquely linked alliance Discord for the same profile. Their weekly milestones appear under one **STATE EVENTS** section in state and alliance roundups. Pre-alerts never create roundup milestones, while a phase remains in roundups even if its exact-time post is disabled. State-event management includes edit, phase add/edit/remove, next-occurrence preview, isolated TEST send, pause/resume, and soft delete.

See [Event Scheduler](docs/event-scheduler.md) for setup and workflow details.

## Player Accounts And Gift Codes

Set `PLAYER_GIFT_CODES_ENABLED=true` to register the profile-specific `/player-register`, `/gift-codes`, and `/gift-codes-admin` panels. They support multiple Whiteout Survival or Kingshot characters per Discord user, one active primary character per user/profile, private account management, public candidate submission, transactional State/Kingdom changes, and soft deactivation. PostgreSQL is canonical; the existing `/register`, `/my-info`, and `/unregister` booking identities remain unchanged in Apps Script.

Migration 011 also establishes case-sensitive gift codes, source/submission provenance, and restart-safe redemption identities with initial Player ID and State/Kingdom snapshots. It does not automatically redeem, discover, scrape, announce, or bulk-process codes. See [Player Accounts And Gift Codes](docs/player-gift-codes.md).

Migration 012 adds durable verification claims, immutable API-attempt history, notification state, and concurrency-safe redemption workers. Migration 013 adds profile-scoped guild settings, account/guild links, and restart-safe engagement delivery state. `/gift-codes` provides candidate submission, per-account opt-in/out, history, and private status. `/gift-codes-admin` provides authorized diagnostics, native channel configuration, contributor-role health, community statistics, and controlled one-code verification. Both live Century workers default to disabled.

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
- Player/gift-code variables: `PLAYER_GIFT_CODES_ENABLED`, `GIFT_CODE_MAX_AUTO_REDEEM_ACCOUNTS_PER_USER` (default `2` per user/profile), separately gated verification/redemption workers, profile-specific optional verifier characters, the shared standard `DATABASE_URL`, optional signing-suffix overrides, and conservative Century delay/backoff controls.
- Optional scheduler tuning: `EVENT_SCHEDULER_LOOKAHEAD_MINUTES`, `EVENT_SCHEDULER_GRACE_MINUTES`, `EVENT_SCHEDULER_POLL_INTERVAL_MS`, `EVENT_SCHEDULER_BATCH_SIZE`, `EVENT_SCHEDULER_CLAIM_LEASE_SECONDS`, and `EVENT_SCHEDULER_HANDLER_TIMEOUT_MS`.
- Test-only database variable: `TEST_DATABASE_URL`, which must reference a disposable Postgres database.

Invalid scheduler tuning values safely fall back to code defaults. `EVENT_SCHEDULER_ENABLED` must be exactly `true` to enable the subsystem.

## Migrations

Migrations run during scheduler initialization under a session-level Postgres advisory lock. The lock covers creating/checking `schema_migrations`, applying each pending migration transactionally, and recording its completed version. Concurrent services therefore serialize the complete migration decision and application path. Migrations can also be checked explicitly:

```bash
EVENT_SCHEDULER_ENABLED=true npm run migrate
```

When only the player/gift-code subsystem is enabled, the same portable command is:

```bash
PLAYER_GIFT_CODES_ENABLED=true npm run migrate
```

Migration files are append-only. Migrations 001 through 010 own the scheduler and state-event history described above. Migration 011 adds canonical profile-scoped player accounts, transfer history, gift codes, submissions, sources, and durable redemption identities. Migration 012 adds controlled verification/redemption claims and per-call attempt history. Migration 013 adds gift-code guild settings, account/guild links, and idempotent community-engagement events. Existing channels, schedules, state links, sent history, booking sheets, and Apps Script records are preserved. Run migrations again to confirm that zero files reapply.

## Testing

```bash
npm run check
npm test
```

Database tests use `TEST_DATABASE_URL` and must point only to a disposable PostgreSQL database. The suite includes two simultaneous migration runners and verifies that one applies the probe migration, the other observes it, and both finish with one valid schema record:

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

The application is not tied to Railway APIs. The same services can use a standard `DATABASE_URL` on a Linux VPS, Docker Compose, externally hosted PostgreSQL, or another managed PostgreSQL provider. Durable state remains in PostgreSQL rather than deployment-local files.

## Safe Rollout And Rollback

Back up Postgres, validate migrations on a disposable database, deploy one service at a time, and verify scheduler health before enabling public use. Database idempotency, leases, profile scope, and a compatibility trigger protect ordinary restarts and overlapping deployments.

The immediate rollback is `EVENT_SCHEDULER_ENABLED=false`. This disables scheduler commands and polling without affecting existing bot features. Leave additive migrations applied; reverting to old scheduler code after migrations 006 through 008 preserves existing rows, but old code does not understand sub-alliance management, custom messages, native state destinations, link codes, or independent roundup controls.

## Troubleshooting

- Scheduler says unavailable: use `/event-scheduler-help`, then inspect startup logs for missing `DATABASE_URL`, invalid `GAME_PROFILE`, missing `BOT_INSTANCE_NAME`, authentication, network, or migration errors.
- Commands are absent: confirm `EVENT_SCHEDULER_ENABLED=true` and that global command registration completed for that service's `CLIENT_ID`.
- Messages are not delivered: verify the configured guild/channel, bot membership, View Channel, Send Messages, Embed Links, and Attach Files permissions.
- State roundup is empty: verify the profile-scoped scheduler state link, sharing status, weekly-roundup inclusion, canonical state-event status, and weekly window.
- State link code is rejected: generate a fresh code in the state Discord and verify both sides use the same bot/game profile; codes expire after 15 minutes and work once.
- Duplicate alliance rejected: names are unique case-insensitively within one guild/profile.

## Security And Limitations

- Credentials stay in environment variables; no secrets belong in Git.
- Custom messages are plain text only, bounded to 500 characters, mention-neutralized, and cannot provide raw embeds or Discord payloads.
- Delivery is database-idempotent and transactionally claimed, but Discord itself does not offer a transactional send. A crash after Discord accepts a message and before its ID is stored can cause an at-least-once duplicate on retry.
- All scheduler dates and times are UTC; there is no per-user timezone conversion during creation.
- Alliance deletion is intentionally conservative: the main alliance cannot be deleted, and any alliance with event history must be retained.
- State links in this subsystem are independent from the existing Apps Script state-linking contract.
