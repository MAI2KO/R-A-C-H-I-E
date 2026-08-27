# Architecture

## Deployment Shape

The repository serves two independent Discord applications:

| Service | Game profile | Bot instance | Discord/App Script ownership |
| --- | --- | --- | --- |
| R.A.C.H.I.E | `wos` | `rachie-wos` | Its own token, client ID, Apps Script URL, deployment, and booking sheet |
| P.E.G.G.I.E | `kingshot` | `peggie-kingshot` | Its own token, client ID, Apps Script URL, deployment, and booking sheet |

The bot services may share standard Postgres. `game_profile` is included in scheduler, player, gift-code, and redemption ownership. A process handles only its own profile. `BOT_INSTANCE_NAME` records worker ownership; it is not a secret.

## Remaining Apps Script Boundary

The only registered live Apps Script dependency is `/sheet-link`, which reads
the legacy Sheet URL. Legacy handler source and responsibility clients remain
for audit/rollback reference but booking, `/register`, bot-manager
authorization, and booking-open announcements do not call them. Banter command
registration and passive message handling are dormant.

Website manager authorization uses the existing signed internal listener. The
website sends the exact linked guild and authenticated Discord user to the
matching profile bot; the bot evaluates current owner/Administrator state and
the native PostgreSQL bot-manager role. Role decisions are live and uncached.

Historically, `index.js` routed Apps Script payloads through four internal clients:

- booking: registration, availability, booking dates and links, open/close, user/admin bookings, cancellation, clearing, and reservations;
- state/registry: setup, linking/unlinking, linked-server lookup, join-password reset, announcement configuration, and Sheet links/access;
- configuration/admin-role: booking requirement/settings reads and writes plus bot-admin role lookup/update;
- banter: channel and spice reads/writes, still wrapped by the existing five-minute cache.

Each client has an optional responsibility-specific URL and otherwise resolves to the existing `APPS_SCRIPT_URL`. The action names, bodies, admin key, Axios JSON headers, response handling, and profile-independent code path are unchanged. The scheduler neither calls these actions nor reuses their channel fields. Each bot keeps its own Apps Script deployment and sheets.

The canonical player-account subsystem is also separate from Apps Script booking identities. Its optional mirror is an interface only; no Sheet write is guessed. A mirror failure occurs after the PostgreSQL commit and is safe to retry.

## Portable Process Boundaries

All durable application state uses standard PostgreSQL through `DATABASE_URL` and append-only SQL migrations. Discord interaction modules, repositories, the Century client, and rate limiter have explicit dependency boundaries. The Century client and any future redemption worker can run in a separate Node.js process without a Discord client. No business logic uses Railway hostnames, service identifiers, internal DNS, database APIs, or deployment-local durable files.

A future container deployment can run PostgreSQL, R.A.C.H.I.E, P.E.G.G.I.E, a website/API, and independently scaled background workers. A future external cache or queue must sit behind an interface; PostgreSQL remains the current durable queue/idempotency layer.

Gift-code behavior follows `Discord -> service -> repository/client`. Candidate verification and player redemption processors are importable modules independent of Discord interaction handlers. PostgreSQL leases provide cross-process ownership; the in-process limiter controls only request pacing. `gift_code_attempts` stores each completed API call while the parent code/redemption rows hold current queue state.

Scheduler failure is isolated. Discord login and existing slash-command registration do not wait for Postgres. When enabled, scheduler initialization validates ownership, connects, runs migrations, and records health. Failure closes the scheduler pool and leaves the rest of the bot running. Database-backed management returns a private unavailable response; static scheduler help remains usable.

## Scheduler Components

- `eventSchedulerInteractions.js`: authorised ephemeral entry point and guild/channel configuration.
- `eventSchedulerHelp.js`: database-independent, ephemeral help pages and controls.
- `stateLinkCodes.js`: random, normalized and hashed one-time state destination codes.
- `allianceManagementInteractions.js`: opaque, profile-scoped alliance add/list/rename/delete flows.
- `eventCreationInteractions.js`: alliance selection, modal input, options, preview, and confirmation.
- `eventSchedulerRepository.js`: transactional event/alliance writes and claim reconciliation.
- `occurrenceCalculation.js`: deterministic UTC recurrence arithmetic.
- `eventDeliveryGeneration.js`: bounded claim generation for advance and one-minute final deliveries.
- `eventDeliveryRepository.js`: idempotent insertion, reconciliation, leases, and `FOR UPDATE SKIP LOCKED` claiming.
- `discordEventDelivery.js`: target/permission validation and alliance-only sends.
- `weeklyRoundup*`: UTC scheduling, selection, formatting, multipart persistence, and retry.

## Data Ownership

`event_guild_settings` is keyed by `(guild_id, game_profile)` and owns shared alliance reminder and weekly-roundup channels, the UTC roundup weekday/time, independent alliance/state enablement, and a not-before replay boundary. Its reminder channel may be null during identity-first setup. `event_state_destinations` records a channel selected natively inside a state guild. `event_state_link_codes` stores only a hash, profile, destination, expiry and one-time consumption state. Consuming a valid code writes the existing `event_state_links` contract used by weekly roundups, so older stored links remain compatible.

`event_alliances` gives a guild/profile one main alliance and any number of sub-alliances. Names are unique case-insensitively within that scope. The same name may exist in another guild or game profile.

`scheduled_events` stores `alliance_id`, `guild_id`, and `game_profile`. Its composite foreign key requires all three to match one alliance row. Legacy `alliance_name` remains denormalized for compatibility and is synchronized by a database trigger. Existing rows were backfilled by migration 006.

Event groups and images retain event/profile foreign keys. Delivery claims include event/profile, schedule version, occurrence, delivery kind, target, and lease ownership. Weekly roundup claims are profile-scoped and store each successful Discord part in `weekly_roundup_messages`.

## Delivery Guarantees

Claim uniqueness prevents duplicate database rows. Pollers claim transactionally with leases and `SKIP LOCKED`, so restarts, overlapping ticks, and multiple instances do not concurrently own one row. A sender revalidates event status, schedule version, target channel, profile, and final-announcement window before sending.

Migration runners acquire the same session-level PostgreSQL advisory lock before creating or reading `schema_migrations`. The lock remains held while pending files are checked, each migration and version record commit in one transaction, and final state is returned. It is released in `finally`, including failures. Simultaneous R.A.C.H.I.E and P.E.G.G.I.E startups therefore cannot both apply the same migration.

Sent history is never deleted by edits. Schedule-affecting changes increment `schedule_version` and terminally fail every unsent old claim. Grouped claims retain immutable group identity/name snapshots when group rows are replaced. State individual and exact-start claims are blocked by migration 005's policy constraint; historical sent rows remain readable.

Roundup configuration changes preserve sent claims, terminally fail unsent claims for that source guild/profile, and advance `weekly_roundup_not_before`. Claim generation accepts only a scheduled time after that boundary, preventing an enable or schedule change from replaying an elapsed roundup. Alliance and state roundup enablement are evaluated independently.

Discord send and database completion cannot be one transaction. If Discord accepts a message and the process exits before storing its ID, a retry may duplicate it. Multipart roundups reduce this window by storing each part immediately and skipping recorded parts.

## Security

Management uses the existing `userCanManageServer(interaction)` authorization. Interaction sessions bind opaque IDs to user, guild, and profile and expire after 15 minutes. Internal alliance and event IDs are not shown in public controls. Current-guild channels use Discord channel selectors and are revalidated for guild ownership, type and bot permissions before storage. State codes carry no credentials, are hashed at rest, expire after 15 minutes, work once, and are profile scoped.

Custom messages are trimmed, length-limited, plain text, checked for unsafe URL schemes, mention-neutralized, and sent with `allowedMentions.parse=[]`. Image downloads accept only trusted Discord attachment URLs, supported image signatures, and bounded sizes.
