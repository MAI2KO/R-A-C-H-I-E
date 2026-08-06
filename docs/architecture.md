# Architecture

## Deployment Shape

The repository serves two independent Discord applications:

| Service | Game profile | Bot instance | Discord/App Script ownership |
| --- | --- | --- | --- |
| R.A.C.H.I.E | `wos` | `rachie-wos` | Its own token, client ID, Apps Script URL, deployment, and booking sheet |
| P.E.G.G.I.E | `kingshot` | `peggie-kingshot` | Its own token, client ID, Apps Script URL, deployment, and booking sheet |

The Railway services may share Postgres. `game_profile` is included in settings, events, alliances, claims, links, images, and every ownership query. A worker claims only its own profile. `BOT_INSTANCE_NAME` records and checks sender ownership; it is not a secret.

## Existing Apps Script Boundary

`index.js` retains the existing Apps Script client and actions for booking, sheet-backed state behavior, administration, announcements, and related commands. The scheduler neither calls those actions nor reuses their channel fields. Each bot keeps its own `APPS_SCRIPT_URL` and sheet.

Scheduler failure is isolated. Discord login and existing slash-command registration do not wait for Postgres. When enabled, scheduler initialization validates ownership, connects, runs migrations, and records health. Failure closes the scheduler pool and leaves the rest of the bot running.

## Scheduler Components

- `eventSchedulerInteractions.js`: authorised ephemeral entry point and guild/channel configuration.
- `allianceManagementInteractions.js`: opaque, profile-scoped alliance add/list/rename/delete flows.
- `eventCreationInteractions.js`: alliance selection, modal input, options, preview, and confirmation.
- `eventSchedulerRepository.js`: transactional event/alliance writes and claim reconciliation.
- `occurrenceCalculation.js`: deterministic UTC recurrence arithmetic.
- `eventDeliveryGeneration.js`: bounded claim generation for advance and one-minute final deliveries.
- `eventDeliveryRepository.js`: idempotent insertion, reconciliation, leases, and `FOR UPDATE SKIP LOCKED` claiming.
- `discordEventDelivery.js`: target/permission validation and alliance-only sends.
- `weeklyRoundup*`: UTC scheduling, selection, formatting, multipart persistence, and retry.

## Data Ownership

`event_guild_settings` is keyed by `(guild_id, game_profile)` and owns shared alliance reminder and weekly-roundup channels. `event_state_links` uses the same scope and supplies only the state weekly-roundup destination.

`event_alliances` gives a guild/profile one main alliance and any number of sub-alliances. Names are unique case-insensitively within that scope. The same name may exist in another guild or game profile.

`scheduled_events` stores `alliance_id`, `guild_id`, and `game_profile`. Its composite foreign key requires all three to match one alliance row. Legacy `alliance_name` remains denormalized for compatibility and is synchronized by a database trigger. Existing rows were backfilled by migration 006.

Event groups and images retain event/profile foreign keys. Delivery claims include event/profile, schedule version, occurrence, delivery kind, target, and lease ownership. Weekly roundup claims are profile-scoped and store each successful Discord part in `weekly_roundup_messages`.

## Delivery Guarantees

Claim uniqueness prevents duplicate database rows. Pollers claim transactionally with leases and `SKIP LOCKED`, so restarts, overlapping ticks, and multiple instances do not concurrently own one row. A sender revalidates event status, schedule version, target channel, profile, and final-announcement window before sending.

Sent history is never deleted by edits. Schedule-affecting changes increment `schedule_version` and terminally fail every unsent old claim. Grouped claims retain immutable group identity/name snapshots when group rows are replaced. State individual and exact-start claims are blocked by migration 005's policy constraint; historical sent rows remain readable.

Discord send and database completion cannot be one transaction. If Discord accepts a message and the process exits before storing its ID, a retry may duplicate it. Multipart roundups reduce this window by storing each part immediately and skipping recorded parts.

## Security

Management uses the existing `userCanManageServer(interaction)` authorization. Interaction sessions bind opaque IDs to user, guild, and profile and expire after 15 minutes. Internal alliance and event IDs are not shown in public controls.

Custom messages are trimmed, length-limited, plain text, checked for unsafe URL schemes, mention-neutralized, and sent with `allowedMentions.parse=[]`. Image downloads accept only trusted Discord attachment URLs, supported image signatures, and bounded sizes.
