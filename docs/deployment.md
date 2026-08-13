# Deployment

## Service Variables

Keep the two Railway services and Discord applications separate.

R.A.C.H.I.E:

```text
GAME_PROFILE=wos
BOT_INSTANCE_NAME=rachie-wos
EVENT_SCHEDULER_ENABLED=true
DATABASE_URL=<shared or dedicated Postgres URL>
```

P.E.G.G.I.E:

```text
GAME_PROFILE=kingshot
BOT_INSTANCE_NAME=peggie-kingshot
EVENT_SCHEDULER_ENABLED=true
DATABASE_URL=<shared or dedicated Postgres URL>
```

Optional on either profile:

```text
PLAYER_GIFT_CODES_ENABLED=true
BOT_OWNER_IDS=<comma-separated Discord user IDs>
GIFT_CODE_VERIFICATION_ENABLED=false
GIFT_CODE_REDEMPTION_WORKER_ENABLED=false
GIFT_CODE_SOURCE_POLLING_ENABLED=false
WOS_REWARDS_SOURCE_ENABLED=false
KINGSHOT_REWARDS_SOURCE_ENABLED=false
GIFT_CODE_SOURCE_POLL_INTERVAL_SECONDS=900
CENTURY_MINIMUM_DELAY_MS=1000
CENTURY_MAXIMUM_RETRIES=2
CENTURY_BACKOFF_BASE_MS=2000
CENTURY_BACKOFF_CAP_MS=60000
```

The adapters include signing suffixes observed in the current official public browser clients and independently validated against real official requests. These are public-client implementation details, not secrets or credentials, and may change. `CENTURY_WOS_SIGNING_SUFFIX` and `CENTURY_KINGSHOT_SIGNING_SUFFIX` are optional emergency compatibility overrides; a non-empty environment value takes precedence over the built-in value.

Verifier characters are independently optional per profile: `WOS_GIFT_VERIFY_FID` with `WOS_GIFT_VERIFY_KID`, or `KINGSHOT_GIFT_VERIFY_FID` with `KINGSHOT_GIFT_VERIFY_KID`. Never print these identifiers in normal logs. Enabling verification without a complete valid pair leaves candidates pending and reports the verifier as unavailable. No registered player is selected as an implicit verifier.

Each service keeps its own `BOT_TOKEN`, `CLIENT_ID`, `APPS_SCRIPT_URL`, `ADMIN_API_KEY`, and any `OPENAI_API_KEY`/`BANTER_PROFILE` configuration. Do not combine Apps Script deployments or booking sheets.

Every variable read by the repository is listed in [`.env.example`](../.env.example). Scheduler tuning variables are optional; invalid or out-of-range values fall back to defaults.

## Pre-Deployment Gate

1. Back up the scheduler database.
2. Run `npm install` from the committed lockfile.
3. Run `npm run check` and `npm test`.
4. Apply all migrations to a clean disposable PostgreSQL database.
5. Run the migration command a second time and verify `applied 0`.
6. Run the full suite with `TEST_DATABASE_URL` pointing only to that disposable database.
7. Confirm the concurrent migration-runner test passes: one runner applies and both finish with one valid migration record.
8. Confirm `.env` is ignored, `.env.example` contains placeholders, and `git status --short` is clean.
9. Confirm no test used live Discord, Railway, Apps Script, or production Postgres.

## Startup And Degraded Mode

`EVENT_SCHEDULER_ENABLED` and `PLAYER_GIFT_CODES_ENABLED` independently gate their subsystems. A Postgres pool is created only when at least one database-backed subsystem is enabled; otherwise `DATABASE_URL` is not required.

When enabled, scheduler initialization validates `GAME_PROFILE` and `BOT_INSTANCE_NAME`, checks Postgres, and runs migrations under an advisory lock. The lock encloses migration-state creation/checking, transactional application, and version recording, so simultaneous services serialize safely. A failure marks only database-backed scheduler management unavailable and closes its pool. Discord login and existing bot behavior continue; `/event-scheduler-help` remains available and the polling worker does not start.

## Safe Rollout

1. Deploy the committed code without changing Apps Script or existing bot variables.
2. Deploy one service first and inspect scheduler health/migration logs.
3. Verify `/event-scheduler` privately in a test guild/channel for that profile.
4. Verify `/event-scheduler-help` and the management **Help** control are private and complete.
5. Deploy the second service and verify it sees only its own profile data.
6. Configure the main alliance identity, select alliance channels natively, and add any sub-alliances.
7. For optional state sharing, configure the destination from inside the state Discord and consume its one-time code in the alliance Discord.
8. Create a test event far enough ahead to observe the chosen advance reminder and one-minute final announcement.
9. Verify no exact-start or individual state message appears.

Migration 006 includes an old-writer compatibility trigger: an overlapping older process that omits `alliance_id` is assigned the profile's default alliance. Migration 007 preserves existing channel/state-link IDs while adding identity-first setup, state destinations, and one-time link codes. Migration 008 preserves existing roundup channel/schedule values, backfills prior state-roundup enablement, and adds schedule-change replay protection. New event code always selects an explicit alliance.

## Rollback

Set `EVENT_SCHEDULER_ENABLED=false` on the affected service and redeploy/restart. This disables scheduler command registration and polling while preserving existing booking, state, administration, banter, and Apps Script-backed behavior.

Leave additive migrations applied. Do not drop alliance, custom-message, state-destination, link-code, or roundup-history data during an incident. Older code can continue default-alliance inserts because of the compatibility trigger and can read existing state links, but it cannot manage sub-alliances, custom messages, native destinations, link codes, or independent roundup controls.

If a release must be reverted, revert only application code, keep the database backup, and monitor scheduler logs. Re-enable the scheduler after the corrected application is deployed and migration state reports zero pending files.

## Railway Checks

- Confirm each service has the correct token/client pair and its own Apps Script URL.
- Confirm `BOT_INSTANCE_NAME` values are distinct.
- Confirm both services intentionally use the same or separate `DATABASE_URL`.
- Never place `TEST_DATABASE_URL` in a production service.
- Do not print tokens, URLs with credentials, or full environment dumps in logs.

## Infrastructure Portability

Railway is one deployment target, not an application dependency. `npm run migrate`, the bot processes, and future worker/API processes use ordinary environment variables and standard PostgreSQL connectivity. They can run under systemd, Docker, Docker Compose, or another scheduler against local, external, or managed PostgreSQL.

Do not introduce Railway service IDs, private hostnames, deployment APIs, or persistent local-file assumptions into business logic. Services should remain stateless outside PostgreSQL. Future cache/queue providers must be accessed through an abstraction so deployment can move without rewriting player, gift-code, or redemption rules.

## Gift-Code Rollout

1. Stage A: deploy migrations 012 through 016 and code with verification, redemption, and catalogue polling off. Migration 014 reconciles earlier applied 013 schemas; migration 016 adds source observations and channel configuration. Verify both bots and existing commands before enabling any external traffic.
2. Stage B: configure only the matching profile verifier, enable verification only, submit one known valid or already-redeemed code with `/gift-codes-admin` controlled verification, and inspect its classification and observed headers.
3. Stage C: register and opt in one controlled player, enable the redemption worker, and verify one code/account path including its DM result.
4. Stage D: expand to a small opted-in group while reviewing retries, unknown responses, and rate-limit observations.
5. Stage E: permit broader voluntary opt-in only after production evidence supports it.
6. Stage F: configure a mirrored source channel independently in each intended guild/profile. Enable one matching catalogue adapter only after fixture parsing has been reviewed against the current public page; monitor source health without treating catalogue observations as valid codes.

Command registration occurs during normal bot startup through the existing global Discord command registration call. Deploying commands live is a separate operational act and is not performed by tests. Keep `GIFT_CODE_VERIFICATION_ENABLED=false` and `GIFT_CODE_REDEMPTION_WORKER_ENABLED=false` until their respective rollout stages.
