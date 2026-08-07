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

`EVENT_SCHEDULER_ENABLED` is checked before `DATABASE_URL`. When it is not exactly `true`, the scheduler remains disabled and no Postgres pool is created.

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
