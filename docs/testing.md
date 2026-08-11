# Testing

## Fast Checks

```bash
npm run check
npm test
git diff --check
```

Without `TEST_DATABASE_URL`, database suites report as skipped at their inner test level while all parsing, calculation, formatting, interaction, worker, and mocked Discord tests run locally.

## Disposable PostgreSQL

Use PostgreSQL 16 or another supported disposable Postgres instance. Never use Railway or production credentials.

```bash
EVENT_SCHEDULER_ENABLED=true \
DATABASE_URL=postgresql://user:password@127.0.0.1:5432/disposable \
npm run migrate

TEST_DATABASE_URL=postgresql://user:password@127.0.0.1:5432/disposable \
npm test
```

Run `npm run migrate` again with the same scheduler variables and require `applied 0`. Recreate the database from empty whenever a not-yet-deployed migration changes during development.

The database suite starts two `runMigrations` calls concurrently against one isolated schema and a deliberately slow probe migration. The advisory lock must allow exactly one application; the other runner must finish after observing the recorded version. The test then requires one migration row and one probe row.

## Coverage Areas

- Apps Script-independent startup and scheduler degraded mode.
- Profile-scoped settings and WOS/Kingshot isolation.
- UTC parsing and deterministic occurrence arithmetic.
- Advance choices, one-minute final boundaries, and no exact-start claims.
- Guided single/group timing, shared group times, structured add/edit/remove, and case-insensitive group-name validation.
- Ordinary edit ownership/image retention, separate alliance changes, and explicit image actions.
- Concise reminder/custom-message isolation, mention safety, and absence of schedule fields.
- Image validation and advance-only attachment behavior.
- Alliance-only individual delivery and terminal legacy-state reconciliation.
- Main/sub-alliance ownership, duplicate names, opaque controls, editing, and deletion blocks.
- Native reminder, alliance-roundup, and state-destination channel selectors; current-guild ownership; channel type; and permission validation.
- Hashed, expiring, one-time, profile-scoped state link codes and preservation of existing ID-backed links.
- Reminder cancellation through editing, custom-message clearing, and sent-history preservation.
- Alliance and combined state roundup eligibility, independent enablement, editable UTC schedules, replay protection, ordering, splitting, and profile isolation.
- Transactional claiming, idempotency, leases, retries, stale-worker rejection, and multipart recovery.
- Migrations 006/007/008 backfill, old-writer, stored-channel/state-link compatibility, and roundup-setting preservation.
- Ephemeral scheduler help content, controls, limits, degraded-mode availability, and accurate state policy.
- Concurrent migration-runner serialization and exactly-once migration application.

## Test Safety

Discord clients and channels are mocked. Attachment downloads use controlled mock responses. Database test IDs and credentials are synthetic. Tests must not register live commands, call Apps Script, send Discord messages, alter Railway, or connect to production Postgres.

Remove disposable containers/databases and inspect these commands before committing:

```bash
git status --short
git ls-files --others --exclude-standard
git diff --check
```

Confirm `.env` remains ignored and `.env.example` contains placeholders only.

## One Controlled Century Comparison

Normal tests never contact Century Games. To issue exactly one Whiteout verification request from a developer machine, use the dedicated harness with all opt-ins explicit:

```bash
ALLOW_ONE_LIVE_CENTURY_REQUEST=true \
GAME_PROFILE=wos \
LIVE_CENTURY_GIFT_CODE=gogoWOS \
WOS_GIFT_VERIFY_FID=<controlled verifier FID> \
WOS_GIFT_VERIFY_KID=<controlled verifier State> \
npm run test:live-century:wos
```

The harness has no database or Discord dependency, fixes retries at zero, invokes the shared adapter once, and prints only bounded sanitized diagnostics. Do not put the opt-in variable in production configuration and do not repeat the request merely to gather more samples.
