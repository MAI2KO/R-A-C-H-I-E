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

## Coverage Areas

- Apps Script-independent startup and scheduler degraded mode.
- Profile-scoped settings and WOS/Kingshot isolation.
- UTC parsing and deterministic occurrence arithmetic.
- Advance choices, one-minute final boundaries, and no exact-start claims.
- Custom-message normalization, isolation, mention safety, and standard details.
- Image validation and advance-only attachment behavior.
- Alliance-only individual delivery and terminal legacy-state reconciliation.
- Main/sub-alliance ownership, duplicate names, opaque controls, editing, and deletion blocks.
- Alliance and combined state roundup eligibility, ordering, splitting, and profile isolation.
- Transactional claiming, idempotency, leases, retries, stale-worker rejection, and multipart recovery.
- Migration 006 backfill and old-writer compatibility.

## Test Safety

Discord clients and channels are mocked. Attachment downloads use controlled mock responses. Database test IDs and credentials are synthetic. Tests must not register live commands, call Apps Script, send Discord messages, alter Railway, or connect to production Postgres.

Remove disposable containers/databases and inspect these commands before committing:

```bash
git status --short
git ls-files --others --exclude-standard
git diff --check
```

Confirm `.env` remains ignored and `.env.example` contains placeholders only.
