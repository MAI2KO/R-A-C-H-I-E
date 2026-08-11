# Player Accounts And Gift Codes

## Purpose And Scope

PostgreSQL is the canonical identity and gift-code store for both bot profiles. Whiteout Survival uses **State** and Kingshot uses **Kingdom** in every player-facing response. One Discord user can own multiple characters in one or both profiles; Discord is an external owner identity, not the character's primary key.

This phase implements `/player register`, `/player view`, `/player location`, and `/player remove`. The existing Apps Script-backed `/register`, `/my-info`, and `/unregister` booking commands are unchanged. Automatic redemption, scraping, global announcements, website identity, and bulk workers are intentionally absent.

## Data Model

```text
Discord user
  -> player_accounts (current character and current State/Kingdom)
       -> player_location_history (append-only transfers)
       -> gift_code_redemptions (immutable player/location snapshots)

gift_code_sources
  -> gift_code_submissions
  -> gift_codes (exact case-sensitive code and verification state)
       -> gift_code_redemptions
```

`player_accounts` uses UUID character identities and profile-scoped Player ID uniqueness. One partial unique index permits one active primary account per Discord user/profile. Registration makes the first active account primary. Removing an account is a soft deactivation; removing the primary promotes the oldest active account by creation time and UUID.

A location change locks the owned active account, updates its current number, and inserts history in the same transaction. Failed history insertion rolls back the account update. Redemption rows snapshot Player ID and State/Kingdom, so later transfers never rewrite audit history.

## Gift-Code Provenance And Idempotency

Gift codes preserve exact case in `gift_codes.code`; no code sent to Century is lowercased. Sources and submissions retain provenance, but `trusted=true` never marks a code verified. Every candidate remains eligible for a future official Century verifier.

One unique redemption row exists per profile, gift code, and player account. Its attempt count, status, retry schedule, and lease fields support restart/reclaim behavior without creating duplicate work. No worker consumes those rows in this phase.

## Century Client

The shared client uses small Whiteout Survival and Kingshot adapters for official frontend/API URLs. Their built-in signing suffixes were observed in the current official public browser clients and independently validated against real official requests. They are shipped to browsers as public-client implementation details, not secrets or credentials, and may change. Non-empty `CENTURY_WOS_SIGNING_SUFFIX` and `CENTURY_KINGSHOT_SIGNING_SUFFIX` environment values override the defaults for emergency compatibility.

Signing sorts and URL-encodes unsigned fields, appends the profile suffix, and computes MD5 once. The exact gift-code case is preserved. Logs must not include the suffix or complete signing material.

The response classifier recognizes observed success, already-received, expired, missing-code, and player-error responses without treating that set as exhaustive. Unknown `code`, `err_code`, and `msg` values are returned and can be persisted by a future worker.

No CAPTCHA, anti-bot, or rate-limit bypass exists. Unit tests inject transport and never contact Century.

## Conservative Rate Handling

Each profile limiter runs one request at a time. Defaults are a 1-second minimum spacing, two retries, 2-second exponential backoff, and a 60-second cap. `Retry-After` takes precedence when longer. HTTP 429 and temporary failures can retry; permanent classifications do not.

The observed public limit is 30, but its reset interval and scope are **UNKNOWN**. The limiter captures HTTP status, `x-ratelimit-limit`, `x-ratelimit-remaining`, `x-ratelimit-reset`, `retry-after`, and request/response timestamps in a bounded 100-entry in-memory diagnostic buffer. This avoids unbounded database growth. A future worker may emit these records as structured logs with retention.

## Google Sheet Compatibility

`playerMirror.js` defines the optional post-commit mirror boundary. The repository contains no confirmed Apps Script action for safely mirroring canonical Player IDs into the existing bot-users sheet, so the default implementation is intentionally a no-op. A future adapter must document the exact Apps Script contract, authenticate through existing environment configuration, log safe retryable failures, and never make Sheet success part of the PostgreSQL transaction.

## Deployment And Portability

Set `PLAYER_GIFT_CODES_ENABLED=true` with `GAME_PROFILE`, `BOT_INSTANCE_NAME`, and an ordinary `DATABASE_URL`. Initialization runs the same portable advisory-locked migrations used elsewhere. Database failure disables only `/player`; booking, minister, banter, and other Apps Script features continue.

Repositories, the Century client, and limiter do not depend on Discord or Railway. They can be imported by a website/API or independently runnable worker process on Linux, Docker Compose, or another platform using standard PostgreSQL. No durable files are written to deployment-local storage.
