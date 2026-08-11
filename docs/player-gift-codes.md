# Player Accounts And Gift Codes

## Purpose And Scope

PostgreSQL is canonical for both profiles. Whiteout Survival uses **State** and Kingshot uses **Kingdom** in player-facing responses. Discord owns commands, PostgreSQL owns durable workflow state, and Apps Script booking identities remain unchanged.

The subsystem supports canonical player accounts, controlled candidate submission, verification through an optional profile verifier, explicit per-account redemption opt-in, durable workers, private results, and authorized diagnostics. It does not scrape, import unknown feeds, bypass Century controls, publish mass announcements, or build the future website.

## Data And Service Boundaries

```text
Discord user
  -> player_accounts (current character and current State/Kingdom)
       -> player_location_history (append-only transfers)
       -> gift_code_redemptions (one durable identity per player/code)
            -> gift_code_attempts (immutable API calls and location snapshots)

gift_code_sources
  -> gift_code_submissions
  -> gift_codes (exact case and verification state)
```

Player registration, location changes, primary selection, and soft removal remain transactional and profile scoped. A transfer updates current location and append-only history together. A redemption refreshes its Player ID and State/Kingdom snapshot when each attempt is claimed, so later transfers do not rewrite attempt history and future codes use the new location.

Business logic follows `Discord -> service -> repository/client`. Repositories, services, and worker processors are reusable by a future HTTP API, website, or independently runnable process.

## Submission And Opt-In

`/gift submit` trims surrounding whitespace but preserves exact code case. Every user/admin submission gets immutable provenance. Duplicate submissions reuse the existing candidate and do not create duplicate verification work; source trust never implies validity. New source adapters must enter through this same pipeline.

`/gift auto-enable`, `/gift auto-disable`, and `/gift status` operate on one owned account at a time. An omitted Player ID selects the current primary account; it never changes every character. Default is disabled. Enabling queues existing active codes for that account. Disabling or deactivating it disables unfinished work while retaining completed history.

## Controlled Workflow

```text
User/admin submission
        |
        v
gift_code_submissions -> candidate gift_code
        |
        v
verification worker
        |----> invalid / expired / restricted / unknown
        v
      active
        |
        v
active opted-in player redemption rows
        |
        v
serial rate-limited redemption worker
        |----> retry / player issue / restricted / unknown
        v
success / already redeemed -> private DM attempt
```

The optional verifier is configured independently per profile and is never selected from registered users. Verification `success` consumes the reward on that verifier; this is expected and is not bypassed. `already_redeemed` also proves Century recognises the code. Both outcomes activate the code and transactionally fan out idempotent work to active opted-in accounts.

Expired and invalid codes stop before fan-out. Invalid verifier information blocks the candidate without calling it invalid. Eligibility/redeeming limits remain restricted review states. Rate limits and temporary failures get durable backoff. Unknown responses retain HTTP status, Century `code`, `err_code`, `msg`, profile, endpoint metadata, and timestamps; they never trigger fan-out or crash the loop.

## Century Client And Classification

Whiteout Survival and Kingshot adapters own their official URLs, independently verified current browser-client signing defaults, optional emergency overrides, and profile-specific response mappings. The signing values are public-client implementation details, not credentials. Logs never contain them or complete signing material.

Classification prioritizes `err_code`, then top-level code, then explicitly mapped stable messages, then `unknown_response`. WOS observations currently include `20000` success, `40008` already received, `40007` expired, `40014` not found, and `40020` user information error. Kingshot currently maps only independently verified success semantics; other codes remain unknown until observed and reviewed.

No CAPTCHA, anti-bot, proxy, IP rotation, concurrency, or rate-limit bypass exists. Tests inject transport and never contact Century.

## Retry And Rate Handling

One limiter serializes verification and redemption requests per profile/process. Workers disable the client's internal retries so each response becomes a durable attempt before another request. `Retry-After` is obeyed, HTTP 429 and temporary transport/server failures use exponential database backoff, and `GIFT_CODE_MAX_ATTEMPTS` bounds automatic work. Exhaustion, restrictions, unknown responses, and configuration problems become terminal review states rather than tight loops.

Bounded observations retain request/response timestamps, endpoint, HTTP status, `x-ratelimit-limit`, `x-ratelimit-remaining`, `x-ratelimit-reset`, and `retry-after`. Remaining capacity near zero increases spacing conservatively; only explicit `Retry-After` defines a server-requested pause. The observed limit is 30, but the reset interval and whether scope is fixed/rolling or IP/endpoint/account based remain unknown. Future analysis should use ordinary retained traffic and must not deliberately exhaust the limit.

PostgreSQL leases and `FOR UPDATE SKIP LOCKED` provide cross-process work ownership. A profile advisory lock plus one-row pacing state serializes Century request starts across overlapping same-profile processes. Stale claims and locks recover after crashes. Completed redemption rows are not claimable again, and uniqueness prevents a Railway/VPS restart from creating another player/code identity.

## Notifications And Administration

Successful, already-redeemed, invalid-player, restricted, and unknown player results may produce one concise DM attempt. Notification ownership is recorded before Discord delivery. DM failure is stored separately and never changes or retries the Century result. Invalid player information marks account verification failed and asks the owner to confirm their current State/Kingdom without guessing or deactivating it.

`/gift-admin status`, `/gift-admin queue`, `/gift-admin code`, and `/gift-admin verify` use the existing server-management authorization callback. They show bounded counts and classifications, not verifier identifiers or player dumps. Controlled verification uses the same configured verifier, durable claim, idempotency, and rate limiter as automatic work; intentionally activating a candidate performs normal opt-in fan-out.

## Deployment And Portability

Set `PLAYER_GIFT_CODES_ENABLED=true` with `GAME_PROFILE`, `BOT_INSTANCE_NAME`, and ordinary `DATABASE_URL`. `GIFT_CODE_VERIFICATION_ENABLED` and `GIFT_CODE_REDEMPTION_WORKER_ENABLED` independently default to false. Configure `WOS_GIFT_VERIFY_FID`/`WOS_GIFT_VERIFY_KID` or `KINGSHOT_GIFT_VERIFY_FID`/`KINGSHOT_GIFT_VERIFY_KID` only for the matching verifier profile.

Database, verifier, Century, or individual job failures remain contained to this subsystem. Booking, minister, scheduler, banter, and Apps Script behavior continue. Durable state stays in standard PostgreSQL; no Railway API or deployment-local persistent file is required. Follow the staged rollout in [Deployment](deployment.md) and never enable both workers automatically on first deployment.
