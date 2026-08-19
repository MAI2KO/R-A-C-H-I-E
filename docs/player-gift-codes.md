# Player Accounts And Gift Codes

## Purpose And Scope

PostgreSQL is canonical for both profiles. Whiteout Survival uses **State** and Kingshot uses **Kingdom** in player-facing responses. Discord owns commands, PostgreSQL owns durable workflow state, and Apps Script booking identities remain unchanged.

The subsystem supports canonical player accounts, manual and configured-source candidate discovery, verification through an optional profile verifier, explicit per-account redemption opt-in, durable workers, private results, authorized diagnostics, and optional guild-scoped community updates. It does not import arbitrary page text, bypass Century controls, publish per-player redemption traffic, or build the future website.

## Data And Service Boundaries

```text
Discord user
  -> player_accounts (current character and current State/Kingdom)
       -> player_location_history (append-only transfers)
       -> gift_code_redemptions (one durable identity per player/code)
            -> gift_code_attempts (immutable API calls and location snapshots)

gift_code_sources
  -> gift_code_submissions
  -> gift_code_source_channels
  -> gift_code_source_observations
  -> gift_codes (exact case and verification state)

player_account_guilds
gift_code_guild_settings
  -> gift_code_engagement_events (restart-safe announcements, progress and rewards)
```

Player registration, location changes, primary selection, and soft removal remain transactional and profile scoped. A transfer updates current location and append-only history together. A redemption refreshes its Player ID and State/Kingdom snapshot when each attempt is claimed, so later transfers do not rewrite attempt history and future codes use the new location.

Business logic follows `Discord -> service -> repository/client`. Repositories, services, and worker processors are reusable by a future HTTP API, website, or independently runnable process.

## Discord Panels

`/player-register` opens a private account-management panel for registration, account selection, State/Kingdom changes, gift-code settings, and soft removal. `/gift-code-add code:<code>` is the quick ordinary-user submission path. `/gift-codes` keeps the private player-aware panel submission flow, per-account auto-redemption, and redemption history. `/gift-codes-admin` retains the existing server-management authorization and provides diagnostics, controlled verification, separate native announcement/source channel configuration, contributor-role health, queue status, source health, and community statistics. The former `/player`, `/gift`, and `/gift-admin` command trees are not registered.

The pre-existing `/register` command is separate: it collects alliance tag, in-game name, and Player ID, then calls the Apps Script `register_player_for_server` action used by booking/minister identity workflows. `/my-info`, `/unregister`, and booking continue to use that Apps Script record. `/player-register` overlaps in player identity only and is the intended future canonical surface, but merging the two stores is deferred to avoid changing working booking behavior.

## Submission And Opt-In

Any ordinary guild user can submit through `/gift-code-add` or `/gift-codes`; no administrator permission is required. Both call the same service. Submission trims surrounding whitespace but preserves exact code case. Every user/admin submission gets immutable provenance. Duplicate submissions reuse the existing candidate and do not create duplicate verification work; source trust never implies validity. New source adapters must enter through this same pipeline.

## Discovery Sources

All sources are discovery only: `source -> candidate ingestion -> case-sensitive deduplication -> existing Century verifier -> existing redemption pipeline`. No source can mark a code active, expired, or valid, and source priority is attribution/analytics metadata only. Automatic sources never receive the human contributor role. A later human duplicate of an automatically discovered code is not misreported as first discovery.

The generic Discord mirror adapter reads configured guild/profile source channels and conservatively recognises explicit `Code:`, `Gift Code:`, and `Redeem Code:` labels. It accepts ordinary or webhook-authored mirrored messages without needing the incoming webhook URL or token. Provenance records guild, channel, message, optional webhook ID, display name, source type, and original message time. Configure this with **Configure Source Channel** in `/gift-codes-admin`. This is separate from **Configure Channel**, which controls verified announcements. WOS and Kingshot source settings remain independent.

Explicit UTC expiry text is stored as source-reported expiry on the observation. It never overwrites the canonical Century status or canonical expiry. Earliest observation time is preserved, repeated message/catalogue observations are idempotent, and a catalogue disappearance is recorded without marking the code expired.

WOSRewards (`https://www.wosrewards.com/`) and KingshotRewards (`https://kingshotrewards.com/`) are secondary public catalogue adapters. Each running game-profile process has at most one logical poller, independent of guild count. Polling uses a normal bounded HTTP GET, ten-second timeout, descriptive User-Agent, and explicit active-entry parsing. Markup changes, oversized responses, timeouts, and repository failures are contained and reported as concise source health codes. Tests inject transport and make no website requests.

Catalogue polling defaults off. `GIFT_CODE_SOURCE_POLLING_ENABLED=true` is the global gate; the matching `WOS_REWARDS_SOURCE_ENABLED=true` or `KINGSHOT_REWARDS_SOURCE_ENABLED=true` gate must also be enabled. `GIFT_CODE_SOURCE_POLL_INTERVAL_SECONDS` defaults to `900` and is clamped to at least five minutes. Discord source-channel ingestion operates whenever the player/gift-code subsystem is running and does not require catalogue polling.

The admin panel shows mirror configuration, latest observations, catalogue enablement, last poll/success, current source presence, unresolved new candidates, and a sanitized last error. Current source presence counts each canonical code once across enabled profile sources whose observation has not disappeared; guild panels include only that guild's enabled mirrors, and a disabled catalogue or source does not contribute. New candidates are currently present source codes that have not begun their first verification attempt. Historical observations and source activity timestamps remain stored, and neither disappearance nor source disablement changes canonical code validity. Panel Refresh recalculates these database-derived values without polling external services. It never shows raw HTML. Facebook ingestion is intentionally deferred because its access behavior is unstable. Whiteout Wiki is only a possible future low-priority fallback because it can lag. Future social, Telegram, RSS, webhook-receiver, or website adapters should emit observations through the same provider-neutral ingestion service.

No Discord webhook credential is required or stored. If a webhook URL/token was exposed independently during development, rotate it in Discord; this feature cannot recover or use it.

Auto-redemption controls operate on one selected owned account at a time; they never change every character. Default is disabled. Enabling queues existing active codes for that account. Disabling or deactivating it disables unfinished work while retaining completed history. A transaction-scoped advisory lock enforces `GIFT_CODE_MAX_AUTO_REDEEM_ACCOUNTS_PER_USER` (default `2`) independently per Discord user and game profile, including concurrent enable attempts. Inactive and opted-out accounts do not consume a slot, and accounts beyond the cap are retained.

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

Production has verified a verifier-level `success` response and a player-level `already_redeemed` response. A player-level `success` has intentionally not been manufactured by repeatedly sourcing an account that has not claimed a known code. The next naturally occurring new code should provide that observation; no bypass or special test path exists for it.

Expired and invalid codes stop before fan-out. Invalid verifier information blocks the candidate without calling it invalid. Account-specific claim, level, eligibility, age, and same-type limits prove that Century recognises the code, so verifier results activate and fan out while an individual player's result remains terminal and restricted. Rate limits and simultaneous-action failures use durable bounded backoff. Unknown responses retain HTTP status, Century `code`, `err_code`, `msg`, profile, endpoint metadata, and timestamps; they never trigger fan-out or crash the loop.

## Century Client And Classification

Whiteout Survival and Kingshot adapters own their official URLs, independently verified current browser-client signing defaults, optional emergency overrides, and profile-specific response mappings. The signing values are public-client implementation details, not credentials. Logs never contain them or complete signing material.

Classification prioritizes `err_code`, then top-level code, then explicitly mapped stable messages, then `unknown_response`. Both adapters use their independently extracted official frontend maps; profile differences and conservative unresolved outcomes are documented rather than inferred.

The evidence-level protocol tables, retry decisions, and unresolved keys are maintained in [Century Gift-Code Response Protocol](century-response-protocol.md).

Century requests use the official profile endpoint, form encoding, public-client Origin and Referer, a descriptive application User-Agent, and a Unix timestamp in seconds. An unrecognized HTTP 401/403 is retained as `upstream_rejection`; it remains in manual review and cannot activate or fan out a code. Stored edge diagnostics are limited to HTTP status, response type, bounded Content-Type/Server and allowlisted CDN/rate-limit headers, plus a sanitized response summary of at most 2 KB. Cookies, authorization values, signing material, verifier identifiers, and raw HTML are not exposed through Discord.

For a single developer-machine comparison with production, use the explicitly gated harness documented in [Testing](testing.md). It performs one request through the same Whiteout adapter with retries disabled and is never part of the normal test suite or worker startup.

No CAPTCHA, anti-bot, proxy, IP rotation, concurrency, or rate-limit bypass exists. Tests inject transport and never contact Century.

## Retry And Rate Handling

One limiter serializes verification and redemption requests per profile/process. Workers disable the client's internal retries so each response becomes a durable attempt before another request. `Retry-After` is obeyed, HTTP 429 and temporary transport/server failures use exponential database backoff, and `GIFT_CODE_MAX_ATTEMPTS` bounds automatic work. Exhaustion, restrictions, unknown responses, and configuration problems become terminal review states rather than tight loops.

Bounded observations retain request/response timestamps, endpoint, HTTP status, `x-ratelimit-limit`, `x-ratelimit-remaining`, `x-ratelimit-reset`, and `retry-after`. Remaining capacity near zero increases spacing conservatively; only explicit `Retry-After` defines a server-requested pause. The observed limit is 30, but the reset interval and whether scope is fixed/rolling or IP/endpoint/account based remain unknown. Future analysis should use ordinary retained traffic and must not deliberately exhaust the limit.

PostgreSQL leases and `FOR UPDATE SKIP LOCKED` provide cross-process work ownership. A profile advisory lock plus one-row pacing state serializes Century request starts across overlapping same-profile processes. Stale claims and locks recover after crashes. Completed redemption rows are not claimable again, and uniqueness prevents a Railway/VPS restart from creating another player/code identity.

## Community Engagement

Migration 013 stores a profile-scoped activity channel and bot-managed contributor role per guild. Channel configuration uses Discord's native selector. The contributor role is created automatically when first earned, is named exactly `🍭`, has no permissions, is not mentionable, and is reused through its persisted ID. A database provisioning lease prevents overlapping processes from creating duplicate roles. Deleted roles are safely recovered; missing Manage Roles permission or hierarchy failures are retained as concise admin-visible health errors and retried later without changing verification or redemption results.

Migration 014 reconciles databases that recorded an earlier revision of migration 013 before contributor-role leasing and failed-event retry fields were finalized. It is idempotent and is a no-op for already-current columns and constraints.

The first auto-redemption enable for a user/guild can create one public join event. A genuinely new candidate that becomes active can create one announcement and one contributor-role event for its earliest qualifying submitter. Duplicate, invalid, expired, restricted, and unknown candidates do not qualify. Event uniqueness, durable claims, stored Discord message IDs, and deterministic Discord nonces make delivery restart-safe. Role assignment checks current membership and Discord role manageability.

Each active code uses one public aggregate progress message per guild. Updates are coalesced by count/time thresholds and edit the stored message, with a final update when no work remains. No public message is emitted for each player/API request. Reusable profile-scoped statistics include registered users/accounts, enabled accounts, successful and already-redeemed totals, monthly successes, verified codes, latest verified code, and unique contributors; account statistics are guild-filtered through durable account/guild links where applicable.

Public output may contain gift codes, user mentions, State/Kingdom, and aggregate counts. It never contains full Player IDs, raw Discord snowflakes as display text, signing material, Century raw responses, exact account errors, transfer history, or admin diagnostics.

## Notifications And Administration

Successful, already-redeemed, invalid-player, restricted, and unknown player results may produce one concise DM attempt. Notification ownership is recorded before Discord delivery. DM failure is stored separately and never changes or retries the Century result. Invalid player information marks account verification failed and asks the owner to confirm their current State/Kingdom without guessing or deactivating it.

The `/gift-codes-admin` panel uses the existing server-management authorization callback. It shows bounded counts and classifications, not verifier identifiers or player dumps. Controlled verification uses the same configured verifier, durable claim, idempotency, and rate limiter as automatic work; intentionally activating a candidate performs normal opt-in fan-out. Global Player ID recovery is separate: `/player-admin release` requires exact membership in the comma-separated `BOT_OWNER_IDS` allowlist and never accepts guild ownership, roles, Administrator, Manage Guild, or Apps Script authorization as a substitute.

## Deployment And Portability

Set `PLAYER_GIFT_CODES_ENABLED=true` with `GAME_PROFILE`, `BOT_INSTANCE_NAME`, and ordinary `DATABASE_URL`. Set `BOT_OWNER_IDS` to a comma-separated list of Discord user IDs allowed to perform global ownership recovery; unset, empty, and malformed values fail closed without disabling player self-release. `GIFT_CODE_VERIFICATION_ENABLED` and `GIFT_CODE_REDEMPTION_WORKER_ENABLED` independently default to false. Configure `WOS_GIFT_VERIFY_FID`/`WOS_GIFT_VERIFY_KID` or `KINGSHOT_GIFT_VERIFY_FID`/`KINGSHOT_GIFT_VERIFY_KID` only for the matching verifier profile.

Database, verifier, Century, or individual job failures remain contained to this subsystem. Booking, minister, scheduler, banter, and Apps Script behavior continue. Durable state stays in standard PostgreSQL; no Railway API or deployment-local persistent file is required. Follow the staged rollout in [Deployment](deployment.md) and never enable both workers automatically on first deployment.
