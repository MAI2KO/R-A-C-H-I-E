# Invalid player-account audit and cleanup

Migration `020_canonical_setup_and_player_identity.sql` added `in_game_name` and
`alliance_abbreviation` as nullable columns. That preserved active accounts created by the
older `/player-register` path, which required only Discord owner, Player ID, and State/Kingdom.
It did not invent metadata or deactivate those accounts. These are the expected source of
legacy accounts with missing name/alliance.

The original schema has always constrained Player ID and State/Kingdom to numeric values.
Discord owner was numeric and non-null until migration `018`; release then made it nullable,
but the release transaction also sets `is_active=false`. Consequently an active account with
an invalid location or no owner cannot be produced by the historical or current domain paths;
it indicates direct SQL, a failed/altered constraint, or external corruption and requires
manual review.

Current `/register` validates owner, Player ID, State/Kingdom, in-game name, and alliance before
the repository is called. Retired registration controls are not registered and are intercepted
before legacy dispatch. Migration `022` adds a `NOT VALID` active-identity check: PostgreSQL
enforces it for every new or updated row while allowing existing legacy-invalid rows to remain
available for explicit review.

## Safe procedure

Run independently in each deployment. `GAME_PROFILE=wos` audits WOS; `GAME_PROFILE=kingshot`
audits Kingshot. The configured website URL and integration secret must match that profile.

```bash
npm run audit:player-accounts -- --dry-run
```

The report contains aggregate reason counts and opaque 16-character account references. It
does not print Discord IDs or Player IDs. For each invalid account it asks the matching website
for counts of participant mirrors, bookings, approval/points history, and whether ownership is
safe. Preview uses the read-only signed endpoint and consumes no nonce.

Prefer asking a reachable owner to rerun `/register`; that supplies real metadata and preserves
the account. If an operator has independently confirmed that an invalid registration must be
released, execute exactly one opaque reference:

```bash
npm run audit:player-accounts -- --release=<account-ref> --operator=<operator-discord-user-id>
```

Execute first deactivates matching active website participants (`status='inactive'` and
`is_primary=false`) while retaining participant rows, bookings, approval records, points, and
foreign keys. A repeated request for the same inactive participant returns
`already_completed`; absence of any participant returns `safe_noop`. It then invokes the bot's
existing guarded `operator_release`: the bot row becomes
inactive and ownerless, pending gift work is disabled, guild enrollment is removed, ownership
history is recorded, and another primary is promoted if necessary. No booking, points, participant,
ownership-history, redemption-history, or audit-history row is hard-deleted. Any promoted valid
replacement is then synchronized to the website as authoritative MAIN.

The action refuses ambiguous ownership, changed account identity, multiple targets, or an unsafe
website response. An active bot row with no valid owner is deactivated only when the website
confirms there is no active participant mirror; otherwise it remains report-only for manual
review. The ownerless soft-deactivation records the operator, opaque account reference, reasons,
and action in account metadata.

## Failure and retry semantics

The two databases cannot commit atomically. The command therefore makes each side independently
idempotent and always performs website deactivation first:

- If website deactivation commits but its response is lost, or the subsequent bot transaction
  fails, the temporary state is an inactive website participant with the bot account still active.
  Rerun the same opaque reference. The website returns `already_completed`, then the guarded bot
  release runs normally.
- If the bot transaction commits but its database response is lost, or replacement MAIN mirroring
  fails afterward, the target is already inactive in both databases. The ownership-history row
  stores the cleanup reference (ownerless cleanup stores it in account metadata). Rerunning finds
  that durable marker, returns the website's idempotent result, does not execute a second release or
  ownership transition, and retries only the authoritative MAIN mirror.
- Concurrent website attempts lock the matching participant rows. Concurrent bot attempts lock the
  exact account row; only the first transaction can release it, and later attempts resolve the
  durable completion marker. Repeating an already-applied website update has a zero-row mutation.

Never switch to manual SQL to repair either temporary state. Rerun the same command and account
reference until it reports `botAccountReleased: true` and, when a replacement exists,
`replacementPrimaryMirrored: true`.
