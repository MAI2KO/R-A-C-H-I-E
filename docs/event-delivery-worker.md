# Event Delivery Worker

Phase 5 provides durable delivery generation and worker infrastructure without
registering a production delivery handler. `index.js` does not start the worker.

## Delivery Semantics

- Claim creation is idempotent through the database uniqueness index.
- PostgreSQL row locking and leases allow at most one active worker to own a
  claim at a time.
- Completion and failure updates succeed only while the same game profile, bot
  instance and opaque worker ID still own the claim.
- Expired leases are reclaimed through the normal transactional claim query.
- Delivery is at least once after an ambiguous process crash, not exactly once.

There is a narrow unavoidable ambiguity when an external service accepts a
message and the process exits before the database records `status='sent'`.
After the lease expires, another attempt may send a duplicate. Database
idempotency prevents duplicate claim rows and concurrent sends, but it cannot
provide idempotency for an external Discord send. Phase 6 must preserve this
documented limitation when it supplies the real handler.

## Windows And Retries

Claims are generated for a half-open delivery window from the grace boundary
through the configured lookahead boundary. Defaults are 60 minutes behind and
1,440 minutes ahead. Occurrences are calculated from their original UTC anchor.

Attempts are incremented when a claim is acquired. Retryable failures use
delays of 1, 5, 15, 30 and 60 minutes for attempts one through five. A permanent
failure, or a fifth failed attempt, remains failed without another retry time.
