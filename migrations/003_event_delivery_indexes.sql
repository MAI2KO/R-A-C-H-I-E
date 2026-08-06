CREATE INDEX event_delivery_claims_pending_due_idx
  ON event_delivery_claims (game_profile, deliver_at, id)
  WHERE status = 'pending';

CREATE INDEX event_delivery_claims_failed_due_idx
  ON event_delivery_claims (game_profile, next_attempt_at, id)
  WHERE status = 'failed' AND next_attempt_at IS NOT NULL;

CREATE INDEX event_delivery_claims_expired_lease_idx
  ON event_delivery_claims (game_profile, claimed_until, id)
  WHERE status = 'claimed';
