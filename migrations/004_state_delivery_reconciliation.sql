CREATE INDEX event_delivery_claims_state_unsent_idx
  ON event_delivery_claims (
    game_profile,
    status,
    event_id,
    target_guild_id,
    target_channel_id
  )
  WHERE target_kind = 'state' AND status IN ('pending', 'failed', 'claimed');
