ALTER TABLE player_accounts
  ADD CONSTRAINT player_accounts_active_identity_check
  CHECK (
    NOT is_active OR (
      discord_user_id IS NOT NULL
      AND discord_user_id ~ '^[0-9]{1,32}$'
      AND player_id ~ '^[0-9]{1,32}$'
      AND state_or_kingdom_number ~ '^[0-9]{1,10}$'
      AND in_game_name IS NOT NULL
      AND in_game_name = btrim(in_game_name)
      AND in_game_name <> ''
      AND in_game_name !~ '[[:cntrl:]]'
      AND alliance_abbreviation IS NOT NULL
      AND alliance_abbreviation ~ '^[A-Z0-9]{3}$'
    )
  ) NOT VALID;

COMMENT ON CONSTRAINT player_accounts_active_identity_check ON player_accounts IS
  'Enforced for new and updated rows; NOT VALID preserves legacy active rows until operator review.';
