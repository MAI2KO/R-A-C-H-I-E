ALTER TABLE player_accounts
  ALTER COLUMN discord_user_id DROP NOT NULL;

CREATE TABLE player_account_ownership_history (
  id uuid PRIMARY KEY,
  game_profile varchar(32) NOT NULL,
  player_account_id uuid NOT NULL,
  previous_discord_user_id varchar(32),
  new_discord_user_id varchar(32),
  action_type varchar(32) NOT NULL,
  performed_by_discord_user_id varchar(32) NOT NULL,
  source_metadata jsonb NOT NULL DEFAULT '{}'::jsonb,
  changed_at_utc timestamptz NOT NULL DEFAULT now(),

  CONSTRAINT player_account_ownership_history_account_fk
    FOREIGN KEY (player_account_id, game_profile)
    REFERENCES player_accounts (id, game_profile),
  CONSTRAINT player_account_ownership_history_profile_check
    CHECK (game_profile IN ('wos', 'kingshot')),
  CONSTRAINT player_account_ownership_history_previous_check
    CHECK (previous_discord_user_id IS NULL OR previous_discord_user_id ~ '^[0-9]{1,32}$'),
  CONSTRAINT player_account_ownership_history_new_check
    CHECK (new_discord_user_id IS NULL OR new_discord_user_id ~ '^[0-9]{1,32}$'),
  CONSTRAINT player_account_ownership_history_actor_check
    CHECK (performed_by_discord_user_id ~ '^[0-9]{1,32}$'),
  CONSTRAINT player_account_ownership_history_action_check
    CHECK (action_type IN ('release', 'operator_release', 'claim')),
  CONSTRAINT player_account_ownership_history_metadata_check
    CHECK (jsonb_typeof(source_metadata) = 'object')
);

CREATE INDEX player_account_ownership_history_account_idx
  ON player_account_ownership_history
    (game_profile, player_account_id, changed_at_utc, id);

ALTER TABLE gift_code_redemptions
  ADD COLUMN discord_owner_id_snapshot varchar(32);

UPDATE gift_code_redemptions r
   SET discord_owner_id_snapshot = a.discord_user_id
  FROM player_accounts a
 WHERE a.id = r.player_account_id
   AND a.game_profile = r.game_profile;

CREATE FUNCTION gift_code_redemptions_set_owner_snapshot()
RETURNS trigger AS $$
BEGIN
  IF NEW.discord_owner_id_snapshot IS NULL THEN
    SELECT discord_user_id INTO NEW.discord_owner_id_snapshot
      FROM player_accounts
     WHERE id = NEW.player_account_id AND game_profile = NEW.game_profile;
  END IF;
  RETURN NEW;
END;
$$ LANGUAGE plpgsql;

CREATE TRIGGER gift_code_redemptions_owner_snapshot_trigger
BEFORE INSERT ON gift_code_redemptions
FOR EACH ROW EXECUTE FUNCTION gift_code_redemptions_set_owner_snapshot();

ALTER TABLE gift_code_redemptions
  ALTER COLUMN discord_owner_id_snapshot SET NOT NULL,
  ADD CONSTRAINT gift_code_redemptions_owner_snapshot_check
    CHECK (discord_owner_id_snapshot ~ '^[0-9]{1,32}$');

CREATE INDEX gift_code_redemptions_owner_history_idx
  ON gift_code_redemptions
    (game_profile, discord_owner_id_snapshot, created_at_utc, id);

COMMENT ON COLUMN player_accounts.discord_user_id IS
  'Current global Discord owner within game_profile. NULL only after explicit release.';

COMMENT ON COLUMN gift_code_redemptions.discord_owner_id_snapshot IS
  'Discord owner when this immutable redemption record was created; never follows later ownership changes.';
