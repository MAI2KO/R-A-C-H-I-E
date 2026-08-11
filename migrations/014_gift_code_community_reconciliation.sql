ALTER TABLE gift_code_guild_settings
  ADD COLUMN IF NOT EXISTS contributor_role_status varchar(24),
  ADD COLUMN IF NOT EXISTS contributor_role_last_error varchar(200),
  ADD COLUMN IF NOT EXISTS contributor_role_claimed_by varchar(200),
  ADD COLUMN IF NOT EXISTS contributor_role_claimed_until_utc timestamptz;

ALTER TABLE gift_code_guild_settings
  ALTER COLUMN contributor_role_status SET DEFAULT 'unconfigured';

UPDATE gift_code_guild_settings
   SET contributor_role_status = CASE
     WHEN contributor_role_id IS NOT NULL THEN 'ready'
     ELSE 'unconfigured'
   END
 WHERE contributor_role_status IS NULL
    OR contributor_role_status NOT IN ('unconfigured', 'claiming', 'ready', 'error')
    OR (contributor_role_id IS NOT NULL AND contributor_role_status = 'unconfigured');

ALTER TABLE gift_code_guild_settings
  ALTER COLUMN contributor_role_status SET NOT NULL;

DO $reconcile_role_status$
BEGIN
  IF NOT EXISTS (
    SELECT 1
      FROM pg_constraint
     WHERE conrelid = 'gift_code_guild_settings'::regclass
       AND conname = 'gift_code_guild_settings_role_status_check'
  ) THEN
    ALTER TABLE gift_code_guild_settings
      ADD CONSTRAINT gift_code_guild_settings_role_status_check
      CHECK (contributor_role_status IN ('unconfigured', 'claiming', 'ready', 'error'))
      NOT VALID;
  END IF;
END
$reconcile_role_status$;

ALTER TABLE gift_code_guild_settings
  VALIDATE CONSTRAINT gift_code_guild_settings_role_status_check;

ALTER TABLE gift_code_engagement_events
  ADD COLUMN IF NOT EXISTS next_attempt_at_utc timestamptz;

DO $reconcile_pending_index$
DECLARE
  pending_index regclass := to_regclass('gift_code_engagement_pending_idx');
  pending_definition text;
BEGIN
  IF pending_index IS NOT NULL THEN
    pending_definition := pg_get_indexdef(pending_index);
    IF position('next_attempt_at_utc' IN pending_definition) = 0
       OR position('failed' IN pending_definition) = 0 THEN
      EXECUTE 'DROP INDEX gift_code_engagement_pending_idx';
    END IF;
  END IF;
END
$reconcile_pending_index$;

CREATE INDEX IF NOT EXISTS player_account_guilds_account_idx
  ON player_account_guilds (game_profile, player_account_id, guild_id);

CREATE UNIQUE INDEX IF NOT EXISTS gift_code_engagement_join_once_idx
  ON gift_code_engagement_events (game_profile, guild_id, discord_user_id, event_type)
  WHERE event_type = 'auto_redeem_join';

CREATE UNIQUE INDEX IF NOT EXISTS gift_code_engagement_code_once_idx
  ON gift_code_engagement_events (game_profile, guild_id, gift_code_id, event_type)
  WHERE event_type IN ('code_progress', 'contributor_role');

CREATE INDEX IF NOT EXISTS gift_code_engagement_pending_idx
  ON gift_code_engagement_events (game_profile, status, next_attempt_at_utc, created_at_utc, id)
  WHERE status IN ('pending', 'claimed', 'failed');

CREATE INDEX IF NOT EXISTS gift_code_engagement_progress_idx
  ON gift_code_engagement_events (game_profile, gift_code_id, last_update_at_utc, id)
  WHERE event_type = 'code_progress';
