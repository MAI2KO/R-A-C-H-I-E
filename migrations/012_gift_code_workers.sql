ALTER TABLE gift_codes
  ADD COLUMN verification_state varchar(24) NOT NULL DEFAULT 'pending',
  ADD COLUMN verification_attempt_count integer NOT NULL DEFAULT 0,
  ADD COLUMN verification_next_retry_at_utc timestamptz,
  ADD COLUMN verification_claimed_by_worker varchar(200),
  ADD COLUMN verification_claimed_at_utc timestamptz,
  ADD COLUMN verification_claimed_until_utc timestamptz,
  ADD COLUMN verification_http_status integer,
  ADD COLUMN verification_metadata jsonb NOT NULL DEFAULT '{}'::jsonb,
  ADD COLUMN last_verification_at_utc timestamptz;

UPDATE gift_codes
   SET verification_state = CASE
     WHEN status = 'unknown' THEN 'review'
     WHEN status IN ('active', 'expired', 'invalid', 'restricted', 'disabled') THEN 'complete'
     ELSE 'pending'
   END;

ALTER TABLE gift_codes
  ADD CONSTRAINT gift_codes_verification_state_check
    CHECK (verification_state IN ('pending', 'claimed', 'retry', 'complete', 'blocked', 'review')),
  ADD CONSTRAINT gift_codes_verification_attempt_check
    CHECK (verification_attempt_count >= 0),
  ADD CONSTRAINT gift_codes_verification_http_check
    CHECK (verification_http_status IS NULL OR verification_http_status BETWEEN 100 AND 599),
  ADD CONSTRAINT gift_codes_verification_metadata_check
    CHECK (jsonb_typeof(verification_metadata) = 'object');

CREATE INDEX gift_codes_verification_work_idx
  ON gift_codes (
    game_profile, verification_state, verification_next_retry_at_utc,
    first_seen_at_utc, id
  )
  WHERE verification_state IN ('pending', 'retry', 'claimed');

ALTER TABLE gift_code_redemptions
  DROP CONSTRAINT gift_code_redemptions_status_check,
  ADD CONSTRAINT gift_code_redemptions_status_check
    CHECK (status IN (
      'queued', 'claimed', 'success', 'already_redeemed', 'expired',
      'invalid_code', 'invalid_player', 'restricted', 'rate_limited',
      'temporary_error', 'retry_exhausted', 'unknown', 'disabled'
    )),
  ADD COLUMN notification_status varchar(24) NOT NULL DEFAULT 'pending',
  ADD COLUMN notification_attempted_at_utc timestamptz,
  ADD COLUMN notified_at_utc timestamptz,
  ADD COLUMN notification_error varchar(200),
  ADD CONSTRAINT gift_code_redemptions_notification_status_check
    CHECK (notification_status IN ('pending', 'sending', 'sent', 'failed', 'suppressed'));

CREATE TABLE gift_code_attempts (
  id uuid NOT NULL,
  game_profile varchar(32) NOT NULL,
  gift_code_id uuid NOT NULL,
  redemption_id uuid,
  attempt_type varchar(24) NOT NULL,
  attempt_number integer NOT NULL,
  player_id_snapshot varchar(32),
  location_number_snapshot varchar(32),
  classification varchar(40) NOT NULL,
  api_code integer,
  err_code integer,
  api_message varchar(500),
  http_status integer,
  response_metadata jsonb NOT NULL DEFAULT '{}'::jsonb,
  request_started_at_utc timestamptz NOT NULL,
  response_received_at_utc timestamptz NOT NULL,
  created_at_utc timestamptz NOT NULL DEFAULT now(),

  PRIMARY KEY (id),
  UNIQUE (id, game_profile),
  CONSTRAINT gift_code_attempts_code_fk
    FOREIGN KEY (gift_code_id, game_profile)
    REFERENCES gift_codes (id, game_profile),
  CONSTRAINT gift_code_attempts_redemption_fk
    FOREIGN KEY (redemption_id, game_profile)
    REFERENCES gift_code_redemptions (id, game_profile),
  CONSTRAINT gift_code_attempts_profile_check
    CHECK (game_profile IN ('wos', 'kingshot')),
  CONSTRAINT gift_code_attempts_type_check
    CHECK (attempt_type IN ('verification', 'redemption')),
  CONSTRAINT gift_code_attempts_number_check
    CHECK (attempt_number > 0),
  CONSTRAINT gift_code_attempts_player_check
    CHECK (player_id_snapshot IS NULL OR player_id_snapshot ~ '^[0-9]{1,32}$'),
  CONSTRAINT gift_code_attempts_location_check
    CHECK (location_number_snapshot IS NULL OR location_number_snapshot ~ '^[0-9]{1,10}$'),
  CONSTRAINT gift_code_attempts_http_check
    CHECK (http_status IS NULL OR http_status BETWEEN 100 AND 599),
  CONSTRAINT gift_code_attempts_metadata_check
    CHECK (jsonb_typeof(response_metadata) = 'object'),
  CONSTRAINT gift_code_attempts_target_check
    CHECK (
      (attempt_type = 'verification' AND redemption_id IS NULL)
      OR
      (attempt_type = 'redemption' AND redemption_id IS NOT NULL
        AND player_id_snapshot IS NOT NULL AND location_number_snapshot IS NOT NULL)
    )
);

CREATE UNIQUE INDEX gift_code_attempts_verification_number_idx
  ON gift_code_attempts (game_profile, gift_code_id, attempt_number)
  WHERE attempt_type = 'verification';

CREATE UNIQUE INDEX gift_code_attempts_redemption_number_idx
  ON gift_code_attempts (game_profile, redemption_id, attempt_number)
  WHERE redemption_id IS NOT NULL;

CREATE INDEX gift_code_attempts_code_history_idx
  ON gift_code_attempts (game_profile, gift_code_id, created_at_utc, id);

CREATE INDEX gift_code_attempts_redemption_history_idx
  ON gift_code_attempts (game_profile, redemption_id, created_at_utc, id)
  WHERE redemption_id IS NOT NULL;

CREATE TABLE gift_code_rate_limit_state (
  game_profile varchar(32) PRIMARY KEY,
  last_request_started_at_utc timestamptz,
  updated_at_utc timestamptz NOT NULL DEFAULT now(),

  CONSTRAINT gift_code_rate_limit_state_profile_check
    CHECK (game_profile IN ('wos', 'kingshot'))
);

COMMENT ON COLUMN gift_code_redemptions.location_number_snapshot IS
  'State or Kingdom number refreshed from the account when an attempt is claimed; remains historical after that attempt.';
