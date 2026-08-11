CREATE TABLE player_accounts (
  id uuid NOT NULL,
  game_profile varchar(32) NOT NULL,
  discord_user_id varchar(32) NOT NULL,
  player_id varchar(32) NOT NULL,
  state_or_kingdom_number varchar(32) NOT NULL,
  display_name varchar(100),
  avatar_url text,
  furnace_or_town_level integer,
  account_metadata jsonb NOT NULL DEFAULT '{}'::jsonb,
  is_primary boolean NOT NULL DEFAULT false,
  is_active boolean NOT NULL DEFAULT true,
  gift_redemption_enabled boolean NOT NULL DEFAULT false,
  verification_status varchar(24) NOT NULL DEFAULT 'unverified',
  verification_error_code varchar(100),
  verification_error_message varchar(500),
  last_verified_at_utc timestamptz,
  created_at_utc timestamptz NOT NULL DEFAULT now(),
  updated_at_utc timestamptz NOT NULL DEFAULT now(),

  PRIMARY KEY (id),
  UNIQUE (id, game_profile),
  UNIQUE (game_profile, player_id),

  CONSTRAINT player_accounts_profile_check
    CHECK (game_profile IN ('wos', 'kingshot')),
  CONSTRAINT player_accounts_discord_user_check
    CHECK (discord_user_id ~ '^[0-9]{1,32}$'),
  CONSTRAINT player_accounts_player_id_check
    CHECK (player_id ~ '^[0-9]{1,32}$'),
  CONSTRAINT player_accounts_location_check
    CHECK (state_or_kingdom_number ~ '^[0-9]{1,10}$'),
  CONSTRAINT player_accounts_display_name_check
    CHECK (display_name IS NULL OR (display_name = btrim(display_name) AND display_name <> '')),
  CONSTRAINT player_accounts_level_check
    CHECK (furnace_or_town_level IS NULL OR furnace_or_town_level >= 0),
  CONSTRAINT player_accounts_metadata_check
    CHECK (jsonb_typeof(account_metadata) = 'object'),
  CONSTRAINT player_accounts_verification_check
    CHECK (verification_status IN ('unverified', 'pending', 'verified', 'failed'))
);

CREATE UNIQUE INDEX player_accounts_one_active_primary_idx
  ON player_accounts (game_profile, discord_user_id)
  WHERE is_primary = true AND is_active = true;

CREATE INDEX player_accounts_owner_idx
  ON player_accounts (game_profile, discord_user_id, is_active, created_at_utc, id);

CREATE TABLE player_location_history (
  id uuid PRIMARY KEY,
  player_account_id uuid NOT NULL,
  game_profile varchar(32) NOT NULL,
  previous_number varchar(32),
  new_number varchar(32) NOT NULL,
  changed_by_discord_user_id varchar(32),
  change_source varchar(32) NOT NULL,
  changed_at_utc timestamptz NOT NULL DEFAULT now(),

  CONSTRAINT player_location_history_account_fk
    FOREIGN KEY (player_account_id, game_profile)
    REFERENCES player_accounts (id, game_profile),
  CONSTRAINT player_location_history_profile_check
    CHECK (game_profile IN ('wos', 'kingshot')),
  CONSTRAINT player_location_history_previous_check
    CHECK (previous_number IS NULL OR previous_number ~ '^[0-9]{1,10}$'),
  CONSTRAINT player_location_history_new_check
    CHECK (new_number ~ '^[0-9]{1,10}$'),
  CONSTRAINT player_location_history_actor_check
    CHECK (changed_by_discord_user_id IS NULL OR changed_by_discord_user_id ~ '^[0-9]{1,32}$'),
  CONSTRAINT player_location_history_source_check
    CHECK (change_source IN (
      'user_command', 'admin', 'api_verification', 'website_future', 'migration'
    ))
);

CREATE INDEX player_location_history_account_idx
  ON player_location_history (game_profile, player_account_id, changed_at_utc, id);

CREATE TABLE gift_code_sources (
  id uuid NOT NULL,
  game_profile varchar(32) NOT NULL,
  source_type varchar(32) NOT NULL,
  source_name varchar(100) NOT NULL,
  source_reference text,
  trusted boolean NOT NULL DEFAULT false,
  enabled boolean NOT NULL DEFAULT true,
  metadata jsonb NOT NULL DEFAULT '{}'::jsonb,
  created_at_utc timestamptz NOT NULL DEFAULT now(),
  updated_at_utc timestamptz NOT NULL DEFAULT now(),

  PRIMARY KEY (id),
  UNIQUE (id, game_profile),
  CONSTRAINT gift_code_sources_profile_check
    CHECK (game_profile IN ('wos', 'kingshot')),
  CONSTRAINT gift_code_sources_type_check
    CHECK (source_type = btrim(source_type) AND source_type <> ''),
  CONSTRAINT gift_code_sources_name_check
    CHECK (source_name = btrim(source_name) AND source_name <> ''),
  CONSTRAINT gift_code_sources_metadata_check
    CHECK (jsonb_typeof(metadata) = 'object')
);

CREATE INDEX gift_code_sources_enabled_idx
  ON gift_code_sources (game_profile, enabled, source_type, source_name);

CREATE TABLE gift_codes (
  id uuid NOT NULL,
  game_profile varchar(32) NOT NULL,
  code varchar(128) NOT NULL,
  normalized_code varchar(128) NOT NULL,
  status varchar(24) NOT NULL DEFAULT 'candidate',
  first_seen_at_utc timestamptz NOT NULL DEFAULT now(),
  verified_at_utc timestamptz,
  expires_at_utc timestamptz,
  discovered_by_source_id uuid,
  last_api_code integer,
  last_err_code integer,
  last_api_message varchar(500),
  metadata jsonb NOT NULL DEFAULT '{}'::jsonb,
  created_at_utc timestamptz NOT NULL DEFAULT now(),
  updated_at_utc timestamptz NOT NULL DEFAULT now(),

  PRIMARY KEY (id),
  UNIQUE (id, game_profile),
  UNIQUE (game_profile, code),
  CONSTRAINT gift_codes_source_fk
    FOREIGN KEY (discovered_by_source_id, game_profile)
    REFERENCES gift_code_sources (id, game_profile),
  CONSTRAINT gift_codes_profile_check
    CHECK (game_profile IN ('wos', 'kingshot')),
  CONSTRAINT gift_codes_code_check
    CHECK (code = btrim(code) AND code <> ''),
  CONSTRAINT gift_codes_normalized_check
    CHECK (normalized_code = btrim(normalized_code) AND normalized_code <> ''),
  CONSTRAINT gift_codes_status_check
    CHECK (status IN (
      'candidate', 'verifying', 'active', 'expired', 'invalid',
      'restricted', 'unknown', 'disabled'
    )),
  CONSTRAINT gift_codes_metadata_check
    CHECK (jsonb_typeof(metadata) = 'object')
);

CREATE INDEX gift_codes_status_idx
  ON gift_codes (game_profile, status, first_seen_at_utc, id);

CREATE INDEX gift_codes_discovery_lookup_idx
  ON gift_codes (game_profile, normalized_code);

CREATE TABLE gift_code_submissions (
  id uuid PRIMARY KEY,
  game_profile varchar(32) NOT NULL,
  submitted_code varchar(128) NOT NULL,
  submitted_by_discord_user_id varchar(32),
  source_id uuid,
  received_at_utc timestamptz NOT NULL DEFAULT now(),
  duplicate_of_gift_code_id uuid,
  processing_status varchar(24) NOT NULL DEFAULT 'received',
  raw_metadata jsonb NOT NULL DEFAULT '{}'::jsonb,

  CONSTRAINT gift_code_submissions_source_fk
    FOREIGN KEY (source_id, game_profile)
    REFERENCES gift_code_sources (id, game_profile),
  CONSTRAINT gift_code_submissions_duplicate_fk
    FOREIGN KEY (duplicate_of_gift_code_id, game_profile)
    REFERENCES gift_codes (id, game_profile),
  CONSTRAINT gift_code_submissions_profile_check
    CHECK (game_profile IN ('wos', 'kingshot')),
  CONSTRAINT gift_code_submissions_code_check
    CHECK (submitted_code = btrim(submitted_code) AND submitted_code <> ''),
  CONSTRAINT gift_code_submissions_user_check
    CHECK (submitted_by_discord_user_id IS NULL OR submitted_by_discord_user_id ~ '^[0-9]{1,32}$'),
  CONSTRAINT gift_code_submissions_status_check
    CHECK (processing_status IN ('received', 'duplicate', 'pending_verification', 'processed', 'rejected')),
  CONSTRAINT gift_code_submissions_metadata_check
    CHECK (jsonb_typeof(raw_metadata) = 'object')
);

CREATE INDEX gift_code_submissions_pending_idx
  ON gift_code_submissions (game_profile, processing_status, received_at_utc, id);

CREATE TABLE gift_code_redemptions (
  id uuid NOT NULL,
  game_profile varchar(32) NOT NULL,
  gift_code_id uuid NOT NULL,
  player_account_id uuid NOT NULL,
  player_id_snapshot varchar(32) NOT NULL,
  location_number_snapshot varchar(32) NOT NULL,
  attempt_number integer NOT NULL DEFAULT 0,
  status varchar(24) NOT NULL DEFAULT 'queued',
  api_code integer,
  err_code integer,
  api_message varchar(500),
  http_status integer,
  attempted_at_utc timestamptz,
  completed_at_utc timestamptz,
  next_retry_at_utc timestamptz,
  retryable boolean NOT NULL DEFAULT false,
  response_metadata jsonb NOT NULL DEFAULT '{}'::jsonb,
  bot_instance_name varchar(100),
  claimed_by_worker varchar(200),
  claimed_at_utc timestamptz,
  claimed_until_utc timestamptz,
  created_at_utc timestamptz NOT NULL DEFAULT now(),
  updated_at_utc timestamptz NOT NULL DEFAULT now(),

  PRIMARY KEY (id),
  UNIQUE (id, game_profile),
  UNIQUE (game_profile, gift_code_id, player_account_id),
  CONSTRAINT gift_code_redemptions_code_fk
    FOREIGN KEY (gift_code_id, game_profile)
    REFERENCES gift_codes (id, game_profile),
  CONSTRAINT gift_code_redemptions_account_fk
    FOREIGN KEY (player_account_id, game_profile)
    REFERENCES player_accounts (id, game_profile),
  CONSTRAINT gift_code_redemptions_profile_check
    CHECK (game_profile IN ('wos', 'kingshot')),
  CONSTRAINT gift_code_redemptions_player_check
    CHECK (player_id_snapshot ~ '^[0-9]{1,32}$'),
  CONSTRAINT gift_code_redemptions_location_check
    CHECK (location_number_snapshot ~ '^[0-9]{1,10}$'),
  CONSTRAINT gift_code_redemptions_attempt_check
    CHECK (attempt_number >= 0),
  CONSTRAINT gift_code_redemptions_status_check
    CHECK (status IN (
      'queued', 'claimed', 'success', 'already_redeemed', 'expired',
      'invalid_code', 'invalid_player', 'restricted', 'rate_limited',
      'temporary_error', 'unknown', 'disabled'
    )),
  CONSTRAINT gift_code_redemptions_http_check
    CHECK (http_status IS NULL OR (http_status >= 100 AND http_status <= 599)),
  CONSTRAINT gift_code_redemptions_metadata_check
    CHECK (jsonb_typeof(response_metadata) = 'object')
);

CREATE INDEX gift_code_redemptions_work_idx
  ON gift_code_redemptions (game_profile, status, next_retry_at_utc, created_at_utc, id)
  WHERE status IN ('queued', 'rate_limited', 'temporary_error');

CREATE INDEX gift_code_redemptions_lease_idx
  ON gift_code_redemptions (game_profile, claimed_until_utc, id)
  WHERE status = 'claimed';

CREATE INDEX gift_code_redemptions_account_history_idx
  ON gift_code_redemptions (game_profile, player_account_id, created_at_utc, id);

COMMENT ON COLUMN gift_codes.code IS
  'Exact case-sensitive Century Games gift code. Never lowercase this value.';

COMMENT ON COLUMN gift_code_redemptions.location_number_snapshot IS
  'State or Kingdom number at queue creation; remains unchanged after player transfer.';
