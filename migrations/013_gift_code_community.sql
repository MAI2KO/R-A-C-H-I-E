CREATE TABLE player_account_guilds (
  game_profile varchar(32) NOT NULL,
  guild_id varchar(32) NOT NULL,
  player_account_id uuid NOT NULL,
  linked_at_utc timestamptz NOT NULL DEFAULT now(),

  PRIMARY KEY (game_profile, guild_id, player_account_id),
  CONSTRAINT player_account_guilds_account_fk
    FOREIGN KEY (player_account_id, game_profile)
    REFERENCES player_accounts (id, game_profile),
  CONSTRAINT player_account_guilds_profile_check
    CHECK (game_profile IN ('wos', 'kingshot')),
  CONSTRAINT player_account_guilds_guild_check
    CHECK (guild_id ~ '^[0-9]{1,32}$')
);

CREATE INDEX player_account_guilds_account_idx
  ON player_account_guilds (game_profile, player_account_id, guild_id);

CREATE TABLE gift_code_guild_settings (
  game_profile varchar(32) NOT NULL,
  guild_id varchar(32) NOT NULL,
  gift_code_channel_id varchar(32),
  contributor_role_id varchar(32),
  contributor_role_status varchar(24) NOT NULL DEFAULT 'unconfigured',
  contributor_role_last_error varchar(200),
  contributor_role_claimed_by varchar(200),
  contributor_role_claimed_until_utc timestamptz,
  announcements_enabled boolean NOT NULL DEFAULT true,
  created_at_utc timestamptz NOT NULL DEFAULT now(),
  updated_at_utc timestamptz NOT NULL DEFAULT now(),

  PRIMARY KEY (game_profile, guild_id),
  CONSTRAINT gift_code_guild_settings_profile_check
    CHECK (game_profile IN ('wos', 'kingshot')),
  CONSTRAINT gift_code_guild_settings_guild_check
    CHECK (guild_id ~ '^[0-9]{1,32}$'),
  CONSTRAINT gift_code_guild_settings_channel_check
    CHECK (gift_code_channel_id IS NULL OR gift_code_channel_id ~ '^[0-9]{1,32}$'),
  CONSTRAINT gift_code_guild_settings_role_check
    CHECK (contributor_role_id IS NULL OR contributor_role_id ~ '^[0-9]{1,32}$'),
  CONSTRAINT gift_code_guild_settings_role_status_check
    CHECK (contributor_role_status IN ('unconfigured', 'claiming', 'ready', 'error'))
);

CREATE TABLE gift_code_engagement_events (
  id uuid NOT NULL,
  game_profile varchar(32) NOT NULL,
  guild_id varchar(32) NOT NULL,
  event_type varchar(32) NOT NULL,
  gift_code_id uuid,
  player_account_id uuid,
  discord_user_id varchar(32),
  status varchar(24) NOT NULL DEFAULT 'pending',
  channel_id varchar(32),
  message_id varchar(32),
  claimed_by_worker varchar(200),
  claimed_at_utc timestamptz,
  claimed_until_utc timestamptz,
  attempt_count integer NOT NULL DEFAULT 0,
  last_progress_count integer NOT NULL DEFAULT 0,
  progress_successful integer NOT NULL DEFAULT 0,
  progress_already_redeemed integer NOT NULL DEFAULT 0,
  progress_account_issues integer NOT NULL DEFAULT 0,
  progress_restricted integer NOT NULL DEFAULT 0,
  progress_remaining integer NOT NULL DEFAULT 0,
  last_update_at_utc timestamptz,
  finalized_at_utc timestamptz,
  last_error varchar(200),
  next_attempt_at_utc timestamptz,
  metadata jsonb NOT NULL DEFAULT '{}'::jsonb,
  created_at_utc timestamptz NOT NULL DEFAULT now(),
  updated_at_utc timestamptz NOT NULL DEFAULT now(),

  PRIMARY KEY (id),
  UNIQUE (id, game_profile),
  CONSTRAINT gift_code_engagement_events_code_fk
    FOREIGN KEY (gift_code_id, game_profile)
    REFERENCES gift_codes (id, game_profile),
  CONSTRAINT gift_code_engagement_events_account_fk
    FOREIGN KEY (player_account_id, game_profile)
    REFERENCES player_accounts (id, game_profile),
  CONSTRAINT gift_code_engagement_events_profile_check
    CHECK (game_profile IN ('wos', 'kingshot')),
  CONSTRAINT gift_code_engagement_events_guild_check
    CHECK (guild_id ~ '^[0-9]{1,32}$'),
  CONSTRAINT gift_code_engagement_events_user_check
    CHECK (discord_user_id IS NULL OR discord_user_id ~ '^[0-9]{1,32}$'),
  CONSTRAINT gift_code_engagement_events_type_check
    CHECK (event_type IN ('auto_redeem_join', 'code_progress', 'contributor_role')),
  CONSTRAINT gift_code_engagement_events_status_check
    CHECK (status IN ('pending', 'claimed', 'completed', 'failed', 'disabled')),
  CONSTRAINT gift_code_engagement_events_attempt_check
    CHECK (
      attempt_count >= 0 AND last_progress_count >= 0
      AND progress_successful >= 0 AND progress_already_redeemed >= 0
      AND progress_account_issues >= 0 AND progress_restricted >= 0
      AND progress_remaining >= 0
    ),
  CONSTRAINT gift_code_engagement_events_metadata_check
    CHECK (jsonb_typeof(metadata) = 'object')
);

CREATE UNIQUE INDEX gift_code_engagement_join_once_idx
  ON gift_code_engagement_events (game_profile, guild_id, discord_user_id, event_type)
  WHERE event_type = 'auto_redeem_join';

CREATE UNIQUE INDEX gift_code_engagement_code_once_idx
  ON gift_code_engagement_events (game_profile, guild_id, gift_code_id, event_type)
  WHERE event_type IN ('code_progress', 'contributor_role');

CREATE INDEX gift_code_engagement_pending_idx
  ON gift_code_engagement_events (game_profile, status, next_attempt_at_utc, created_at_utc, id)
  WHERE status IN ('pending', 'claimed', 'failed');

CREATE INDEX gift_code_engagement_progress_idx
  ON gift_code_engagement_events (game_profile, gift_code_id, last_update_at_utc, id)
  WHERE event_type = 'code_progress';
