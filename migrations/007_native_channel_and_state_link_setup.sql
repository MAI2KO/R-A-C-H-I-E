ALTER TABLE event_guild_settings
  ALTER COLUMN event_channel_id DROP NOT NULL;

CREATE TABLE event_state_destinations (
  state_guild_id varchar(32) NOT NULL,
  game_profile varchar(32) NOT NULL,
  configured_by_bot_instance varchar(100) NOT NULL,
  state_roundup_channel_id varchar(32) NOT NULL,
  enabled boolean NOT NULL DEFAULT true,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now(),

  PRIMARY KEY (state_guild_id, game_profile),

  CONSTRAINT event_state_destinations_game_profile_check
    CHECK (game_profile IN ('wos', 'kingshot')),

  CONSTRAINT event_state_destinations_bot_instance_check
    CHECK (btrim(configured_by_bot_instance) <> ''),

  CONSTRAINT event_state_destinations_channel_check
    CHECK (btrim(state_roundup_channel_id) <> '')
);

CREATE TABLE event_state_link_codes (
  id bigint GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
  game_profile varchar(32) NOT NULL,
  state_guild_id varchar(32) NOT NULL,
  code_hash char(64) NOT NULL UNIQUE,
  created_by_bot_instance varchar(100) NOT NULL,
  created_by_user_id varchar(32) NOT NULL,
  expires_at timestamptz NOT NULL,
  consumed_at timestamptz,
  consumed_by_alliance_guild_id varchar(32),
  created_at timestamptz NOT NULL DEFAULT now(),

  CONSTRAINT event_state_link_codes_destination_fk
    FOREIGN KEY (state_guild_id, game_profile)
    REFERENCES event_state_destinations (state_guild_id, game_profile)
    ON DELETE CASCADE,

  CONSTRAINT event_state_link_codes_game_profile_check
    CHECK (game_profile IN ('wos', 'kingshot')),

  CONSTRAINT event_state_link_codes_bot_instance_check
    CHECK (btrim(created_by_bot_instance) <> ''),

  CONSTRAINT event_state_link_codes_user_check
    CHECK (btrim(created_by_user_id) <> ''),

  CONSTRAINT event_state_link_codes_hash_check
    CHECK (code_hash ~ '^[0-9a-f]{64}$'),

  CONSTRAINT event_state_link_codes_expiry_check
    CHECK (expires_at > created_at),

  CONSTRAINT event_state_link_codes_consumption_check
    CHECK (
      (consumed_at IS NULL AND consumed_by_alliance_guild_id IS NULL)
      OR (
        consumed_at IS NOT NULL
        AND consumed_by_alliance_guild_id IS NOT NULL
        AND btrim(consumed_by_alliance_guild_id) <> ''
      )
    )
);

CREATE INDEX event_state_link_codes_lookup_idx
  ON event_state_link_codes (game_profile, code_hash, expires_at)
  WHERE consumed_at IS NULL;
