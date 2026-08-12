ALTER TABLE gift_code_sources
  ADD COLUMN IF NOT EXISTS priority integer NOT NULL DEFAULT 100,
  ADD COLUMN IF NOT EXISTS last_observation_at_utc timestamptz,
  ADD COLUMN IF NOT EXISTS last_candidate_at_utc timestamptz,
  ADD COLUMN IF NOT EXISTS last_poll_at_utc timestamptz,
  ADD COLUMN IF NOT EXISTS last_successful_poll_at_utc timestamptz,
  ADD COLUMN IF NOT EXISTS last_error varchar(300),
  ADD COLUMN IF NOT EXISTS observations_count bigint NOT NULL DEFAULT 0,
  ADD COLUMN IF NOT EXISTS candidates_count bigint NOT NULL DEFAULT 0;

CREATE TABLE IF NOT EXISTS gift_code_source_channels (
  game_profile varchar(32) NOT NULL,
  guild_id varchar(32) NOT NULL,
  channel_id varchar(32) NOT NULL,
  source_id uuid NOT NULL,
  enabled boolean NOT NULL DEFAULT true,
  require_webhook boolean NOT NULL DEFAULT false,
  created_at_utc timestamptz NOT NULL DEFAULT now(),
  updated_at_utc timestamptz NOT NULL DEFAULT now(),

  PRIMARY KEY (game_profile, guild_id, channel_id),
  CONSTRAINT gift_code_source_channels_source_fk
    FOREIGN KEY (source_id, game_profile)
    REFERENCES gift_code_sources (id, game_profile),
  CONSTRAINT gift_code_source_channels_profile_check
    CHECK (game_profile IN ('wos', 'kingshot')),
  CONSTRAINT gift_code_source_channels_guild_check
    CHECK (guild_id ~ '^[0-9]{1,32}$'),
  CONSTRAINT gift_code_source_channels_channel_check
    CHECK (channel_id ~ '^[0-9]{1,32}$')
);

CREATE INDEX IF NOT EXISTS gift_code_source_channels_enabled_idx
  ON gift_code_source_channels (game_profile, channel_id, enabled);

CREATE TABLE IF NOT EXISTS gift_code_source_observations (
  id uuid PRIMARY KEY,
  game_profile varchar(32) NOT NULL,
  source_id uuid NOT NULL,
  gift_code_id uuid NOT NULL,
  observed_code varchar(128) NOT NULL,
  observation_key varchar(300) NOT NULL,
  observed_at_utc timestamptz NOT NULL,
  source_reported_expiry_at_utc timestamptz,
  no_longer_observed_at_utc timestamptz,
  provenance jsonb NOT NULL DEFAULT '{}'::jsonb,
  created_at_utc timestamptz NOT NULL DEFAULT now(),
  updated_at_utc timestamptz NOT NULL DEFAULT now(),

  UNIQUE (game_profile, source_id, observation_key),
  CONSTRAINT gift_code_source_observations_source_fk
    FOREIGN KEY (source_id, game_profile)
    REFERENCES gift_code_sources (id, game_profile),
  CONSTRAINT gift_code_source_observations_code_fk
    FOREIGN KEY (gift_code_id, game_profile)
    REFERENCES gift_codes (id, game_profile),
  CONSTRAINT gift_code_source_observations_profile_check
    CHECK (game_profile IN ('wos', 'kingshot')),
  CONSTRAINT gift_code_source_observations_code_check
    CHECK (observed_code = btrim(observed_code) AND observed_code <> ''),
  CONSTRAINT gift_code_source_observations_provenance_check
    CHECK (jsonb_typeof(provenance) = 'object')
);

CREATE INDEX IF NOT EXISTS gift_code_source_observations_code_idx
  ON gift_code_source_observations (game_profile, gift_code_id, observed_at_utc, id);
