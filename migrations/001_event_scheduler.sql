CREATE TABLE event_guild_settings (
  guild_id varchar(32) NOT NULL,
  game_profile varchar(32) NOT NULL,
  bot_instance_name varchar(100) NOT NULL,
  alliance_name varchar(100) NOT NULL,
  event_channel_id varchar(32) NOT NULL,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now(),

  PRIMARY KEY (guild_id, game_profile),

  CONSTRAINT event_guild_settings_game_profile_check
    CHECK (game_profile IN ('wos', 'kingshot')),

  CONSTRAINT event_guild_settings_bot_instance_name_check
    CHECK (btrim(bot_instance_name) <> ''),

  CONSTRAINT event_guild_settings_alliance_name_check
    CHECK (btrim(alliance_name) <> '')
);

CREATE TABLE event_state_links (
  alliance_guild_id varchar(32) NOT NULL,
  game_profile varchar(32) NOT NULL,
  configured_by_bot_instance varchar(100) NOT NULL,
  state_guild_id varchar(32) NOT NULL,
  state_event_channel_id varchar(32) NOT NULL,
  sharing_enabled boolean NOT NULL DEFAULT false,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now(),

  PRIMARY KEY (alliance_guild_id, game_profile),

  CONSTRAINT event_state_links_guild_settings_fk
    FOREIGN KEY (alliance_guild_id, game_profile)
    REFERENCES event_guild_settings (guild_id, game_profile)
    ON DELETE CASCADE,

  CONSTRAINT event_state_links_game_profile_check
    CHECK (game_profile IN ('wos', 'kingshot')),

  CONSTRAINT event_state_links_bot_instance_check
    CHECK (btrim(configured_by_bot_instance) <> '')
);

CREATE TABLE scheduled_events (
  id bigint GENERATED ALWAYS AS IDENTITY,
  guild_id varchar(32) NOT NULL,
  game_profile varchar(32) NOT NULL,
  created_by_bot_instance varchar(100) NOT NULL,
  alliance_name varchar(100) NOT NULL,
  event_name varchar(100) NOT NULL,
  first_occurrence_date date NOT NULL,
  event_time_utc time without time zone,
  recurrence_days smallint NOT NULL,
  image_url text,
  advance_reminder_minutes smallint,
  reminder_at_start boolean NOT NULL DEFAULT false,
  publish_to_alliance boolean NOT NULL DEFAULT true,
  publish_to_state boolean NOT NULL DEFAULT false,
  include_in_weekly_roundup boolean NOT NULL DEFAULT false,
  status varchar(16) NOT NULL DEFAULT 'active',
  created_by_user_id varchar(32) NOT NULL,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now(),

  PRIMARY KEY (id, game_profile),

  CONSTRAINT scheduled_events_guild_settings_fk
    FOREIGN KEY (guild_id, game_profile)
    REFERENCES event_guild_settings (guild_id, game_profile)
    ON DELETE CASCADE,

  CONSTRAINT scheduled_events_game_profile_check
    CHECK (game_profile IN ('wos', 'kingshot')),

  CONSTRAINT scheduled_events_bot_instance_check
    CHECK (btrim(created_by_bot_instance) <> ''),

  CONSTRAINT scheduled_events_alliance_name_check
    CHECK (btrim(alliance_name) <> ''),

  CONSTRAINT scheduled_events_event_name_check
    CHECK (btrim(event_name) <> ''),

  CONSTRAINT scheduled_events_recurrence_check
    CHECK (recurrence_days IN (3, 7, 14, 28)),

  CONSTRAINT scheduled_events_advance_reminder_check
    CHECK (
      advance_reminder_minutes IS NULL
      OR advance_reminder_minutes IN (10, 30)
    ),

  CONSTRAINT scheduled_events_status_check
    CHECK (status IN ('active', 'paused', 'deleted'))
);

CREATE TABLE scheduled_event_groups (
  id bigint GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
  event_id bigint NOT NULL,
  game_profile varchar(32) NOT NULL,
  group_name varchar(100) NOT NULL,
  event_time_utc time without time zone NOT NULL,
  sort_order integer NOT NULL DEFAULT 0,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now(),

  CONSTRAINT scheduled_event_groups_event_fk
    FOREIGN KEY (event_id, game_profile)
    REFERENCES scheduled_events (id, game_profile)
    ON DELETE CASCADE,

  CONSTRAINT scheduled_event_groups_game_profile_check
    CHECK (game_profile IN ('wos', 'kingshot')),

  CONSTRAINT scheduled_event_groups_name_check
    CHECK (btrim(group_name) <> ''),

  CONSTRAINT scheduled_event_groups_identity_scope_unique
    UNIQUE (id, event_id, game_profile),

  CONSTRAINT scheduled_event_groups_name_unique
    UNIQUE (event_id, game_profile, group_name)
);

CREATE TABLE event_delivery_claims (
  id bigint GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
  event_id bigint NOT NULL,
  group_id bigint,
  game_profile varchar(32) NOT NULL,
  occurrence_at timestamptz NOT NULL,
  deliver_at timestamptz NOT NULL,
  delivery_kind varchar(32) NOT NULL,
  target_kind varchar(16) NOT NULL,
  target_guild_id varchar(32) NOT NULL,
  target_channel_id varchar(32) NOT NULL,
  status varchar(16) NOT NULL DEFAULT 'pending',
  claimed_by_bot_instance varchar(100),
  claimed_by_worker varchar(200),
  claimed_at timestamptz,
  claimed_until timestamptz,
  attempt_count integer NOT NULL DEFAULT 0,
  next_attempt_at timestamptz,
  sent_at timestamptz,
  sent_message_id varchar(32),
  last_error text,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now(),

  CONSTRAINT event_delivery_claims_event_fk
    FOREIGN KEY (event_id, game_profile)
    REFERENCES scheduled_events (id, game_profile)
    ON DELETE CASCADE,

  CONSTRAINT event_delivery_claims_group_fk
    FOREIGN KEY (group_id, event_id, game_profile)
    REFERENCES scheduled_event_groups (id, event_id, game_profile)
    ON DELETE CASCADE,

  CONSTRAINT event_delivery_claims_game_profile_check
    CHECK (game_profile IN ('wos', 'kingshot')),

  CONSTRAINT event_delivery_claims_delivery_kind_check
    CHECK (delivery_kind IN ('advance_reminder', 'event_start')),

  CONSTRAINT event_delivery_claims_target_kind_check
    CHECK (target_kind IN ('alliance', 'state')),

  CONSTRAINT event_delivery_claims_status_check
    CHECK (status IN ('pending', 'claimed', 'sent', 'failed')),

  CONSTRAINT event_delivery_claims_claimed_bot_instance_check
    CHECK (
      claimed_by_bot_instance IS NULL
      OR btrim(claimed_by_bot_instance) <> ''
    ),

  CONSTRAINT event_delivery_claims_claimed_worker_check
    CHECK (
      claimed_by_worker IS NULL
      OR btrim(claimed_by_worker) <> ''
    ),

  CONSTRAINT event_delivery_claims_attempt_count_check
    CHECK (attempt_count >= 0)
);

CREATE UNIQUE INDEX event_delivery_claims_idempotency_idx
  ON event_delivery_claims (
    event_id,
    game_profile,
    COALESCE(group_id, 0::bigint),
    occurrence_at,
    delivery_kind,
    target_kind,
    target_channel_id
  );

CREATE INDEX event_delivery_claims_polling_idx
  ON event_delivery_claims (
    game_profile,
    status,
    deliver_at,
    next_attempt_at,
    claimed_until
  );

CREATE TABLE weekly_roundup_claims (
  id bigint GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
  week_start_date date NOT NULL,
  game_profile varchar(32) NOT NULL,
  target_kind varchar(16) NOT NULL,
  target_guild_id varchar(32) NOT NULL,
  target_channel_id varchar(32) NOT NULL,
  status varchar(16) NOT NULL DEFAULT 'pending',
  claimed_by_bot_instance varchar(100),
  claimed_by_worker varchar(200),
  claimed_at timestamptz,
  claimed_until timestamptz,
  attempt_count integer NOT NULL DEFAULT 0,
  next_attempt_at timestamptz,
  sent_at timestamptz,
  sent_message_id varchar(32),
  last_error text,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now(),

  CONSTRAINT weekly_roundup_claims_game_profile_check
    CHECK (game_profile IN ('wos', 'kingshot')),

  CONSTRAINT weekly_roundup_claims_target_kind_check
    CHECK (target_kind IN ('alliance', 'state')),

  CONSTRAINT weekly_roundup_claims_status_check
    CHECK (status IN ('pending', 'claimed', 'sent', 'failed')),

  CONSTRAINT weekly_roundup_claims_claimed_bot_instance_check
    CHECK (
      claimed_by_bot_instance IS NULL
      OR btrim(claimed_by_bot_instance) <> ''
    ),

  CONSTRAINT weekly_roundup_claims_claimed_worker_check
    CHECK (
      claimed_by_worker IS NULL
      OR btrim(claimed_by_worker) <> ''
    ),

  CONSTRAINT weekly_roundup_claims_attempt_count_check
    CHECK (attempt_count >= 0),

  CONSTRAINT weekly_roundup_claims_idempotency_unique
    UNIQUE (
      week_start_date,
      game_profile,
      target_kind,
      target_channel_id
    )
);

CREATE INDEX weekly_roundup_claims_polling_idx
  ON weekly_roundup_claims (
    game_profile,
    status,
    week_start_date,
    next_attempt_at,
    claimed_until
  );
