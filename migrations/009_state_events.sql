CREATE EXTENSION IF NOT EXISTS pgcrypto;

ALTER TABLE scheduled_events
  DROP CONSTRAINT scheduled_events_recurrence_check,
  ADD CONSTRAINT scheduled_events_recurrence_check
    CHECK (recurrence_days IN (2, 3, 7, 14, 21, 28, 35, 42));

ALTER TABLE event_state_destinations
  ADD COLUMN state_number varchar(32),
  ADD CONSTRAINT event_state_destinations_state_number_check
    CHECK (
      state_number IS NULL
      OR (
        state_number = btrim(state_number)
        AND char_length(state_number) BETWEEN 1 AND 32
      )
    );

CREATE TABLE state_events (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  state_guild_id varchar(32) NOT NULL,
  game_profile varchar(32) NOT NULL,
  created_by_bot_instance varchar(100) NOT NULL,
  event_name varchar(100) NOT NULL,
  first_occurrence_date date NOT NULL,
  recurrence_days smallint,
  status varchar(16) NOT NULL DEFAULT 'active',
  schedule_version integer NOT NULL DEFAULT 1,
  created_by_user_id varchar(32) NOT NULL,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now(),

  CONSTRAINT state_events_destination_fk
    FOREIGN KEY (state_guild_id, game_profile)
    REFERENCES event_state_destinations (state_guild_id, game_profile)
    ON DELETE CASCADE,

  CONSTRAINT state_events_game_profile_check
    CHECK (game_profile IN ('wos', 'kingshot')),

  CONSTRAINT state_events_bot_instance_check
    CHECK (btrim(created_by_bot_instance) <> ''),

  CONSTRAINT state_events_name_check
    CHECK (event_name = btrim(event_name) AND btrim(event_name) <> ''),

  CONSTRAINT state_events_recurrence_check
    CHECK (recurrence_days IS NULL OR recurrence_days IN (2, 3, 7, 14, 21, 28, 35, 42)),

  CONSTRAINT state_events_status_check
    CHECK (status IN ('active', 'paused', 'deleted')),

  CONSTRAINT state_events_schedule_version_check
    CHECK (schedule_version > 0),

  CONSTRAINT state_events_user_check
    CHECK (btrim(created_by_user_id) <> ''),

  CONSTRAINT state_events_identity_scope_unique
    UNIQUE (id, game_profile)
);

CREATE INDEX state_events_owner_idx
  ON state_events (state_guild_id, game_profile, status, first_occurrence_date, id);

CREATE TABLE state_event_phases (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  state_event_id uuid NOT NULL REFERENCES state_events (id) ON DELETE CASCADE,
  game_profile varchar(32) NOT NULL,
  phase_name varchar(100) NOT NULL,
  phase_time_utc time without time zone NOT NULL,
  pre_alert_minutes smallint,
  pre_alert_message text,
  announce_exact boolean NOT NULL DEFAULT true,
  exact_message text,
  sort_order integer NOT NULL DEFAULT 0,
  status varchar(16) NOT NULL DEFAULT 'active',
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now(),

  CONSTRAINT state_event_phases_profile_fk
    FOREIGN KEY (state_event_id, game_profile)
    REFERENCES state_events (id, game_profile)
    ON DELETE CASCADE,

  CONSTRAINT state_event_phases_game_profile_check
    CHECK (game_profile IN ('wos', 'kingshot')),

  CONSTRAINT state_event_phases_name_check
    CHECK (phase_name = btrim(phase_name) AND btrim(phase_name) <> ''),

  CONSTRAINT state_event_phases_pre_alert_check
    CHECK (pre_alert_minutes IS NULL OR pre_alert_minutes IN (5, 10, 15, 20, 30)),

  CONSTRAINT state_event_phases_status_check
    CHECK (status IN ('active', 'deleted')),

  CONSTRAINT state_event_phases_pre_message_check
    CHECK (
      pre_alert_message IS NULL
      OR (
        pre_alert_message = btrim(pre_alert_message)
        AND char_length(pre_alert_message) BETWEEN 1 AND 500
      )
    ),

  CONSTRAINT state_event_phases_exact_message_check
    CHECK (
      exact_message IS NULL
      OR (
        exact_message = btrim(exact_message)
        AND char_length(exact_message) BETWEEN 1 AND 500
      )
    )
);

CREATE UNIQUE INDEX state_event_phases_name_unique_ci
  ON state_event_phases (state_event_id, game_profile, lower(btrim(phase_name)))
  WHERE status = 'active';

CREATE INDEX state_event_phases_event_idx
  ON state_event_phases (state_event_id, game_profile, sort_order, phase_time_utc, id);

ALTER TABLE state_event_phases
  ADD CONSTRAINT state_event_phases_identity_scope_unique
    UNIQUE (id, state_event_id, game_profile);

CREATE TABLE state_event_media (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  state_event_id uuid NOT NULL,
  phase_id uuid NOT NULL,
  game_profile varchar(32) NOT NULL,
  delivery_kind varchar(16) NOT NULL,
  original_filename varchar(255) NOT NULL,
  content_type varchar(100) NOT NULL,
  byte_size integer NOT NULL,
  image_data bytea NOT NULL,
  created_at timestamptz NOT NULL DEFAULT now(),

  CONSTRAINT state_event_media_phase_fk
    FOREIGN KEY (phase_id, state_event_id, game_profile)
    REFERENCES state_event_phases (id, state_event_id, game_profile)
    ON DELETE CASCADE,

  CONSTRAINT state_event_media_game_profile_check
    CHECK (game_profile IN ('wos', 'kingshot')),

  CONSTRAINT state_event_media_kind_check
    CHECK (delivery_kind IN ('pre_alert', 'exact')),

  CONSTRAINT state_event_media_filename_check
    CHECK (btrim(original_filename) <> ''),

  CONSTRAINT state_event_media_content_type_check
    CHECK (content_type IN ('image/png', 'image/jpeg', 'image/gif', 'image/webp')),

  CONSTRAINT state_event_media_size_check
    CHECK (byte_size > 0 AND byte_size <= 8388608)
);

CREATE UNIQUE INDEX state_event_media_one_per_kind_idx
  ON state_event_media (phase_id, game_profile, delivery_kind);

CREATE TABLE state_event_delivery_claims (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  state_event_id uuid NOT NULL,
  phase_id uuid NOT NULL,
  game_profile varchar(32) NOT NULL,
  schedule_version integer NOT NULL,
  occurrence_at timestamptz NOT NULL,
  deliver_at timestamptz NOT NULL,
  delivery_kind varchar(16) NOT NULL,
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

  CONSTRAINT state_event_claims_phase_fk
    FOREIGN KEY (phase_id, state_event_id, game_profile)
    REFERENCES state_event_phases (id, state_event_id, game_profile)
    ON DELETE CASCADE,

  CONSTRAINT state_event_claims_game_profile_check
    CHECK (game_profile IN ('wos', 'kingshot')),

  CONSTRAINT state_event_claims_schedule_version_check
    CHECK (schedule_version > 0),

  CONSTRAINT state_event_claims_kind_check
    CHECK (delivery_kind IN ('pre_alert', 'exact')),

  CONSTRAINT state_event_claims_target_kind_check
    CHECK (target_kind IN ('state', 'alliance')),

  CONSTRAINT state_event_claims_status_check
    CHECK (status IN ('pending', 'claimed', 'sent', 'failed')),

  CONSTRAINT state_event_claims_attempt_count_check
    CHECK (attempt_count >= 0)
);

CREATE UNIQUE INDEX state_event_claims_idempotency_idx
  ON state_event_delivery_claims (
    state_event_id,
    game_profile,
    schedule_version,
    phase_id,
    occurrence_at,
    delivery_kind,
    target_guild_id
  );

CREATE UNIQUE INDEX state_event_claims_sent_history_idx
  ON state_event_delivery_claims (
    state_event_id,
    game_profile,
    phase_id,
    occurrence_at,
    delivery_kind,
    target_guild_id
  )
  WHERE status = 'sent';

CREATE INDEX state_event_claims_pending_due_idx
  ON state_event_delivery_claims (game_profile, deliver_at, id)
  WHERE status = 'pending';

CREATE INDEX state_event_claims_failed_due_idx
  ON state_event_delivery_claims (game_profile, next_attempt_at, id)
  WHERE status = 'failed' AND next_attempt_at IS NOT NULL;

CREATE INDEX state_event_claims_expired_lease_idx
  ON state_event_delivery_claims (game_profile, claimed_until, id)
  WHERE status = 'claimed';

COMMENT ON COLUMN event_state_destinations.state_number IS
  'Profile-scoped human-friendly state number or name displayed on state events.';

COMMENT ON COLUMN state_events.recurrence_days IS
  'NULL means one-time. Non-null values are deterministic day intervals anchored to first_occurrence_date plus phase time.';
