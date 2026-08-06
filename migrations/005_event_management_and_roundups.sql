ALTER TABLE event_guild_settings
  ADD COLUMN weekly_roundup_enabled boolean NOT NULL DEFAULT false,
  ADD COLUMN weekly_roundup_day smallint NOT NULL DEFAULT 1,
  ADD COLUMN weekly_roundup_time_utc time without time zone NOT NULL DEFAULT '09:00',
  ADD COLUMN weekly_roundup_channel_id varchar(32),
  ADD COLUMN roundup_when_empty boolean NOT NULL DEFAULT false,
  ADD CONSTRAINT event_guild_settings_roundup_day_check
    CHECK (weekly_roundup_day BETWEEN 0 AND 6),
  ADD CONSTRAINT event_guild_settings_roundup_channel_check
    CHECK (weekly_roundup_channel_id IS NULL OR btrim(weekly_roundup_channel_id) <> '');

ALTER TABLE scheduled_events
  ADD COLUMN schedule_version integer NOT NULL DEFAULT 1,
  ADD CONSTRAINT scheduled_events_schedule_version_check CHECK (schedule_version > 0);

ALTER TABLE event_delivery_claims
  ADD COLUMN schedule_version integer NOT NULL DEFAULT 1,
  ADD CONSTRAINT event_delivery_claims_schedule_version_check CHECK (schedule_version > 0);

ALTER TABLE event_delivery_claims
  DROP CONSTRAINT event_delivery_claims_delivery_kind_check,
  ADD CONSTRAINT event_delivery_claims_delivery_kind_check
    CHECK (delivery_kind IN ('advance_reminder', 'final_reminder', 'event_start'));

UPDATE event_delivery_claims
   SET status = 'failed',
       claimed_by_bot_instance = NULL,
       claimed_by_worker = NULL,
       claimed_at = NULL,
       claimed_until = NULL,
       next_attempt_at = NULL,
       last_error = CASE
         WHEN target_kind = 'state'
           THEN 'Individual state reminders are disabled.'
         ELSE 'Exact-start reminders are disabled.'
       END,
       updated_at = now()
 WHERE status <> 'sent'
   AND (target_kind = 'state' OR delivery_kind = 'event_start');

ALTER TABLE event_delivery_claims
  ADD CONSTRAINT event_delivery_claims_individual_policy_check
    CHECK (
      (target_kind = 'alliance' AND delivery_kind <> 'event_start')
      OR status = 'sent'
      OR (status = 'failed' AND next_attempt_at IS NULL)
    );

DROP INDEX event_delivery_claims_idempotency_idx;

CREATE UNIQUE INDEX event_delivery_claims_idempotency_idx
  ON event_delivery_claims (
    event_id,
    game_profile,
    schedule_version,
    COALESCE(group_id, 0::bigint),
    occurrence_at,
    delivery_kind,
    target_kind,
    target_channel_id
  );

ALTER TABLE weekly_roundup_claims
  ADD COLUMN source_guild_id varchar(32),
  ADD COLUMN scheduled_for timestamptz,
  ADD COLUMN post_when_empty boolean NOT NULL DEFAULT false,
  ADD COLUMN part_count integer,
  ADD CONSTRAINT weekly_roundup_claims_part_count_check
    CHECK (part_count IS NULL OR part_count >= 0);

UPDATE weekly_roundup_claims
   SET source_guild_id = target_guild_id,
       scheduled_for = (week_start_date::timestamp AT TIME ZONE 'UTC') + interval '9 hours'
 WHERE source_guild_id IS NULL OR scheduled_for IS NULL;

ALTER TABLE weekly_roundup_claims
  ALTER COLUMN source_guild_id SET NOT NULL,
  ALTER COLUMN scheduled_for SET NOT NULL;

CREATE INDEX weekly_roundup_claims_pending_due_idx
  ON weekly_roundup_claims (game_profile, scheduled_for, id)
  WHERE status = 'pending';

CREATE INDEX weekly_roundup_claims_failed_due_idx
  ON weekly_roundup_claims (game_profile, next_attempt_at, id)
  WHERE status = 'failed' AND next_attempt_at IS NOT NULL;

CREATE INDEX weekly_roundup_claims_expired_lease_idx
  ON weekly_roundup_claims (game_profile, claimed_until, id)
  WHERE status = 'claimed';

CREATE TABLE weekly_roundup_messages (
  roundup_claim_id bigint NOT NULL REFERENCES weekly_roundup_claims (id) ON DELETE CASCADE,
  message_index integer NOT NULL,
  sent_message_id varchar(32) NOT NULL,
  payload_hash varchar(64) NOT NULL,
  created_at timestamptz NOT NULL DEFAULT now(),

  PRIMARY KEY (roundup_claim_id, message_index),

  CONSTRAINT weekly_roundup_messages_index_check CHECK (message_index >= 0),
  CONSTRAINT weekly_roundup_messages_id_check CHECK (btrim(sent_message_id) <> ''),
  CONSTRAINT weekly_roundup_messages_hash_check CHECK (payload_hash ~ '^[0-9a-f]{64}$')
);
