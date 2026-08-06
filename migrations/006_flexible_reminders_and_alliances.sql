CREATE TABLE event_alliances (
  id bigint GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
  guild_id varchar(32) NOT NULL,
  game_profile varchar(32) NOT NULL,
  alliance_name varchar(100) NOT NULL,
  is_default boolean NOT NULL DEFAULT false,
  created_by_bot_instance varchar(100) NOT NULL,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now(),

  CONSTRAINT event_alliances_guild_settings_fk
    FOREIGN KEY (guild_id, game_profile)
    REFERENCES event_guild_settings (guild_id, game_profile)
    ON DELETE CASCADE,

  CONSTRAINT event_alliances_game_profile_check
    CHECK (game_profile IN ('wos', 'kingshot')),

  CONSTRAINT event_alliances_name_check
    CHECK (btrim(alliance_name) <> ''),

  CONSTRAINT event_alliances_bot_instance_check
    CHECK (btrim(created_by_bot_instance) <> ''),

  CONSTRAINT event_alliances_ownership_unique
    UNIQUE (id, guild_id, game_profile)
);

CREATE UNIQUE INDEX event_alliances_name_unique_ci
  ON event_alliances (guild_id, game_profile, lower(btrim(alliance_name)));

CREATE UNIQUE INDEX event_alliances_one_default_idx
  ON event_alliances (guild_id, game_profile)
  WHERE is_default = true;

INSERT INTO event_alliances (
  guild_id,
  game_profile,
  alliance_name,
  is_default,
  created_by_bot_instance
)
SELECT
  guild_id,
  game_profile,
  btrim(alliance_name),
  true,
  bot_instance_name
FROM event_guild_settings;

INSERT INTO event_alliances (
  guild_id,
  game_profile,
  alliance_name,
  is_default,
  created_by_bot_instance
)
SELECT
  e.guild_id,
  e.game_profile,
  min(btrim(e.alliance_name)),
  false,
  min(e.created_by_bot_instance)
FROM scheduled_events e
WHERE NOT EXISTS (
  SELECT 1
    FROM event_alliances a
   WHERE a.guild_id = e.guild_id
     AND a.game_profile = e.game_profile
     AND lower(btrim(a.alliance_name)) = lower(btrim(e.alliance_name))
)
GROUP BY e.guild_id, e.game_profile, lower(btrim(e.alliance_name));

ALTER TABLE scheduled_events
  ADD COLUMN alliance_id bigint,
  ADD COLUMN advance_reminder_message text,
  ADD COLUMN final_reminder_message text;

UPDATE scheduled_events e
   SET alliance_id = a.id
  FROM event_alliances a
 WHERE a.guild_id = e.guild_id
   AND a.game_profile = e.game_profile
   AND lower(btrim(a.alliance_name)) = lower(btrim(e.alliance_name));

ALTER TABLE scheduled_events
  ALTER COLUMN alliance_id SET NOT NULL,
  ADD CONSTRAINT scheduled_events_alliance_fk
    FOREIGN KEY (alliance_id, guild_id, game_profile)
    REFERENCES event_alliances (id, guild_id, game_profile),
  ADD CONSTRAINT scheduled_events_advance_message_check
    CHECK (
      advance_reminder_message IS NULL
      OR (
        advance_reminder_message = btrim(advance_reminder_message)
        AND char_length(advance_reminder_message) BETWEEN 1 AND 500
      )
    ),
  ADD CONSTRAINT scheduled_events_final_message_check
    CHECK (
      final_reminder_message IS NULL
      OR (
        final_reminder_message = btrim(final_reminder_message)
        AND char_length(final_reminder_message) BETWEEN 1 AND 500
      )
    );

ALTER TABLE scheduled_events
  DROP CONSTRAINT scheduled_events_advance_reminder_check,
  ADD CONSTRAINT scheduled_events_advance_reminder_check
    CHECK (
      advance_reminder_minutes IS NULL
      OR advance_reminder_minutes IN (5, 10, 15, 20, 30)
    );

ALTER TABLE event_delivery_claims
  ADD COLUMN group_id_snapshot bigint,
  ADD COLUMN group_name_snapshot varchar(100),
  ADD CONSTRAINT event_delivery_claims_group_snapshot_check
    CHECK (
      (group_id_snapshot IS NULL AND group_name_snapshot IS NULL)
      OR (
        group_id_snapshot IS NOT NULL
        AND group_name_snapshot IS NOT NULL
        AND
        group_id_snapshot > 0
        AND group_name_snapshot = btrim(group_name_snapshot)
        AND btrim(group_name_snapshot) <> ''
      )
    );

UPDATE event_delivery_claims d
   SET group_id_snapshot = g.id,
       group_name_snapshot = btrim(g.group_name)
  FROM scheduled_event_groups g
 WHERE d.group_id = g.id
   AND d.event_id = g.event_id
   AND d.game_profile = g.game_profile;

DROP INDEX event_delivery_claims_idempotency_idx;

CREATE UNIQUE INDEX event_delivery_claims_idempotency_idx
  ON event_delivery_claims (
    event_id,
    game_profile,
    schedule_version,
    COALESCE(group_id_snapshot, group_id, 0::bigint),
    occurrence_at,
    delivery_kind,
    target_kind,
    target_channel_id
  );

CREATE INDEX scheduled_events_alliance_idx
  ON scheduled_events (guild_id, game_profile, alliance_id, status, id);

CREATE FUNCTION event_scheduler_sync_event_alliance()
RETURNS trigger
LANGUAGE plpgsql
AS $$
DECLARE
  selected_alliance event_alliances%ROWTYPE;
BEGIN
  IF NEW.alliance_id IS NULL THEN
    SELECT *
      INTO selected_alliance
      FROM event_alliances
     WHERE guild_id = NEW.guild_id
       AND game_profile = NEW.game_profile
       AND is_default = true;
  ELSE
    SELECT *
      INTO selected_alliance
      FROM event_alliances
     WHERE id = NEW.alliance_id
       AND guild_id = NEW.guild_id
       AND game_profile = NEW.game_profile;
  END IF;

  IF selected_alliance.id IS NULL THEN
    RAISE EXCEPTION 'A valid profile-scoped event alliance is required.'
      USING ERRCODE = '23503';
  END IF;

  NEW.alliance_id = selected_alliance.id;
  NEW.alliance_name = selected_alliance.alliance_name;
  RETURN NEW;
END;
$$;

CREATE TRIGGER scheduled_events_sync_alliance_trigger
BEFORE INSERT OR UPDATE OF alliance_id, alliance_name, guild_id, game_profile
ON scheduled_events
FOR EACH ROW
EXECUTE FUNCTION event_scheduler_sync_event_alliance();

COMMENT ON COLUMN scheduled_events.reminder_at_start IS
  'Legacy column name: true schedules the final announcement one minute before occurrence_at; no exact-start delivery is created.';

COMMENT ON COLUMN scheduled_events.alliance_name IS
  'Legacy denormalized alliance name retained for compatibility; alliance_id is authoritative.';

COMMENT ON COLUMN event_guild_settings.alliance_name IS
  'Legacy default-alliance name retained for compatibility; event_alliances is authoritative.';

COMMENT ON COLUMN event_delivery_claims.group_id_snapshot IS
  'Immutable group identity used for delivery history and idempotency if the source group is replaced.';

COMMENT ON COLUMN event_delivery_claims.group_name_snapshot IS
  'Immutable group name retained when event edits replace the source group row.';
