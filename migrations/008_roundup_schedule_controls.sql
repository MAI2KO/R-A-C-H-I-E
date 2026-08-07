ALTER TABLE event_guild_settings
  ADD COLUMN state_roundup_enabled boolean NOT NULL DEFAULT false,
  ADD COLUMN weekly_roundup_not_before timestamptz NOT NULL DEFAULT '-infinity';

UPDATE event_guild_settings
   SET state_roundup_enabled = weekly_roundup_enabled;

COMMENT ON COLUMN event_guild_settings.state_roundup_enabled IS
  'Controls whether this alliance contributes eligible events to its linked state weekly roundup.';

COMMENT ON COLUMN event_guild_settings.weekly_roundup_not_before IS
  'Prevents a newly enabled or changed schedule from replaying a roundup whose scheduled time has already passed.';
