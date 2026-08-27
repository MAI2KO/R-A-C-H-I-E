ALTER TABLE bot_managed_discord_setups
  ADD COLUMN bot_manager_role_id varchar(32),
  ADD CONSTRAINT bot_managed_discord_setups_manager_role_check
    CHECK (bot_manager_role_id IS NULL OR bot_manager_role_id ~ '^[0-9]{1,32}$');

COMMENT ON COLUMN bot_managed_discord_setups.bot_manager_role_id IS
  'Optional additional manager role for this exact game_profile and Discord guild.';
