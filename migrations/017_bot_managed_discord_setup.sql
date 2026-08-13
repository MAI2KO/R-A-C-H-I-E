CREATE TABLE bot_managed_discord_setups (
  game_profile varchar(32) NOT NULL,
  guild_id varchar(32) NOT NULL,
  category_id varchar(32),
  gift_auto_redeem_channel_id varchar(32),
  gift_announcements_channel_id varchar(32),
  minister_sign_up_channel_id varchar(32),
  event_scheduler_channel_id varchar(32),
  event_announcements_channel_id varchar(32),
  gift_auto_redeem_message_id varchar(32),
  minister_sign_up_message_id varchar(32),
  event_scheduler_message_id varchar(32),
  created_at_utc timestamptz NOT NULL DEFAULT now(),
  updated_at_utc timestamptz NOT NULL DEFAULT now(),

  PRIMARY KEY (game_profile, guild_id),
  CONSTRAINT bot_managed_discord_setups_profile_check
    CHECK (game_profile IN ('wos', 'kingshot')),
  CONSTRAINT bot_managed_discord_setups_guild_check
    CHECK (guild_id ~ '^[0-9]{1,32}$'),
  CONSTRAINT bot_managed_discord_setups_category_check
    CHECK (category_id IS NULL OR category_id ~ '^[0-9]{1,32}$'),
  CONSTRAINT bot_managed_discord_setups_gift_auto_channel_check
    CHECK (gift_auto_redeem_channel_id IS NULL OR gift_auto_redeem_channel_id ~ '^[0-9]{1,32}$'),
  CONSTRAINT bot_managed_discord_setups_gift_announcement_channel_check
    CHECK (gift_announcements_channel_id IS NULL OR gift_announcements_channel_id ~ '^[0-9]{1,32}$'),
  CONSTRAINT bot_managed_discord_setups_minister_channel_check
    CHECK (minister_sign_up_channel_id IS NULL OR minister_sign_up_channel_id ~ '^[0-9]{1,32}$'),
  CONSTRAINT bot_managed_discord_setups_event_scheduler_channel_check
    CHECK (event_scheduler_channel_id IS NULL OR event_scheduler_channel_id ~ '^[0-9]{1,32}$'),
  CONSTRAINT bot_managed_discord_setups_event_announcement_channel_check
    CHECK (event_announcements_channel_id IS NULL OR event_announcements_channel_id ~ '^[0-9]{1,32}$'),
  CONSTRAINT bot_managed_discord_setups_gift_message_check
    CHECK (gift_auto_redeem_message_id IS NULL OR gift_auto_redeem_message_id ~ '^[0-9]{1,32}$'),
  CONSTRAINT bot_managed_discord_setups_minister_message_check
    CHECK (minister_sign_up_message_id IS NULL OR minister_sign_up_message_id ~ '^[0-9]{1,32}$'),
  CONSTRAINT bot_managed_discord_setups_event_message_check
    CHECK (event_scheduler_message_id IS NULL OR event_scheduler_message_id ~ '^[0-9]{1,32}$')
);
