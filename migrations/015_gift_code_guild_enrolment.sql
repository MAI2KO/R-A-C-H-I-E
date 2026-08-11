ALTER TABLE player_account_guilds
  ADD COLUMN IF NOT EXISTS gift_code_enrolled boolean NOT NULL DEFAULT false,
  ADD COLUMN IF NOT EXISTS gift_code_first_enabled_at_utc timestamptz,
  ADD COLUMN IF NOT EXISTS gift_code_updated_at_utc timestamptz;

UPDATE player_account_guilds ag
   SET gift_code_enrolled = true,
       gift_code_first_enabled_at_utc = COALESCE(
         gift_code_first_enabled_at_utc,
         (
           SELECT MIN(e.created_at_utc)
             FROM gift_code_engagement_events e
            WHERE e.game_profile = ag.game_profile
              AND e.guild_id = ag.guild_id
              AND e.player_account_id = ag.player_account_id
              AND e.event_type = 'auto_redeem_join'
         ),
         ag.linked_at_utc
       ),
       gift_code_updated_at_utc = COALESCE(gift_code_updated_at_utc, now())
 WHERE EXISTS (
   SELECT 1
     FROM gift_code_engagement_events e
    WHERE e.game_profile = ag.game_profile
      AND e.guild_id = ag.guild_id
      AND e.player_account_id = ag.player_account_id
      AND e.event_type = 'auto_redeem_join'
 );

CREATE INDEX IF NOT EXISTS player_account_guilds_gift_enrolment_idx
  ON player_account_guilds (game_profile, guild_id, player_account_id)
  WHERE gift_code_enrolled = true;
