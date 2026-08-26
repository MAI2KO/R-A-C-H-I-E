ALTER TABLE bot_managed_discord_setups
  ADD COLUMN community_number varchar(10),
  ADD COLUMN discord_guild_name varchar(100),
  ADD COLUMN alliance_abbreviation varchar(3),
  ADD CONSTRAINT bot_managed_discord_setups_community_number_check
    CHECK (community_number IS NULL OR community_number ~ '^[0-9]{1,10}$'),
  ADD CONSTRAINT bot_managed_discord_setups_guild_name_check
    CHECK (discord_guild_name IS NULL OR (
      discord_guild_name = btrim(discord_guild_name) AND discord_guild_name <> ''
    )),
  ADD CONSTRAINT bot_managed_discord_setups_alliance_check
    CHECK (alliance_abbreviation IS NULL OR alliance_abbreviation ~ '^[A-Z0-9]{3}$');

ALTER TABLE player_accounts
  ADD COLUMN in_game_name varchar(30),
  ADD COLUMN alliance_abbreviation varchar(3),
  ADD CONSTRAINT player_accounts_in_game_name_check
    CHECK (in_game_name IS NULL OR (
      in_game_name = btrim(in_game_name) AND in_game_name <> ''
      AND in_game_name !~ '[[:cntrl:]]'
    )),
  ADD CONSTRAINT player_accounts_alliance_check
    CHECK (alliance_abbreviation IS NULL OR alliance_abbreviation ~ '^[A-Z0-9]{3}$');

COMMENT ON COLUMN player_accounts.in_game_name IS
  'Canonical player-entered in-game name shared with native booking registration.';
COMMENT ON COLUMN player_accounts.alliance_abbreviation IS
  'Canonical three-character alliance abbreviation; it is not globally unique.';
