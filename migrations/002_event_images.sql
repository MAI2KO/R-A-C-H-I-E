CREATE TABLE scheduled_event_images (
  event_id bigint NOT NULL,
  game_profile varchar(32) NOT NULL,
  original_filename varchar(255) NOT NULL,
  content_type varchar(100) NOT NULL,
  byte_size integer NOT NULL,
  image_data bytea NOT NULL,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now(),

  PRIMARY KEY (event_id, game_profile),

  CONSTRAINT scheduled_event_images_event_fk
    FOREIGN KEY (event_id, game_profile)
    REFERENCES scheduled_events (id, game_profile)
    ON DELETE CASCADE,

  CONSTRAINT scheduled_event_images_game_profile_check
    CHECK (game_profile IN ('wos', 'kingshot')),

  CONSTRAINT scheduled_event_images_filename_check
    CHECK (btrim(original_filename) <> ''),

  CONSTRAINT scheduled_event_images_content_type_check
    CHECK (content_type IN ('image/png', 'image/jpeg', 'image/gif', 'image/webp')),

  CONSTRAINT scheduled_event_images_byte_size_check
    CHECK (byte_size > 0 AND byte_size <= 8388608),

  CONSTRAINT scheduled_event_images_data_size_check
    CHECK (octet_length(image_data) = byte_size)
);
