ALTER TABLE scheduled_event_groups
  ADD COLUMN first_occurrence_date date;

COMMENT ON COLUMN scheduled_event_groups.first_occurrence_date IS
  'UTC date of this group in the first parent occurrence. NULL preserves legacy parent-date behaviour.';
