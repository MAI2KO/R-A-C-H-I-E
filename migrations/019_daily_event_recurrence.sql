ALTER TABLE scheduled_events
  DROP CONSTRAINT scheduled_events_recurrence_check,
  ADD CONSTRAINT scheduled_events_recurrence_check
    CHECK (recurrence_days IN (1, 2, 3, 7, 14, 21, 28, 35, 42));

ALTER TABLE state_events
  DROP CONSTRAINT state_events_recurrence_check,
  ADD CONSTRAINT state_events_recurrence_check
    CHECK (recurrence_days IS NULL OR recurrence_days IN (1, 2, 3, 7, 14, 21, 28, 35, 42));
