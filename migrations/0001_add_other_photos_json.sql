-- Already applied manually on 2026-07-13. Kept here for record.
alter table client_submissions add column if not exists other_photos_json text;
