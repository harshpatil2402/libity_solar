-- ═══════════════════════════════════════════════════════════════
-- LibityInfotech Solar SaaS — Supabase Schema
-- Run this entire file in Supabase SQL Editor once.
-- ═══════════════════════════════════════════════════════════════

-- ── Agencies ─────────────────────────────────────────────────────
create table if not exists agencies (
  id              uuid primary key default gen_random_uuid(),
  agency_name     text not null,
  director_name   text,
  email           text unique not null,
  username        text unique not null,
  password        text not null,           -- plain text (you control this)
  contact_number  text,
  agency_address  text,
  logo_b64        text,                    -- base64 agency logo
  stamp_b64       text,                    -- base64 agency stamp
  plan            text not null default 'full',   -- 'full' | 'docs_only'
  is_active       boolean not null default true,
  expires_at      date,
  created_at      timestamptz default now()
);

-- ── Clients ───────────────────────────────────────────────────────
create table if not exists clients (
  id              uuid primary key default gen_random_uuid(),
  agency_id       uuid not null references agencies(id) on delete cascade,
  name            text not null,
  mobile          text not null,
  kw_capacity     numeric(6,2),
  final_amount    numeric(12,2),
  consumer_number text,
  address         text,
  city            text,
  status          text not null default 'New Lead',
  -- pipeline statuses:
  -- 'New Lead' → 'Link Sent' → 'Info Collected' → 'Documents Generated'
  -- → 'Portal Applied' → 'Sanctioned' → 'Installation Done'
  -- → 'GPS Photos Received' → 'Final Submission Done' → 'Completed'
  seen            boolean not null default false,   -- for notification badge
  notes           text,
  portal_applied_at timestamptz,
  net_meter_number  text,
  created_at      timestamptz default now(),
  updated_at      timestamptz default now()
);

-- ── Client Submissions (info collected via link) ──────────────────
create table if not exists client_submissions (
  id              uuid primary key default gen_random_uuid(),
  client_id       uuid not null references clients(id) on delete cascade,
  agency_id       uuid not null references agencies(id) on delete cascade,
  token           text unique not null,
  status          text not null default 'pending',  -- 'pending' | 'submitted' | 'expired'
  -- Consumer details filled by client
  consumer_name   text,
  consumer_number text,
  consumer_address text,
  consumer_email  text,
  consumer_aadhar_num text,
  city            text,
  -- Images as base64
  aadhar_b64      text,
  pan_b64         text,
  cheque_b64      text,
  signature_b64   text,
  expires_at      timestamptz,
  submitted_at    timestamptz,
  created_at      timestamptz default now()
);

-- ── Doc Generation History ────────────────────────────────────────
create table if not exists doc_jobs (
  id              uuid primary key default gen_random_uuid(),
  agency_id       uuid not null references agencies(id) on delete cascade,
  client_id       uuid references clients(id) on delete set null,
  consumer_name   text,
  consumer_number text,
  status          text not null default 'running',  -- 'running'|'done'|'error'
  zip_name        text,
  created_at      timestamptz default now()
);

-- ── GPS Photos ────────────────────────────────────────────────────
create table if not exists gps_photos (
  id          uuid primary key default gen_random_uuid(),
  client_id   uuid not null references clients(id) on delete cascade,
  agency_id   uuid not null references agencies(id) on delete cascade,
  token       text unique not null,
  status      text not null default 'pending',   -- 'pending' | 'submitted'
  photo_b64   text,
  latitude    numeric(10,7),
  longitude   numeric(10,7),
  taken_at    timestamptz,
  created_at  timestamptz default now()
);

-- ── Indexes ───────────────────────────────────────────────────────
create index if not exists idx_clients_agency    on clients(agency_id);
create index if not exists idx_clients_status    on clients(status);
create index if not exists idx_submissions_token on client_submissions(token);
create index if not exists idx_gps_token         on gps_photos(token);
create index if not exists idx_docjobs_agency    on doc_jobs(agency_id);
