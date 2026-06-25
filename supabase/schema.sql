-- ============================================================
-- Benchstock — Supabase schema
-- Paste this whole file into Supabase ▸ SQL Editor ▸ New query ▸ Run.
-- ============================================================

-- ---- Tables -------------------------------------------------

create table if not exists profiles (
  id           uuid primary key references auth.users(id) on delete cascade,
  display_name text not null,
  role         text not null default 'member',
  created_at   timestamptz default now()
);

create table if not exists items (
  id                uuid primary key default gen_random_uuid(),
  name              text not null,
  category          text,
  lot_number        text,
  catalog_number    text,
  supplier          text,
  quantity          numeric default 0,
  unit              text,
  reorder_threshold numeric default 0,
  location          text,
  expiration_date   date,
  notes             text,
  created_by        text,
  created_at        timestamptz default now()
);

create table if not exists usage_log (
  id      uuid primary key default gen_random_uuid(),
  item_id uuid,
  name    text,
  lot     text,
  amount  numeric,
  unit    text,
  who     text,
  ts      timestamptz default now()
);

-- ---- Row Level Security ------------------------------------
-- Everyone who is signed in shares one lab inventory.

alter table profiles  enable row level security;
alter table items     enable row level security;
alter table usage_log enable row level security;

-- profiles: anyone signed in can read the member list; you manage only your own row
drop policy if exists "profiles_read"   on profiles;
drop policy if exists "profiles_insert" on profiles;
drop policy if exists "profiles_update" on profiles;
create policy "profiles_read"   on profiles for select to authenticated using (true);
create policy "profiles_insert" on profiles for insert to authenticated with check (auth.uid() = id);
create policy "profiles_update" on profiles for update to authenticated using (auth.uid() = id);

-- items: any signed-in lab member has full access
drop policy if exists "items_read"   on items;
drop policy if exists "items_insert" on items;
drop policy if exists "items_update" on items;
drop policy if exists "items_delete" on items;
create policy "items_read"   on items for select to authenticated using (true);
create policy "items_insert" on items for insert to authenticated with check (true);
create policy "items_update" on items for update to authenticated using (true);
create policy "items_delete" on items for delete to authenticated using (true);

-- usage log: read all, add your own entries
drop policy if exists "log_read"   on usage_log;
drop policy if exists "log_insert" on usage_log;
create policy "log_read"   on usage_log for select to authenticated using (true);
create policy "log_insert" on usage_log for insert to authenticated with check (true);

-- ---- Realtime ----------------------------------------------
-- Push live updates so the lab sees changes instantly.

alter publication supabase_realtime add table items;
alter publication supabase_realtime add table usage_log;
