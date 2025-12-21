-- Checkmeg Supabase schema: bookmarks table
--
-- How to apply:
-- 1) Open Supabase Dashboard → SQL Editor
-- 2) Paste this file and run
--
-- This schema matches the fields the app reads/writes via PostgREST:
-- id, type, typeExplicit, content, binaryData, mimeType, tags,
-- timestamp, lastUsed, deviceId, validOnAnyDevice

-- Needed for gen_random_uuid()
create extension if not exists pgcrypto;

create table if not exists public.bookmarks (
  -- Stable client-generated UUID (the app generates RFC4122 v4 strings).
  id uuid primary key default gen_random_uuid(),

  -- Row ownership (used by RLS)
  user_id uuid not null default auth.uid(),

  -- Bookmark fields
  type text not null check (type in ('text','url','file','command','binary')),
  "typeExplicit" boolean not null default false,

  content text not null,

  -- For binary bookmarks (base64)
  "binaryData" text not null default '',
  "mimeType" text not null default '',

  -- Tags
  tags text[] not null default '{}'::text[],

  -- Epoch seconds (matches local file format)
  "timestamp" bigint not null,
  "lastUsed" bigint not null,

  "deviceId" text not null default '',
  "validOnAnyDevice" boolean not null default true,

  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

-- Keep updated_at fresh
create or replace function public.set_updated_at()
returns trigger
language plpgsql
as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

drop trigger if exists trg_bookmarks_updated_at on public.bookmarks;
create trigger trg_bookmarks_updated_at
before update on public.bookmarks
for each row
execute function public.set_updated_at();

-- Helpful index for per-user queries
create index if not exists idx_bookmarks_user_lastused
  on public.bookmarks (user_id, "lastUsed" desc);

-- Enable Row Level Security
alter table public.bookmarks enable row level security;

-- Policies: users can only access their own rows
-- SELECT
drop policy if exists "bookmarks_select_own" on public.bookmarks;
create policy "bookmarks_select_own"
  on public.bookmarks
  for select
  using (auth.uid() = user_id);

-- INSERT
drop policy if exists "bookmarks_insert_own" on public.bookmarks;
create policy "bookmarks_insert_own"
  on public.bookmarks
  for insert
  with check (auth.uid() = user_id);

-- UPDATE
drop policy if exists "bookmarks_update_own" on public.bookmarks;
create policy "bookmarks_update_own"
  on public.bookmarks
  for update
  using (auth.uid() = user_id)
  with check (auth.uid() = user_id);

-- DELETE
drop policy if exists "bookmarks_delete_own" on public.bookmarks;
create policy "bookmarks_delete_own"
  on public.bookmarks
  for delete
  using (auth.uid() = user_id);

-- Optional: prevent a user from creating two rows with the same id
-- (id is already PRIMARY KEY, so this is implied)
