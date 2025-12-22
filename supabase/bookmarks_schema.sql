create extension if not exists pgcrypto;

create table if not exists public.bookmarks (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null default auth.uid(),
  type text not null check (type in ('text','url','file','command','binary')),
  "typeExplicit" boolean not null default false,

  sensitive boolean not null default false,
  "contentEnc" text not null default '',

  content text not null,
  "binaryData" text not null default '',
  "mimeType" text not null default '',
  tags text[] not null default '{}'::text[],
  "timestamp" bigint not null,
  "lastUsed" bigint not null,

  "deviceId" text not null default '',
  "validOnAnyDevice" boolean not null default true,

  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);
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
create index if not exists idx_bookmarks_user_lastused
  on public.bookmarks (user_id, "lastUsed" desc);
alter table public.bookmarks enable row level security;
drop policy if exists "bookmarks_select_own" on public.bookmarks;
create policy "bookmarks_select_own"
  on public.bookmarks
  for select
  using (auth.uid() = user_id);
drop policy if exists "bookmarks_insert_own" on public.bookmarks;
create policy "bookmarks_insert_own"
  on public.bookmarks
  for insert
  with check (auth.uid() = user_id);
drop policy if exists "bookmarks_update_own" on public.bookmarks;
create policy "bookmarks_update_own"
  on public.bookmarks
  for update
  using (auth.uid() = user_id)
  with check (auth.uid() = user_id);
drop policy if exists "bookmarks_delete_own" on public.bookmarks;
create policy "bookmarks_delete_own"
  on public.bookmarks
  for delete
  using (auth.uid() = user_id);
