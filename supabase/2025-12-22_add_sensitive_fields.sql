-- Migration: Add Sensitive bookmark fields
-- Date: 2025-12-22
-- Applies to an existing public.bookmarks table created by supabase/bookmarks_schema.sql

begin;

alter table if exists public.bookmarks
  add column if not exists sensitive boolean not null default false;

alter table if exists public.bookmarks
  add column if not exists "contentEnc" text not null default '';

-- Optional backfill for legacy rows written before the schema migration.
-- The app may have stored DPAPI-encrypted ciphertext inline as: content = 'enc:v1:<base64>'
-- This migrates those rows to the new columns.
update public.bookmarks
set
  sensitive = true,
  "contentEnc" = substring(content from 8),
  content = ''
where
  content like 'enc:v1:%'
  and ("contentEnc" = '');

commit;
