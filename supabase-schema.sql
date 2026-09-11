create extension if not exists pgcrypto;

create table if not exists public.dashboard_settings (
  key text primary key,
  value text not null,
  updated_at timestamptz not null default now()
);

create table if not exists public.dashboard_workbooks (
  id uuid primary key default gen_random_uuid(),
  workbook_name text not null,
  workbook jsonb not null,
  uploaded_at timestamptz not null default now(),
  is_published boolean not null default true
);

alter table public.dashboard_settings enable row level security;
alter table public.dashboard_workbooks enable row level security;

drop policy if exists "public can read published workbook" on public.dashboard_workbooks;
create policy "public can read published workbook"
  on public.dashboard_workbooks
  for select
  to anon, authenticated
  using (is_published = true);

create or replace function public.publish_dashboard_workbook(
  p_password text,
  p_workbook_name text,
  p_workbook jsonb
)
returns jsonb
language plpgsql
security definer
set search_path = public, extensions
as $$
declare
  saved_id uuid;
begin
  if not exists (
    select 1
    from public.dashboard_settings
    where key = 'upload_password_hash'
      and value = crypt(p_password, value)
  ) then
    raise exception 'invalid upload password' using errcode = '28000';
  end if;

  update public.dashboard_workbooks
  set is_published = false
  where is_published = true;

  insert into public.dashboard_workbooks (workbook_name, workbook, is_published)
  values (p_workbook_name, p_workbook, true)
  returning id into saved_id;

  return jsonb_build_object('id', saved_id, 'uploaded_at', now());
end;
$$;

revoke all on table public.dashboard_settings from anon, authenticated;
revoke all on table public.dashboard_workbooks from anon, authenticated;
grant select on table public.dashboard_workbooks to anon, authenticated;
grant execute on function public.publish_dashboard_workbook(text, text, jsonb) to anon, authenticated;

-- Run this separately after replacing the placeholder with your upload password:
-- insert into public.dashboard_settings (key, value)
-- values ('upload_password_hash', crypt('__UPLOAD_PASSWORD__', gen_salt('bf')))
-- on conflict (key) do update set value = excluded.value, updated_at = now();
