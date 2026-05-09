create extension if not exists pgcrypto;

alter table if exists public.profiles
  add column if not exists role text not null default 'user',
  add column if not exists access_status text not null default 'pending',
  add column if not exists approved_by uuid,
  add column if not exists approved_at timestamptz,
  add column if not exists denied_at timestamptz;

alter table if exists public.profiles
  drop constraint if exists profiles_role_check;

alter table if exists public.profiles
  add constraint profiles_role_check check (role in ('admin', 'user'));

alter table if exists public.profiles
  drop constraint if exists profiles_access_status_check;

alter table if exists public.profiles
  add constraint profiles_access_status_check check (access_status in ('pending', 'approved', 'rejected'));

create table if not exists public.access_requests (
  id uuid primary key default gen_random_uuid(),
  auth_user_id uuid unique,
  full_name text not null,
  email text not null unique,
  status text not null default 'pending',
  approved_role text not null default 'user',
  reviewed_by uuid,
  reviewed_at timestamptz,
  created_at timestamptz not null default timezone('utc'::text, now())
);

alter table public.access_requests
  drop constraint if exists access_requests_status_check;

alter table public.access_requests
  add constraint access_requests_status_check check (status in ('pending', 'approved', 'rejected'));

alter table public.access_requests
  drop constraint if exists access_requests_role_check;

alter table public.access_requests
  add constraint access_requests_role_check check (approved_role in ('admin', 'user'));

create table if not exists public.admin_requests (
  id uuid primary key default gen_random_uuid(),
  request_type text not null,
  entity_type text,
  entity_id text,
  title text not null,
  description text,
  reason text,
  payload jsonb not null default '{}'::jsonb,
  status text not null default 'pending',
  requested_by_id uuid not null,
  requested_by_name text,
  requested_by_email text,
  reviewed_by uuid,
  reviewed_at timestamptz,
  created_at timestamptz not null default timezone('utc'::text, now())
);

alter table public.admin_requests
  drop constraint if exists admin_requests_status_check;

alter table public.admin_requests
  add constraint admin_requests_status_check check (status in ('pending', 'approved', 'rejected'));

create table if not exists public.lead_source_catalog (
  id uuid primary key default gen_random_uuid(),
  name text not null unique,
  created_by uuid,
  created_at timestamptz not null default timezone('utc'::text, now())
);

insert into public.lead_source_catalog (name)
values ('Organico'), ('Pago')
on conflict (name) do nothing;

create or replace function public.handle_new_user()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
begin
  insert into public.profiles (id, full_name, email)
  values (
    new.id,
    coalesce(new.raw_user_meta_data ->> 'full_name', split_part(new.email, '@', 1)),
    new.email
  )
  on conflict (id) do update
    set full_name = excluded.full_name,
        email = excluded.email;

  return new;
end;
$$;

drop trigger if exists on_auth_user_created on auth.users;
create trigger on_auth_user_created
after insert on auth.users
for each row execute procedure public.handle_new_user();

create or replace function public.current_user_role()
returns text
language sql
stable
security definer
set search_path = public
as $$
  select coalesce((select role from public.profiles where id = auth.uid()), 'user');
$$;

create or replace function public.current_access_status()
returns text
language sql
stable
security definer
set search_path = public
as $$
  select coalesce((select access_status from public.profiles where id = auth.uid()), 'pending');
$$;

create or replace function public.is_admin()
returns boolean
language sql
stable
security definer
set search_path = public
as $$
  select public.current_user_role() = 'admin' and public.current_access_status() = 'approved';
$$;

create or replace function public.is_approved_user()
returns boolean
language sql
stable
security definer
set search_path = public
as $$
  select public.current_access_status() = 'approved';
$$;

create or replace function public.enforce_lead_permissions()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
declare
  current_name text;
begin
  if not public.is_approved_user() then
    raise exception 'Acesso nao aprovado.';
  end if;

  select full_name into current_name
  from public.profiles
  where id = auth.uid();

  if tg_op = 'DELETE' and not public.is_admin() then
    raise exception 'Somente administradores podem excluir leads.';
  end if;

  if tg_op = 'UPDATE' and not public.is_admin() then
    if new.owner is distinct from old.owner then
      raise exception 'Somente administradores podem alterar o responsavel do lead.';
    end if;
    if new.assigned_to is distinct from old.assigned_to then
      raise exception 'Somente administradores podem alterar o responsavel tecnico do lead.';
    end if;
  end if;

  if tg_op = 'INSERT' and not public.is_admin() then
    if coalesce(trim(new.owner), '') = '' then
      new.owner := coalesce(current_name, split_part(coalesce((select email from public.profiles where id = auth.uid()), ''), '@', 1));
    elsif trim(new.owner) is distinct from coalesce(current_name, trim(new.owner)) then
      raise exception 'Somente administradores podem definir o responsavel do lead.';
    end if;
  end if;

  return case when tg_op = 'DELETE' then old else new end;
end;
$$;

drop trigger if exists leads_permission_guard on public.leads;
create trigger leads_permission_guard
before insert or update or delete on public.leads
for each row execute procedure public.enforce_lead_permissions();

alter table if exists public.profiles enable row level security;
alter table if exists public.leads enable row level security;
alter table if exists public.stages enable row level security;
alter table if exists public.change_history enable row level security;
alter table if exists public.stage_type_catalog enable row level security;
alter table if exists public.lead_source_catalog enable row level security;
alter table if exists public.access_requests enable row level security;
alter table if exists public.admin_requests enable row level security;

drop policy if exists profiles_select_approved on public.profiles;
create policy profiles_select_approved on public.profiles
for select
using (public.is_approved_user());

drop policy if exists profiles_insert_own on public.profiles;
create policy profiles_insert_own on public.profiles
for insert
with check (auth.uid() = id);

drop policy if exists profiles_update_own_or_admin on public.profiles;
create policy profiles_update_own_or_admin on public.profiles
for update
using (auth.uid() = id or public.is_admin())
with check (auth.uid() = id or public.is_admin());

drop policy if exists leads_select_approved on public.leads;
create policy leads_select_approved on public.leads
for select
using (public.is_approved_user());

drop policy if exists leads_insert_approved on public.leads;
create policy leads_insert_approved on public.leads
for insert
with check (public.is_approved_user());

drop policy if exists leads_update_approved on public.leads;
create policy leads_update_approved on public.leads
for update
using (public.is_approved_user())
with check (public.is_approved_user());

drop policy if exists leads_delete_admin on public.leads;
create policy leads_delete_admin on public.leads
for delete
using (public.is_admin());

drop policy if exists stages_select_approved on public.stages;
create policy stages_select_approved on public.stages
for select
using (public.is_approved_user());

drop policy if exists stages_mutate_admin on public.stages;
create policy stages_mutate_admin on public.stages
for all
using (public.is_admin())
with check (public.is_admin());

drop policy if exists history_select_admin on public.change_history;
create policy history_select_admin on public.change_history
for select
using (public.is_admin());

drop policy if exists history_insert_approved on public.change_history;
create policy history_insert_approved on public.change_history
for insert
with check (public.is_approved_user());

drop policy if exists stage_type_select_approved on public.stage_type_catalog;
create policy stage_type_select_approved on public.stage_type_catalog
for select
using (public.is_approved_user());

drop policy if exists stage_type_mutate_admin on public.stage_type_catalog;
create policy stage_type_mutate_admin on public.stage_type_catalog
for all
using (public.is_admin())
with check (public.is_admin());

drop policy if exists lead_source_select_approved on public.lead_source_catalog;
create policy lead_source_select_approved on public.lead_source_catalog
for select
using (public.is_approved_user());

drop policy if exists lead_source_mutate_admin on public.lead_source_catalog;
create policy lead_source_mutate_admin on public.lead_source_catalog
for all
using (public.is_admin())
with check (public.is_admin());

drop policy if exists access_requests_insert_authenticated on public.access_requests;
create policy access_requests_insert_authenticated on public.access_requests
for insert
with check (auth.role() = 'authenticated');

drop policy if exists access_requests_update_admin on public.access_requests;
create policy access_requests_update_admin on public.access_requests
for update
using (public.is_admin())
with check (public.is_admin());

drop policy if exists access_requests_select_admin on public.access_requests;
create policy access_requests_select_admin on public.access_requests
for select
using (public.is_admin());

drop policy if exists admin_requests_insert_approved on public.admin_requests;
create policy admin_requests_insert_approved on public.admin_requests
for insert
with check (public.is_approved_user());

drop policy if exists admin_requests_select_admin on public.admin_requests;
create policy admin_requests_select_admin on public.admin_requests
for select
using (public.is_admin());

drop policy if exists admin_requests_update_admin on public.admin_requests;
create policy admin_requests_update_admin on public.admin_requests
for update
using (public.is_admin())
with check (public.is_admin());
