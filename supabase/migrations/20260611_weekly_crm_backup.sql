create extension if not exists pgcrypto;
create extension if not exists pg_cron;
create extension if not exists pg_net;

create table if not exists public.crm_backup_runs (
  id uuid primary key default gen_random_uuid(),
  trigger_source text not null default 'scheduled',
  status text not null default 'running',
  storage_path text,
  file_size_bytes bigint,
  row_counts jsonb not null default '{}'::jsonb,
  error_message text,
  started_at timestamptz not null default timezone('utc'::text, now()),
  finished_at timestamptz
);

alter table public.crm_backup_runs
  drop constraint if exists crm_backup_runs_status_check;

alter table public.crm_backup_runs
  add constraint crm_backup_runs_status_check
  check (status in ('running', 'success', 'error'));

alter table if exists public.crm_backup_runs enable row level security;

drop policy if exists crm_backup_runs_select_admin on public.crm_backup_runs;

create policy crm_backup_runs_select_admin on public.crm_backup_runs
for select
using (public.is_admin());

do $$
declare
  existing_job record;
begin
  for existing_job in
    select jobid
    from cron.job
    where jobname = 'weekly-crm-backup'
  loop
    perform cron.unschedule(existing_job.jobid);
  end loop;
end
$$;

select
  cron.schedule(
    'weekly-crm-backup',
    '0 9 * * 1',
    $job$
    select
      net.http_post(
        url := 'https://qxuuladntzrojngvdfil.supabase.co/functions/v1/weekly-crm-backup',
        headers := jsonb_build_object(
          'Content-Type', 'application/json',
          'Authorization', 'Bearer sb_publishable_oepxAXvlNxop_3tITuiw3Q_yGgWPcmf'
        ),
        body := jsonb_build_object(
          'trigger', 'weekly_cron'
        )
      );
    $job$
  );
