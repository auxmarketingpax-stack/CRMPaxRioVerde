drop policy if exists access_requests_insert_authenticated on public.access_requests;
drop policy if exists access_requests_insert_public on public.access_requests;

create policy access_requests_insert_public on public.access_requests
for insert
with check (
  status = 'pending'
  and coalesce(approved_role, 'user') = 'user'
  and reviewed_by is null
  and reviewed_at is null
  and (
    (auth.role() = 'authenticated' and auth.uid() = auth_user_id)
    or (auth.role() = 'anon' and auth_user_id is null)
  )
);
