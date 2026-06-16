create or replace function public.delete_leads_by_ids(target_ids uuid[])
returns table (deleted_id uuid)
language plpgsql
security definer
set search_path = public
as $$
begin
  if not public.is_approved_user() then
    raise exception 'Acesso nao aprovado.';
  end if;

  if not public.is_admin() then
    raise exception 'Somente administradores podem excluir leads.';
  end if;

  if coalesce(array_length(target_ids, 1), 0) = 0 then
    return;
  end if;

  return query
  with deleted_rows as (
    delete from public.leads
    where id = any(target_ids)
    returning id
  )
  select deleted_rows.id
  from deleted_rows;
end;
$$;

revoke all on function public.delete_leads_by_ids(uuid[]) from public;
grant execute on function public.delete_leads_by_ids(uuid[]) to authenticated;
