-- A run has one payment workbook and one or more independently retained GCash workbooks.
alter table public.verification_source_workbooks add column source_index integer not null default 0 check (source_index >= 0);
alter table public.verification_source_workbooks drop constraint verification_source_workbooks_pkey;
alter table public.verification_source_workbooks add primary key (verification_run_id, kind, source_index);

create function public.create_multi_gcash_import(
  p_business_id uuid, p_session_id uuid, p_payments jsonb, p_gcash jsonb,
  p_payment_source jsonb, p_gcash_sources jsonb
) returns uuid language plpgsql security definer set search_path='' as $$
declare run_id uuid; expected_sources integer;
begin
  if auth.uid() is null or private.member_role(p_business_id, auth.uid()) is null then
    raise exception 'not authorized' using errcode='42501';
  end if;
  if jsonb_typeof(p_gcash_sources) <> 'array' or jsonb_array_length(p_gcash_sources) = 0 then
    raise exception 'at least one GCash source workbook is required' using errcode='22023';
  end if;
  if exists (select 1 from jsonb_array_elements(p_gcash_sources) source where jsonb_typeof(source) <> 'object') then
    raise exception 'invalid GCash source workbook' using errcode='22023';
  end if;

  run_id := public.create_dual_file_import(p_business_id, p_session_id, p_payments, p_gcash);
  expected_sources := 1 + jsonb_array_length(p_gcash_sources);
  if exists (select 1 from public.verification_source_workbooks where verification_run_id=run_id and business_id=p_business_id) then
    if (select count(*) from public.verification_source_workbooks where verification_run_id=run_id and business_id=p_business_id) <> expected_sources
      or not exists (select 1 from public.verification_source_workbooks where verification_run_id=run_id and kind='payments' and source_index=0 and source=p_payment_source)
      or exists (select 1 from jsonb_array_elements(p_gcash_sources) with ordinality source(value, ordinality)
        where not exists (select 1 from public.verification_source_workbooks s where s.verification_run_id=run_id and s.kind='gcash' and s.source_index=source.ordinality-1 and s.source=source.value)) then
      raise exception 'source files changed for this import session' using errcode='22023';
    end if;
  else
    insert into public.verification_source_workbooks (verification_run_id,business_id,kind,source_index,source)
      values (run_id,p_business_id,'payments',0,p_payment_source);
    insert into public.verification_source_workbooks (verification_run_id,business_id,kind,source_index,source)
      select run_id,p_business_id,'gcash',ordinality-1,value
      from jsonb_array_elements(p_gcash_sources) with ordinality;
  end if;
  return run_id;
end;
$$;
revoke all on function public.create_multi_gcash_import(uuid,uuid,jsonb,jsonb,jsonb,jsonb) from public, anon;
grant execute on function public.create_multi_gcash_import(uuid,uuid,jsonb,jsonb,jsonb,jsonb) to authenticated;

create or replace function public.get_verification_export(p_business_id uuid, p_run_id uuid)
returns jsonb language plpgsql security invoker set search_path='' as $$
declare payload jsonb; source_rows bigint;
begin
  if auth.uid() is null or private.member_role(p_business_id,auth.uid()) is null or not exists (
   select 1 from public.verification_runs where id=p_run_id and business_id=p_business_id and status='succeeded'
  ) then raise exception 'completed run not found' using errcode='42501'; end if;
  select (select count(*) from public.payments where business_id=p_business_id and verification_run_id=p_run_id)
   +(select count(*) from public.gcash_transactions where business_id=p_business_id and verification_run_id=p_run_id) into source_rows;
  if source_rows>20000 then raise exception 'Export supports at most 20,000 combined source rows.' using errcode='PT413'; end if;
  select jsonb_build_object(
   'workspaceName',(select name from public.businesses where id=p_business_id),
   'run',(select to_jsonb(r) from public.verification_run_summaries r where r.id=p_run_id and r.business_id=p_business_id),
   'payments',(select coalesce(jsonb_agg(to_jsonb(p) order by p.row_index,p.id),'[]'::jsonb) from public.payment_results p where p.business_id=p_business_id and p.verification_run_id=p_run_id),
   'gcash',(select coalesce(jsonb_agg(to_jsonb(g) order by g.source_order,g.row_index,g.id),'[]'::jsonb) from public.gcash_results g where g.business_id=p_business_id and g.verification_run_id=p_run_id),
   'sources',jsonb_build_object(
     'payments',(select source from public.verification_source_workbooks s where s.business_id=p_business_id and s.verification_run_id=p_run_id and s.kind='payments' and s.source_index=0),
     'gcashSources',(select coalesce(jsonb_agg(jsonb_build_object('sourceIndex',s.source_index,'source',s.source) order by s.source_index),'[]'::jsonb) from public.verification_source_workbooks s where s.business_id=p_business_id and s.verification_run_id=p_run_id and s.kind='gcash')
   )
  ) into payload;
  if jsonb_array_length(payload->'payments') <> (select count(*) from public.payments where business_id=p_business_id and verification_run_id=p_run_id) then
   raise exception 'Run contains payments without persisted reconciliation results.' using errcode='22023'; end if;
  if octet_length(payload::text)>33554432 then raise exception 'Export data exceeds the 32 MB synchronous export limit.' using errcode='PT413'; end if;
  return payload;
end;
$$;
