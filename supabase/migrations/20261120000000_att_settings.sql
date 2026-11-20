-- 근태 적법성 점검: 공유 설정 저장소
-- 엑셀 원본·판정 결과는 저장하지 않습니다. 담당자가 지정한 설정만 저장합니다.
--   code:<코드>          근태 코드 의미·근로시간·야간시간
--   holidays             공휴일 목록 (없으면 기본값 사용)
--   site:<시트명>        사업장 설정 (5인 이상, 탄력제, 주휴일 등)
--   emp:<시트명>|<성명>  직원별 감단·교대·초단기 지정

create table if not exists public.att_members (
  email      text primary key,
  created_at timestamptz not null default now()
);

create table if not exists public.att_settings (
  key        text primary key,
  value      jsonb not null,
  updated_at timestamptz not null default now(),
  updated_by uuid references auth.users(id) on delete set null
);

alter table public.att_members  enable row level security;
alter table public.att_settings enable row level security;

-- 로그인한 이메일이 att_members 에 등록돼 있는지
create or replace function public.is_att_member()
returns boolean
language sql
stable
security definer
set search_path = ''
as $$
  select exists (
    select 1 from public.att_members m
    where lower(m.email) = lower(coalesce(auth.jwt() ->> 'email', ''))
  );
$$;
revoke all on function public.is_att_member() from public, anon;
grant execute on function public.is_att_member() to authenticated;

-- 등록된 담당자만 설정을 읽고 쓸 수 있음
drop policy if exists att_settings_select on public.att_settings;
drop policy if exists att_settings_insert on public.att_settings;
drop policy if exists att_settings_update on public.att_settings;
drop policy if exists att_settings_delete on public.att_settings;
create policy att_settings_select on public.att_settings for select to authenticated using (public.is_att_member());
create policy att_settings_insert on public.att_settings for insert to authenticated with check (public.is_att_member());
create policy att_settings_update on public.att_settings for update to authenticated using (public.is_att_member()) with check (public.is_att_member());
create policy att_settings_delete on public.att_settings for delete to authenticated using (public.is_att_member());

-- 본인 등록 여부만 확인 가능 (명단 추가·삭제는 대시보드/SQL 에서)
drop policy if exists att_members_self on public.att_members;
create policy att_members_self on public.att_members for select to authenticated
  using (lower(email) = lower(coalesce(auth.jwt() ->> 'email', '')));

revoke all on public.att_settings, public.att_members from anon;
grant select, insert, update, delete on public.att_settings to authenticated;
grant select on public.att_members to authenticated;

-- 담당자 등록 예시 (실제 이메일로 바꿔 실행)
-- insert into public.att_members (email) values ('담당자@회사.com');
