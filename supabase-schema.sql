-- 在 Supabase SQL Editor 中执行本文件。
create extension if not exists pgcrypto;

create table if not exists public.templates (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null references auth.users(id) on delete cascade,
  name text not null,
  project text not null default '',
  data jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now()
);

create table if not exists public.history (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null references auth.users(id) on delete cascade,
  filename text not null default '',
  template_name text not null default '',
  rows integer not null default 0 check (rows >= 0),
  cols integer not null default 0 check (cols >= 0),
  created_at timestamptz not null default now()
);

create index if not exists templates_created_at_idx on public.templates (created_at desc);
create index if not exists templates_user_id_idx on public.templates (user_id);
create index if not exists history_user_created_at_idx on public.history (user_id, created_at desc);

alter table public.templates enable row level security;
alter table public.history enable row level security;

-- 已登录用户可读取团队共享模板；管理员邮箱或经 GitHub OAuth 验证的管理员账号可写入模板。
drop policy if exists "authenticated users read templates" on public.templates;
create policy "authenticated users read templates" on public.templates
for select to authenticated using (true);

drop policy if exists "admin inserts templates" on public.templates;
create policy "admin inserts templates" on public.templates
for insert to authenticated
with check (
  auth.uid() = user_id
  and (
    lower(coalesce(auth.jwt() ->> 'email', '')) = lower('hukehuan@bytedance.com')
    or (
      lower(coalesce(auth.jwt() -> 'app_metadata' ->> 'provider', '')) = 'github'
      and lower(coalesce(auth.jwt() -> 'user_metadata' ->> 'user_name', '')) = 'kewenqueen'
    )
  )
);

drop policy if exists "admin updates templates" on public.templates;
create policy "admin updates templates" on public.templates
for update to authenticated
using (
  lower(coalesce(auth.jwt() ->> 'email', '')) = lower('hukehuan@bytedance.com')
  or (
    lower(coalesce(auth.jwt() -> 'app_metadata' ->> 'provider', '')) = 'github'
    and lower(coalesce(auth.jwt() -> 'user_metadata' ->> 'user_name', '')) = 'kewenqueen'
  )
)
with check (
  lower(coalesce(auth.jwt() ->> 'email', '')) = lower('hukehuan@bytedance.com')
  or (
    lower(coalesce(auth.jwt() -> 'app_metadata' ->> 'provider', '')) = 'github'
    and lower(coalesce(auth.jwt() -> 'user_metadata' ->> 'user_name', '')) = 'kewenqueen'
  )
);

drop policy if exists "admin deletes templates" on public.templates;
create policy "admin deletes templates" on public.templates
for delete to authenticated
using (
  lower(coalesce(auth.jwt() ->> 'email', '')) = lower('hukehuan@bytedance.com')
  or (
    lower(coalesce(auth.jwt() -> 'app_metadata' ->> 'provider', '')) = 'github'
    and lower(coalesce(auth.jwt() -> 'user_metadata' ->> 'user_name', '')) = 'kewenqueen'
  )
);

-- 历史记录严格按当前账号隔离。
drop policy if exists "users read own history" on public.history;
create policy "users read own history" on public.history
for select to authenticated using (auth.uid() = user_id);

drop policy if exists "users insert own history" on public.history;
create policy "users insert own history" on public.history
for insert to authenticated with check (auth.uid() = user_id);

drop policy if exists "users delete own history" on public.history;
create policy "users delete own history" on public.history
for delete to authenticated using (auth.uid() = user_id);
