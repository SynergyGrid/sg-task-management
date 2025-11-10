-- TaskFlow Pro Supabase schema and RLS policies
-- Execute inside an empty Supabase project (SQL editor or migration).

begin;

create extension if not exists "pgcrypto";
create extension if not exists "citext";

create or replace function public.touch_updated_at()
returns trigger
language plpgsql
as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

create table if not exists public.profiles (
  id uuid primary key references auth.users(id) on delete cascade,
  full_name text,
  avatar_url text,
  timezone text default 'UTC',
  created_at timestamptz default now(),
  updated_at timestamptz default now()
);

drop trigger if exists profiles_set_updated_at on public.profiles;
create trigger profiles_set_updated_at
before update on public.profiles
for each row
execute procedure public.touch_updated_at();

create table if not exists public.workspaces (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  slug citext unique,
  created_by uuid references public.profiles(id) on delete set null,
  created_at timestamptz default now(),
  updated_at timestamptz default now()
);

drop trigger if exists workspaces_set_updated_at on public.workspaces;
create trigger workspaces_set_updated_at
before update on public.workspaces
for each row
execute procedure public.touch_updated_at();

create table if not exists public.workspace_members (
  workspace_id uuid not null references public.workspaces(id) on delete cascade,
  profile_id uuid not null references public.profiles(id) on delete cascade,
  role text not null check (role in ('owner', 'admin', 'member', 'viewer')) default 'member',
  invited_by uuid references public.profiles(id) on delete set null,
  created_at timestamptz default now(),
  updated_at timestamptz default now(),
  primary key (workspace_id, profile_id)
);

drop trigger if exists workspace_members_set_updated_at on public.workspace_members;
create trigger workspace_members_set_updated_at
before update on public.workspace_members
for each row
execute procedure public.touch_updated_at();

create index if not exists workspace_members_profile_idx
  on public.workspace_members(profile_id);

create table if not exists public.departments (
  id uuid primary key default gen_random_uuid(),
  workspace_id uuid not null references public.workspaces(id) on delete cascade,
  name text not null,
  color text,
  is_default boolean default false,
  created_by uuid references public.profiles(id) on delete set null,
  created_at timestamptz default now(),
  updated_at timestamptz default now()
);

drop trigger if exists departments_set_updated_at on public.departments;
create trigger departments_set_updated_at
before update on public.departments
for each row
execute procedure public.touch_updated_at();

create unique index if not exists departments_name_unique
  on public.departments(workspace_id, lower(name));

create table if not exists public.members (
  id uuid primary key default gen_random_uuid(),
  workspace_id uuid not null references public.workspaces(id) on delete cascade,
  profile_id uuid references public.profiles(id) on delete set null,
  department_id uuid references public.departments(id) on delete set null,
  display_name text not null,
  title text,
  email text,
  avatar_url text,
  created_at timestamptz default now(),
  updated_at timestamptz default now()
);

drop trigger if exists members_set_updated_at on public.members;
create trigger members_set_updated_at
before update on public.members
for each row
execute procedure public.touch_updated_at();

create index if not exists members_workspace_idx
  on public.members(workspace_id);

create table if not exists public.projects (
  id uuid primary key default gen_random_uuid(),
  workspace_id uuid not null references public.workspaces(id) on delete cascade,
  name text not null,
  color text,
  is_default boolean default false,
  created_by uuid references public.profiles(id) on delete set null,
  created_at timestamptz default now(),
  updated_at timestamptz default now()
);

drop trigger if exists projects_set_updated_at on public.projects;
create trigger projects_set_updated_at
before update on public.projects
for each row
execute procedure public.touch_updated_at();

create unique index if not exists projects_name_unique
  on public.projects(workspace_id, lower(name));

create table if not exists public.sub_projects (
  id uuid primary key default gen_random_uuid(),
  workspace_id uuid not null references public.workspaces(id) on delete cascade,
  project_id uuid not null references public.projects(id) on delete cascade,
  name text not null,
  created_by uuid references public.profiles(id) on delete set null,
  created_at timestamptz default now(),
  updated_at timestamptz default now()
);

drop trigger if exists sub_projects_set_updated_at on public.sub_projects;
create trigger sub_projects_set_updated_at
before update on public.sub_projects
for each row
execute procedure public.touch_updated_at();

create index if not exists sub_projects_project_idx
  on public.sub_projects(project_id);

create table if not exists public.sections (
  id uuid primary key default gen_random_uuid(),
  workspace_id uuid not null references public.workspaces(id) on delete cascade,
  project_id uuid not null references public.projects(id) on delete cascade,
  name text not null,
  sort_order integer default 0,
  created_by uuid references public.profiles(id) on delete set null,
  created_at timestamptz default now(),
  updated_at timestamptz default now()
);

drop trigger if exists sections_set_updated_at on public.sections;
create trigger sections_set_updated_at
before update on public.sections
for each row
execute procedure public.touch_updated_at();

create index if not exists sections_project_idx
  on public.sections(project_id, sort_order);

create table if not exists public.tasks (
  id uuid primary key default gen_random_uuid(),
  workspace_id uuid not null references public.workspaces(id) on delete cascade,
  project_id uuid not null references public.projects(id) on delete cascade,
  sub_project_id uuid references public.sub_projects(id) on delete set null,
  section_id uuid references public.sections(id) on delete set null,
  title text not null,
  description text,
  due_date date,
  priority text check (priority in ('critical', 'very-high', 'high', 'medium', 'low', 'optional')) default 'medium',
  assignee_id uuid references public.members(id) on delete set null,
  department_id uuid references public.departments(id) on delete set null,
  completed boolean default false,
  completed_at timestamptz,
  created_by uuid references public.profiles(id) on delete set null,
  updated_by uuid references public.profiles(id) on delete set null,
  created_at timestamptz default now(),
  updated_at timestamptz default now()
);

drop trigger if exists tasks_set_updated_at on public.tasks;
create trigger tasks_set_updated_at
before update on public.tasks
for each row
execute procedure public.touch_updated_at();

create index if not exists tasks_workspace_idx
  on public.tasks(workspace_id);

create index if not exists tasks_project_idx
  on public.tasks(project_id);

create index if not exists tasks_section_idx
  on public.tasks(section_id);

create index if not exists tasks_assignee_idx
  on public.tasks(assignee_id);

create table if not exists public.attachments (
  id uuid primary key default gen_random_uuid(),
  workspace_id uuid not null references public.workspaces(id) on delete cascade,
  task_id uuid not null references public.tasks(id) on delete cascade,
  storage_path text not null,
  name text,
  mime_type text,
  size_bytes integer,
  uploaded_by uuid references public.profiles(id) on delete set null,
  uploaded_at timestamptz default now()
);

create index if not exists attachments_task_idx
  on public.attachments(task_id);

create table if not exists public.settings (
  profile_id uuid not null references public.profiles(id) on delete cascade,
  workspace_id uuid not null references public.workspaces(id) on delete cascade,
  data jsonb not null default '{}'::jsonb,
  created_at timestamptz default now(),
  updated_at timestamptz default now(),
  primary key (profile_id, workspace_id)
);

drop trigger if exists settings_set_updated_at on public.settings;
create trigger settings_set_updated_at
before update on public.settings
for each row
execute procedure public.touch_updated_at();

-- Row level security ------------------------------------------------------

alter table public.profiles force row level security;
alter table public.workspaces force row level security;
alter table public.workspace_members force row level security;
alter table public.departments force row level security;
alter table public.members force row level security;
alter table public.projects force row level security;
alter table public.sub_projects force row level security;
alter table public.sections force row level security;
alter table public.tasks force row level security;
alter table public.attachments force row level security;
alter table public.settings force row level security;

create policy "Users can read own profile"
  on public.profiles
  for select
  using (auth.uid() = id);

create policy "Users can update own profile"
  on public.profiles
  for update
  using (auth.uid() = id)
  with check (auth.uid() = id);

create policy "Workspace visibility"
  on public.workspaces
  for select
  using (
    auth.uid() is not null
    and exists (
      select 1
      from public.workspace_members m
      where m.workspace_id = workspaces.id
        and m.profile_id = auth.uid()
    )
  );

create policy "Workspace owners manage workspaces"
  on public.workspaces
  for all
  using (
    auth.uid() is not null
    and exists (
      select 1
      from public.workspace_members m
      where m.workspace_id = workspaces.id
        and m.profile_id = auth.uid()
        and m.role in ('owner', 'admin')
    )
  )
  with check (
    auth.uid() is not null
    and exists (
      select 1
      from public.workspace_members m
      where m.workspace_id = workspaces.id
        and m.profile_id = auth.uid()
        and m.role in ('owner', 'admin')
    )
  );

create policy "Members can read workspace_members"
  on public.workspace_members
  for select
  using (
    auth.uid() is not null
    and exists (
      select 1
      from public.workspace_members m2
      where m2.workspace_id = workspace_members.workspace_id
        and m2.profile_id = auth.uid()
    )
  );

create policy "Users can join their workspace"
  on public.workspace_members
  for insert
  with check (
    auth.uid() = profile_id
    and auth.uid() is not null
  );

create policy "Members manage own membership"
  on public.workspace_members
  for update
  using (auth.uid() = profile_id)
  with check (auth.uid() = profile_id);

create policy "Members can leave workspace"
  on public.workspace_members
  for delete
  using (auth.uid() = profile_id);

create policy "Members can read departments"
  on public.departments
  for select
  using (
    auth.uid() is not null
    and exists (
      select 1
      from public.workspace_members m
      where m.workspace_id = departments.workspace_id
        and m.profile_id = auth.uid()
    )
  );

create policy "Members manage departments"
  on public.departments
  for all
  using (
    auth.uid() is not null
    and exists (
      select 1
      from public.workspace_members m
      where m.workspace_id = departments.workspace_id
        and m.profile_id = auth.uid()
    )
  )
  with check (
    auth.uid() is not null
    and exists (
      select 1
      from public.workspace_members m
      where m.workspace_id = departments.workspace_id
        and m.profile_id = auth.uid()
    )
  );

create policy "Members can read people directory"
  on public.members
  for select
  using (
    auth.uid() is not null
    and exists (
      select 1
      from public.workspace_members m
      where m.workspace_id = members.workspace_id
        and m.profile_id = auth.uid()
    )
  );

create policy "Members manage people directory"
  on public.members
  for all
  using (
    auth.uid() is not null
    and exists (
      select 1
      from public.workspace_members m
      where m.workspace_id = members.workspace_id
        and m.profile_id = auth.uid()
    )
  )
  with check (
    auth.uid() is not null
    and exists (
      select 1
      from public.workspace_members m
      where m.workspace_id = members.workspace_id
        and m.profile_id = auth.uid()
    )
  );

create policy "Members read projects"
  on public.projects
  for select
  using (
    auth.uid() is not null
    and exists (
      select 1
      from public.workspace_members m
      where m.workspace_id = projects.workspace_id
        and m.profile_id = auth.uid()
    )
  );

create policy "Members manage projects"
  on public.projects
  for all
  using (
    auth.uid() is not null
    and exists (
      select 1
      from public.workspace_members m
      where m.workspace_id = projects.workspace_id
        and m.profile_id = auth.uid()
    )
  )
  with check (
    auth.uid() is not null
    and exists (
      select 1
      from public.workspace_members m
      where m.workspace_id = projects.workspace_id
        and m.profile_id = auth.uid()
    )
  );

create policy "Members read sub-projects"
  on public.sub_projects
  for select
  using (
    auth.uid() is not null
    and exists (
      select 1
      from public.workspace_members m
      where m.workspace_id = sub_projects.workspace_id
        and m.profile_id = auth.uid()
    )
  );

create policy "Members manage sub-projects"
  on public.sub_projects
  for all
  using (
    auth.uid() is not null
    and exists (
      select 1
      from public.workspace_members m
      where m.workspace_id = sub_projects.workspace_id
        and m.profile_id = auth.uid()
    )
  )
  with check (
    auth.uid() is not null
    and exists (
      select 1
      from public.workspace_members m
      where m.workspace_id = sub_projects.workspace_id
        and m.profile_id = auth.uid()
    )
  );

create policy "Members read sections"
  on public.sections
  for select
  using (
    auth.uid() is not null
    and exists (
      select 1
      from public.workspace_members m
      where m.workspace_id = sections.workspace_id
        and m.profile_id = auth.uid()
    )
  );

create policy "Members manage sections"
  on public.sections
  for all
  using (
    auth.uid() is not null
    and exists (
      select 1
      from public.workspace_members m
      where m.workspace_id = sections.workspace_id
        and m.profile_id = auth.uid()
    )
  )
  with check (
    auth.uid() is not null
    and exists (
      select 1
      from public.workspace_members m
      where m.workspace_id = sections.workspace_id
        and m.profile_id = auth.uid()
    )
  );

create policy "Members read tasks"
  on public.tasks
  for select
  using (
    auth.uid() is not null
    and exists (
      select 1
      from public.workspace_members m
      where m.workspace_id = tasks.workspace_id
        and m.profile_id = auth.uid()
    )
  );

create policy "Members manage tasks"
  on public.tasks
  for all
  using (
    auth.uid() is not null
    and exists (
      select 1
      from public.workspace_members m
      where m.workspace_id = tasks.workspace_id
        and m.profile_id = auth.uid()
    )
  )
  with check (
    auth.uid() is not null
    and exists (
      select 1
      from public.workspace_members m
      where m.workspace_id = tasks.workspace_id
        and m.profile_id = auth.uid()
    )
  );

create policy "Members read attachments"
  on public.attachments
  for select
  using (
    auth.uid() is not null
    and exists (
      select 1
      from public.workspace_members m
      where m.workspace_id = attachments.workspace_id
        and m.profile_id = auth.uid()
    )
  );

create policy "Members manage attachments"
  on public.attachments
  for all
  using (
    auth.uid() is not null
    and exists (
      select 1
      from public.workspace_members m
      where m.workspace_id = attachments.workspace_id
        and m.profile_id = auth.uid()
    )
  )
  with check (
    auth.uid() is not null
    and exists (
      select 1
      from public.workspace_members m
      where m.workspace_id = attachments.workspace_id
        and m.profile_id = auth.uid()
    )
  );

create policy "Users read their settings"
  on public.settings
  for select
  using (
    auth.uid() = settings.profile_id
  );

create policy "Users manage their settings"
  on public.settings
  for all
  using (
    auth.uid() = settings.profile_id
  )
  with check (
    auth.uid() = settings.profile_id
  );

commit;
