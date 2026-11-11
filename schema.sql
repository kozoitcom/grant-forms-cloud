create table if not exists forms (
 id uuid primary key default gen_random_uuid(),
 title text not null,
 source_name text,
 fields jsonb not null,
 published boolean not null default false,
 created_at timestamptz not null default now()
);
create table if not exists submissions (
 id uuid primary key default gen_random_uuid(),
 form_id uuid not null references forms(id),
 answers jsonb not null,
 created_at timestamptz not null default now()
);
create table if not exists attachments (
 id uuid primary key default gen_random_uuid(),
 submission_id uuid not null references submissions(id),
 field_id text not null,
 object_path text not null,
 filename text not null
);
alter table forms enable row level security;
alter table submissions enable row level security;
alter table attachments enable row level security;
insert into storage.buckets (id,name,public) values ('application-documents','application-documents',false) on conflict (id) do nothing;
