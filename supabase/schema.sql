create extension if not exists pgcrypto;

create table if not exists public.reports (
    id uuid primary key default gen_random_uuid(),
    driver_name text not null default '',
    plate text not null default '',
    car_model text not null default '',
    date_from date,
    date_to date,
    odometer_start numeric not null default 0,
    odometer_end numeric not null default 0,
    total_km numeric not null default 0,
    trip_purpose text not null default '',
    file_name text not null default '',
    status text not null default 'generated',
    created_at timestamptz not null default now()
);

alter table public.reports enable row level security;

create policy "service_role_only_read_reports" on public.reports
for select using (false);

create policy "service_role_only_insert_reports" on public.reports
for insert with check (false);
