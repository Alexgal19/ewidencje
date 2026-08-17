-- ══════════════════════════════════════════════════════════════════════════
-- Ewidencja Przebiegu Pojazdu – Supabase schema
-- Uruchom cały plik w: Supabase Dashboard → SQL Editor → New query → Run
-- ══════════════════════════════════════════════════════════════════════════

-- ── Pojazdy ────────────────────────────────────────────────────────────────
create table if not exists public.vehicles (
    id         uuid primary key default gen_random_uuid(),
    user_id    uuid not null references auth.users (id) on delete cascade,
    plate      text not null,
    make       text,
    model      text,
    created_at timestamptz not null default now()
);

-- ── Kierowcy ───────────────────────────────────────────────────────────────
create table if not exists public.drivers (
    id         uuid primary key default gen_random_uuid(),
    user_id    uuid not null references auth.users (id) on delete cascade,
    name       text not null,
    created_at timestamptz not null default now()
);

-- ── Historia wygenerowanych ewidencji ──────────────────────────────────────
create table if not exists public.ewidencje (
    id              uuid primary key default gen_random_uuid(),
    user_id         uuid not null references auth.users (id) on delete cascade,
    vehicle_id      uuid references public.vehicles (id) on delete set null,
    driver_id       uuid references public.drivers (id) on delete set null,
    plate           text,
    car_model       text,
    driver_name     text,
    period_start    date,
    period_end      date,
    odometer_start  integer,
    odometer_end    integer,
    total_km        numeric(12, 2),
    file_name       text,
    file_path       text,
    created_at      timestamptz not null default now()
);

create index if not exists ewidencje_user_created_idx on public.ewidencje (user_id, created_at desc);
create index if not exists vehicles_user_idx on public.vehicles (user_id);
create index if not exists drivers_user_idx on public.drivers (user_id);

-- ── Row Level Security (każdy użytkownik widzi tylko swoje dane) ────────────
alter table public.vehicles  enable row level security;
alter table public.drivers   enable row level security;
alter table public.ewidencje enable row level security;

drop policy if exists "vehicles_user_all"  on public.vehicles;
drop policy if exists "drivers_user_all"   on public.drivers;
drop policy if exists "ewidencje_user_all" on public.ewidencje;

create policy "vehicles_user_all" on public.vehicles
    for all to authenticated
    using (auth.uid() = user_id)
    with check (auth.uid() = user_id);

create policy "drivers_user_all" on public.drivers
    for all to authenticated
    using (auth.uid() = user_id)
    with check (auth.uid() = user_id);

create policy "ewidencje_user_all" on public.ewidencje
    for all to authenticated
    using (auth.uid() = user_id)
    with check (auth.uid() = user_id);

-- ── Storage (bucket na pliki Excel, prywatny) ──────────────────────────────
insert into storage.buckets (id, name, public)
values ('ewidencje', 'ewidencje', false)
on conflict (id) do nothing;

-- Pliki leżą w folderze {user_id}/... – użytkownik może zarządzać swoim folderem
drop policy if exists "ewidencje_storage_user" on storage.objects;

create policy "ewidencje_storage_user" on storage.objects
    for all to authenticated
    using (
        bucket_id = 'ewidencje'
        and auth.uid()::text = (storage.foldername(name))[1]
    )
    with check (
        bucket_id = 'ewidencje'
        and auth.uid()::text = (storage.foldername(name))[1]
    );
