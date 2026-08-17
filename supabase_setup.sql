-- ============================================================================
--  SKILL INBOUND — DASHBOARD · Configuración completa de Supabase
--  Banco Agrario de Colombia
-- ----------------------------------------------------------------------------
--  Ejecutar en: Supabase → SQL Editor → New query → pegar todo → RUN
--  Es idempotente: puedes volver a ejecutarlo sin romper nada.
--
--  QUÉ RESUELVE
--   · Saca las contraseñas del texto plano → quedan hasheadas en auth.users
--   · Enciende RLS en todas las tablas (hoy están UNRESTRICTED = abiertas)
--   · Migra los usuarios que YA tienes, con sus mismas contraseñas
--   · Deja que UNO cargue los Excel y TODOS los demás vean esa data
--   · Conserva tu comodidad: dar de alta usuarios con una línea de SQL
-- ============================================================================


-- ============================================================================
--  0. EXTENSIONES
-- ============================================================================
create extension if not exists pgcrypto with schema extensions;


-- ============================================================================
--  1. PERFILES — datos NO sensibles (rol, nombre, permisos). Sin contraseñas.
-- ============================================================================
create table if not exists public.perfiles (
  id            uuid primary key references auth.users(id) on delete cascade,
  usuario       text unique not null,            -- 'oscar', '6802889'
  nombre        text not null,
  rol           text not null default 'agente'
                check (rol in ('supervisor','agente')),
  puede_cargar  boolean not null default false,  -- true = puede subir los Excel
  activo        boolean not null default true,
  creado_en     timestamptz not null default now()
);


-- ============================================================================
--  2. CORTES — cada carga de Excel
-- ============================================================================
-- indicadores_def guarda las columnas dinámicas del Excel, por eso Inbound usa
-- AHT, Chat TMO y RRSS TMR sin tener que migrar la base de datos.
create table if not exists public.cortes (
  id               uuid primary key default gen_random_uuid(),
  skill            text not null check (skill in ('inbound','chat','rrss','email')),
  archivo          text,
  periodo          text,
  fecha_corte      date not null,
  indicadores_def  jsonb not null default '[]'::jsonb,
  extras_def       jsonb not null default '[]'::jsonb,
  cargado_por      uuid references auth.users(id) on delete set null,
  creado_en        timestamptz not null default now(),
  unique (skill, fecha_corte)
);
create index if not exists idx_cortes_skill_fecha on public.cortes (skill, fecha_corte desc);


-- ============================================================================
--  3. INDICADORES — una fila por asesor. Valores en jsonb (flexible).
-- ============================================================================
create table if not exists public.indicadores (
  id         bigserial primary key,
  corte_id   uuid not null references public.cortes(id) on delete cascade,
  cedula     text,
  asesor     text not null,
  supervisor text default '',
  valores    jsonb not null default '{}'::jsonb,
  creado_en  timestamptz not null default now()
);
create index if not exists idx_indicadores_corte  on public.indicadores (corte_id);
create index if not exists idx_indicadores_asesor on public.indicadores (asesor);


-- ============================================================================
--  4. PENDIENTES DE TIPIFICACIÓN
-- ============================================================================
create table if not exists public.pendientes_tipificacion (
  id            bigserial primary key,
  agente        text not null,
  id_llamada    text,
  cola          text,
  macroproceso  text,
  proceso       text,
  subproceso_1  text,
  subproceso_2  text,
  observaciones text,
  link_gestion  text,
  fecha_carga   date not null default current_date,
  cargado_por   uuid references auth.users(id) on delete set null,
  creado_en     timestamptz not null default now()
);
create index if not exists idx_pendientes_agente on public.pendientes_tipificacion (agente);
create index if not exists idx_pendientes_fecha  on public.pendientes_tipificacion (fecha_carga desc);


-- ============================================================================
--  5. CALIDAD  (si ya existe, se respeta tal como está)
-- ============================================================================
create table if not exists public.calidad (
  id               bigserial primary key,
  skill            text not null default 'inbound',
  asesor           text not null,
  supervisor       text default '',
  fecha            text default '',
  descripcion      text default '',
  analisis         text default '',
  oportunidad      text default '',
  plan_accion      text default '',
  tipo_alerta      text default '',
  id_llamada       text default '',
  nota             numeric,
  estado           text default 'Pendiente',
  from_compromisos boolean default false,
  firma_asesor     text,
  firma_supervisor text,
  created_at       timestamptz not null default now()
);


-- ============================================================================
--  6. CHAT INTERNO  (estructura real que usa el dashboard)
-- ============================================================================
create table if not exists public.mensajes (
  id           bigserial primary key,
  skill        text,
  tipo         text not null default 'grupo',   -- 'grupo' | 'privado'
  remitente    text not null,
  nombre_rem   text,
  destinatario text,
  nombre_dest  text,
  mensaje      text not null,
  created_at   timestamptz not null default now()
);


-- ============================================================================
--  7. FUNCIÓN DE PERMISOS — ¿el usuario conectado puede escribir?
-- ============================================================================
create or replace function public.puede_cargar()
returns boolean language sql security definer stable set search_path = public as $$
  select coalesce((select p.puede_cargar and p.activo
                     from public.perfiles p where p.id = auth.uid()), false);
$$;
revoke all on function public.puede_cargar() from public;
grant execute on function public.puede_cargar() to authenticated;


-- ============================================================================
--  8. ALTA DE USUARIOS — así reemplazas tu tabla `usuarios`
-- ============================================================================
--  Crear un usuario nuevo es UNA línea desde el SQL Editor:
--
--     select public.crear_usuario('ximena', 'ClaveSegura1*', 'XIMENA', 'supervisor', false);
--
--  La contraseña entra en texto plano aquí pero se guarda hasheada con bcrypt;
--  nadie —ni tú— puede volver a leerla después. Para cambiarla:
--
--     select public.cambiar_clave('ximena', 'NuevaClave1*');

create or replace function public.crear_usuario(
  p_usuario      text,
  p_clave        text,
  p_nombre       text,
  p_rol          text    default 'agente',
  p_puede_cargar boolean default false
) returns uuid
language plpgsql security definer set search_path = public, auth, extensions as $$
declare
  v_id    uuid := gen_random_uuid();
  v_email text := lower(trim(p_usuario)) || '@skillinbound.local';
  v_rol   text := lower(coalesce(p_rol,'agente'));
begin
  if v_rol not in ('supervisor','agente') then v_rol := 'agente'; end if;

  -- Si ya existe, solo actualiza la contraseña y el perfil
  if exists (select 1 from auth.users u where u.email = v_email) then
    select u.id into v_id from auth.users u where u.email = v_email;
    update auth.users
       set encrypted_password = extensions.crypt(p_clave, extensions.gen_salt('bf')),
           updated_at = now()
     where id = v_id;
    update public.perfiles
       set nombre = p_nombre, rol = v_rol, puede_cargar = p_puede_cargar, activo = true
     where id = v_id;
    return v_id;
  end if;

  insert into auth.users (
    instance_id, id, aud, role, email, encrypted_password,
    email_confirmed_at, created_at, updated_at,
    raw_app_meta_data, raw_user_meta_data,
    confirmation_token, recovery_token, email_change_token_new, email_change
  ) values (
    '00000000-0000-0000-0000-000000000000', v_id, 'authenticated', 'authenticated',
    v_email, extensions.crypt(p_clave, extensions.gen_salt('bf')),
    now(), now(), now(),
    '{"provider":"email","providers":["email"]}'::jsonb,
    jsonb_build_object('nombre', p_nombre, 'rol', v_rol, 'puede_cargar', p_puede_cargar),
    '', '', '', ''
  );

  insert into auth.identities (
    id, user_id, provider_id, identity_data, provider,
    last_sign_in_at, created_at, updated_at
  ) values (
    gen_random_uuid(), v_id, v_id::text,
    jsonb_build_object('sub', v_id::text, 'email', v_email),
    'email', now(), now(), now()
  );

  insert into public.perfiles (id, usuario, nombre, rol, puede_cargar)
  values (v_id, lower(trim(p_usuario)), p_nombre, v_rol, p_puede_cargar)
  on conflict (id) do nothing;

  return v_id;
end $$;

revoke all on function public.crear_usuario(text,text,text,text,boolean) from public, anon, authenticated;


create or replace function public.cambiar_clave(p_usuario text, p_clave text)
returns void
language plpgsql security definer set search_path = public, auth, extensions as $$
begin
  update auth.users
     set encrypted_password = extensions.crypt(p_clave, extensions.gen_salt('bf')),
         updated_at = now()
   where email = lower(trim(p_usuario)) || '@skillinbound.local';
end $$;

revoke all on function public.cambiar_clave(text,text) from public, anon, authenticated;


-- ============================================================================
--  9. MIGRACIÓN — pasa los usuarios que YA tienes, con sus mismas contraseñas
-- ============================================================================
--  Recorre tu tabla `usuarios` actual y crea cada uno en Supabase Auth.
--  Marca puede_cargar = true SOLO para 'oscar' (ajusta la lista si necesitas).
--  Si la tabla `usuarios` ya no existe, este bloque se salta solo.

do $$
declare u record;
begin
  if to_regclass('public.usuarios') is null then
    raise notice 'No hay tabla `usuarios` que migrar — se omite.';
    return;
  end if;

  for u in select * from public.usuarios where coalesce(activo, true) loop
    perform public.crear_usuario(
      u.user_id,
      u.password,
      coalesce(u.nombre, u.user_id),
      lower(coalesce(u.rol, 'agente')),          -- corrige 'Agente' → 'agente'
      lower(u.user_id) in ('oscar')              -- ← quién puede cargar los Excel
    );
  end loop;

  raise notice 'Usuarios migrados correctamente.';
end $$;

--  Verifica que quedaron bien ANTES de borrar la tabla vieja:
--     select usuario, nombre, rol, puede_cargar, activo from public.perfiles order by rol, usuario;
--
--  Cuando confirmes que todos están y que puedes entrar al dashboard,
--  ELIMINA la tabla con las contraseñas en texto plano ejecutando:
--
--     drop table if exists public.usuarios cascade;
--
--  (Se deja comentado a propósito: hazlo solo después de verificar.)


-- ============================================================================
--  10. ROW LEVEL SECURITY
-- ============================================================================
--  Esto apaga el estado "UNRESTRICTED" que tienes hoy.
--     LEER    → cualquier usuario autenticado (todos ven lo que tú cargas)
--     ESCRIBIR → solo quien tenga puede_cargar = true
--     Sin sesión iniciada → no se ve absolutamente nada.

alter table public.perfiles                enable row level security;
alter table public.cortes                  enable row level security;
alter table public.indicadores             enable row level security;
alter table public.pendientes_tipificacion enable row level security;
alter table public.calidad                 enable row level security;
alter table public.mensajes                enable row level security;

-- ── PERFILES ────────────────────────────────────────────────────────────────
drop policy if exists perfiles_select on public.perfiles;
create policy perfiles_select on public.perfiles for select to authenticated using (true);

drop policy if exists perfiles_update_propio on public.perfiles;
create policy perfiles_update_propio on public.perfiles for update to authenticated
  using (id = auth.uid()) with check (id = auth.uid());

-- ── CORTES ──────────────────────────────────────────────────────────────────
drop policy if exists cortes_select on public.cortes;
create policy cortes_select on public.cortes for select to authenticated using (true);
drop policy if exists cortes_insert on public.cortes;
create policy cortes_insert on public.cortes for insert to authenticated with check (public.puede_cargar());
drop policy if exists cortes_update on public.cortes;
create policy cortes_update on public.cortes for update to authenticated
  using (public.puede_cargar()) with check (public.puede_cargar());
drop policy if exists cortes_delete on public.cortes;
create policy cortes_delete on public.cortes for delete to authenticated using (public.puede_cargar());

-- ── INDICADORES ─────────────────────────────────────────────────────────────
drop policy if exists indicadores_select on public.indicadores;
create policy indicadores_select on public.indicadores for select to authenticated using (true);
drop policy if exists indicadores_insert on public.indicadores;
create policy indicadores_insert on public.indicadores for insert to authenticated with check (public.puede_cargar());
drop policy if exists indicadores_update on public.indicadores;
create policy indicadores_update on public.indicadores for update to authenticated
  using (public.puede_cargar()) with check (public.puede_cargar());
drop policy if exists indicadores_delete on public.indicadores;
create policy indicadores_delete on public.indicadores for delete to authenticated using (public.puede_cargar());

-- ── PENDIENTES DE TIPIFICACIÓN ──────────────────────────────────────────────
drop policy if exists pendientes_select on public.pendientes_tipificacion;
create policy pendientes_select on public.pendientes_tipificacion for select to authenticated using (true);
drop policy if exists pendientes_insert on public.pendientes_tipificacion;
create policy pendientes_insert on public.pendientes_tipificacion for insert to authenticated with check (public.puede_cargar());
drop policy if exists pendientes_delete on public.pendientes_tipificacion;
create policy pendientes_delete on public.pendientes_tipificacion for delete to authenticated using (public.puede_cargar());

-- ── CALIDAD — los supervisores registran retroalimentaciones ────────────────
drop policy if exists calidad_select on public.calidad;
create policy calidad_select on public.calidad for select to authenticated using (true);
drop policy if exists calidad_insert on public.calidad;
create policy calidad_insert on public.calidad for insert to authenticated with check (
  exists (select 1 from public.perfiles p where p.id = auth.uid() and p.rol = 'supervisor' and p.activo));
drop policy if exists calidad_update on public.calidad;
create policy calidad_update on public.calidad for update to authenticated using (
  exists (select 1 from public.perfiles p where p.id = auth.uid() and p.rol = 'supervisor' and p.activo));
drop policy if exists calidad_delete on public.calidad;
create policy calidad_delete on public.calidad for delete to authenticated using (public.puede_cargar());

-- ── MENSAJES (chat) ─────────────────────────────────────────────────────────
drop policy if exists mensajes_select on public.mensajes;
create policy mensajes_select on public.mensajes for select to authenticated using (true);
drop policy if exists mensajes_insert on public.mensajes;
create policy mensajes_insert on public.mensajes for insert to authenticated with check (true);


-- ============================================================================
--  11. REALTIME (chat en vivo)
-- ============================================================================
do $$
begin
  if not exists (select 1 from pg_publication_tables
                  where pubname='supabase_realtime' and schemaname='public' and tablename='mensajes')
  then alter publication supabase_realtime add table public.mensajes;
  end if;
end $$;


-- ============================================================================
--  VERIFICACIÓN FINAL — ejecuta esto aparte
-- ============================================================================
--  1) Que RLS quedó encendido en todas (debe decir true):
--       select tablename, rowsecurity as rls_activo
--         from pg_tables where schemaname='public' order by tablename;
--
--  2) Que los usuarios migraron bien:
--       select usuario, nombre, rol, puede_cargar, activo
--         from public.perfiles order by rol, usuario;
--
--  3) Entra al dashboard con tu usuario y clave de siempre (oscar / super123).
--     Si entra y ves los datos, ya puedes borrar la tabla vieja:
--       drop table if exists public.usuarios cascade;
--
--  ── Uso diario a partir de ahora ──
--     Crear usuario:     select public.crear_usuario('6804111','Gomitas1*','PEREZ JUAN','agente',false);
--     Cambiar clave:     select public.cambiar_clave('ximena','NuevaClave1*');
--     Dar permiso carga: update public.perfiles set puede_cargar=true  where usuario='ximena';
--     Desactivar:        update public.perfiles set activo=false       where usuario='6804111';
-- ============================================================================
