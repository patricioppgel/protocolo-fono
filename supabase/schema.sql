-- ============================================================
-- Protocolo de Atendimento Fonoaudiológico — schema Supabase
-- ============================================================
-- Modelo: um único "kv_store" que reproduz a mesma API de
-- armazenamento usada no protótipo (get/set/delete/list por
-- chave). Isso evita reescrever toda a lógica do app — só troca
-- o que grava os dados por baixo.
--
-- Por padrão os registros são "compartilhados": qualquer usuário
-- autenticado do time (estagiários + supervisores) vê e edita
-- todos os atendimentos, estudos de caso etc. — como uma ficha
-- clínica de equipe. Isso é intencional para este caso de uso.
-- ============================================================

create table if not exists public.kv_store (
  id bigint generated always as identity primary key,
  owner uuid references auth.users(id) not null default auth.uid(),
  key text not null,
  value text not null,
  shared boolean not null default true,
  updated_at timestamptz not null default now(),
  created_at timestamptz not null default now()
);

-- Uma chave "compartilhada" é única globalmente (todo o time edita o
-- mesmo registro); uma chave "pessoal" é única por usuário.
create unique index if not exists kv_store_shared_key_idx
  on public.kv_store (key) where shared = true;
create unique index if not exists kv_store_personal_key_idx
  on public.kv_store (owner, key) where shared = false;

create index if not exists kv_store_key_prefix_idx on public.kv_store (key text_pattern_ops);

alter table public.kv_store enable row level security;

-- Qualquer usuário autenticado lê tudo que é compartilhado,
-- e também os próprios registros pessoais.
create policy "kv_select" on public.kv_store
  for select using (auth.role() = 'authenticated' and (shared = true or owner = auth.uid()));

create policy "kv_insert" on public.kv_store
  for insert with check (auth.role() = 'authenticated');

-- Qualquer membro do time pode atualizar registros compartilhados
-- (ficha de equipe); registros pessoais só o dono edita.
create policy "kv_update" on public.kv_store
  for update using (auth.role() = 'authenticated' and (shared = true or owner = auth.uid()));

create policy "kv_delete" on public.kv_store
  for delete using (auth.role() = 'authenticated' and (shared = true or owner = auth.uid()));

-- ============================================================
-- (Opcional, recomendado) Perfis — nome/papel de cada usuário,
-- útil para mostrar "quem criou" cada atendimento e para
-- restringir cadastro a e-mails da instituição, se desejar.
-- ============================================================
create table if not exists public.profiles (
  id uuid references auth.users(id) primary key,
  nome text,
  papel text check (papel in ('estagiario', 'supervisor', 'admin')) default 'estagiario',
  created_at timestamptz not null default now()
);

alter table public.profiles enable row level security;

create policy "profiles_select_all_authenticated" on public.profiles
  for select using (auth.role() = 'authenticated');

create policy "profiles_update_own" on public.profiles
  for update using (id = auth.uid());

-- Cria automaticamente um perfil quando alguém se cadastra
create or replace function public.handle_new_user()
returns trigger as $$
begin
  insert into public.profiles (id, nome)
  values (new.id, coalesce(new.raw_user_meta_data->>'nome', new.email));
  return new;
end;
$$ language plpgsql security definer;

drop trigger if exists on_auth_user_created on auth.users;
create trigger on_auth_user_created
  after insert on auth.users
  for each row execute procedure public.handle_new_user();

-- ============================================================
-- Passo a passo para aplicar este schema:
-- 1. No painel do Supabase, vá em "SQL Editor" → "New query"
-- 2. Cole todo este arquivo e clique em "Run"
-- 3. (Opcional) Em "Authentication → Providers", habilite apenas
--    "Email" e, se quiser restringir a cadastros institucionais,
--    em "Authentication → URL Configuration" configure o domínio
--    permitido, ou valide manualmente cada novo usuário.
-- ============================================================
