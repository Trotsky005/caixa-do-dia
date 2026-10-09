-- Caixa do Dia — estrutura do banco no Supabase.
-- Cole tudo no Supabase › SQL Editor › New query e clique em Run (só uma vez).

create table if not exists lancamentos (
  id text primary key, seq bigserial,
  frente text, data text, hora text, tipo text, sentido text, forma text,
  valor numeric, descricao text, excluido boolean default false,
  dados jsonb not null
);
create index if not exists lancamentos_frente_data on lancamentos (frente, data);

create table if not exists fechamentos (
  id text primary key, seq bigserial,
  frente text, data text, status text, fechado_por text,
  contado numeric, diferenca_dinheiro numeric, pdf text,
  dados jsonb not null
);

create table if not exists contas_pagar (
  id text primary key, seq bigserial,
  vencimento text, fornecedor text, valor numeric, status text, loja text,
  dados jsonb not null
);

create table if not exists boletos_receber (
  id text primary key, seq bigserial,
  vencimento text, cliente text, valor numeric, status text, loja text,
  dados jsonb not null
);

create table if not exists contas_fixas (
  id text primary key, seq bigserial,
  fornecedor text, valor numeric, dia int, ativa boolean,
  dados jsonb not null
);

create table if not exists cheques_recebidos (
  id text primary key, seq bigserial,
  bom_para text, numero text, emitente text, valor numeric, status text,
  dados jsonb not null
);

create table if not exists cheques_emitidos (
  id text primary key, seq bigserial,
  bom_para text, numero text, nominal text, valor numeric, status text,
  dados jsonb not null
);

create table if not exists config (
  chave text primary key,
  valor text
);

-- Senhas de acesso ao site: edite a coluna "senha" pelo Table Editor.
create table if not exists senhas (
  perfil text primary key,
  acesso text not null,
  senha text default ''
);
insert into senhas (perfil, acesso, senha) values
  ('admin', 'Administrador', ''),
  ('cx-areias', 'Caixa Areias', ''),
  ('fin-areias', 'Financeiro Areias', ''),
  ('cx-barreiro', 'Caixa Barreiro', ''),
  ('fin-barreiro', 'Financeiro Barreiro', ''),
  ('banco', 'Banco', '')
on conflict (perfil) do nothing;

create table if not exists sessoes (
  token text primary key,
  perfil text not null,
  marca text not null,
  expira timestamptz not null
);

-- Segurança: ninguém lê as tabelas direto pela internet; só a função "api" acessa.
alter table lancamentos       enable row level security;
alter table fechamentos       enable row level security;
alter table contas_pagar      enable row level security;
alter table boletos_receber   enable row level security;
alter table contas_fixas      enable row level security;
alter table cheques_recebidos enable row level security;
alter table cheques_emitidos  enable row level security;
alter table config            enable row level security;
alter table senhas            enable row level security;
alter table sessoes           enable row level security;

-- Pasta privada para os PDFs dos fechamentos.
insert into storage.buckets (id, name, public)
values ('fechamentos', 'fechamentos', false)
on conflict (id) do nothing;
