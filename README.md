# Protocolo de Atendimento Fonoaudiológico — versão web (multiusuário)

App React completo (todas as áreas, audiometria com gráficos, PDF, planilhas)
já testado e compilando (`npm run build` passou sem erros). O que falta é só
criar as contas gratuitas do Supabase e da Vercel e apontar o projeto para elas.

## 1. Criar o banco de dados (Supabase) — ~5 min

1. Crie uma conta grátis em **https://supabase.com** e clique em "New project".
2. Escolha um nome (ex.: `protocolo-fono`) e uma senha de banco (guarde-a).
3. Quando o projeto terminar de criar, vá em **SQL Editor → New query**.
4. Abra o arquivo `supabase/schema.sql` deste projeto, copie todo o conteúdo,
   cole no editor e clique em **Run**. Isso cria a tabela de dados e as
   permissões (RLS) — cada usuário autenticado passa a ver/editar os
   atendimentos da equipe.
5. Vá em **Authentication → Providers** e confirme que **Email** está ativado
   (vem ativado por padrão).
6. (Opcional, recomendado) Em **Authentication → Settings**, você pode
   desativar a confirmação por e-mail se quiser que os estagiários entrem
   direto após o cadastro, sem clicar em link de confirmação.
7. Vá em **Project Settings → API** e anote dois valores:
   - **Project URL**
   - **anon public key**

## 2. Configurar o projeto localmente

1. Copie `.env.example` para `.env`.
2. Cole ali a **Project URL** e a **anon public key** que você anotou:
   ```
   VITE_SUPABASE_URL=https://SEU-PROJETO.supabase.co
   VITE_SUPABASE_ANON_KEY=sua-chave-anon-aqui
   ```
3. Instale as dependências e teste localmente:
   ```
   npm install
   npm run dev
   ```
   Abra o link que aparecer (geralmente `http://localhost:5173`) — você deve
   ver a tela de login. Crie uma conta de teste para conferir se tudo
   funciona antes de publicar.

## 3. Publicar (Vercel) — ~5 min

1. Crie uma conta grátis em **https://vercel.com** (dá para entrar direto
   com GitHub).
2. Suba este projeto para um repositório no GitHub (crie um repo novo e
   faça `git push`), **ou** instale a Vercel CLI e rode `vercel` direto
   nesta pasta (ela guia o deploy sem precisar do GitHub).
3. Ao importar o projeto na Vercel, ela detecta que é um app Vite
   automaticamente. Antes de finalizar, adicione as variáveis de ambiente
   (mesmo nome do `.env`):
   - `VITE_SUPABASE_URL`
   - `VITE_SUPABASE_ANON_KEY`
4. Clique em **Deploy**. Em ~1 minuto você recebe um link do tipo
   `https://protocolo-fono.vercel.app` — esse é o endereço que você
   compartilha com estagiários e supervisores.

## 4. Uso do dia a dia

- Cada pessoa (estagiário ou supervisor) cria sua própria conta (e-mail +
  senha) na tela de login.
- Por padrão, **todos os atendimentos, estudos de caso etc. são
  compartilhados entre todos os usuários autenticados** — é uma ficha
  clínica de equipe, não uma conta pessoal isolada. Se depois vocês
  quiserem restringir o que cada estagiário vê, dá para ajustar as
  políticas de RLS no `schema.sql` (por exemplo, um supervisor vendo tudo
  e cada estagiário só os próprios pacientes) — me avise que eu ajusto.
- Para adicionar mais estagiários depois, basta compartilhar o link e cada
  um se cadastra.

## O que já foi validado aqui

- ✅ `npm install` e `npm run build` rodam sem erro.
- ✅ Todas as 11 áreas, audiograma com gráfico, cálculo de mascaramento,
  PDF do atendimento, PDF do laudo, exportação em planilha — nada foi
  reescrito, é o mesmo `ProtocoloApp.jsx` que já usávamos, só que agora
  gravando em um banco de dados real via `storageAdapter.js`.
- ⚠️ O que **não** foi testado por mim (preciso que você confirme): login
  de verdade contra um projeto Supabase real, e o deploy na Vercel — isso
  exige uma conta que só você pode criar. Siga os passos acima; se algo
  der erro, me mande a mensagem de erro que eu ajudo a resolver.
