import { supabase } from "./supabaseClient";

/**
 * Reproduz exatamente a API `window.storage` usada pelo app
 * (get/set/delete/list, com o mesmo formato de retorno), mas
 * gravando de verdade no Supabase em vez de um armazenamento
 * temporário. Assim o componente principal (ProtocoloApp.jsx)
 * não precisa ser reescrito — só passamos a apontar
 * `window.storage` para este adaptador depois do login.
 *
 * Por padrão `shared = true`: este sistema é uma ficha clínica de
 * equipe (estagiários + supervisores veem os mesmos atendimentos).
 *
 * IMPORTANTE: os índices únicos de kv_store são PARCIAIS
 * (`where shared = true` / `where shared = false`), e o Postgres
 * não resolve `ON CONFLICT (coluna)` contra um índice parcial só
 * citando as colunas — por isso `set()` NÃO usa `.upsert()` nativo
 * aqui. Em vez disso, faz um select pra ver se a linha já existe e
 * decide entre update/insert manualmente.
 */

async function currentUserId() {
  const { data, error } = await supabase.auth.getUser();
  if (error || !data?.user) throw new Error("Usuário não autenticado.");
  return data.user.id;
}

async function get(key, shared = true) {
  const uid = await currentUserId();
  let query = supabase.from("kv_store").select("key,value,shared").eq("key", key).eq("shared", shared);
  if (!shared) query = query.eq("owner", uid);
  const { data, error } = await query.maybeSingle();
  if (error) throw error;
  if (!data) throw new Error(`Chave não encontrada: ${key}`);
  return { key: data.key, value: data.value, shared: data.shared };
}

async function set(key, value, shared = true) {
  const uid = await currentUserId();
  const strValue = String(value);
  const now = new Date().toISOString();

  // 1) Verifica se já existe uma linha com essa chave (respeitando
  //    a mesma regra de unicidade usada pelos índices parciais do banco).
  let existingQuery = supabase.from("kv_store").select("id").eq("key", key).eq("shared", shared);
  if (!shared) existingQuery = existingQuery.eq("owner", uid);
  const { data: existing, error: selectError } = await existingQuery.maybeSingle();
  if (selectError) throw selectError;

  // 2) Atualiza se existir, insere se não existir — nunca usa
  //    ON CONFLICT, então não depende do índice ser parcial ou não.
  if (existing) {
    const { data, error } = await supabase
      .from("kv_store")
      .update({ value: strValue, updated_at: now })
      .eq("id", existing.id)
      .select("key,value,shared")
      .maybeSingle();
    if (error) throw error;
    return data || { key, value: strValue, shared };
  }

  const { data, error } = await supabase
    .from("kv_store")
    .insert({ owner: uid, key, value: strValue, shared, updated_at: now })
    .select("key,value,shared")
    .maybeSingle();
  if (error) throw error;
  return data || { key, value: strValue, shared };
}

async function del(key, shared = true) {
  const uid = await currentUserId();
  let query = supabase.from("kv_store").delete().eq("key", key).eq("shared", shared);
  if (!shared) query = query.eq("owner", uid);
  const { error } = await query;
  if (error) throw error;
  return { key, deleted: true, shared };
}

async function list(prefix = "", shared = true) {
  const uid = await currentUserId();
  let query = supabase.from("kv_store").select("key").eq("shared", shared).like("key", `${prefix}%`);
  if (!shared) query = query.eq("owner", uid);
  const { data, error } = await query;
  if (error) throw error;
  return { keys: (data || []).map((r) => r.key), prefix, shared };
}

export const storageAdapter = { get, set, delete: del, list };
