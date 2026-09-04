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
  const row = { owner: uid, key, value: String(value), shared, updated_at: new Date().toISOString() };
  const onConflict = shared ? "key" : "owner,key";
  const { data, error } = await supabase.from("kv_store").upsert(row, { onConflict }).select("key,value,shared").maybeSingle();
  if (error) throw error;
  return data || { key, value: row.value, shared };
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
