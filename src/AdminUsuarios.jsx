import React, { useEffect, useState } from "react";
import { supabase } from "./supabaseClient";

const PAPEIS = ["estagiario", "fonoaudiologo", "supervisor", "admin"];
const LABELS = { estagiario: "Estagiário(a)", fonoaudiologo: "Fonoaudiólogo(a)", supervisor: "Supervisor(a)", admin: "Administrador(a)" };

export default function AdminUsuarios({ onClose }) {
  const [profiles, setProfiles] = useState(null);
  const [error, setError] = useState("");
  const [savingId, setSavingId] = useState(null);

  useEffect(() => {
    (async () => {
      const { data, error } = await supabase.from("profiles").select("id,nome,email,papel").order("nome");
      if (error) setError(error.message);
      else setProfiles(data);
    })();
  }, []);

  const changeRole = async (id, papel) => {
    setSavingId(id);
    const { error } = await supabase.from("profiles").update({ papel }).eq("id", id);
    if (error) {
      setError(error.message);
    } else {
      setProfiles((prev) => prev.map((p) => (p.id === id ? { ...p, papel } : p)));
    }
    setSavingId(null);
  };

  return (
    <div className="fixed inset-0 flex items-center justify-center px-4" style={{ background: "rgba(29,42,40,0.5)", zIndex: 200 }}>
      <div style={{ background: "white", borderRadius: 12, padding: 24, width: "100%", maxWidth: 560, maxHeight: "80vh", overflowY: "auto" }}>
        <div className="flex items-center justify-between" style={{ marginBottom: 12 }}>
          <div style={{ fontFamily: "'Fraunces', serif", fontSize: 18, fontWeight: 600, color: "#1D2A28" }}>Administração de Usuários</div>
          <button onClick={onClose} style={{ fontSize: 13, color: "#54655F" }}>Fechar</button>
        </div>
        <div style={{ fontSize: 12.5, color: "#54655F", marginBottom: 14 }}>
          Só quem tem papel "Administrador(a)" consegue alterar o papel de outras pessoas — isso é verificado pelo próprio banco de dados, não só pela tela.
        </div>
        {error && <div style={{ background: "#F3DFDA", color: "#A6432F", borderRadius: 7, padding: "8px 10px", fontSize: 12.5, marginBottom: 10 }}>{error}</div>}
        {!profiles ? (
          <div style={{ fontSize: 13, color: "#8B9B95" }}>Carregando…</div>
        ) : profiles.length === 0 ? (
          <div style={{ fontSize: 13, color: "#8B9B95" }}>Nenhum usuário cadastrado ainda.</div>
        ) : (
          <div className="flex flex-col gap-2">
            {profiles.map((p) => (
              <div key={p.id} className="flex items-center gap-3 p-2" style={{ border: "1px solid #E4EAE8", borderRadius: 8 }}>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontSize: 13.5, color: "#1D2A28", fontWeight: 500 }}>{p.nome || "(sem nome)"}</div>
                  <div style={{ fontSize: 11.5, color: "#8B9B95", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{p.email}</div>
                </div>
                <select
                  value={p.papel}
                  disabled={savingId === p.id}
                  onChange={(e) => changeRole(p.id, e.target.value)}
                  style={{ fontSize: 12.5, padding: "6px 8px", borderRadius: 6, border: "1px solid #D2DBD8" }}
                >
                  {PAPEIS.map((papel) => <option key={papel} value={papel}>{LABELS[papel]}</option>)}
                </select>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}
