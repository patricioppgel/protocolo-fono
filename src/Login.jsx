import React, { useState } from "react";
import { supabase } from "./supabaseClient";

const T = {
  paper: "#EEF2F1", surface: "#FFFFFF", ink: "#1D2A28", inkSoft: "#54655F",
  accent: "#2C6E62", line: "#D2DBD8", warn: "#A6432F", warnSoft: "#F3DFDA",
};

export default function Login() {
  const [mode, setMode] = useState("signin"); // signin | signup
  const [nome, setNome] = useState("");
  const [email, setEmail] = useState("");
  const [senha, setSenha] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [info, setInfo] = useState("");

  const submit = async (e) => {
    e.preventDefault();
    setError(""); setInfo(""); setLoading(true);
    try {
      if (mode === "signin") {
        const { error } = await supabase.auth.signInWithPassword({ email, password: senha });
        if (error) throw error;
      } else {
        const { error } = await supabase.auth.signUp({
          email, password: senha, options: { data: { nome } },
        });
        if (error) throw error;
        setInfo("Conta criada. Verifique seu e-mail para confirmar o cadastro antes de entrar.");
      }
    } catch (err) {
      setError(err.message || "Não foi possível concluir. Tente novamente.");
    } finally {
      setLoading(false);
    }
  };

  const inputStyle = {
    width: "100%", fontSize: 14.5, color: T.ink, background: T.surface,
    border: "1px solid " + T.line, borderRadius: 8, padding: "10px 12px", outline: "none",
  };

  return (
    <div style={{ minHeight: "100vh", background: T.paper, display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'IBM Plex Sans', sans-serif", padding: 16 }}>
      <form onSubmit={submit} style={{ width: "100%", maxWidth: 360, background: T.surface, border: "1px solid " + T.line, borderRadius: 12, padding: 28, display: "flex", flexDirection: "column", gap: 14 }}>
        <div>
          <div style={{ fontFamily: "'Fraunces', serif", fontSize: 20, fontWeight: 600, color: T.ink }}>Protocolo de Atendimento</div>
          <div style={{ fontSize: 12.5, color: T.inkSoft }}>Fonoaudiologia · acesso da equipe</div>
        </div>

        {mode === "signup" && (
          <label style={{ display: "flex", flexDirection: "column", gap: 4, fontSize: 12.5, color: T.inkSoft }}>
            Nome completo
            <input required value={nome} onChange={(e) => setNome(e.target.value)} style={inputStyle} />
          </label>
        )}
        <label style={{ display: "flex", flexDirection: "column", gap: 4, fontSize: 12.5, color: T.inkSoft }}>
          E-mail institucional
          <input required type="email" value={email} onChange={(e) => setEmail(e.target.value)} style={inputStyle} />
        </label>
        <label style={{ display: "flex", flexDirection: "column", gap: 4, fontSize: 12.5, color: T.inkSoft }}>
          Senha
          <input required type="password" minLength={6} value={senha} onChange={(e) => setSenha(e.target.value)} style={inputStyle} />
        </label>

        {error && <div style={{ background: T.warnSoft, color: T.warn, borderRadius: 7, padding: "8px 10px", fontSize: 12.5 }}>{error}</div>}
        {info && <div style={{ background: "#DCEAE5", color: T.accent, borderRadius: 7, padding: "8px 10px", fontSize: 12.5 }}>{info}</div>}

        <button type="submit" disabled={loading} style={{ background: T.accent, color: "white", border: "none", borderRadius: 8, padding: "10px 0", fontSize: 14, fontWeight: 500, cursor: "pointer" }}>
          {loading ? "Aguarde…" : mode === "signin" ? "Entrar" : "Criar conta"}
        </button>

        <button type="button" onClick={() => { setMode(mode === "signin" ? "signup" : "signin"); setError(""); setInfo(""); }} style={{ background: "none", border: "none", color: T.accent, fontSize: 12.5, cursor: "pointer" }}>
          {mode === "signin" ? "Não tem conta? Cadastre-se" : "Já tem conta? Entrar"}
        </button>
      </form>
    </div>
  );
}
