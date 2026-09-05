import React, { useEffect, useState } from "react";
import { supabase } from "./supabaseClient";
import { storageAdapter } from "./storageAdapter";
import Login from "./Login";
import ProtocoloApp from "./ProtocoloApp";
import AdminUsuarios from "./AdminUsuarios";

export default function App() {
  const [session, setSession] = useState(undefined); // undefined = carregando
  const [authProfile, setAuthProfile] = useState(undefined); // undefined = carregando, null = sem perfil ainda
  const [showAdmin, setShowAdmin] = useState(false);

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => setSession(data.session));
    const { data: sub } = supabase.auth.onAuthStateChange((_event, s) => setSession(s));
    return () => sub.subscription.unsubscribe();
  }, []);

  const loadProfile = async (userId) => {
    const { data, error } = await supabase.from("profiles").select("id,nome,email,papel").eq("id", userId).maybeSingle();
    if (!error) setAuthProfile(data || null);
  };

  useEffect(() => {
    if (session?.user?.id) loadProfile(session.user.id);
  }, [session?.user?.id]);

  // Expõe o papel real (verificado no banco) para o ProtocoloApp.jsx usar
  // ao decidir quem pode validar um atendimento.
  useEffect(() => {
    window.authProfile = authProfile || null;
  }, [authProfile]);

  if (session === undefined || (session && authProfile === undefined)) {
    return <div style={{ padding: 40, fontFamily: "sans-serif", color: "#54655F" }}>Carregando…</div>;
  }
  if (!session) {
    return <Login />;
  }

  // A partir daqui, todo window.storage.get/set/delete/list usado
  // dentro de ProtocoloApp.jsx passa a gravar de verdade no Supabase.
  window.storage = storageAdapter;

  return (
    <div style={{ position: "relative" }}>
      <div style={{ position: "fixed", top: 10, right: 10, zIndex: 100, display: "flex", gap: 8 }}>
        {authProfile?.papel === "admin" && (
          <button
            onClick={() => setShowAdmin(true)}
            style={{ fontSize: 12, background: "white", border: "1px solid #D2DBD8", borderRadius: 7, padding: "6px 10px", color: "#2C6E62", cursor: "pointer" }}
          >
            Administração
          </button>
        )}
        <button
          onClick={() => supabase.auth.signOut()}
          style={{ fontSize: 12, background: "white", border: "1px solid #D2DBD8", borderRadius: 7, padding: "6px 10px", color: "#54655F", cursor: "pointer" }}
        >
          Sair
        </button>
      </div>
      {showAdmin && <AdminUsuarios onClose={() => { setShowAdmin(false); loadProfile(session.user.id); }} />}
      <ProtocoloApp />
    </div>
  );
}
