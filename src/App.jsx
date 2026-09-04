import React, { useEffect, useState } from "react";
import { supabase } from "./supabaseClient";
import { storageAdapter } from "./storageAdapter";
import Login from "./Login";
import ProtocoloApp from "./ProtocoloApp";

export default function App() {
  const [session, setSession] = useState(undefined); // undefined = carregando

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => setSession(data.session));
    const { data: sub } = supabase.auth.onAuthStateChange((_event, s) => setSession(s));
    return () => sub.subscription.unsubscribe();
  }, []);

  if (session === undefined) {
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
      <button
        onClick={() => supabase.auth.signOut()}
        style={{
          position: "fixed", top: 10, right: 10, zIndex: 100, fontSize: 12,
          background: "white", border: "1px solid #D2DBD8", borderRadius: 7,
          padding: "6px 10px", color: "#54655F", cursor: "pointer",
        }}
      >
        Sair
      </button>
      <ProtocoloApp />
    </div>
  );
}
