import { useState, useEffect } from "react";
import { supabase } from "./supabase.js";

// ── Login screen ──────────────────────────────────────────────────
function LoginScreen({ onLogin }) {
  const [email, setEmail]       = useState("");
  const [password, setPassword] = useState("");
  const [loading, setLoading]   = useState(false);
  const [error, setError]       = useState("");
  const [showPass, setShowPass] = useState(false);

  const handleLogin = async (e) => {
    e.preventDefault();
    setLoading(true);
    setError("");
    try {
      const { data, error } = await supabase.auth.signInWithPassword({ email, password });
      if (error) throw error;
      onLogin(data.user);
    } catch (err) {
      setError(err.message === "Invalid login credentials"
        ? "Email o contraseña incorrectos"
        : "Error al iniciar sesión. Intentá de nuevo.");
    }
    setLoading(false);
  };

  return (
    <div style={{
      minHeight: "100vh", background: "#f8fafc",
      display: "flex", alignItems: "center", justifyContent: "center",
      fontFamily: "'Inter', system-ui, sans-serif", padding: 20,
      position: "relative", overflow: "hidden",
    }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
        @keyframes fadeUp { from{opacity:0;transform:translateY(20px)} to{opacity:1;transform:none} }
        @keyframes pulse { 0%,100%{opacity:.4} 50%{opacity:.8} }
        .login-input { background:rgba(255,255,255,.06); border:1px solid rgba(255,255,255,.1); border-radius:10px; padding:12px 14px; color:#e2e8f0; font-size:14px; width:100%; box-sizing:border-box; font-family:inherit; outline:none; transition:border-color .15s; }
        .login-input:focus { border-color:#10b981; }
        .login-input::placeholder { color:#475569; }
        .login-btn { width:100%; padding:13px; border-radius:10px; border:none; background:linear-gradient(135deg,#065f46,#10b981); color:#fff; font-size:15px; font-weight:700; cursor:pointer; font-family:inherit; transition:opacity .15s,transform .1s; }
        .login-btn:hover:not(:disabled) { opacity:.92; transform:translateY(-1px); }
        .login-btn:active { transform:translateY(0); }
        .login-btn:disabled { opacity:.6; cursor:not-allowed; }
      `}</style>

      {/* Background decoration */}
      <div style={{ position:"absolute", inset:0, overflow:"hidden", pointerEvents:"none" }}>
        <div style={{ position:"absolute", top:-200, right:-200, width:600, height:600, borderRadius:"50%", background:"radial-gradient(circle, rgba(16,185,129,.06) 0%, transparent 70%)" }}/>
        <div style={{ position:"absolute", bottom:-300, left:-200, width:700, height:700, borderRadius:"50%", background:"radial-gradient(circle, rgba(59,130,246,.04) 0%, transparent 70%)" }}/>
        <div style={{ position:"absolute", top:"30%", left:"10%", width:2, height:2, borderRadius:"50%", background:"#10b981", animation:"pulse 2s infinite", boxShadow:"0 0 20px #10b981" }}/>
        <div style={{ position:"absolute", top:"60%", right:"15%", width:2, height:2, borderRadius:"50%", background:"#3b82f6", animation:"pulse 3s infinite", boxShadow:"0 0 15px #3b82f6" }}/>
      </div>

      <div style={{ width:"100%", maxWidth:400, animation:"fadeUp .4s ease" }}>
        {/* Logo */}
        <div style={{ textAlign:"center", marginBottom:36 }}>
          <div style={{ width:56, height:56, borderRadius:16, background:"linear-gradient(135deg,#065f46,#10b981)", display:"flex", alignItems:"center", justifyContent:"center", margin:"0 auto 16px", boxShadow:"0 8px 32px rgba(16,185,129,.3)" }}>
            <span style={{ color:"#fff", fontWeight:800, fontSize:24 }}>S</span>
          </div>
          <h1 style={{ color:"#f1f5f9", fontSize:24, fontWeight:700, margin:"0 0 6px" }}>SASOC Merch</h1>
          <p style={{ color:"#475569", fontSize:14, margin:0 }}>CRM Empresarial</p>
        </div>

        {/* Card */}
        <div style={{ background:"#131f35", border:"1px solid rgba(255,255,255,.08)", borderRadius:18, padding:"32px 28px", boxShadow:"0 24px 64px rgba(0,0,0,.4)" }}>
          <h2 style={{ color:"#e2e8f0", fontSize:17, fontWeight:700, margin:"0 0 24px" }}>Iniciar sesión</h2>

          <form onSubmit={handleLogin}>
            <div style={{ marginBottom:16 }}>
              <label style={{ display:"block", fontSize:12, color:"#64748b", fontWeight:700, textTransform:"uppercase", letterSpacing:".5px", marginBottom:6 }}>Email</label>
              <input
                className="login-input"
                type="email"
                value={email}
                onChange={e => setEmail(e.target.value)}
                placeholder="tu@email.com"
                required
                autoFocus
              />
            </div>

            <div style={{ marginBottom:24 }}>
              <label style={{ display:"block", fontSize:12, color:"#64748b", fontWeight:700, textTransform:"uppercase", letterSpacing:".5px", marginBottom:6 }}>Contraseña</label>
              <div style={{ position:"relative" }}>
                <input
                  className="login-input"
                  type={showPass ? "text" : "password"}
                  value={password}
                  onChange={e => setPassword(e.target.value)}
                  placeholder="••••••••"
                  required
                  style={{ paddingRight:44 }}
                />
                <button
                  type="button"
                  onClick={() => setShowPass(!showPass)}
                  style={{ position:"absolute", right:12, top:"50%", transform:"translateY(-50%)", background:"transparent", border:"none", color:"#475569", cursor:"pointer", fontSize:16, padding:4 }}>
                  {showPass ? "🙈" : "👁"}
                </button>
              </div>
            </div>

            {error && (
              <div style={{ background:"rgba(239,68,68,.1)", border:"1px solid rgba(239,68,68,.2)", borderRadius:8, padding:"10px 12px", marginBottom:16, fontSize:13, color:"#fca5a5" }}>
                ⚠ {error}
              </div>
            )}

            <button className="login-btn" type="submit" disabled={loading}>
              {loading ? "Ingresando..." : "Ingresar →"}
            </button>
          </form>
        </div>

        <p style={{ textAlign:"center", color:"#1e3a5f", fontSize:12, marginTop:20 }}>
          SASOC Merch CRM · Acceso restringido
        </p>
      </div>
    </div>
  );
}

// ── Auth wrapper ──────────────────────────────────────────────────
export default function Auth({ children }) {
  const [user, setUser]       = useState(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    // Check active session
    supabase.auth.getSession().then(({ data: { session } }) => {
      setUser(session?.user ?? null);
      setLoading(false);
    });

    // Listen for auth changes
    const { data: { subscription } } = supabase.auth.onAuthStateChange((_event, session) => {
      setUser(session?.user ?? null);
    });

    return () => subscription.unsubscribe();
  }, []);

  const handleLogout = async () => {
    await supabase.auth.signOut();
    setUser(null);
  };

  if (loading) {
    return (
      <div style={{ minHeight:"100vh", background:"#0b1120", display:"flex", alignItems:"center", justifyContent:"center", fontFamily:"'DM Sans',system-ui,sans-serif" }}>
        <div style={{ textAlign:"center" }}>
          <div style={{ width:36, height:36, border:"3px solid rgba(16,185,129,.3)", borderTopColor:"#10b981", borderRadius:"50%", margin:"0 auto 12px", animation:"spin .8s linear infinite" }}/>
          <p style={{ color:"#64748b", fontSize:13 }}>Verificando sesión...</p>
          <style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style>
        </div>
      </div>
    );
  }

  if (!user) {
    return <LoginScreen onLogin={setUser} />;
  }

  // Pass user and logout to children
  return children({ user, onLogout: handleLogout });
}
