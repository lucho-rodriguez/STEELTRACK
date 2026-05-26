import { useState, useRef, useEffect, useCallback } from "react";
import { createClient } from "@supabase/supabase-js";
import * as XLSX from "xlsx";

// ─── SUPABASE ─────────────────────────────────────────────────────────────────
const SUPABASE_URL = "https://otiytntcfkoinxwgbcpq.supabase.co";
const SUPABASE_KEY = "sb_publishable_q1JM9AS9dynSOtuzji5RMg_N1g3wXEE";
const supabase = createClient(SUPABASE_URL, SUPABASE_KEY);

// ─── CONSTANTS ────────────────────────────────────────────────────────────────
const TUNDISH_STATE_LIMITS = { obrador_basculador: 1, obrador_enfriando: 2 };
const TUNDISH_ZONES = ["Acería", "Obrador", "Taller"];

// ─── UTILS ────────────────────────────────────────────────────────────────────
const getStateById = (states, id) => states.find(s => s.id === id) || states[0];
const formatDate = ts => new Date(ts).toLocaleDateString("es-AR", { day:"2-digit", month:"2-digit", year:"numeric" });
const formatTime = ts => new Date(ts).toLocaleTimeString("es-AR", { hour:"2-digit", minute:"2-digit" });
function formatElapsed(ts) {
  const diff = Date.now() - new Date(ts).getTime();
  const mins = Math.floor(diff/60000), hrs = Math.floor(mins/60), days = Math.floor(hrs/24);
  if (days>0) return `${days}d ${hrs%24}h`; if (hrs>0) return `${hrs}h ${mins%60}m`;
  if (mins>0) return `${mins}m`; return "justo ahora";
}
function exportToExcel(items, states, label) {
  const rows = [];
  items.forEach(item => {
    (item.history||[]).forEach(h => {
      const st = getStateById(states, h.state_id);
      rows.push({ Nombre: item.name, Estado: st?.label||h.state_id, Fecha: formatDate(h.created_at), Hora: formatTime(h.created_at), Usuario: h.user_name, Comentario: h.comment||"", "Tiene Foto": h.photo_url?"Sí":"No" });
    });
  });
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, label);
  XLSX.writeFile(wb, `historial_${label.toLowerCase()}_${new Date().toISOString().slice(0,10)}.xlsx`);
}

// ─── STYLES ───────────────────────────────────────────────────────────────────
const CSS = `
  @import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@400;600;700;800&family=Barlow:wght@400;500;600&display=swap');
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Barlow', sans-serif; background: #0d1117; color: #e6edf3; }
  ::-webkit-scrollbar { width: 6px; } ::-webkit-scrollbar-track { background: #161b22; } ::-webkit-scrollbar-thumb { background: #30363d; border-radius: 3px; }
  @keyframes fadeIn { from { opacity:0; transform:translateY(8px) } to { opacity:1; transform:translateY(0) } }
  @keyframes spin { to { transform: rotate(360deg) } }
  .card-item { animation: fadeIn .2s ease both; }
  .tab-btn { transition: all .15s; } .tab-btn:hover { background: #21262d !important; } .tab-btn.active { background: #e6edf3 !important; color: #0d1117 !important; }
  .item-card { transition: all .2s; cursor: pointer; } .item-card:hover { transform: translateY(-2px); box-shadow: 0 8px 24px rgba(0,0,0,.4) !important; border-color: #58a6ff !important; }
  input, select, textarea { font-family: 'Barlow', sans-serif; }
  .overlay { position:fixed; inset:0; background:rgba(0,0,0,.7); z-index:200; display:flex; align-items:center; justify-content:center; padding:16px; backdrop-filter:blur(4px); }
  .modal { background:#161b22; border:1px solid #30363d; border-radius:16px; width:100%; max-width:580px; max-height:90vh; overflow-y:auto; animation:fadeIn .2s ease; }
  .inp { width:100%; padding:9px 12px; background:#0d1117; border:1.5px solid #30363d; border-radius:8px; color:#e6edf3; font-size:14px; outline:none; transition:border .15s; }
  .inp:focus { border-color:#58a6ff; }
  .btn-primary { background:#238636; color:#fff; border:none; border-radius:8px; padding:10px 20px; font-weight:600; font-size:14px; cursor:pointer; transition:background .15s; font-family:'Barlow',sans-serif; }
  .btn-primary:hover { background:#2ea043; } .btn-primary:disabled { opacity:.5; cursor:not-allowed; }
  .btn-secondary { background:#21262d; color:#e6edf3; border:1px solid #30363d; border-radius:8px; padding:10px 20px; font-weight:600; font-size:14px; cursor:pointer; transition:all .15s; font-family:'Barlow',sans-serif; }
  .btn-secondary:hover { background:#30363d; }
  .badge-dot { display:inline-block; width:8px; height:8px; border-radius:50%; flex-shrink:0; }
  .admin-badge { display:inline-flex; align-items:center; gap:4px; background:#7c3aed22; border:1px solid #7c3aed66; color:#a78bfa; border-radius:6px; padding:2px 8px; font-size:11px; font-weight:700; }
  .spinner { width:20px; height:20px; border:2px solid #30363d; border-top-color:#58a6ff; border-radius:50%; animation:spin .7s linear infinite; }
`;

// ─── SHARED COMPONENTS ────────────────────────────────────────────────────────
function StateBadge({ state, size="sm" }) {
  if (!state) return null;
  return (
    <span style={{ display:"inline-flex", alignItems:"center", gap:6, padding:size==="sm"?"3px 10px":"5px 14px",
      borderRadius:20, background:state.color+"22", color:state.color, fontSize:size==="sm"?12:13, fontWeight:600, whiteSpace:"nowrap" }}>
      <span className="badge-dot" style={{ background:state.color }} />{state.label}
    </span>
  );
}
function Field({ label, children }) {
  return (
    <div style={{ marginBottom:16 }}>
      <label style={{ display:"block", fontSize:11, fontWeight:700, color:"#8b949e", marginBottom:6, textTransform:"uppercase", letterSpacing:".07em" }}>{label}</label>
      {children}
    </div>
  );
}
function ModalWrap({ title, onClose, children, maxWidth=580 }) {
  return (
    <div className="overlay" onClick={e => e.target===e.currentTarget && onClose()}>
      <div className="modal" style={{ maxWidth }}>
        <div style={{ padding:"18px 24px 14px", borderBottom:"1px solid #21262d", display:"flex", justifyContent:"space-between", alignItems:"center" }}>
          <h3 style={{ fontFamily:"'Barlow Condensed',sans-serif", fontSize:22, fontWeight:800, color:"#e6edf3" }}>{title}</h3>
          <button onClick={onClose} style={{ background:"#21262d", border:"none", borderRadius:8, width:32, height:32, cursor:"pointer", color:"#8b949e", fontSize:15 }}>✕</button>
        </div>
        <div style={{ padding:24 }}>{children}</div>
      </div>
    </div>
  );
}
function Loader({ text="Cargando..." }) {
  return (
    <div style={{ display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", gap:16, padding:64 }}>
      <div className="spinner" />
      <span style={{ color:"#8b949e", fontSize:14 }}>{text}</span>
    </div>
  );
}

// ─── AUTH ─────────────────────────────────────────────────────────────────────
function AuthScreen({ onLogin }) {
  const [mode, setMode] = useState("login");
  const [form, setForm] = useState({ name:"", email:"", password:"" });
  const [err, setErr] = useState(""); const [loading, setLoading] = useState(false);

  async function handleLogin() {
    setLoading(true); setErr("");
    const { data, error } = await supabase.auth.signInWithPassword({ email:form.email, password:form.password });
    if (error) { setErr("Email o contraseña incorrectos."); setLoading(false); return; }
    const { data: profile } = await supabase.from("profiles").select("*").eq("id", data.user.id).single();
    onLogin({ ...data.user, ...profile });
    setLoading(false);
  }
  async function handleRegister() {
    if (!form.name||!form.email||!form.password) return setErr("Completá todos los campos.");
    setLoading(true); setErr("");
    const { data, error } = await supabase.auth.signUp({ email:form.email, password:form.password });
    if (error) { setErr(error.message); setLoading(false); return; }
    await supabase.from("profiles").insert({ id:data.user.id, name:form.name, role:"operario" });
    const profile = { id:data.user.id, name:form.name, role:"operario" };
    onLogin({ ...data.user, ...profile });
    setLoading(false);
  }

  return (
    <div style={{ minHeight:"100vh", display:"flex", alignItems:"center", justifyContent:"center", background:"#0d1117",
      backgroundImage:"radial-gradient(ellipse at 20% 50%, #1a2332 0%, transparent 60%)" }}>
      <div style={{ width:"100%", maxWidth:400, padding:16 }}>
        <div style={{ textAlign:"center", marginBottom:40 }}>
          <div style={{ fontSize:48, marginBottom:12 }}>🏭</div>
          <h1 style={{ fontFamily:"'Barlow Condensed',sans-serif", fontSize:36, fontWeight:800, color:"#e6edf3" }}>
            STEEL<span style={{ color:"#f59e0b" }}>TRACK</span>
          </h1>
          <p style={{ color:"#8b949e", fontSize:14, marginTop:4 }}>Gestión de Tundish y Potes</p>
        </div>
        <div style={{ background:"#161b22", border:"1px solid #30363d", borderRadius:16, padding:28 }}>
          <div style={{ display:"flex", background:"#0d1117", borderRadius:10, padding:4, marginBottom:24 }}>
            {["login","register"].map(m => (
              <button key={m} className={`tab-btn ${mode===m?"active":""}`} onClick={()=>{setMode(m);setErr("");}}
                style={{ flex:1, padding:"8px", border:"none", borderRadius:8, cursor:"pointer", fontSize:13, fontWeight:600, background:"transparent", color:mode===m?"#0d1117":"#8b949e", fontFamily:"'Barlow',sans-serif" }}>
                {m==="login"?"Ingresar":"Registrarse"}
              </button>
            ))}
          </div>
          {mode==="register" && <Field label="Nombre completo"><input className="inp" value={form.name} onChange={e=>setForm({...form,name:e.target.value})} placeholder="Tu nombre" /></Field>}
          <Field label="Email"><input className="inp" type="email" value={form.email} onChange={e=>setForm({...form,email:e.target.value})} placeholder="tu@email.com" /></Field>
          <Field label="Contraseña">
            <input className="inp" type="password" value={form.password} onChange={e=>setForm({...form,password:e.target.value})}
              placeholder="••••••••" onKeyDown={e=>e.key==="Enter"&&(mode==="login"?handleLogin():handleRegister())} />
          </Field>
          {err && <div style={{ background:"#da363322", border:"1px solid #da3633", borderRadius:8, padding:"8px 12px", fontSize:13, color:"#ff7b72", marginBottom:16 }}>{err}</div>}
          <button className="btn-primary" style={{ width:"100%", marginTop:4 }} onClick={mode==="login"?handleLogin:handleRegister} disabled={loading}>
            {loading ? <span style={{ display:"flex", alignItems:"center", justifyContent:"center", gap:8 }}><div className="spinner" style={{ width:16,height:16 }} />Espera...</span> : mode==="login"?"Ingresar":"Crear cuenta"}
          </button>
        </div>
      </div>
    </div>
  );
}

// ─── DETAIL MODAL ─────────────────────────────────────────────────────────────
function DetailModal({ item, states, currentUser, allTundish, isTundish, onClose, onUpdated }) {
  const [newStateId, setNewStateId] = useState(item.state_id);
  const [comment, setComment] = useState("");
  const [photo, setPhoto] = useState(null);
  const [photoPreview, setPhotoPreview] = useState(null);
  const [tab, setTab] = useState("change");
  const [editingName, setEditingName] = useState(false);
  const [nameVal, setNameVal] = useState(item.name);
  const [limitErr, setLimitErr] = useState("");
  const [saving, setSaving] = useState(false);
  const fileRef = useRef();
  const isAdmin = currentUser.role === "admin";
  const table = isTundish ? "tundish" : "potes";
  const histTable = isTundish ? "tundish_history" : "pote_history";
  const fkCol = isTundish ? "tundish_id" : "pote_id";

  async function saveName() {
    if (nameVal.trim() && nameVal.trim() !== item.name) {
      await supabase.from(table).update({ name: nameVal.trim() }).eq("id", item.id);
      onUpdated();
    }
    setEditingName(false);
  }
  function handleStateChange(sid) {
    setLimitErr("");
    if (isTundish && TUNDISH_STATE_LIMITS[sid] !== undefined) {
      const count = (allTundish||[]).filter(t => t.id !== item.id && t.state_id === sid).length;
      if (count >= TUNDISH_STATE_LIMITS[sid]) {
        const st = states.find(s => s.id===sid);
        setLimitErr(`⚠️ Límite: máx. ${TUNDISH_STATE_LIMITS[sid]} tundish en "${st?.label}".`);
        return;
      }
    }
    setNewStateId(sid);
  }
  async function handleSave() {
    if (limitErr || (!comment.trim() && newStateId===item.state_id)) return;
    setSaving(true);
    let photo_url = null;
    if (photo) {
      const ext = photo.split(";")[0].split("/")[1];
      const path = `${table}/${item.id}/${Date.now()}.${ext}`;
      const blob = await (await fetch(photo)).blob();
      const { data: upData } = await supabase.storage.from("photos").upload(path, blob, { contentType: blob.type });
      if (upData) {
        const { data: urlData } = supabase.storage.from("photos").getPublicUrl(path);
        photo_url = urlData.publicUrl;
      }
    }
    await supabase.from(table).update({ state_id: newStateId }).eq("id", item.id);
    await supabase.from(histTable).insert({ [fkCol]: item.id, state_id: newStateId, user_name: currentUser.name, comment: comment.trim(), photo_url });
    setSaving(false);
    onUpdated();
    onClose();
  }

  const curState = getStateById(states, item.state_id);
  const curEntry = (item.history||[]).find(h => h.state_id===item.state_id) || (item.history||[])[0];

  return (
    <div className="overlay" onClick={e=>e.target===e.currentTarget&&onClose()}>
      <div className="modal">
        <div style={{ padding:"20px 24px 14px", borderBottom:"1px solid #21262d", display:"flex", justifyContent:"space-between", alignItems:"flex-start" }}>
          <div style={{ flex:1 }}>
            {isAdmin && editingName ? (
              <div style={{ display:"flex", alignItems:"center", gap:8 }}>
                <input className="inp" value={nameVal} onChange={e=>setNameVal(e.target.value)} autoFocus
                  onKeyDown={e=>{if(e.key==="Enter")saveName();if(e.key==="Escape"){setNameVal(item.name);setEditingName(false);}}}
                  style={{ fontFamily:"'Barlow Condensed',sans-serif", fontSize:20, fontWeight:800, padding:"4px 10px", flex:1 }} />
                <button onClick={saveName} style={{ background:"#238636", border:"none", borderRadius:7, padding:"5px 12px", color:"#fff", fontWeight:700, fontSize:13, cursor:"pointer", fontFamily:"inherit" }}>✓</button>
                <button onClick={()=>{setNameVal(item.name);setEditingName(false);}} style={{ background:"#21262d", border:"none", borderRadius:7, padding:"5px 10px", color:"#8b949e", fontSize:13, cursor:"pointer" }}>✕</button>
              </div>
            ) : (
              <div style={{ display:"flex", alignItems:"center", gap:8 }}>
                <h2 style={{ fontFamily:"'Barlow Condensed',sans-serif", fontSize:24, fontWeight:800, color:"#e6edf3", margin:0 }}>{item.name}</h2>
                {isAdmin && <button onClick={()=>setEditingName(true)} style={{ background:"none", border:"none", cursor:"pointer", color:"#6e7681", fontSize:15, padding:"2px 4px" }}>✏️</button>}
              </div>
            )}
            <div style={{ marginTop:8 }}><StateBadge state={curState} size="md" /></div>
            {curEntry && (
              <div style={{ marginTop:12, display:"flex", flexWrap:"wrap", gap:8 }}>
                {[["📅","Inicio",formatDate(curEntry.created_at),"#c9d1d9"],["🕐","Hora",formatTime(curEntry.created_at),"#58a6ff"],
                  ["👤","Usuario",curEntry.user_name,"#c9d1d9"],["⏱","Tiempo",formatElapsed(curEntry.created_at),"#f59e0b"]
                ].map(([ic,lb,val,col]) => (
                  <div key={lb} style={{ background:"#0d1117", borderRadius:7, padding:"5px 10px" }}>
                    <div style={{ fontSize:10, color:"#6e7681", textTransform:"uppercase", letterSpacing:".05em" }}>{ic} {lb}</div>
                    <div style={{ fontSize:13, fontWeight:700, color:col }}>{val}</div>
                  </div>
                ))}
              </div>
            )}
          </div>
          <button onClick={onClose} style={{ background:"#21262d", border:"none", borderRadius:8, width:34, height:34, cursor:"pointer", color:"#8b949e", fontSize:16, flexShrink:0, marginLeft:12 }}>✕</button>
        </div>
        <div style={{ display:"flex", padding:"10px 24px 0", borderBottom:"1px solid #21262d" }}>
          {[["change","Actualizar"],["history","Historial"]].map(([t,l]) => (
            <button key={t} onClick={()=>setTab(t)} style={{ padding:"8px 16px", border:"none", background:"transparent", cursor:"pointer",
              color:tab===t?"#58a6ff":"#8b949e", fontWeight:600, fontSize:14, fontFamily:"'Barlow',sans-serif",
              borderBottom:tab===t?"2px solid #58a6ff":"2px solid transparent" }}>{l}</button>
          ))}
        </div>
        <div style={{ padding:24 }}>
          {tab==="change" && (
            <>
              <Field label="Nuevo Estado">
                <select className="inp" value={newStateId} onChange={e=>handleStateChange(e.target.value)}>
                  {states.map(s => <option key={s.id} value={s.id}>{s.label}</option>)}
                </select>
              </Field>
              {limitErr && <div style={{ background:"#f59e0b18", border:"1px solid #f59e0b66", borderRadius:8, padding:"8px 12px", fontSize:13, color:"#f59e0b", marginBottom:14 }}>{limitErr}</div>}
              <Field label="Comentario">
                <textarea className="inp" rows={3} value={comment} onChange={e=>setComment(e.target.value)} placeholder="Agrega un comentario..." style={{ resize:"vertical" }} />
              </Field>
              <Field label="Foto (opcional)">
                <input ref={fileRef} type="file" accept="image/*" style={{ display:"none" }} onChange={async e=>{
                  const f=e.target.files[0]; if(!f) return;
                  const r=new FileReader(); r.onload=()=>{setPhoto(r.result);setPhotoPreview(r.result);}; r.readAsDataURL(f);
                }} />
                <button className="btn-secondary" style={{ width:"100%", textAlign:"left" }} onClick={()=>fileRef.current.click()}>
                  📷 {photoPreview?"Foto seleccionada — clic para cambiar":"Seleccionar foto"}
                </button>
                {photoPreview && (
                  <div style={{ marginTop:10, position:"relative" }}>
                    <img src={photoPreview} alt="preview" style={{ width:"100%", maxHeight:200, objectFit:"cover", borderRadius:8, border:"1px solid #30363d" }} />
                    <button onClick={()=>{setPhoto(null);setPhotoPreview(null);}} style={{ position:"absolute", top:6, right:6, background:"#0d1117cc", border:"none", borderRadius:6, color:"#ff7b72", cursor:"pointer", padding:"4px 8px", fontSize:12 }}>✕</button>
                  </div>
                )}
              </Field>
              <div style={{ display:"flex", gap:10, justifyContent:"flex-end" }}>
                <button className="btn-secondary" onClick={onClose}>Cancelar</button>
                <button className="btn-primary" onClick={handleSave}
                  disabled={saving||!!limitErr||(!comment.trim()&&newStateId===item.state_id)}>
                  {saving?<span style={{ display:"flex",alignItems:"center",gap:8 }}><div className="spinner" style={{width:14,height:14}}/>Guardando...</span>:"Guardar"}
                </button>
              </div>
            </>
          )}
          {tab==="history" && (
            <div style={{ display:"flex", flexDirection:"column" }}>
              {(!item.history||item.history.length===0) && <p style={{ color:"#8b949e", textAlign:"center", padding:32 }}>Sin historial</p>}
              {(item.history||[]).map((h,i) => {
                const st = getStateById(states, h.state_id);
                return (
                  <div key={h.id} style={{ display:"flex", gap:14, paddingBottom:16, borderBottom:i<item.history.length-1?"1px solid #21262d":"none", paddingTop:i===0?0:16 }}>
                    <div style={{ display:"flex", flexDirection:"column", alignItems:"center", flexShrink:0 }}>
                      <div style={{ width:12, height:12, borderRadius:"50%", background:st?.color||"#555", marginTop:4, boxShadow:`0 0 6px ${st?.color||"#555"}66` }} />
                      {i<item.history.length-1 && <div style={{ width:2, flex:1, background:"#21262d", marginTop:4 }} />}
                    </div>
                    <div style={{ flex:1, minWidth:0 }}>
                      <StateBadge state={st} />
                      <div style={{ display:"flex", flexWrap:"wrap", gap:6, marginTop:7 }}>
                        {[["📅",formatDate(h.created_at),"#c9d1d9"],["🕐",formatTime(h.created_at),"#58a6ff"],["👤",h.user_name,"#c9d1d9"]].map(([ic,val,col])=>(
                          <span key={ic} style={{ display:"inline-flex", alignItems:"center", gap:4, fontSize:12, color:col, background:"#21262d", padding:"2px 8px", borderRadius:6 }}>{ic} {val}</span>
                        ))}
                      </div>
                      {h.comment && <div style={{ marginTop:8, padding:"8px 10px", background:"#0d1117", borderRadius:7, borderLeft:`3px solid ${st?.color||"#555"}`, fontSize:13, color:"#c9d1d9", lineHeight:1.5 }}>{h.comment}</div>}
                      {h.photo_url && <img src={h.photo_url} alt="foto" style={{ marginTop:8, width:"100%", maxHeight:180, objectFit:"cover", borderRadius:8, border:"1px solid #30363d" }} />}
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// ─── ITEM CARD ────────────────────────────────────────────────────────────────
function ItemCard({ item, states, onClick, compact=false }) {
  const state = getStateById(states, item.state_id);
  const entry = (item.history||[]).find(h=>h.state_id===item.state_id)||(item.history||[])[0];
  if (compact) return (
    <div className="item-card card-item" onClick={onClick}
      style={{ background:"#161b22", border:"1px solid #30363d", borderRadius:9, padding:"10px 12px", position:"relative", overflow:"hidden" }}>
      <div style={{ position:"absolute", top:0, left:0, right:0, height:2, background:state?.color||"#555" }} />
      <div style={{ fontFamily:"'Barlow Condensed',sans-serif", fontSize:17, fontWeight:800, color:"#e6edf3", marginBottom:6 }}>{item.name}</div>
      <StateBadge state={state} />
    </div>
  );
  return (
    <div className="item-card card-item" onClick={onClick}
      style={{ background:"#161b22", border:"1px solid #30363d", borderRadius:12, padding:"14px 16px", position:"relative", overflow:"hidden" }}>
      <div style={{ position:"absolute", top:0, left:0, right:0, height:3, background:state?.color||"#555" }} />
      <div style={{ fontFamily:"'Barlow Condensed',sans-serif", fontSize:22, fontWeight:800, color:"#e6edf3", marginBottom:8 }}>{item.name}</div>
      <StateBadge state={state} />
      <div style={{ marginTop:10, display:"flex", alignItems:"center", gap:5 }}>
        <span style={{ fontSize:13 }}>⏱</span>
        <span style={{ fontSize:13, fontWeight:700, color:"#f59e0b" }}>{entry?formatElapsed(entry.created_at):"—"}</span>
        <span style={{ fontSize:11, color:"#6e7681" }}>en este estado</span>
      </div>
    </div>
  );
}

// ─── ADD ITEM MODAL ───────────────────────────────────────────────────────────
function AddItemModal({ type, states, onClose, onAdded }) {
  const [form, setForm] = useState({ name:"", stateId:states[0]?.id||"" });
  const [saving, setSaving] = useState(false);
  async function handle() {
    if (!form.name.trim()) return;
    setSaving(true);
    const table = type==="tundish"?"tundish":"potes";
    const histTable = type==="tundish"?"tundish_history":"pote_history";
    const fkCol = type==="tundish"?"tundish_id":"pote_id";
    const { data } = await supabase.from(table).insert({ name:form.name.trim(), state_id:form.stateId }).select().single();
    if (data) await supabase.from(histTable).insert({ [fkCol]:data.id, state_id:form.stateId, user_name:"Sistema", comment:"Ingreso al sistema" });
    setSaving(false); onAdded(); onClose();
  }
  return (
    <ModalWrap title={`Agregar ${type==="tundish"?"Tundish":"Pote"}`} onClose={onClose} maxWidth={420}>
      <Field label="Nombre"><input className="inp" value={form.name} onChange={e=>setForm({...form,name:e.target.value})} placeholder={type==="tundish"?"Ej: Tundish T-09":"Ej: Pote 96"} /></Field>
      <Field label="Estado inicial">
        <select className="inp" value={form.stateId} onChange={e=>setForm({...form,stateId:e.target.value})}>
          {states.map(s=><option key={s.id} value={s.id}>{s.label}</option>)}
        </select>
      </Field>
      <div style={{ display:"flex", gap:10, justifyContent:"flex-end" }}>
        <button className="btn-secondary" onClick={onClose}>Cancelar</button>
        <button className="btn-primary" onClick={handle} disabled={saving}>{saving?"Guardando...":"Agregar"}</button>
      </div>
    </ModalWrap>
  );
}

// ─── MANAGE STATES MODAL ─────────────────────────────────────────────────────
function ManageStatesModal({ type, states, onClose, onSaved }) {
  const [list, setList] = useState(states.map(s=>({...s})));
  const [newS, setNewS] = useState({ label:"", color:"#58a6ff", zone:"Acería" });
  const [saving, setSaving] = useState(false);
  const COLORS = ["#ef4444","#f97316","#f59e0b","#eab308","#84cc16","#22c55e","#10b981","#14b8a6","#3b82f6","#6366f1","#8b5cf6","#ec4899","#6b7280","#58a6ff","#4ade80"];
  const table = type==="tundish"?"tundish_states":"pote_states";

  function addState() {
    if (!newS.label.trim()) return;
    const id = newS.label.toLowerCase().replace(/\s+/g,"_").replace(/[^a-z0-9_]/g,"")+`_${Date.now()}`;
    setList([...list, { id, label:newS.label.trim(), color:newS.color, ...(type==="tundish"?{zone:newS.zone}:{}), position:list.length }]);
    setNewS({ label:"", color:"#58a6ff", zone:"Acería" });
  }
  async function saveAll() {
    setSaving(true);
    for (const s of list) {
      await supabase.from(table).upsert({ id:s.id, label:s.label, color:s.color, position:s.position||0, ...(type==="tundish"?{zone:s.zone}:{}) });
    }
    setSaving(false); onSaved(); onClose();
  }

  return (
    <ModalWrap title={`Estados — ${type==="tundish"?"Tundish":"Potes"}`} onClose={onClose} maxWidth={640}>
      <div style={{ maxHeight:300, overflowY:"auto", display:"flex", flexDirection:"column", gap:8, marginBottom:20 }}>
        {list.map(s => (
          <div key={s.id} style={{ display:"flex", alignItems:"center", gap:8, padding:"8px 10px", background:"#0d1117", borderRadius:8 }}>
            <input type="color" value={s.color} onChange={e=>setList(list.map(x=>x.id===s.id?{...x,color:e.target.value}:x))}
              style={{ width:28, height:28, border:"none", borderRadius:4, cursor:"pointer", padding:0 }} />
            <input className="inp" value={s.label} onChange={e=>setList(list.map(x=>x.id===s.id?{...x,label:e.target.value}:x))}
              style={{ flex:1, padding:"6px 10px", fontSize:13 }} />
            {type==="tundish" && (
              <select className="inp" value={s.zone||"Acería"} onChange={e=>setList(list.map(x=>x.id===s.id?{...x,zone:e.target.value}:x))}
                style={{ width:110, padding:"6px 8px", fontSize:12 }}>
                {TUNDISH_ZONES.map(z=><option key={z}>{z}</option>)}
              </select>
            )}
            <button onClick={()=>setList(list.filter(x=>x.id!==s.id))} style={{ background:"#da363322", border:"1px solid #da363366", borderRadius:6, color:"#ff7b72", cursor:"pointer", padding:"5px 10px", fontSize:12 }}>✕</button>
          </div>
        ))}
      </div>
      <div style={{ padding:14, background:"#0d1117", borderRadius:10, border:"1px solid #30363d", marginBottom:16 }}>
        <div style={{ fontSize:12, fontWeight:700, color:"#8b949e", marginBottom:10, textTransform:"uppercase" }}>Nuevo Estado</div>
        <div style={{ display:"flex", gap:8, alignItems:"center", flexWrap:"wrap" }}>
          <input type="color" value={newS.color} onChange={e=>setNewS({...newS,color:e.target.value})}
            style={{ width:36, height:36, border:"none", borderRadius:6, cursor:"pointer", padding:0, flexShrink:0 }} />
          <input className="inp" value={newS.label} onChange={e=>setNewS({...newS,label:e.target.value})} placeholder="Nombre del estado" style={{ flex:1, minWidth:140 }} />
          {type==="tundish" && (
            <select className="inp" value={newS.zone} onChange={e=>setNewS({...newS,zone:e.target.value})} style={{ width:110 }}>
              {TUNDISH_ZONES.map(z=><option key={z}>{z}</option>)}
            </select>
          )}
          <button className="btn-primary" onClick={addState} style={{ whiteSpace:"nowrap" }}>+ Agregar</button>
        </div>
        <div style={{ display:"flex", flexWrap:"wrap", gap:6, marginTop:10 }}>
          {COLORS.map(c=>(
            <div key={c} onClick={()=>setNewS({...newS,color:c})}
              style={{ width:20, height:20, borderRadius:"50%", background:c, cursor:"pointer", border:newS.color===c?"2px solid #fff":"2px solid transparent" }} />
          ))}
        </div>
      </div>
      <div style={{ display:"flex", gap:10, justifyContent:"flex-end" }}>
        <button className="btn-secondary" onClick={onClose}>Cancelar</button>
        <button className="btn-primary" onClick={saveAll} disabled={saving}>{saving?"Guardando...":"Guardar Cambios"}</button>
      </div>
    </ModalWrap>
  );
}

// ─── NOTIF MODAL ──────────────────────────────────────────────────────────────
function NotifModal({ config, onClose, onSaved }) {
  const [form, setForm] = useState(config);
  const [saving, setSaving] = useState(false);
  async function save() {
    setSaving(true);
    await supabase.from("notif_config").update({ email_list:form.email_list, smtp_host:form.smtp_host, smtp_user:form.smtp_user, smtp_pass:form.smtp_pass, wa_token:form.wa_token, wa_group_id:form.wa_group_id, updated_at:new Date().toISOString() }).eq("id",1);
    setSaving(false); onSaved(form); onClose();
  }
  return (
    <ModalWrap title="🔔 Notificaciones" onClose={onClose} maxWidth={500}>
      <div style={{ background:"#0d1117", border:"1px solid #58a6ff44", borderRadius:10, padding:14, marginBottom:20, fontSize:13, color:"#8b949e" }}>
        💡 Se notifica cuando alguien agrega un comentario.
      </div>
      <div style={{ fontFamily:"'Barlow Condensed',sans-serif", fontSize:16, fontWeight:700, color:"#f59e0b", marginBottom:12 }}>📧 EMAIL</div>
      <Field label="Destinatarios (separados por coma)"><input className="inp" value={form.email_list||""} onChange={e=>setForm({...form,email_list:e.target.value})} placeholder="juan@planta.com, maria@planta.com" /></Field>
      <Field label="SMTP Host"><input className="inp" value={form.smtp_host||""} onChange={e=>setForm({...form,smtp_host:e.target.value})} placeholder="smtp.gmail.com" /></Field>
      <Field label="SMTP User"><input className="inp" value={form.smtp_user||""} onChange={e=>setForm({...form,smtp_user:e.target.value})} placeholder="notif@planta.com" /></Field>
      <Field label="SMTP Password"><input className="inp" type="password" value={form.smtp_pass||""} onChange={e=>setForm({...form,smtp_pass:e.target.value})} placeholder="••••••••" /></Field>
      <div style={{ fontFamily:"'Barlow Condensed',sans-serif", fontSize:16, fontWeight:700, color:"#25d366", margin:"20px 0 12px" }}>📱 WHATSAPP</div>
      <Field label="API Token"><input className="inp" value={form.wa_token||""} onChange={e=>setForm({...form,wa_token:e.target.value})} placeholder="Tu API token" /></Field>
      <Field label="ID del Grupo"><input className="inp" value={form.wa_group_id||""} onChange={e=>setForm({...form,wa_group_id:e.target.value})} placeholder="120363xxxxxxx@g.us" /></Field>
      <div style={{ display:"flex", gap:10, justifyContent:"flex-end", marginTop:8 }}>
        <button className="btn-secondary" onClick={onClose}>Cancelar</button>
        <button className="btn-primary" onClick={save} disabled={saving}>{saving?"Guardando...":"Guardar"}</button>
      </div>
    </ModalWrap>
  );
}

// ─── MAIN APP ─────────────────────────────────────────────────────────────────
export default function App() {
  const [currentUser, setCurrentUser] = useState(null);
  const [activeTab, setActiveTab] = useState("tundish");
  const [tundishList, setTundishList] = useState([]);
  const [potesList, setPotesList] = useState([]);
  const [tundishStates, setTundishStates] = useState([]);
  const [poteStates, setPoteStates] = useState([]);
  const [notifConfig, setNotifConfig] = useState({});
  const [loading, setLoading] = useState(true);
  const [selected, setSelected] = useState(null);
  const [showAdd, setShowAdd] = useState(false);
  const [showNotif, setShowNotif] = useState(false);
  const [showManageStates, setShowManageStates] = useState(false);
  const [tView, setTView] = useState("zone");
  const [tZone, setTZone] = useState(null);
  const [tState, setTState] = useState(null);
  const [pView, setPView] = useState("state");
  const [pState, setPState] = useState(null);

  const isAdmin = currentUser?.role === "admin";
  const isTundish = activeTab === "tundish";
  const items = isTundish ? tundishList : potesList;
  const states = isTundish ? tundishStates : poteStates;

  useEffect(() => {
    supabase.auth.getSession().then(async ({ data: { session } }) => {
      if (session) {
        const { data: profile } = await supabase.from("profiles").select("*").eq("id", session.user.id).single();
        if (profile) setCurrentUser({ ...session.user, ...profile });
      }
    });
  }, []);

  const loadData = useCallback(async () => {
    setLoading(true);
    const [tsRes, psRes, tRes, pRes, ncRes] = await Promise.all([
      supabase.from("tundish_states").select("*").order("position"),
      supabase.from("pote_states").select("*").order("position"),
      supabase.from("tundish").select("*").order("name"),
      supabase.from("potes").select("*").order("name"),
      supabase.from("notif_config").select("*").eq("id",1).single(),
    ]);
    setTundishStates(tsRes.data||[]);
    setPoteStates(psRes.data||[]);
    setNotifConfig(ncRes.data||{});
    const tItems = tRes.data||[];
    const pItems = pRes.data||[];
    const [thRes, phRes] = await Promise.all([
      supabase.from("tundish_history").select("*").order("created_at", { ascending:false }),
      supabase.from("pote_history").select("*").order("created_at", { ascending:false }),
    ]);
    const thAll = thRes.data||[];
    const phAll = phRes.data||[];
    setTundishList(tItems.map(t => ({ ...t, history: thAll.filter(h=>h.tundish_id===t.id) })));
    setPotesList(pItems.map(p => ({ ...p, history: phAll.filter(h=>h.pote_id===p.id) })));
    setLoading(false);
  }, []);

  useEffect(() => { if (currentUser) loadData(); }, [currentUser, loadData]);

  useEffect(() => {
    if (!currentUser) return;
    const ch = supabase.channel("realtime-all")
      .on("postgres_changes", { event:"*", schema:"public", table:"tundish" }, () => loadData())
      .on("postgres_changes", { event:"*", schema:"public", table:"potes" }, () => loadData())
      .on("postgres_changes", { event:"*", schema:"public", table:"tundish_history" }, () => loadData())
      .on("postgres_changes", { event:"*", schema:"public", table:"pote_history" }, () => loadData())
      .subscribe();
    return () => supabase.removeChannel(ch);
  }, [currentUser, loadData]);

  if (!currentUser) return <><style>{CSS}</style><AuthScreen onLogin={setCurrentUser} /></>;
  if (loading) return <><style>{CSS}</style><div style={{ minHeight:"100vh", background:"#0d1117" }}><Loader text="Cargando datos..." /></div></>;

  return (
    <>
      <style>{CSS}</style>
      <div style={{ minHeight:"100vh", background:"#0d1117" }}>
        <div style={{ background:"#161b22", borderBottom:"1px solid #21262d", padding:"0 24px", position:"sticky", top:0, zIndex:100 }}>
          <div style={{ maxWidth:1280, margin:"0 auto", display:"flex", alignItems:"center", height:56, gap:12 }}>
            <div style={{ display:"flex", alignItems:"center", gap:10, marginRight:"auto" }}>
              <span style={{ fontSize:22 }}>🏭</span>
              <span style={{ fontFamily:"'Barlow Condensed',sans-serif", fontSize:22, fontWeight:800, color:"#e6edf3" }}>
                STEEL<span style={{ color:"#f59e0b" }}>TRACK</span>
              </span>
            </div>
            <div style={{ display:"flex", background:"#0d1117", borderRadius:10, padding:4, gap:2 }}>
              {[["tundish","⚗️ Tundish"],["potes","🪣 Potes"]].map(([t,l]) => (
                <button key={t} className={`tab-btn ${activeTab===t?"active":""}`} onClick={()=>setActiveTab(t)}
                  style={{ padding:"6px 16px", border:"none", borderRadius:8, cursor:"pointer", fontSize:13, fontWeight:600,
                    background:"transparent", color:activeTab===t?"#0d1117":"#8b949e", fontFamily:"'Barlow',sans-serif" }}>{l}</button>
              ))}
            </div>
            <button onClick={()=>exportToExcel(items, states, isTundish?"Tundish":"Potes")}
              style={{ background:"#21262d", border:"1px solid #30363d", borderRadius:8, padding:"6px 12px", cursor:"pointer", color:"#8b949e", fontSize:13, fontFamily:"'Barlow',sans-serif" }}>
              📥 Excel
            </button>
            {isAdmin && <>
              <button onClick={()=>setShowManageStates(true)}
                style={{ background:"#21262d", border:"1px solid #7c3aed66", borderRadius:8, padding:"6px 12px", cursor:"pointer", color:"#a78bfa", fontSize:13, fontFamily:"'Barlow',sans-serif" }}>
                ⚙️ Estados
              </button>
              <button onClick={()=>setShowNotif(true)}
                style={{ background:"#21262d", border:"1px solid #30363d", borderRadius:8, padding:"6px 12px", cursor:"pointer", color:"#8b949e", fontSize:13, fontFamily:"'Barlow',sans-serif" }}>
                🔔 Notif.
              </button>
            </>}
            <div style={{ display:"flex", alignItems:"center", gap:8, padding:"5px 12px", background:"#21262d", borderRadius:8 }}>
              <div style={{ width:28, height:28, borderRadius:"50%", background:isAdmin?"#7c3aed":"#58a6ff", display:"flex", alignItems:"center", justifyContent:"center", fontSize:13, fontWeight:700, color:"#fff" }}>
                {currentUser.name?.[0]?.toUpperCase()||"U"}
              </div>
              <div>
                <div style={{ fontSize:13, color:"#c9d1d9", fontWeight:500, lineHeight:1 }}>{currentUser.name}</div>
                {isAdmin && <div className="admin-badge" style={{ marginTop:3 }}>👑 Admin</div>}
              </div>
              <button onClick={async()=>{await supabase.auth.signOut();setCurrentUser(null);}}
                style={{ background:"none", border:"none", color:"#6e7681", cursor:"pointer", fontSize:11, marginLeft:4 }}>Salir</button>
            </div>
          </div>
        </div>

        <div style={{ maxWidth:1280, margin:"0 auto", padding:"28px 24px" }}>
          <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:20, flexWrap:"wrap", gap:12 }}>
            <div>
              <h2 style={{ fontFamily:"'Barlow Condensed',sans-serif", fontSize:28, fontWeight:800, color:"#e6edf3" }}>
                {isTundish?"⚗️ Tundish":"🪣 Potes"}
              </h2>
              <p style={{ color:"#8b949e", fontSize:13, marginTop:2 }}>{items.length} {isTundish?"tundish":"potes"} en seguimiento</p>
            </div>
            <div style={{ display:"flex", gap:10, alignItems:"center" }}>
              <div style={{ display:"flex", background:"#161b22", border:"1px solid #30363d", borderRadius:8, overflow:"hidden" }}>
                {(isTundish?[["grid","Grilla"],["zone","Por Zona"]]:[["grid","Grilla"],["state","Por Estado"]]).map(([v,l]) => (
                  <button key={v} onClick={()=>{ isTundish?(setTView(v),setTZone(null),setTState(null)):(setPView(v),setPState(null)); }}
                    style={{ padding:"7px 16px", border:"none", cursor:"pointer", fontSize:12, fontWeight:600, fontFamily:"'Barlow',sans-serif",
                      background:(isTundish?tView:pView)===v?"#58a6ff":"transparent", color:(isTundish?tView:pView)===v?"#0d1117":"#8b949e" }}>{l}</button>
                ))}
              </div>
              {isAdmin && <button className="btn-primary" onClick={()=>setShowAdd(true)}>+ Agregar {isTundish?"Tundish":"Pote"}</button>}
            </div>
          </div>

          {isTundish && tView==="grid" && (
            <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill, minmax(160px,1fr))", gap:8 }}>
              {items.map(item=><ItemCard key={item.id} item={item} states={states} onClick={()=>setSelected(item)} compact />)}
            </div>
          )}

          {isTundish && tView==="zone" && (
            <>
              <div style={{ marginBottom:20 }}>
                <div style={{ fontSize:11, fontWeight:700, color:"#6e7681", textTransform:"uppercase", letterSpacing:".07em", marginBottom:10 }}>Seleccioná una zona</div>
                <div style={{ display:"flex", gap:10, flexWrap:"wrap" }}>
                  {TUNDISH_ZONES.map(zone => {
                    const zoneStates = tundishStates.filter(s=>s.zone===zone);
                    const count = items.filter(i=>zoneStates.find(s=>s.id===i.state_id)).length;
                    const active = tZone===zone;
                    return (
                      <button key={zone} onClick={()=>{setTZone(active?null:zone);setTState(null);}}
                        style={{ padding:"10px 20px", border:active?"2px solid #f59e0b":"2px solid #30363d", borderRadius:10, cursor:"pointer",
                          fontFamily:"'Barlow Condensed',sans-serif", fontSize:17, fontWeight:700,
                          background:active?"#f59e0b18":"#161b22", color:active?"#f59e0b":"#8b949e", transition:"all .15s" }}>
                        {zone} <span style={{ fontSize:12, fontWeight:400, opacity:.6 }}>{count}</span>
                      </button>
                    );
                  })}
                </div>
              </div>
              {tZone && (
                <div style={{ marginBottom:20 }}>
                  <div style={{ fontSize:11, fontWeight:700, color:"#6e7681", textTransform:"uppercase", letterSpacing:".07em", marginBottom:10 }}>Estado en {tZone}</div>
                  <div style={{ display:"flex", gap:8, flexWrap:"wrap" }}>
                    {tundishStates.filter(s=>s.zone===tZone).map(s => {
                      const count = items.filter(i=>i.state_id===s.id).length;
                      const active = tState===s.id;
                      const limit = TUNDISH_STATE_LIMITS[s.id];
                      return (
                        <button key={s.id} onClick={()=>setTState(active?null:s.id)}
                          style={{ display:"inline-flex", alignItems:"center", gap:7, padding:"7px 14px",
                            border:`2px solid ${active?s.color:"#30363d"}`, borderRadius:8, cursor:"pointer",
                            fontFamily:"'Barlow',sans-serif", fontSize:13, fontWeight:600,
                            background:active?s.color+"22":"#161b22", color:active?s.color:"#8b949e", transition:"all .15s" }}>
                          <span style={{ width:8, height:8, borderRadius:"50%", background:s.color, display:"inline-block" }} />
                          {s.label} <span style={{ fontSize:11, opacity:.7 }}>({count}{limit!==undefined?`/${limit}`:""})</span>
                        </button>
                      );
                    })}
                  </div>
                </div>
              )}
              {tZone && tState && (()=>{
                const filtered = items.filter(i=>i.state_id===tState);
                const st = tundishStates.find(s=>s.id===tState);
                return (
                  <div>
                    <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:14 }}>
                      <span style={{ width:10, height:10, borderRadius:"50%", background:st?.color, display:"inline-block" }} />
                      <span style={{ fontFamily:"'Barlow Condensed',sans-serif", fontSize:16, fontWeight:700, color:"#e6edf3" }}>{st?.label}</span>
                      <span style={{ fontSize:12, color:"#6e7681", background:"#21262d", padding:"2px 8px", borderRadius:10 }}>{filtered.length} tundish</span>
                    </div>
                    {filtered.length===0
                      ? <p style={{ color:"#6e7681", fontSize:14, padding:"24px 0" }}>No hay tundish en este estado.</p>
                      : <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill, minmax(200px,1fr))", gap:12 }}>
                          {filtered.map(item=><ItemCard key={item.id} item={item} states={states} onClick={()=>setSelected(item)} />)}
                        </div>}
                  </div>
                );
              })()}
              {tZone && !tState && <p style={{ color:"#6e7681", fontSize:13, padding:"8px 0" }}>Seleccioná un estado para ver los tundish.</p>}
            </>
          )}

          {!isTundish && pView==="grid" && (
            <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill, minmax(160px,1fr))", gap:8 }}>
              {items.map(item=><ItemCard key={item.id} item={item} states={states} onClick={()=>setSelected(item)} compact />)}
            </div>
          )}

          {!isTundish && pView==="state" && (
            <>
              <div style={{ marginBottom:20 }}>
                <div style={{ fontSize:11, fontWeight:700, color:"#6e7681", textTransform:"uppercase", letterSpacing:".07em", marginBottom:10 }}>Seleccioná un estado</div>
                <div style={{ display:"flex", gap:8, flexWrap:"wrap" }}>
                  {poteStates.map(s => {
                    const count = items.filter(i=>i.state_id===s.id).length;
                    const active = pState===s.id;
                    return (
                      <button key={s.id} onClick={()=>setPState(active?null:s.id)}
                        style={{ display:"inline-flex", alignItems:"center", gap:7, padding:"8px 16px",
                          border:`2px solid ${active?s.color:"#30363d"}`, borderRadius:8, cursor:"pointer",
                          fontFamily:"'Barlow',sans-serif", fontSize:13, fontWeight:600,
                          background:active?s.color+"22":"#161b22", color:active?s.color:"#8b949e", transition:"all .15s" }}>
                        <span style={{ width:8, height:8, borderRadius:"50%", background:s.color, display:"inline-block" }} />
                        {s.label} <span style={{ fontSize:11, opacity:.7 }}>({count})</span>
                      </button>
                    );
                  })}
                </div>
              </div>
              {pState && (()=>{
                const filtered = items.filter(i=>i.state_id===pState);
                const st = poteStates.find(s=>s.id===pState);
                return (
                  <div>
                    <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:14 }}>
                      <span style={{ width:10, height:10, borderRadius:"50%", background:st?.color, display:"inline-block" }} />
                      <span style={{ fontFamily:"'Barlow Condensed',sans-serif", fontSize:16, fontWeight:700, color:"#e6edf3" }}>{st?.label}</span>
                      <span style={{ fontSize:12, color:"#6e7681", background:"#21262d", padding:"2px 8px", borderRadius:10 }}>{filtered.length} potes</span>
                    </div>
                    {filtered.length===0
                      ? <p style={{ color:"#6e7681", fontSize:14, padding:"24px 0" }}>No hay potes en este estado.</p>
                      : <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill, minmax(200px,1fr))", gap:12 }}>
                          {filtered.map(item=><ItemCard key={item.id} item={item} states={states} onClick={()=>setSelected(item)} />)}
                        </div>}
                  </div>
                );
              })()}
              {!pState && <p style={{ color:"#6e7681", fontSize:13, padding:"8px 0" }}>Seleccioná un estado para ver los potes.</p>}
            </>
          )}
        </div>
      </div>

      {selected && (
        <DetailModal item={selected} states={states} currentUser={currentUser}
          allTundish={tundishList} isTundish={isTundish}
          onClose={()=>setSelected(null)} onUpdated={()=>{ loadData(); setSelected(null); }} />
      )}
      {showAdd && isAdmin && (
        <AddItemModal type={activeTab} states={states} onClose={()=>setShowAdd(false)} onAdded={loadData} />
      )}
      {showNotif && isAdmin && (
        <NotifModal config={notifConfig} onClose={()=>setShowNotif(false)} onSaved={cfg=>setNotifConfig(cfg)} />
      )}
      {showManageStates && isAdmin && (
        <ManageStatesModal type={activeTab} states={states} onClose={()=>setShowManageStates(false)} onSaved={loadData} />
      )}
    </>
  );
}
