import React, { useState, useEffect, useMemo, useRef } from "react";
import { supabase } from "./supabaseClient.js";
import * as XLSX from "xlsx";

/* ------------------------------------------------------------------ *
 * Benchstock — Shared Lab Inventory (Supabase backend)
 *   • real email + password accounts, password reset
 *   • one shared inventory for the whole lab, live-synced
 *   • lot tracking, printable Code 39 barcode labels
 *   • consume stock per use, attributed to the signed-in user
 *   • low-stock and expiration alerts
 * ------------------------------------------------------------------ */

const T = {
  paper:"#F6F7F8", card:"#FFFFFF", ink:"#16202B", inkSoft:"#5C6A78", line:"#E4E8EC",
  teal:"#0F6E6A", tealDeep:"#0A5450", amber:"#A9680A", amberBg:"#FBF1DD",
  red:"#AE382D", redBg:"#FAE8E5", green:"#2C7A57", greenBg:"#E6F2EB",
};
const CATEGORIES = ["Antibody","Reagent","Kit","Consumable","Chemical","Media","Other"];

/* ---------------------- row <-> object mapping -------------------- */
const fromItem = (r) => ({
  id:r.id, name:r.name, category:r.category||"Other", lotNumber:r.lot_number||"",
  catalogNumber:r.catalog_number||"", supplier:r.supplier||"", quantity:Number(r.quantity)||0,
  unit:r.unit||"", reorderThreshold:Number(r.reorder_threshold)||0, location:r.location||"",
  expirationDate:r.expiration_date||"", notes:r.notes||"", createdBy:r.created_by||"", createdAt:r.created_at,
});
const toItem = (d) => ({
  name:d.name, category:d.category, lot_number:d.lotNumber||null, catalog_number:d.catalogNumber||null,
  supplier:d.supplier||null, quantity:d.quantity, unit:d.unit||null, reorder_threshold:d.reorderThreshold,
  location:d.location||null, expiration_date:d.expirationDate||null, notes:d.notes||null, created_by:d.createdBy||null,
});
const fromLog = (r) => ({ id:r.id, name:r.name, lot:r.lot, amount:Number(r.amount), unit:r.unit, who:r.who, ts:new Date(r.ts).getTime() });
const fromSlide = (r) => ({ id:r.id, subjectId:r.subject_id||"", species:r.species||"Mouse", region:r.region||"", staining:r.staining||"Unstained", bigBox:r.big_box||"", box:r.box||"", project:r.project||"", notes:r.notes||"", createdBy:r.created_by||"", createdAt:r.created_at });
const toSlide = (d) => ({ subject_id:d.subjectId||null, species:d.species, region:d.region||null, staining:d.staining||"Unstained", big_box:d.bigBox||null, box:d.box||null, project:d.project||null, notes:d.notes||null, created_by:d.createdBy||null });
const fromSlideLog = (r) => ({ id:r.id, subjectId:r.subject_id, region:r.region, fromStain:r.from_stain, toStain:r.to_stain, who:r.who, ts:new Date(r.ts).getTime() });

/* --------------------------- Code 39 ------------------------------ */
const C39={ "0":"nnnwwnwnn","1":"wnnwnnnnw","2":"nnwwnnnnw","3":"wnwwnnnnn","4":"nnnwwnnnw","5":"wnnwwnnnn","6":"nnwwwnnnn","7":"nnnwnnwnw","8":"wnnwnnwnn","9":"nnwwnnwnn","A":"wnnnnwnnw","B":"nnwnnwnnw","C":"wnwnnwnnn","D":"nnnnwwnnw","E":"wnnnwwnnn","F":"nnwnwwnnn","G":"nnnnnwwnw","H":"wnnnnwwnn","I":"nnwnnwwnn","J":"nnnnwwwnn","K":"wnnnnnnww","L":"nnwnnnnww","M":"wnwnnnnwn","N":"nnnnwnnww","O":"wnnnwnnwn","P":"nnwnwnnwn","Q":"nnnnnnwww","R":"wnnnnnwwn","S":"nnwnnnwwn","T":"nnnnwnwwn","U":"wwnnnnnnw","V":"nwwnnnnnw","W":"wwwnnnnnn","X":"nwnnwnnnw","Y":"wwnnwnnnn","Z":"nwwnwnnnn","-":"nwnnnnwnw",".":"wwnnnnwnn"," ":"nwwnnnwnn","$":"nwnwnwnnn","/":"nwnwnnnwn","+":"nwnnnwnwn","%":"nnnwnwnwn","*":"nwnnwnwnn" };
function Barcode({ value, height=56, narrow=2 }){
  const text=(value||"").toUpperCase().replace(/[^0-9A-Z\-. $/+%]/g,"");
  const wide=narrow*3, gap=narrow; const rects=[]; let x=0;
  `*${text}*`.split("").forEach((ch,ci)=>{ const pat=C39[ch]; if(!pat) return;
    for(let i=0;i<pat.length;i++){ const w=pat[i]==="w"?wide:narrow; if(i%2===0) rects.push(<rect key={`${ci}-${i}`} x={x} y={0} width={w} height={height} fill="#16202B"/>); x+=w; } x+=gap; });
  return <svg width={x} height={height} viewBox={`0 0 ${x} ${height}`} style={{maxWidth:"100%"}}>{rects}</svg>;
}

/* --------------------------- helpers ------------------------------ */
function daysUntil(s){ if(!s) return null; const d=new Date(s+"T00:00:00"); const n=new Date(); n.setHours(0,0,0,0); return Math.round((d-n)/86400000); }
function itemStatus(it){ const d=daysUntil(it.expirationDate);
  if(d!==null && d<0) return "expired"; if(it.quantity<=0) return "out";
  if(it.quantity<=it.reorderThreshold) return "low"; if(d!==null && d<=60) return "expiring"; return "ok"; }
const STATUS_META={ ok:{label:"In stock",c:T.green,bg:T.greenBg}, low:{label:"Low stock",c:T.amber,bg:T.amberBg}, out:{label:"Out of stock",c:T.red,bg:T.redBg}, expiring:{label:"Expiring soon",c:T.amber,bg:T.amberBg}, expired:{label:"Expired",c:T.red,bg:T.redBg} };
function fmtDate(s){ if(!s) return "—"; return new Date(s+"T00:00:00").toLocaleDateString(undefined,{year:"numeric",month:"short",day:"numeric"}); }

function LabMark(){
  return (
    <svg width="62%" height="62%" viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <path d="M9 3.5H15" stroke="#fff" strokeWidth="1.6" strokeLinecap="round"/>
      <path d="M10 3.5V9.5L5.6 18.4C5.1 19.5 5.9 20.6 7 20.6H17C18.1 20.6 18.9 19.5 18.4 18.4L14 9.5V3.5"
            stroke="#fff" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"/>
      <path d="M7.35 15H16.65L18.4 18.4C18.9 19.5 18.1 20.6 17 20.6H7C5.9 20.6 5.1 19.5 5.6 18.4L7.35 15Z"
            fill="#fff" fillOpacity="0.9"/>
      <circle cx="10.4" cy="17.6" r="0.85" fill="#0F6E6A"/>
      <circle cx="13.4" cy="18.7" r="0.6" fill="#0F6E6A"/>
    </svg>
  );
}

/* ============================== APP =============================== */
export default function App(){
  const [session,setSession]=useState(undefined); // undefined = checking
  const [me,setMe]=useState(null);
  const [items,setItems]=useState([]);
  const [log,setLog]=useState([]);
  const [members,setMembers]=useState([]);
  const [slides,setSlides]=useState([]);
  const [slideLog,setSlideLog]=useState([]);

  const [view,setView]=useState("dashboard");
  const [query,setQuery]=useState("");
  const [catFilter,setCatFilter]=useState("All");
  const [editing,setEditing]=useState(null);
  const [labelItem,setLabelItem]=useState(null);
  const [consumeItem,setConsumeItem]=useState(null);
  const [editingSlide,setEditingSlide]=useState(null);
  const [stainSlide,setStainSlide]=useState(null);

  // auth bootstrap
  useEffect(()=>{
    supabase.auth.getSession().then(({data})=>setSession(data.session));
    const { data:sub } = supabase.auth.onAuthStateChange((_e,s)=>setSession(s));
    return ()=>sub.subscription.unsubscribe();
  },[]);

  async function loadItems(){ const { data } = await supabase.from("items").select("*").order("created_at",{ascending:false}); setItems((data||[]).map(fromItem)); }
  async function loadLog(){ const { data } = await supabase.from("usage_log").select("*").order("ts",{ascending:false}).limit(300); setLog((data||[]).map(fromLog)); }
  async function loadMembers(){ const { data } = await supabase.from("profiles").select("*").order("created_at"); setMembers(data||[]); }
  async function loadSlides(){ const { data } = await supabase.from("slides").select("*").order("created_at",{ascending:false}); setSlides((data||[]).map(fromSlide)); }
  async function loadSlideLog(){ const { data } = await supabase.from("slide_log").select("*").order("ts",{ascending:false}).limit(200); setSlideLog((data||[]).map(fromSlideLog)); }

  // when signed in: ensure profile, load data, subscribe to live changes
  useEffect(()=>{
    if(session===undefined) return;
    if(!session){ setMe(null); return; }
    const u = session.user;
    const displayName = u.user_metadata?.display_name || u.email;
    setMe({ id:u.id, email:u.email, displayName });
    (async()=>{
      await supabase.from("profiles").upsert({ id:u.id, display_name:displayName }, { onConflict:"id" });
      await Promise.all([loadItems(), loadLog(), loadMembers(), loadSlides(), loadSlideLog()]);
    })();
    const ch = supabase.channel("benchstock-live")
      .on("postgres_changes",{event:"*",schema:"public",table:"items"}, loadItems)
      .on("postgres_changes",{event:"*",schema:"public",table:"usage_log"}, loadLog)
      .on("postgres_changes",{event:"*",schema:"public",table:"profiles"}, loadMembers)
      .on("postgres_changes",{event:"*",schema:"public",table:"slides"}, loadSlides)
      .on("postgres_changes",{event:"*",schema:"public",table:"slide_log"}, loadSlideLog)
      .subscribe();
    return ()=>supabase.removeChannel(ch);
  },[session]);

  async function upsert(data){
    if(data.id){ await supabase.from("items").update(toItem(data)).eq("id",data.id); }
    else { await supabase.from("items").insert({ ...toItem(data), created_by:me.displayName }); }
    setEditing(null); loadItems();
  }
  async function remove(id){ await supabase.from("items").delete().eq("id",id); loadItems(); }
  async function consume(item, amount){
    const amt = Math.min(amount, item.quantity);
    await supabase.from("items").update({ quantity:+(item.quantity-amt).toFixed(3) }).eq("id",item.id);
    await supabase.from("usage_log").insert({ item_id:item.id, name:item.name, lot:item.lotNumber, amount:amt, unit:item.unit, who:me.displayName });
    setConsumeItem(null); loadItems(); loadLog();
  }
  async function signOut(){ await supabase.auth.signOut(); setView("dashboard"); }

  async function upsertSlide(data){
    if(data.id){ await supabase.from("slides").update(toSlide(data)).eq("id",data.id); }
    else { await supabase.from("slides").insert({ ...toSlide(data), created_by:me.displayName }); }
    setEditingSlide(null); loadSlides();
  }
  async function removeSlide(id){ await supabase.from("slides").delete().eq("id",id); loadSlides(); }
  async function bulkAddSlides(rows){
    const payload = rows.map(r=>({ ...toSlide(r), created_by:me.displayName }));
    for(let i=0;i<payload.length;i+=400){ await supabase.from("slides").insert(payload.slice(i,i+400)); }
    loadSlides();
  }
  async function recordStain(slide, newStain){
    const from = slide.staining || "Unstained";
    await supabase.from("slides").update({ staining:newStain }).eq("id",slide.id);
    await supabase.from("slide_log").insert({ slide_id:slide.id, subject_id:slide.subjectId, region:slide.region, from_stain:from, to_stain:newStain, who:me.displayName });
    setStainSlide(null); loadSlides(); loadSlideLog();
  }

  const filtered = useMemo(()=>{ const q=query.trim().toLowerCase();
    return items.filter(it=>{ if(catFilter!=="All" && it.category!==catFilter) return false;
      if(!q) return true; return [it.name,it.lotNumber,it.catalogNumber,it.supplier,it.location].filter(Boolean).join(" ").toLowerCase().includes(q); });
  },[items,query,catFilter]);

  const alerts = useMemo(()=>{ const low=[],expiring=[],expired=[];
    items.forEach(it=>{ const s=itemStatus(it); if(s==="low"||s==="out") low.push(it); if(s==="expiring") expiring.push(it); if(s==="expired") expired.push(it); });
    return { low, expiring, expired };
  },[items]);

  function exportCSV(){
    const cols=["name","category","lotNumber","catalogNumber","supplier","quantity","unit","reorderThreshold","location","expirationDate","notes"];
    const esc=v=>`"${String(v??"").replace(/"/g,'""')}"`;
    const rows=[cols.join(",")].concat(items.map(it=>cols.map(c=>esc(it[c])).join(",")));
    const blob=new Blob([rows.join("\n")],{type:"text/csv"}); const url=URL.createObjectURL(blob);
    const a=document.createElement("a"); a.href=url; a.download="benchstock-inventory.csv"; a.click(); URL.revokeObjectURL(url);
  }

  if(session===undefined) return (<div style={{...styles.app, display:"flex", alignItems:"center", justifyContent:"center"}}><Style/><div style={{color:T.inkSoft, fontFamily:"var(--mono)"}}>Loading…</div></div>);
  if(!session || !me) return <Auth/>;

  const totalAlerts = alerts.low.length+alerts.expiring.length+alerts.expired.length;

  return (
    <div style={styles.app}>
      <Style/>
      <header style={styles.header}>
        <div style={styles.brand}>
          <span style={styles.mark}><LabMark/></span>
          <div><div style={styles.brandName}>Benchstock</div><div style={styles.brandSub}>Shared lab inventory</div></div>
        </div>
        <div style={styles.search}>
          <input className="bs-input" placeholder="Search name, lot, catalog #, supplier…" value={query}
            onChange={e=>{setQuery(e.target.value); setView("inventory");}} style={styles.searchInput}/>
        </div>
        <button className="bs-btn-primary" onClick={()=>setEditing({})}>+ Add item</button>
        <div style={styles.userChip}>
          <span style={styles.avatar}>{me.displayName.slice(0,1).toUpperCase()}</span>
          <span style={{fontSize:13, fontWeight:600}}>{me.displayName}</span>
          <button className="bs-link" onClick={signOut}>Sign out</button>
        </div>
      </header>

      <nav style={styles.tabs}>
        {[["dashboard","Dashboard"],["inventory","Inventory"],["histology","Histology slides"],["log","Usage log"],["members","Members"]].map(([k,l])=>(
          <button key={k} onClick={()=>setView(k)} className="bs-tab" style={{...styles.tab, color:view===k?T.ink:T.inkSoft, borderBottom:view===k?`2px solid ${T.teal}`:"2px solid transparent"}}>
            {l}{k==="dashboard"&&totalAlerts>0&&<span style={styles.tabBadge}>{totalAlerts}</span>}
          </button>
        ))}
        <div style={{flex:1}}/>
        <button className="bs-btn-ghost" onClick={exportCSV} style={styles.exportBtn}>Export CSV</button>
      </nav>

      <main style={styles.main}>
        {view==="dashboard" && <Dashboard items={items} alerts={alerts} memberCount={members.length} onOpen={it=>{setQuery(it.name);setView("inventory");}}/>}
        {view==="inventory" && <Inventory items={filtered} catFilter={catFilter} setCatFilter={setCatFilter} onEdit={setEditing} onDelete={remove} onLabel={setLabelItem} onConsume={setConsumeItem}/>}
        {view==="histology" && <Histology slides={slides} slideLog={slideLog} onAdd={()=>setEditingSlide({})} onEdit={setEditingSlide} onDelete={removeSlide} onStain={setStainSlide} onImport={bulkAddSlides}/>}
        {view==="log" && <UsageLog log={log}/>}
        {view==="members" && <Members members={members} me={me}/>}
      </main>

      {editing && <ItemForm initial={editing} onSave={upsert} onClose={()=>setEditing(null)}/>}
      {labelItem && <LabelModal item={labelItem} onClose={()=>setLabelItem(null)}/>}
      {consumeItem && <ConsumeModal item={consumeItem} user={me} onConsume={consume} onClose={()=>setConsumeItem(null)}/>}
      {editingSlide && <SlideForm initial={editingSlide} onSave={upsertSlide} onClose={()=>setEditingSlide(null)}/>}
      {stainSlide && <StainModal slide={stainSlide} user={me} onConfirm={recordStain} onClose={()=>setStainSlide(null)}/>}
    </div>
  );
}

/* ----------------------------- Auth ------------------------------- */
function Auth(){
  const [mode,setMode]=useState("login"); // login | register | forgot
  const [f,setF]=useState({ email:"", displayName:"", password:"", confirm:"" });
  const [err,setErr]=useState(""); const [msg,setMsg]=useState(""); const [busy,setBusy]=useState(false);
  const set=k=>e=>setF({...f,[k]:e.target.value});

  async function submit(){
    setErr(""); setMsg("");
    if(!f.email.trim()) return setErr("Enter your email.");
    if(mode!=="forgot" && !f.password) return setErr("Enter a password.");
    setBusy(true);
    try{
      if(mode==="register"){
        if(f.password.length<6) throw new Error("Password must be at least 6 characters.");
        if(f.password!==f.confirm) throw new Error("Passwords don't match.");
        const { error } = await supabase.auth.signUp({
          email:f.email.trim(), password:f.password,
          options:{ data:{ display_name:f.displayName.trim()||f.email.trim() } },
        });
        if(error) throw error;
        setMsg("Account created. If email confirmation is on, check your inbox, then sign in.");
        setMode("login");
      } else if(mode==="login"){
        const { error } = await supabase.auth.signInWithPassword({ email:f.email.trim(), password:f.password });
        if(error) throw error;
      } else { // forgot
        const { error } = await supabase.auth.resetPasswordForEmail(f.email.trim(), { redirectTo: window.location.origin });
        if(error) throw error;
        setMsg("Password reset link sent. Check your email.");
      }
    } catch(e){ setErr(e.message||"Something went wrong."); }
    finally{ setBusy(false); }
  }

  return (
    <div style={styles.authWrap}>
      <Style/>
      <div style={styles.authCard}>
        <div style={styles.authBrand}>
          <span style={{...styles.mark, width:40, height:40}}><LabMark/></span>
          <div><div style={{...styles.brandName, fontSize:20}}>Benchstock</div><div style={styles.brandSub}>Shared lab inventory</div></div>
        </div>
        {mode!=="forgot" && (
          <div style={styles.authTabs}>
            <button className="bs-tab" onClick={()=>{setMode("login");setErr("");setMsg("");}} style={{...styles.authTab, ...(mode==="login"?styles.authTabOn:{})}}>Sign in</button>
            <button className="bs-tab" onClick={()=>{setMode("register");setErr("");setMsg("");}} style={{...styles.authTab, ...(mode==="register"?styles.authTabOn:{})}}>Register</button>
          </div>
        )}
        <div style={{display:"flex", flexDirection:"column", gap:12}}>
          <Field label="Email"><input className="bs-input" type="email" value={f.email} onChange={set("email")} placeholder="you@lab.org" autoFocus/></Field>
          {mode==="register" && <Field label="Display name"><input className="bs-input" value={f.displayName} onChange={set("displayName")} placeholder="Georgios S."/></Field>}
          {mode!=="forgot" && <Field label="Password"><input className="bs-input" type="password" value={f.password} onChange={set("password")} placeholder="••••••••" onKeyDown={e=>e.key==="Enter"&&submit()}/></Field>}
          {mode==="register" && <Field label="Confirm password"><input className="bs-input" type="password" value={f.confirm} onChange={set("confirm")} placeholder="••••••••"/></Field>}
          {err && <div style={styles.authErr}>{err}</div>}
          {msg && <div style={styles.authMsg}>{msg}</div>}
          <button className="bs-btn-primary" onClick={submit} disabled={busy} style={{padding:"11px", marginTop:4}}>
            {busy ? "…" : mode==="login" ? "Sign in" : mode==="register" ? "Create account" : "Send reset link"}
          </button>
        </div>
        <div style={styles.authNote}>
          {mode==="forgot"
            ? <button className="bs-link" onClick={()=>{setMode("login");setErr("");setMsg("");}}>← Back to sign in</button>
            : <button className="bs-link" onClick={()=>{setMode("forgot");setErr("");setMsg("");}}>Forgot password?</button>}
        </div>
      </div>
      <div style={styles.authFooter}>Everyone signed in shares the same lab inventory.</div>
    </div>
  );
}

/* -------------------------- Dashboard ----------------------------- */
function Dashboard({ items, alerts, memberCount, onOpen }){
  const cats = CATEGORIES.map(c=>({c, n:items.filter(i=>i.category===c).length})).filter(x=>x.n);
  return (
    <div>
      <div style={styles.statGrid}>
        <Stat n={items.length} label="Items tracked" c={T.teal}/>
        <Stat n={alerts.low.length} label="Low / out of stock" c={alerts.low.length?T.amber:T.inkSoft}/>
        <Stat n={alerts.expiring.length} label="Expiring ≤ 60 days" c={alerts.expiring.length?T.amber:T.inkSoft}/>
        <Stat n={alerts.expired.length} label="Expired" c={alerts.expired.length?T.red:T.inkSoft}/>
      </div>
      <div style={styles.dashCols}>
        <div style={styles.panel}>
          <div style={styles.panelHead}>Needs attention</div>
          {[...alerts.expired,...alerts.low,...alerts.expiring].length===0
            ? <div style={styles.empty}>Everything's stocked and in date. Nothing to reorder.</div>
            : <div>
                {alerts.expired.map(it=><AlertRow key={it.id} it={it} kind="expired" onOpen={onOpen}/>)}
                {alerts.low.map(it=><AlertRow key={it.id} it={it} kind="low" onOpen={onOpen}/>)}
                {alerts.expiring.map(it=><AlertRow key={it.id} it={it} kind="expiring" onOpen={onOpen}/>)}
              </div>}
        </div>
        <div style={styles.panel}>
          <div style={styles.panelHead}>By category</div>
          {cats.map(({c,n})=>{ const max=Math.max(...cats.map(x=>x.n));
            return (<div key={c} style={styles.catRow}>
              <span style={styles.catLabel}>{c}</span>
              <div style={styles.catBarTrack}><div style={{...styles.catBarFill, width:`${(n/max)*100}%`}}/></div>
              <span style={styles.catNum}>{n}</span>
            </div>); })}
          <div style={{...styles.catRow, marginTop:8, borderTop:`1px solid ${T.line}`, paddingTop:12}}>
            <span style={styles.catLabel}>Lab members</span><div style={{flex:1}}/><span style={styles.catNum}>{memberCount}</span>
          </div>
        </div>
      </div>
    </div>
  );
}
function Stat({n,label,c}){ return <div style={styles.statCard}><div style={{...styles.statNum,color:c}}>{n}</div><div style={styles.statLabel}>{label}</div></div>; }
function AlertRow({it,kind,onOpen}){
  const meta=STATUS_META[kind==="low"&&it.quantity<=0?"out":kind];
  const sub = kind==="low" ? `${it.quantity} ${it.unit} left · reorder at ${it.reorderThreshold}`
    : kind==="expired" ? `Expired ${fmtDate(it.expirationDate)}` : `Expires ${fmtDate(it.expirationDate)} · ${daysUntil(it.expirationDate)} days`;
  return (<div style={styles.alertRow} onClick={()=>onOpen(it)}>
    <span style={{...styles.dot, background:meta.c}}/>
    <div style={{flex:1, minWidth:0}}><div style={styles.alertName}>{it.name}</div><div style={styles.alertSub}>{sub}</div></div>
    <span style={{...styles.pill, color:meta.c, background:meta.bg}}>{meta.label}</span>
  </div>);
}

/* -------------------------- Inventory ----------------------------- */
function Inventory({ items, catFilter, setCatFilter, onEdit, onDelete, onLabel, onConsume }){
  return (
    <div>
      <div style={styles.filterBar}>
        {["All",...CATEGORIES].map(c=>(
          <button key={c} onClick={()=>setCatFilter(c)} className="bs-chip"
            style={{...styles.chip, background:catFilter===c?T.ink:"transparent", color:catFilter===c?"#fff":T.inkSoft, borderColor:catFilter===c?T.ink:T.line}}>{c}</button>
        ))}
      </div>
      {items.length===0 ? <div style={{...styles.panel,...styles.empty}}>No items match. Adjust your search or add a new item.</div> : (
        <div style={styles.tableWrap}>
          <table style={styles.table}>
            <thead><tr>
              <th style={styles.th}>Item</th><th style={styles.th}>Lot</th>
              <th style={{...styles.th,textAlign:"right"}}>On hand</th><th style={styles.th}>Location</th>
              <th style={styles.th}>Expires</th><th style={styles.th}>Status</th>
              <th style={{...styles.th,textAlign:"right"}}>Actions</th>
            </tr></thead>
            <tbody>
              {items.map(it=>{ const s=itemStatus(it); const meta=STATUS_META[s];
                return (<tr key={it.id} className="bs-row">
                  <td style={styles.td}><div style={styles.itemName}>{it.name}</div><div style={styles.itemMeta}>{it.category} · {it.supplier||"—"}{it.catalogNumber?` · #${it.catalogNumber}`:""}</div></td>
                  <td style={{...styles.td, fontFamily:"var(--mono)", fontSize:12}}>{it.lotNumber||"—"}</td>
                  <td style={{...styles.td, textAlign:"right", fontFamily:"var(--mono)"}}><span style={{fontWeight:600}}>{it.quantity}</span> <span style={{color:T.inkSoft}}>{it.unit}</span></td>
                  <td style={{...styles.td, color:T.inkSoft, fontSize:12}}>{it.location||"—"}</td>
                  <td style={{...styles.td, fontSize:12}}>{fmtDate(it.expirationDate)}</td>
                  <td style={styles.td}><span style={{...styles.pill, color:meta.c, background:meta.bg}}>{meta.label}</span></td>
                  <td style={{...styles.td, textAlign:"right", whiteSpace:"nowrap"}}>
                    <button className="bs-act" onClick={()=>onConsume(it)}>Use</button>
                    <button className="bs-act" onClick={()=>onLabel(it)}>Label</button>
                    <button className="bs-act" onClick={()=>onEdit(it)}>Edit</button>
                    <button className="bs-act bs-act-danger" onClick={()=>onDelete(it.id)}>✕</button>
                  </td>
                </tr>); })}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

/* -------------------------- Usage log ----------------------------- */
function UsageLog({ log }){
  if(!log.length) return <div style={{...styles.panel,...styles.empty}}>No usage recorded yet. Hit “Use” on any item to log consumption.</div>;
  return (
    <div style={styles.tableWrap}>
      <table style={styles.table}>
        <thead><tr><th style={styles.th}>When</th><th style={styles.th}>Item</th><th style={styles.th}>Lot</th><th style={{...styles.th,textAlign:"right"}}>Used</th><th style={styles.th}>By</th></tr></thead>
        <tbody>{log.map(e=>(<tr key={e.id} className="bs-row">
          <td style={{...styles.td, fontSize:12, color:T.inkSoft}}>{new Date(e.ts).toLocaleString()}</td>
          <td style={styles.td}>{e.name}</td>
          <td style={{...styles.td, fontFamily:"var(--mono)", fontSize:12}}>{e.lot||"—"}</td>
          <td style={{...styles.td, textAlign:"right", fontFamily:"var(--mono)"}}>{e.amount} {e.unit}</td>
          <td style={{...styles.td, fontSize:12}}>{e.who||"—"}</td>
        </tr>))}</tbody>
      </table>
    </div>
  );
}

/* --------------------------- Members ------------------------------ */
function Members({ members, me }){
  return (
    <div style={styles.tableWrap}>
      <table style={styles.table}>
        <thead><tr><th style={styles.th}>Member</th><th style={styles.th}>Role</th><th style={styles.th}>Joined</th></tr></thead>
        <tbody>{members.map(u=>(<tr key={u.id} className="bs-row">
          <td style={styles.td}><div style={styles.itemName}>{u.display_name}{u.id===me.id?" (you)":""}</div></td>
          <td style={styles.td}><span style={{...styles.pill, color:u.role==="admin"?T.teal:T.inkSoft, background:u.role==="admin"?T.greenBg:T.paper}}>{u.role}</span></td>
          <td style={{...styles.td, fontSize:12, color:T.inkSoft}}>{u.created_at?new Date(u.created_at).toLocaleDateString():"—"}</td>
        </tr>))}</tbody>
      </table>
    </div>
  );
}

/* -------------------------- Item form ----------------------------- */
function ItemForm({ initial, onSave, onClose }){
  const [f,setF]=useState({ name:"", category:"Reagent", lotNumber:"", catalogNumber:"", supplier:"", quantity:1, unit:"vials", reorderThreshold:1, location:"", expirationDate:"", notes:"", ...initial });
  const [busy,setBusy]=useState(false);
  const set=k=>e=>setF({...f,[k]:e.target.value});
  const valid=f.name.trim().length>0;
  async function submit(){ if(!valid) return; setBusy(true); await onSave({...f, quantity:+f.quantity, reorderThreshold:+f.reorderThreshold}); setBusy(false); }
  return (
    <Modal title={initial.id?"Edit item":"Add inventory item"} onClose={onClose} wide>
      <div style={styles.formGrid}>
        <Field label="Item name" span={2}><input className="bs-input" value={f.name} onChange={set("name")} placeholder="e.g. Anti-mouse TNF-α (clone XT3.11)"/></Field>
        <Field label="Category"><select className="bs-input" value={f.category} onChange={set("category")}>{CATEGORIES.map(c=><option key={c}>{c}</option>)}</select></Field>
        <Field label="Supplier"><input className="bs-input" value={f.supplier} onChange={set("supplier")} placeholder="Bio X Cell"/></Field>
        <Field label="Lot number"><input className="bs-input" value={f.lotNumber} onChange={set("lotNumber")} placeholder="BP0058-7745"/></Field>
        <Field label="Catalog number"><input className="bs-input" value={f.catalogNumber} onChange={set("catalogNumber")} placeholder="BP0058"/></Field>
        <Field label="Quantity on hand"><input className="bs-input" type="number" step="any" value={f.quantity} onChange={set("quantity")}/></Field>
        <Field label="Unit"><input className="bs-input" value={f.unit} onChange={set("unit")} placeholder="vials, g, mL, boxes…"/></Field>
        <Field label="Reorder at"><input className="bs-input" type="number" step="any" value={f.reorderThreshold} onChange={set("reorderThreshold")}/></Field>
        <Field label="Expiration date"><input className="bs-input" type="date" value={f.expirationDate} onChange={set("expirationDate")}/></Field>
        <Field label="Storage location" span={2}><input className="bs-input" value={f.location} onChange={set("location")} placeholder="−20 °C, Freezer B, Shelf 2"/></Field>
        <Field label="Notes" span={2}><textarea className="bs-input" rows={2} value={f.notes} onChange={set("notes")} placeholder="Handling, dosing, hazards…"/></Field>
      </div>
      <div style={styles.modalActions}>
        <button className="bs-btn-ghost" onClick={onClose}>Cancel</button>
        <button className="bs-btn-primary" onClick={submit} disabled={!valid||busy} style={{opacity:valid?1:.5}}>{busy?"…":initial.id?"Save changes":"Add item"}</button>
      </div>
    </Modal>
  );
}

/* ------------------------- Consume modal -------------------------- */
function ConsumeModal({ item, user, onConsume, onClose }){
  const [amt,setAmt]=useState(1); const [busy,setBusy]=useState(false);
  const remaining=+(item.quantity-(+amt||0)).toFixed(3);
  async function go(){ setBusy(true); await onConsume(item,+amt); setBusy(false); }
  return (
    <Modal title="Record usage" onClose={onClose}>
      <div style={{marginBottom:14}}><div style={styles.itemName}>{item.name}</div><div style={styles.itemMeta}>Lot {item.lotNumber||"—"} · {item.quantity} {item.unit} on hand</div></div>
      <div style={styles.formGrid}>
        <Field label={`Amount used (${item.unit})`}><input className="bs-input" type="number" step="any" min="0" max={item.quantity} value={amt} onChange={e=>setAmt(e.target.value)} autoFocus/></Field>
        <Field label="Logged as"><input className="bs-input" value={user.displayName} disabled style={{background:T.paper, color:T.inkSoft}}/></Field>
      </div>
      <div style={{...styles.itemMeta, marginTop:10}}>After this: <strong style={{color:remaining<=item.reorderThreshold?T.amber:T.ink}}>{remaining} {item.unit}</strong> remaining{remaining<=item.reorderThreshold&&remaining>=0?" — will flag for reorder":""}</div>
      <div style={styles.modalActions}>
        <button className="bs-btn-ghost" onClick={onClose}>Cancel</button>
        <button className="bs-btn-primary" onClick={go} disabled={!amt||+amt<=0||busy}>{busy?"…":"Deduct & log"}</button>
      </div>
    </Modal>
  );
}

/* -------------------------- Label modal --------------------------- */
function LabelModal({ item, onClose }){
  const code=(item.lotNumber||item.id).toUpperCase();
  return (
    <Modal title="Barcode label" onClose={onClose}>
      <div className="bs-label-print" style={styles.label}>
        <div style={styles.labelName}>{item.name}</div>
        <div style={styles.labelGrid}>
          <span style={styles.labelKey}>Lot</span><span style={styles.labelVal}>{item.lotNumber||"—"}</span>
          <span style={styles.labelKey}>Cat #</span><span style={styles.labelVal}>{item.catalogNumber||"—"}</span>
          <span style={styles.labelKey}>Exp</span><span style={styles.labelVal}>{fmtDate(item.expirationDate)}</span>
          <span style={styles.labelKey}>Loc</span><span style={styles.labelVal}>{item.location||"—"}</span>
        </div>
        <div style={styles.barcodeWrap}><Barcode value={code}/><div style={styles.barcodeText}>{code}</div></div>
      </div>
      <div style={styles.modalActions}>
        <button className="bs-btn-ghost" onClick={onClose}>Close</button>
        <button className="bs-btn-primary" onClick={()=>window.print()}>Print label</button>
      </div>
      <div style={{...styles.itemMeta, marginTop:8, textAlign:"center"}}>Real Code 39 symbology — scannable with any barcode reader.</div>
    </Modal>
  );
}

/* ------------------------- Histology ------------------------------ */
function Histology({ slides, slideLog, onAdd, onEdit, onDelete, onStain, onImport }){
  const [species,setSpecies]=useState("All");
  const [q,setQ]=useState("");
  const [importOpen,setImportOpen]=useState(false);
  const list = slides.filter(s=>{
    if(species!=="All" && s.species!==species) return false;
    const t=q.trim().toLowerCase(); if(!t) return true;
    return [s.subjectId,s.region,s.staining,s.bigBox,s.box,s.project].filter(Boolean).join(" ").toLowerCase().includes(t);
  });
  const isUnstained = (s)=> !s.staining || s.staining.toLowerCase()==="unstained";

  function downloadTemplate(){
    const rows = [
      ["Subject ID","Project","Species","Region","Staining","Big box","Box","Notes"],
      ["NEC-001","NEC-BBB","Mouse","Cerebellum","Unstained","Big box 1","Box A / page 1","example row — delete me"],
      ["H-2207","Cerebellar dysmaturation","Human","Pons","H&E","Big box 2","Box C / page 4",""],
    ];
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Slides");
    XLSX.writeFile(wb, "histology-slides-template.xlsx");
  }

  return (
    <div>
      <div style={{display:"flex", gap:10, flexWrap:"wrap", alignItems:"center", marginBottom:14}}>
        <div style={styles.filterBar}>
          {["All","Mouse","Human"].map(c=>(
            <button key={c} onClick={()=>setSpecies(c)} className="bs-chip"
              style={{...styles.chip, background:species===c?T.ink:"transparent", color:species===c?"#fff":T.inkSoft, borderColor:species===c?T.ink:T.line}}>{c}</button>
          ))}
        </div>
        <input className="bs-input" placeholder="Search subject, region, staining, box…" value={q} onChange={e=>setQ(e.target.value)} style={{flex:1, minWidth:160, maxWidth:320}}/>
        <div style={{flex:1}}/>
        <button className="bs-btn-ghost" onClick={downloadTemplate} style={{padding:"8px 12px", fontSize:12.5}}>Download template</button>
        <button className="bs-btn-ghost" onClick={()=>setImportOpen(true)} style={{padding:"8px 12px", fontSize:12.5}}>Import Excel</button>
        <button className="bs-btn-primary" onClick={onAdd}>+ Add slide</button>
      </div>

      {list.length===0 ? <div style={{...styles.panel,...styles.empty}}>No slides yet. Click “Add slide” to log one.</div> : (
        <div style={styles.tableWrap}>
          <table style={styles.table}>
            <thead><tr>
              <th style={styles.th}>Subject ID</th><th style={styles.th}>Project</th><th style={styles.th}>Species</th><th style={styles.th}>Region</th>
              <th style={styles.th}>Staining</th><th style={styles.th}>Location</th>
              <th style={{...styles.th,textAlign:"right"}}>Actions</th>
            </tr></thead>
            <tbody>
              {list.map(s=>(
                <tr key={s.id} className="bs-row">
                  <td style={{...styles.td, fontFamily:"var(--mono)", fontWeight:600, fontSize:12.5}}>{s.subjectId||"—"}</td>
                  <td style={{...styles.td, fontSize:12.5}}>{s.project||"—"}</td>
                  <td style={styles.td}><span style={{...styles.pill, color:s.species==="Human"?T.teal:T.inkSoft, background:s.species==="Human"?T.greenBg:T.paper}}>{s.species}</span></td>
                  <td style={styles.td}>{s.region||"—"}</td>
                  <td style={styles.td}><span style={{...styles.pill, color:isUnstained(s)?T.amber:T.green, background:isUnstained(s)?T.amberBg:T.greenBg}}>{s.staining||"Unstained"}</span></td>
                  <td style={{...styles.td, fontSize:12, color:T.inkSoft}}>{[s.bigBox,s.box].filter(Boolean).join(" · ")||"—"}</td>
                  <td style={{...styles.td, textAlign:"right", whiteSpace:"nowrap"}}>
                    <button className="bs-act" onClick={()=>onStain(s)}>Stain</button>
                    <button className="bs-act" onClick={()=>onEdit(s)}>Edit</button>
                    <button className="bs-act bs-act-danger" onClick={()=>onDelete(s.id)}>✕</button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {slideLog.length>0 && (
        <div style={{...styles.panel, marginTop:16}}>
          <div style={styles.panelHead}>Staining history</div>
          {slideLog.slice(0,12).map(e=>(
            <div key={e.id} style={{display:"flex", alignItems:"center", gap:10, padding:"8px 0", borderTop:`1px solid ${T.line}`}}>
              <div style={{flex:1, minWidth:0}}>
                <div style={{fontSize:13, fontWeight:600}}>{e.subjectId||"—"} · {e.region||"—"}</div>
                <div style={{fontSize:12, color:T.inkSoft}}>{e.fromStain} → <strong style={{color:T.ink}}>{e.toStain}</strong> · {e.who}</div>
              </div>
              <span style={{fontSize:11.5, color:T.inkSoft, fontFamily:"var(--mono)"}}>{new Date(e.ts).toLocaleDateString()}</span>
            </div>
          ))}
        </div>
      )}

      {importOpen && <ImportModal onImport={onImport} onClose={()=>setImportOpen(false)}/>}
    </div>
  );
}

function ImportModal({ onImport, onClose }){
  const [rows,setRows]=useState(null);   // {valid:[], skipped:int, flagged:int, total:int}
  const [fileName,setFileName]=useState("");
  const [busy,setBusy]=useState(false);
  const [err,setErr]=useState("");
  const fileRef = useRef(null);

  function normKey(k){ return String(k).toLowerCase().replace(/[^a-z]/g,""); }

  function parseRow(raw){
    const m = {};
    Object.keys(raw).forEach(k=>{ m[normKey(k)] = typeof raw[k]==="string" ? raw[k].trim() : raw[k]; });
    const subjectId = String(m.subjectid ?? "").trim();
    if(!subjectId) return null;
    let species = String(m.species ?? "").trim();
    let flagged = false;
    if(/human/i.test(species)) species = "Human";
    else if(/mouse/i.test(species)) species = "Mouse";
    else { species = "Mouse"; flagged = true; }
    let staining = String(m.staining ?? "").trim(); if(!staining) staining = "Unstained";
    return { row:{ subjectId, species, region:String(m.region??"").trim(), staining, bigBox:String(m.bigbox??"").trim(), box:String(m.box??"").trim(), project:String(m.project??"").trim(), notes:String(m.notes??"").trim() }, flagged };
  }

  async function handleFile(e){
    setErr(""); setRows(null);
    const file = e.target.files?.[0]; if(!file) return;
    setFileName(file.name);
    try{
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type:"array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(ws, { defval:"" });
      const valid=[]; let skipped=0, flagged=0;
      json.forEach(r=>{ const p=parseRow(r); if(!p){ skipped++; return; } valid.push(p.row); if(p.flagged) flagged++; });
      if(valid.length===0 && skipped===0) setErr("No rows found. Make sure the first row has the column headers from the template.");
      setRows({ valid, skipped, flagged, total:json.length });
    }catch(e2){ setErr("Couldn't read that file. Make sure it's an .xlsx file."); }
  }

  async function confirm(){
    if(!rows?.valid.length) return;
    setBusy(true);
    try{ await onImport(rows.valid); onClose(); }
    catch(e3){ setErr("Import failed while saving. Please try again."); setBusy(false); }
  }

  return (
    <Modal title="Import slides from Excel" onClose={onClose}>
      <div style={{...styles.itemMeta, marginBottom:12, lineHeight:1.6}}>
        Pick an .xlsx file with these column headers in the first row: <strong>Subject ID, Project, Species, Region, Staining, Big box, Box, Notes</strong>. Each row becomes one slide. Use “Download template” if you need the format.
      </div>
      <input ref={fileRef} type="file" accept=".xlsx,.xls" onChange={handleFile} style={{display:"none"}}/>
      <button className="bs-btn-ghost" onClick={()=>fileRef.current?.click()} style={{width:"100%"}}>
        {fileName ? `Selected: ${fileName}` : "Choose Excel file…"}
      </button>

      {err && <div style={{...styles.authErr, marginTop:12}}>{err}</div>}

      {rows && (
        <div style={{marginTop:14}}>
          <div style={{display:"flex", gap:8, flexWrap:"wrap"}}>
            <span style={{...styles.pill, color:T.green, background:T.greenBg}}>{rows.valid.length} ready to import</span>
            {rows.skipped>0 && <span style={{...styles.pill, color:T.red, background:T.redBg}}>{rows.skipped} skipped (no Subject ID)</span>}
            {rows.flagged>0 && <span style={{...styles.pill, color:T.amber, background:T.amberBg}}>{rows.flagged} defaulted to Mouse</span>}
          </div>
          {rows.valid.length>0 && (
            <div style={{...styles.tableWrap, marginTop:12, maxHeight:220, overflowY:"auto"}}>
              <table style={styles.table}>
                <thead><tr><th style={styles.th}>Subject</th><th style={styles.th}>Species</th><th style={styles.th}>Region</th><th style={styles.th}>Staining</th><th style={styles.th}>Box</th></tr></thead>
                <tbody>
                  {rows.valid.slice(0,50).map((r,i)=>(
                    <tr key={i} className="bs-row">
                      <td style={{...styles.td, fontFamily:"var(--mono)", fontSize:12}}>{r.subjectId}</td>
                      <td style={styles.td}>{r.species}</td>
                      <td style={styles.td}>{r.region||"—"}</td>
                      <td style={styles.td}>{r.staining}</td>
                      <td style={{...styles.td, fontSize:12, color:T.inkSoft}}>{[r.bigBox,r.box].filter(Boolean).join(" · ")||"—"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {rows.valid.length>50 && <div style={{...styles.itemMeta, padding:"8px 14px"}}>…and {rows.valid.length-50} more</div>}
            </div>
          )}
        </div>
      )}

      <div style={styles.modalActions}>
        <button className="bs-btn-ghost" onClick={onClose}>Cancel</button>
        <button className="bs-btn-primary" onClick={confirm} disabled={!rows?.valid.length||busy}>
          {busy ? "Importing…" : rows?.valid.length ? `Import ${rows.valid.length} slides` : "Import"}
        </button>
      </div>
    </Modal>
  );
}

function SlideForm({ initial, onSave, onClose }){
  const [f,setF]=useState({ subjectId:"", species:"Mouse", region:"", staining:"Unstained", bigBox:"", box:"", project:"", notes:"", ...initial });
  const [busy,setBusy]=useState(false);
  const set=k=>e=>setF({...f,[k]:e.target.value});
  const valid=f.subjectId.trim().length>0;
  async function submit(){ if(!valid) return; setBusy(true); await onSave(f); setBusy(false); }
  return (
    <Modal title={initial.id?"Edit slide":"Add histology slide"} onClose={onClose} wide>
      <div style={styles.formGrid}>
        <Field label="Subject ID"><input className="bs-input" value={f.subjectId} onChange={set("subjectId")} placeholder="e.g. NEC-042" autoFocus/></Field>
        <Field label="Species"><select className="bs-input" value={f.species} onChange={set("species")}><option>Mouse</option><option>Human</option></select></Field>
        <Field label="Project" span={2}><input className="bs-input" value={f.project} onChange={set("project")} placeholder="e.g. NEC-BBB, Cerebellar dysmaturation"/></Field>
        <Field label="Brain region" span={2}><input className="bs-input" value={f.region} onChange={set("region")} placeholder="e.g. Cerebellum, pons, PFC"/></Field>
        <Field label="Staining (type freely)"><input className="bs-input" value={f.staining} onChange={set("staining")} placeholder="Unstained, H&E, IHC anti-TNFα…"/></Field>
        <Field label="Big box"><input className="bs-input" value={f.bigBox} onChange={set("bigBox")} placeholder="Big box 1"/></Field>
        <Field label="Box (slide book)"><input className="bs-input" value={f.box} onChange={set("box")} placeholder="Box A / page 3"/></Field>
        <Field label="Notes"><input className="bs-input" value={f.notes} onChange={set("notes")} placeholder="Optional"/></Field>
      </div>
      <div style={styles.modalActions}>
        <button className="bs-btn-ghost" onClick={onClose}>Cancel</button>
        <button className="bs-btn-primary" onClick={submit} disabled={!valid||busy} style={{opacity:valid?1:.5}}>{busy?"…":initial.id?"Save changes":"Add slide"}</button>
      </div>
    </Modal>
  );
}

function StainModal({ slide, user, onConfirm, onClose }){
  const [stain,setStain]=useState(""); const [busy,setBusy]=useState(false);
  async function go(){ if(!stain.trim()) return; setBusy(true); await onConfirm(slide, stain.trim()); setBusy(false); }
  return (
    <Modal title="Record staining" onClose={onClose}>
      <div style={{marginBottom:14}}>
        <div style={styles.itemName}>{slide.subjectId} · {slide.region||"—"}</div>
        <div style={styles.itemMeta}>{slide.species} · currently <strong>{slide.staining||"Unstained"}</strong></div>
      </div>
      <Field label="New staining (type freely)"><input className="bs-input" value={stain} onChange={e=>setStain(e.target.value)} placeholder="e.g. H&E, IHC anti-TNFα, Nissl" autoFocus onKeyDown={e=>e.key==="Enter"&&go()}/></Field>
      <div style={{...styles.itemMeta, marginTop:10}}>Logged as <strong>{user.displayName}</strong>. This updates the slide and records the change in staining history.</div>
      <div style={styles.modalActions}>
        <button className="bs-btn-ghost" onClick={onClose}>Cancel</button>
        <button className="bs-btn-primary" onClick={go} disabled={!stain.trim()||busy}>{busy?"…":"Record staining"}</button>
      </div>
    </Modal>
  );
}

/* ----------------------------- UI bits ---------------------------- */
function Field({ label, children, span }){
  return (<label style={{...styles.field, gridColumn:span===2?"1 / -1":"auto"}}><span style={styles.fieldLabel}>{label}</span>{children}</label>);
}
function Modal({ title, children, onClose, wide }){
  return (<div style={styles.overlay} onClick={onClose} className="bs-overlay">
    <div style={{...styles.modal, maxWidth:wide?640:460}} onClick={e=>e.stopPropagation()}>
      <div style={styles.modalHead}><span style={styles.modalTitle}>{title}</span><button className="bs-x" onClick={onClose}>✕</button></div>
      <div style={styles.modalBody}>{children}</div>
    </div>
  </div>);
}

/* ----------------------------- styles ----------------------------- */
function Style(){
  return (<style>{`
    @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@400;500;700&display=swap');
    *{ box-sizing:border-box; } body{ margin:0; }
    :root{ --display:'Space Grotesk',sans-serif; --body:'Inter',sans-serif; --mono:'JetBrains Mono',monospace; }
    .bs-input{ width:100%; box-sizing:border-box; border:1px solid ${T.line}; border-radius:7px; padding:9px 11px; font-family:var(--body); font-size:13.5px; color:${T.ink}; background:#fff; outline:none; }
    .bs-input:focus{ border-color:${T.teal}; box-shadow:0 0 0 3px rgba(15,110,106,.12); }
    textarea.bs-input{ resize:vertical; }
    .bs-btn-primary{ background:${T.teal}; color:#fff; border:none; border-radius:7px; padding:9px 15px; font-family:var(--body); font-weight:600; font-size:13.5px; cursor:pointer; transition:background .15s; }
    .bs-btn-primary:hover{ background:${T.tealDeep}; } .bs-btn-primary:disabled{ cursor:default; opacity:.7; }
    .bs-btn-ghost{ background:#fff; color:${T.ink}; border:1px solid ${T.line}; border-radius:7px; padding:9px 15px; font-family:var(--body); font-weight:500; font-size:13.5px; cursor:pointer; }
    .bs-btn-ghost:hover{ background:${T.paper}; }
    .bs-tab{ background:none; border:none; cursor:pointer; }
    .bs-chip{ cursor:pointer; transition:all .12s; } .bs-chip:hover{ border-color:${T.ink}; }
    .bs-row:hover td{ background:${T.paper}; }
    .bs-act{ background:none; border:1px solid ${T.line}; border-radius:6px; padding:4px 9px; margin-left:5px; font-family:var(--body); font-size:12px; color:${T.ink}; cursor:pointer; }
    .bs-act:hover{ background:${T.ink}; color:#fff; border-color:${T.ink}; }
    .bs-act-danger:hover{ background:${T.red}; border-color:${T.red}; }
    .bs-x{ background:none; border:none; font-size:15px; color:${T.inkSoft}; cursor:pointer; }
    .bs-link{ background:none; border:none; color:${T.teal}; font-size:12.5px; font-weight:600; cursor:pointer; font-family:var(--body); padding:0; }
    .bs-link:hover{ text-decoration:underline; }
    .bs-overlay{ animation:bsfade .15s ease; } @keyframes bsfade{ from{opacity:0} to{opacity:1} }
    @media print{ body *{ visibility:hidden; } .bs-label-print, .bs-label-print *{ visibility:visible; } .bs-label-print{ position:fixed; left:0; top:0; } }
  `}</style>);
}

const styles = {
  app:{ minHeight:"100vh", background:T.paper, fontFamily:"var(--body)", color:T.ink },
  header:{ display:"flex", alignItems:"center", gap:16, padding:"14px 20px", background:"#fff", borderBottom:`1px solid ${T.line}`, flexWrap:"wrap" },
  brand:{ display:"flex", alignItems:"center", gap:10 },
  mark:{ width:34, height:34, borderRadius:8, background:T.teal, color:"#fff", display:"grid", placeItems:"center", fontSize:18 },
  brandName:{ fontFamily:"var(--display)", fontWeight:700, fontSize:17, letterSpacing:"-0.01em", lineHeight:1 },
  brandSub:{ fontSize:11, color:T.inkSoft, letterSpacing:".04em", textTransform:"uppercase", marginTop:2 },
  search:{ flex:1, minWidth:180 },
  searchInput:{ width:"100%", boxSizing:"border-box", border:`1px solid ${T.line}`, borderRadius:8, padding:"9px 13px", fontSize:13.5, fontFamily:"var(--body)", outline:"none" },
  userChip:{ display:"flex", alignItems:"center", gap:8, paddingLeft:6 },
  avatar:{ width:28, height:28, borderRadius:20, background:T.ink, color:"#fff", display:"grid", placeItems:"center", fontSize:12, fontWeight:700 },
  tabs:{ display:"flex", alignItems:"center", gap:4, padding:"0 20px", background:"#fff", borderBottom:`1px solid ${T.line}`, overflowX:"auto" },
  tab:{ padding:"12px 4px", marginRight:18, fontSize:13.5, fontWeight:600, fontFamily:"var(--body)", display:"flex", alignItems:"center", gap:7, whiteSpace:"nowrap" },
  tabBadge:{ background:T.red, color:"#fff", borderRadius:20, fontSize:11, padding:"1px 7px", fontFamily:"var(--mono)" },
  exportBtn:{ padding:"6px 12px", fontSize:12.5 },
  main:{ maxWidth:1080, margin:"0 auto", padding:"22px 20px 60px" },

  statGrid:{ display:"grid", gridTemplateColumns:"repeat(auto-fit, minmax(180px,1fr))", gap:14, marginBottom:18 },
  statCard:{ background:"#fff", border:`1px solid ${T.line}`, borderRadius:11, padding:"16px 18px" },
  statNum:{ fontFamily:"var(--display)", fontSize:34, fontWeight:700, lineHeight:1 },
  statLabel:{ fontSize:12.5, color:T.inkSoft, marginTop:6 },

  dashCols:{ display:"grid", gridTemplateColumns:"1.3fr 1fr", gap:16 },
  panel:{ background:"#fff", border:`1px solid ${T.line}`, borderRadius:11, padding:18 },
  panelHead:{ fontFamily:"var(--display)", fontWeight:600, fontSize:14, marginBottom:12, letterSpacing:"-0.01em" },
  empty:{ color:T.inkSoft, fontSize:13.5, padding:"8px 0", lineHeight:1.6 },

  alertRow:{ display:"flex", alignItems:"center", gap:11, padding:"10px 0", borderTop:`1px solid ${T.line}`, cursor:"pointer" },
  dot:{ width:8, height:8, borderRadius:8, flexShrink:0 },
  alertName:{ fontWeight:600, fontSize:13.5, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" },
  alertSub:{ fontSize:12, color:T.inkSoft, marginTop:1 },
  pill:{ fontSize:11.5, fontWeight:600, padding:"3px 9px", borderRadius:20, whiteSpace:"nowrap" },

  catRow:{ display:"flex", alignItems:"center", gap:10, padding:"6px 0" },
  catLabel:{ fontSize:12.5, width:92, color:T.ink },
  catBarTrack:{ flex:1, height:8, background:T.paper, borderRadius:8, overflow:"hidden" },
  catBarFill:{ height:"100%", background:T.teal, borderRadius:8 },
  catNum:{ fontFamily:"var(--mono)", fontSize:12, width:22, textAlign:"right", color:T.inkSoft },

  filterBar:{ display:"flex", gap:7, flexWrap:"wrap", marginBottom:14 },
  chip:{ border:"1px solid", borderRadius:20, padding:"5px 13px", fontSize:12.5, fontWeight:500, fontFamily:"var(--body)" },

  tableWrap:{ background:"#fff", border:`1px solid ${T.line}`, borderRadius:11, overflow:"hidden", overflowX:"auto" },
  table:{ width:"100%", borderCollapse:"collapse", minWidth:760 },
  th:{ textAlign:"left", fontSize:11, textTransform:"uppercase", letterSpacing:".05em", color:T.inkSoft, fontWeight:600, padding:"11px 14px", borderBottom:`1px solid ${T.line}`, background:T.paper },
  td:{ padding:"12px 14px", borderBottom:`1px solid ${T.line}`, fontSize:13.5, verticalAlign:"middle" },
  itemName:{ fontWeight:600, fontSize:13.5 },
  itemMeta:{ fontSize:11.5, color:T.inkSoft, marginTop:2 },

  overlay:{ position:"fixed", inset:0, background:"rgba(22,32,43,.45)", display:"grid", placeItems:"center", padding:16, zIndex:50 },
  modal:{ width:"100%", background:"#fff", borderRadius:13, boxShadow:"0 20px 60px rgba(0,0,0,.25)", maxHeight:"90vh", display:"flex", flexDirection:"column" },
  modalHead:{ display:"flex", alignItems:"center", justifyContent:"space-between", padding:"16px 20px", borderBottom:`1px solid ${T.line}` },
  modalTitle:{ fontFamily:"var(--display)", fontWeight:600, fontSize:15.5 },
  modalBody:{ padding:20, overflowY:"auto" },
  modalActions:{ display:"flex", justifyContent:"flex-end", gap:9, marginTop:18 },

  formGrid:{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:13 },
  field:{ display:"flex", flexDirection:"column", gap:5 },
  fieldLabel:{ fontSize:11.5, fontWeight:600, color:T.inkSoft, textTransform:"uppercase", letterSpacing:".03em" },

  label:{ border:`1.5px solid ${T.ink}`, borderRadius:8, padding:16, background:"#fff" },
  labelName:{ fontFamily:"var(--display)", fontWeight:700, fontSize:15, marginBottom:10, lineHeight:1.25 },
  labelGrid:{ display:"grid", gridTemplateColumns:"auto 1fr", gap:"3px 12px", marginBottom:14 },
  labelKey:{ fontSize:11, textTransform:"uppercase", letterSpacing:".05em", color:T.inkSoft, fontWeight:600 },
  labelVal:{ fontFamily:"var(--mono)", fontSize:12.5 },
  barcodeWrap:{ textAlign:"center", borderTop:`1px solid ${T.line}`, paddingTop:12 },
  barcodeText:{ fontFamily:"var(--mono)", fontSize:12, letterSpacing:".18em", marginTop:5 },

  authWrap:{ minHeight:"100vh", background:T.paper, display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", padding:20, fontFamily:"var(--body)", color:T.ink },
  authCard:{ width:"100%", maxWidth:380, background:"#fff", border:`1px solid ${T.line}`, borderRadius:14, padding:26, boxShadow:"0 14px 40px rgba(22,32,43,.08)" },
  authBrand:{ display:"flex", alignItems:"center", gap:11, marginBottom:20 },
  authTabs:{ display:"flex", gap:4, background:T.paper, borderRadius:9, padding:4, marginBottom:20 },
  authTab:{ flex:1, padding:"8px", borderRadius:7, fontSize:13.5, fontWeight:600, color:T.inkSoft, fontFamily:"var(--body)" },
  authTabOn:{ background:"#fff", color:T.ink, boxShadow:"0 1px 3px rgba(0,0,0,.08)" },
  authErr:{ background:T.redBg, color:T.red, fontSize:12.5, padding:"8px 11px", borderRadius:7 },
  authMsg:{ background:T.greenBg, color:T.green, fontSize:12.5, padding:"8px 11px", borderRadius:7 },
  authNote:{ fontSize:12, color:T.inkSoft, textAlign:"center", marginTop:16 },
  authFooter:{ fontSize:11.5, color:T.inkSoft, marginTop:18, maxWidth:380, textAlign:"center", lineHeight:1.5 },
};
