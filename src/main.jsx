import React, { useState, useEffect, useMemo, useCallback, useRef } from "react";
import ReactDOM from 'react-dom/client'
import * as XLSX from "xlsx";

// ═══════════════════════════════════════════════════════════════════════
// DEFAULT DATA
// ═══════════════════════════════════════════════════════════════════════
const DEFAULT_FLOORS = [
  { id:"PT",  name:"Piano Terra" },
  { id:"PI",  name:"Piano Interrato" },
  { id:"P1",  name:"Primo Piano" },
  { id:"P2",  name:"Secondo Piano" },
  { id:"P3",  name:"Terzo Piano" },
  { id:"RTF", name:"Rooftop" },
];
const DEFAULT_ZONES = [
  { id:"LBB", name:"Lobby & Ingresso",  ambienti:["Ingresso principale","Reception","Corridoio","Bagno ospiti","Concierge"] },
  { id:"CAM", name:"Camere",            ambienti:["Camera Standard","Camera Superior","Junior Suite","Suite Deluxe","Suite Panoramica"] },
  { id:"SPA", name:"SPA & Wellness",    ambienti:["Piscina interna","Sauna","Bagno turco","Sala trattamenti","Zona relax","Spogliatoio uomo","Spogliatoio donna"] },
  { id:"RTF", name:"Rooftop",           ambienti:["Bar","Lounge","Solarium","Cucina rooftop"] },
  { id:"RST", name:"Ristorante",        ambienti:["Sala principale","Terrazza esterna","Cucina","Cantina"] },
  { id:"ACM", name:"Aree Comuni",       ambienti:["Corridoi piani","Ascensore","Scale","Piscina esterna","Giardino"] },
  { id:"BOH", name:"Back of House",     ambienti:["Cucina centrale","Lavanderia","Magazzino","Locali tecnici","Uffici"] },
];
const DEFAULT_CATS = [
  { id:"PAV", name:"Pavimenti",               description:"Tutti i rivestimenti del piano di calpestio, soglie e battiscopa", color:"#8B7355" },
  { id:"PAR", name:"Pareti & Rivestimenti",   description:"Rivestimenti verticali, carta da parati, pannellature e boiserie", color:"#7A8C6E" },
  { id:"SOF", name:"Soffitti",                description:"Controsoffitti, travi a vista, finiture e rivestimenti soffitto", color:"#6B7FA3" },
  { id:"SER", name:"Serramenti",              description:"Porte interne, finestre, vetrate e partizioni vetrate", color:"#9B6B5A" },
  { id:"ARF", name:"Arredo Fisso",            description:"Armadi, reception, banconi, librerie incassate e mobili su misura", color:"#7B6E8E" },
  { id:"ARM", name:"Arredo Mobile",           description:"Letti, sedie, divani, tavoli e complementi mobili", color:"#5B8A8B" },
  { id:"ILL", name:"Illuminazione",           description:"Lampadari, faretti, strip LED, lampade e sistemi di controllo luce", color:"#B8925A" },
  { id:"SAN", name:"Sanitari & Rubinetteria", description:"Sanitari, vasche, docce, miscelatori e accessori bagno", color:"#5A8BAA" },
  { id:"TES", name:"Tessili & Tendaggi",      description:"Tende, tessuti, tappeti, cuscini e biancheria", color:"#9B7B6B" },
  { id:"DEC", name:"Decorazioni & Arte",      description:"Opere d'arte, specchi, piante, vasi e complementi decorativi", color:"#7B8B5A" },
  { id:"ELE", name:"Elettrodomestici",        description:"Frigobar, macchine caffè, forni e piccoli elettrodomestici", color:"#6B7B9B" },
  { id:"TEC", name:"Tecnologia & AV",         description:"Domotica, TV, audio, sistemi chiavi e automazioni", color:"#5B6B8B" },
];
const STATUS = {
  da_definire:    { label:"Da definire",    dot:"#9CA3AF", bg:"#F3F4F6" },
  in_valutazione: { label:"In valutazione", dot:"#3B82F6", bg:"#EFF6FF" },
  approvato:      { label:"Approvato",      dot:"#10B981", bg:"#ECFDF5" },
  ordinato:       { label:"Ordinato",       dot:"#F59E0B", bg:"#FFFBEB" },
  consegnato:     { label:"Consegnato",     dot:"#8B5CF6", bg:"#F5F3FF" },
  installato:     { label:"Installato",     dot:"#059669", bg:"#D1FAE5" },
};
const ROLES = {
  architetto:  { label:"Architetto",  icon:"📐", sub:"Gestione completa del progetto", color:"#C4714A", needsPassword:true,  canEdit:true,  canApprove:true,  canAdd:true,  canDelete:true,  canSeePrice:true,  canSettings:true,  canOffer:false },
  committente: { label:"Committente", icon:"🏛",  sub:"Visualizzazione e approvazione", color:"#B8925A", needsPassword:true,  canEdit:false, canApprove:true,  canAdd:false, canDelete:false, canSeePrice:true,  canSettings:false, canOffer:false },
  fornitore:   { label:"Fornitore",   icon:"📦", sub:"Le tue richieste di preventivo",  color:"#7A8C6E", needsPassword:false, canEdit:false, canApprove:false, canAdd:false, canDelete:false, canSeePrice:false, canSettings:false, canOffer:true  },
};
const UNITS = ["mq","ml","pz","set","kg","lt","h","lotto"];
const CAT_COLORS = ["#8B7355","#7A8C6E","#6B7FA3","#9B6B5A","#7B6E8E","#5B8A8B","#B8925A","#5A8BAA","#9B7B6B","#7B8B5A","#6B7B9B","#5B6B8B","#C4714A","#B8925A","#A67C52","#6B8E6B"];

// ═══════════════════════════════════════════════════════════════════════
// EXCEL FIELD ALIASES
// ═══════════════════════════════════════════════════════════════════════
const FIELD_ALIASES={
  description:      ["descrizione","description","nome","articolo","prodotto","denominazione","oggetto","voce","elemento"],
  referenceSupplier:["fornitore","supplier","marca","brand","produttore","azienda","fornitore rif","fornitore di riferimento"],
  supplierCode:     ["codice fornitore","cod fornitore","codice prodotto","sku","ref fornitore","cod.fornitore","supplier code","part number"],
  webLink:          ["link","url","sito","web","scheda","link prodotto","scheda prodotto","collegamento"],
  qty:              ["quantità","qty","q.tà","quantita","qta","quantity","num","numero","n."],
  unit:             ["u.m.","um","unità","unita","unit","misura","udm","unità di misura"],
  unitPriceInstall: ["posa in opera","prezzo posa","costo posa","prezzo installazione","costo installazione","installazione","montaggio","manodopera","labor","installation","posa"],
  unitPrice:        ["prezzo fornitura","costo fornitura","prezzo unitario","costo unitario","prezzo unit","prezzo un.","costo unit","unit price","fornitura","price","prezzo","costo","p.u."],
  notes:            ["note","notes","annotazioni","commenti","specifiche","descrizione tecnica","osservazioni"],
  status:           ["stato","status","avanzamento","fase"],
  zoneId:           ["zona","zone","area","settore"],
  ambiente:         ["ambiente","locale","stanza","room","spazio"],
  floorId:          ["piano","floor","livello","level"],
  catId:            ["categoria","cat","category","tipologia","tipo"],
  code:             ["codice abaco","cod abaco","codice voce","code","id voce","riferimento"],
};
const OFFER_HEADER_RX=/^\s*(?:offerta|preventivo|quotazione|offer|quote|quotation)\b\s*[\-:–]?\s*(.*?)\s*$/i;
const detectOfferHeader=h=>{
  const m=String(h||"").match(OFFER_HEADER_RX);
  if(!m)return null;
  const supplier=(m[1]||"").trim();
  return supplier||String(h).trim();
};
const autoMatchField=h=>{
  const n=String(h||"").toLowerCase().trim().replace(/\s+/g," ");
  if(!n)return null;
  let best=null,bestLen=0;
  for(const[field,aliases]of Object.entries(FIELD_ALIASES)){
    for(const a of aliases){
      if((n===a||n.includes(a))&&a.length>bestLen){best=field;bestLen=a.length;}
    }
  }
  return best;
};
const statusFromStr=s=>{
  const v=String(s||"").toLowerCase();
  if(v.includes("approv"))return"approvato";
  if(v.includes("ordin")) return"ordinato";
  if(v.includes("conseg"))return"consegnato";
  if(v.includes("install"))return"installato";
  if(v.includes("valut")) return"in_valutazione";
  return"da_definire";
};

// ═══════════════════════════════════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════════════════════════════════
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2,7);
const fmt = n => (n!=null&&n!=="") ? `€\u00A0${Number(n).toLocaleString("it-IT")}` : "—";
const getTotal = i => (i.qty||0)*((i.unitPrice||0)+(i.unitPriceInstall||0));
const normDesc = s => String(s||"").toLowerCase().trim().replace(/\s+/g," ");
const genCode = (zoneId,catId,items) => {
  const pfx=`TAO-${zoneId}-${catId}`;
  const n=items.filter(i=>i.code?.startsWith(pfx)).length;
  return `${pfx}-${String(n+1).padStart(3,"0")}`;
};
const fmtDate = iso => {
  if(!iso) return "";
  const d=new Date(iso), diff=Math.floor((Date.now()-d)/60000);
  if(diff<60)   return diff<1?"ora":`${diff} min fa`;
  if(diff<1440) return `${Math.floor(diff/60)}h fa`;
  return d.toLocaleDateString("it-IT",{day:"numeric",month:"short",year:"numeric"});
};

async function resizeImage(file,maxPx=500,q=0.70){
  return new Promise(res=>{
    const fr=new FileReader();
    fr.onload=e=>{
      const img=new Image();
      img.onload=()=>{
        let w=img.width,h=img.height;
        if(w>maxPx||h>maxPx){if(w>h){h=Math.round(h*maxPx/w);w=maxPx;}else{w=Math.round(w*maxPx/h);h=maxPx;}}
        const c=document.createElement("canvas");c.width=w;c.height=h;
        c.getContext("2d").drawImage(img,0,0,w,h);
        res(c.toDataURL("image/jpeg",q));
      };img.src=e.target.result;
    };fr.readAsDataURL(file);
  });
}
async function readFileB64(file){
  if(file.size>5*1024*1024) throw new Error("File troppo grande. Max 5MB.");
  return new Promise((res,rej)=>{
    const fr=new FileReader();
    fr.onload=e=>res({name:file.name,type:file.type,data:e.target.result.split(",")[1]});
    fr.onerror=rej; fr.readAsDataURL(file);
  });
}
function dlFile(name,type,b64){
  const a=document.createElement("a");
  a.href=`data:${type};base64,${b64}`;a.download=name;a.click();
}

// ═══════════════════════════════════════════════════════════════════════
// STORAGE (shared = true → tutti gli utenti vedono gli stessi dati)
// ═══════════════════════════════════════════════════════════════════════
const stGet=async(k,shared=true)=>{try{const r=await window.storage.get(k,shared);return r?r.value:null;}catch{return null;}};
const stSet=async(k,v,shared=true)=>{try{await window.storage.set(k,v,shared);}catch(e){console.warn("storage:",e);}};
const IDX_KEY="tao-v3-index";
const projKey=id=>`tao-v3-proj-${id}`;
const attKey=id=>`tao-v3-att-${id}`;

async function loadIndex(){
  const v=await stGet(IDX_KEY); try{return v?JSON.parse(v):[];}catch{return [];}
}
async function saveIndex(idx){ await stSet(IDX_KEY,JSON.stringify(idx)); }
async function loadProject(id){
  const v=await stGet(projKey(id)); try{return v?JSON.parse(v):null;}catch{return null;}
}
async function saveProject(proj){ await stSet(projKey(proj.id),JSON.stringify(proj)); }

function makeProject(name="Nuovo Progetto"){
  const SAMPLE_ITEMS=[
    {id:uid(),code:"TAO-LBB-PAV-001",floorId:"PT",  description:"Pavimento in pietra lavica lucida 60×60",          zoneId:"LBB",ambiente:"Ingresso principale",catId:"PAV",referenceSupplier:"Ceramiche Siciliane srl",supplierCode:"PL6060-N",webLink:"https://www.ceramiche.it",qty:85, unit:"mq",unitPrice:180,  status:"approvato",     notes:"Colore nero vulcano. Finitura levigata.",     image:null,offers:[],comments:[],created:new Date().toISOString().slice(0,10)},
    {id:uid(),code:"TAO-LBB-ILL-001",floorId:"PT",  description:"Lampadario sospensione ottone spazzolato Ø120",    zoneId:"LBB",ambiente:"Reception",          catId:"ILL",referenceSupplier:"Apparatus Studio",      supplierCode:"APP-ORB-120", webLink:"https://www.apparatusstudio.com",qty:1,  unit:"pz",unitPrice:4800, status:"in_valutazione",notes:"Verificare altezza soffitto prima dell'ordine.", image:null,offers:[],comments:[],created:new Date().toISOString().slice(0,10)},
    {id:uid(),code:"TAO-CAM-PAV-001",floorId:"P1",  description:"Parquet rovere europeo spazzolato 15×120",         zoneId:"CAM",ambiente:"Camera Standard",    catId:"PAV",referenceSupplier:"Listone Giordano",      supplierCode:"LG-ROV-15",  webLink:"https://www.listonegiordano.com",qty:320,unit:"mq",unitPrice:95,   status:"ordinato",      notes:"Finitura olio naturale. Posa a correre.",     image:null,offers:[{id:uid(),supplierName:"Listone Giordano",unitPrice:95,totalNote:"Trasporto incluso. Consegna 3 sett.",attachmentName:null,attachmentKey:null,attachmentType:null,submittedAt:new Date().toISOString(),isSelected:true}],comments:[],created:new Date().toISOString().slice(0,10)},
    {id:uid(),code:"TAO-SPA-ILL-001",floorId:"PI",  description:"Strip LED RGB dimmerabile IP67",                   zoneId:"SPA",ambiente:"Piscina interna",    catId:"ILL",referenceSupplier:"Artemide",             supplierCode:"ART-RGB-67", webLink:"https://www.artemide.com",qty:85, unit:"ml",unitPrice:65,   status:"consegnato",    notes:"IP67 bordo vasca. Controller DALI.",          image:null,offers:[],comments:[],created:new Date().toISOString().slice(0,10)},
    {id:uid(),code:"TAO-RTF-ARM-001",floorId:"RTF", description:"Poltroncina outdoor teak e corda nautica",         zoneId:"RTF",ambiente:"Lounge",             catId:"ARM",referenceSupplier:"Paola Lenti",           supplierCode:"PL-ROPE-TK", webLink:"https://www.paolalenti.it",qty:24, unit:"pz",unitPrice:1200, status:"da_definire",   notes:"Colore sabbia. Verificare UV.",               image:null,offers:[],comments:[],created:new Date().toISOString().slice(0,10)},
  ];
  return {
    id:uid(), name, archPassword:"arch2024", clientPassword:"client2024",
    floors:[...DEFAULT_FLOORS], zones:[...DEFAULT_ZONES.map(z=>({...z,ambienti:[...z.ambienti]}))],
    categories:[...DEFAULT_CATS.map(c=>({...c}))],
    items:SAMPLE_ITEMS, created:new Date().toISOString().slice(0,10),
  };
}

// ═══════════════════════════════════════════════════════════════════════
// UI ATOMS
// ═══════════════════════════════════════════════════════════════════════
const C={bg:"#F6F1EA",sidebar:"#1C1916",card:"#FFF",text:"#2A2520",muted:"#7A7068",accent:"#C4714A",gold:"#B8925A",border:"#E8E0D4"};

function StatusBadge({status}){
  const cfg=STATUS[status]||{label:status,dot:"#9CA3AF",bg:"#F3F4F6"};
  return <span style={{display:"inline-flex",alignItems:"center",gap:5,padding:"3px 10px",borderRadius:99,background:cfg.bg,fontSize:11,fontWeight:500,color:cfg.dot,whiteSpace:"nowrap"}}>
    <span style={{width:6,height:6,borderRadius:"50%",background:cfg.dot,flexShrink:0}}/>{cfg.label}
  </span>;
}
function CatBadge({catId,cats}){
  const c=cats?.find(x=>x.id===catId)||(DEFAULT_CATS.find(x=>x.id===catId));
  if(!c) return null;
  return <span title={c.description} style={{display:"inline-block",padding:"2px 8px",borderRadius:4,background:c.color+"18",color:c.color,fontSize:11,fontWeight:600,fontFamily:"'DM Mono',monospace",cursor:"help"}}>{c.id}</span>;
}
function Btn({children,onClick,variant="primary",size="md",disabled,style,title}){
  const sz=size==="sm"?{padding:"5px 12px",fontSize:12}:size==="xs"?{padding:"3px 8px",fontSize:11}:{padding:"9px 18px",fontSize:13};
  const vars={primary:{background:C.accent,color:"#FFF"},secondary:{background:"#F0EBE3",color:C.text},
    ghost:{background:"transparent",color:C.muted},danger:{background:"#FEE2E2",color:"#DC2626"},
    green:{background:"#D1FAE5",color:"#059669"},blue:{background:"#DBEAFE",color:"#1D4ED8"},
    outline:{background:"transparent",color:C.accent,border:`1.5px solid ${C.accent}`},
    dark:{background:C.sidebar,color:"#F6F1EA"}};
  return <button title={title} onClick={onClick} disabled={disabled}
    style={{display:"inline-flex",alignItems:"center",gap:6,borderRadius:8,fontFamily:"inherit",fontWeight:500,
      cursor:disabled?"not-allowed":"pointer",border:"none",transition:"all .15s",opacity:disabled?0.5:1,lineHeight:1,
      ...sz,...(vars[variant]||vars.primary),...(style||{})}}>{children}</button>;
}
function Inp({value,onChange,placeholder,style,type="text",readOnly,onKeyDown}){
  return <input type={type} readOnly={readOnly} value={value} onChange={e=>onChange(e.target.value)} placeholder={placeholder} onKeyDown={onKeyDown}
    style={{border:`1.5px solid ${C.border}`,borderRadius:8,padding:"8px 12px",background:readOnly?"#F8F4EF":"#FFF",
      color:C.text,fontSize:13,fontFamily:"inherit",outline:"none",width:"100%",boxSizing:"border-box",...(style||{})}}/>;
}
function Sel({value,onChange,children,style}){
  return <select value={value} onChange={e=>onChange(e.target.value)}
    style={{border:`1.5px solid ${C.border}`,borderRadius:8,padding:"8px 12px",background:"#FFF",color:C.text,
      fontSize:13,fontFamily:"inherit",cursor:"pointer",outline:"none",width:"100%",...(style||{})}}>{children}</select>;
}
function Card({children,style}){
  return <div style={{background:C.card,borderRadius:14,boxShadow:"0 2px 12px rgba(42,37,32,0.06)",...(style||{})}}>{children}</div>;
}
function Lbl({children,badge}){
  return <div style={{fontSize:11,fontWeight:600,color:C.muted,textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:5,display:"flex",alignItems:"center",gap:6}}>
    {children}{badge>0&&<span style={{background:C.accent,color:"#FFF",borderRadius:99,padding:"1px 6px",fontSize:10,fontWeight:700}}>{badge}</span>}
  </div>;
}
function Modal({children,onClose,width=500}){
  return <div style={{position:"fixed",inset:0,zIndex:2000,display:"flex",alignItems:"center",justifyContent:"center",background:"rgba(28,25,22,0.55)",backdropFilter:"blur(3px)"}}>
    <div onClick={e=>e.stopPropagation()} style={{background:C.bg,borderRadius:18,width,maxWidth:"95vw",maxHeight:"90vh",overflow:"auto",boxShadow:"0 24px 64px rgba(0,0,0,0.25)"}}>
      {children}
    </div>
  </div>;
}

// ═══════════════════════════════════════════════════════════════════════
// PROJECT SELECTOR
// ═══════════════════════════════════════════════════════════════════════
function ProjectSelector({onSelect}){
  const [projects,setProjects]=useState(null);
  const [creating,setCreating]=useState(false);
  const [newName,setNewName]=useState("");
  const [loading,setLoading]=useState(false);
  const [hov,setHov]=useState(null);

  useEffect(()=>{ loadIndex().then(setProjects); },[]);

  const create=async()=>{
    if(!newName.trim()) return;
    setLoading(true);
    const proj=makeProject(newName.trim());
    const idx=[...(projects||[]),{id:proj.id,name:proj.name,created:proj.created}];
    await saveProject(proj); await saveIndex(idx);
    setProjects(idx); setCreating(false); setNewName(""); setLoading(false);
    onSelect(proj);
  };

  const open=async(entry)=>{
    setLoading(true);
    const proj=await loadProject(entry.id);
    setLoading(false);
    if(proj) onSelect(proj); else alert("Progetto non trovato. Potrebbe essere stato eliminato.");
  };

  if(projects===null) return <div style={{minHeight:"100vh",background:C.bg,display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"'DM Sans',sans-serif",fontSize:14,color:C.muted}}>Caricamento...</div>;

  return (
    <div style={{minHeight:"100vh",background:C.bg,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",fontFamily:"'DM Sans',sans-serif",backgroundImage:"radial-gradient(circle at 20% 50%,rgba(196,113,74,0.06),transparent 50%),radial-gradient(circle at 80% 20%,rgba(184,146,90,0.06),transparent 50%)"}}>
      <link href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:ital,wght@0,400;0,600;1,400&family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet"/>
      <div style={{textAlign:"center",marginBottom:40}}>
        <div style={{fontFamily:"'DM Mono',monospace",fontSize:11,letterSpacing:"0.25em",textTransform:"uppercase",color:C.gold,marginBottom:8}}>Abaco Forniture · v3</div>
        <h1 style={{fontFamily:"'Cormorant Garamond',serif",fontSize:48,fontWeight:600,color:C.text,margin:0,lineHeight:1}}>I tuoi Progetti</h1>
        <div style={{width:48,height:2,background:`linear-gradient(90deg,${C.accent},${C.gold})`,margin:"16px auto 0",borderRadius:1}}/>
      </div>

      <div style={{width:520,padding:"0 24px"}}>
        {/* Project list */}
        {projects.length>0&&<div style={{display:"flex",flexDirection:"column",gap:10,marginBottom:20}}>
          {projects.map(p=>(
            <div key={p.id} onClick={()=>!loading&&open(p)}
              onMouseEnter={()=>setHov(p.id)} onMouseLeave={()=>setHov(null)}
              style={{padding:"18px 22px",background:hov===p.id?"#FFF":"rgba(255,255,255,0.75)",borderRadius:14,cursor:loading?"wait":"pointer",
                border:hov===p.id?`2px solid ${C.accent}`:"2px solid transparent",transition:"all .2s",
                boxShadow:hov===p.id?`0 8px 24px ${C.accent}18`:"0 2px 8px rgba(42,37,32,0.06)",
                display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <div>
                <div style={{fontFamily:"'Cormorant Garamond',serif",fontSize:20,fontWeight:600,color:C.text}}>{p.name}</div>
                <div style={{fontSize:12,color:C.muted,marginTop:2}}>Creato il {p.created}</div>
              </div>
              <div style={{color:C.accent,fontSize:18,opacity:hov===p.id?1:0,transition:"opacity .2s"}}>→</div>
            </div>
          ))}
        </div>}

        {projects.length===0&&!creating&&(
          <div style={{textAlign:"center",padding:"32px",background:"rgba(255,255,255,0.6)",borderRadius:14,marginBottom:20,color:C.muted}}>
            <div style={{fontSize:32,marginBottom:8}}>📁</div>
            <div style={{fontSize:14}}>Nessun progetto ancora.<br/>Creane uno per iniziare.</div>
          </div>
        )}

        {/* Create form */}
        {creating?(
          <div style={{background:"#FFF",borderRadius:14,padding:24,boxShadow:`0 4px 20px rgba(196,113,74,0.15)`,border:`2px solid ${C.accent}`}}>
            <div style={{fontFamily:"'Cormorant Garamond',serif",fontSize:18,fontWeight:600,color:C.text,marginBottom:14}}>Nuovo Progetto</div>
            <Inp value={newName} onChange={setNewName} placeholder="Es. Hotel Taormina, Villa Messina, Appartamento Milano..."
              onKeyDown={e=>{if(e.key==="Enter") create();}}/>
            <div style={{display:"flex",gap:10,marginTop:12}}>
              <Btn variant="secondary" onClick={()=>{setCreating(false);setNewName("");}}>Annulla</Btn>
              <Btn onClick={create} disabled={!newName.trim()||loading} style={{flex:1}}>
                {loading?"Creazione...":"Crea Progetto"}
              </Btn>
            </div>
          </div>
        ):(
          <Btn onClick={()=>setCreating(true)} style={{width:"100%",justifyContent:"center",padding:"13px",borderRadius:12,fontSize:14}}>
            + Nuovo Progetto
          </Btn>
        )}
      </div>
      <div style={{marginTop:40,fontSize:11,color:"#B0A89A",letterSpacing:"0.05em"}}>
        I dati sono condivisi tra tutti gli utenti che accedono a questo link
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════
// AUTH SCREEN
// ═══════════════════════════════════════════════════════════════════════
function AuthScreen({project,onAuth,onBack}){
  const [step,setStep]=useState("role"); // "role" | "password" | "supplier"
  const [selectedRole,setSelectedRole]=useState(null);
  const [password,setPassword]=useState("");
  const [supplierName,setSupplierName]=useState("");
  const [error,setError]=useState("");
  const [hov,setHov]=useState(null);

  const tryPassword=()=>{
    const correct = selectedRole==="architetto" ? project.archPassword : project.clientPassword;
    if(password===correct){ onAuth(selectedRole,""); }
    else{ setError("Password errata. Riprova."); setPassword(""); }
  };

  if(step==="password") return (
    <div style={{minHeight:"100vh",background:C.bg,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",fontFamily:"'DM Sans',sans-serif"}}>
      <link href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:ital,wght@0,400;0,600;1,400&family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet"/>
      <div style={{width:380,padding:"0 24px"}}>
        <div style={{textAlign:"center",marginBottom:32}}>
          <div style={{fontSize:36,marginBottom:8}}>{ROLES[selectedRole]?.icon}</div>
          <h2 style={{fontFamily:"'Cormorant Garamond',serif",fontSize:32,fontWeight:600,color:C.text,margin:0}}>{ROLES[selectedRole]?.label}</h2>
          <div style={{fontSize:13,color:C.muted,marginTop:6}}>{project.name}</div>
          <div style={{width:40,height:2,background:C.accent,margin:"12px auto 0",borderRadius:1}}/>
        </div>
        <Lbl>Password di accesso</Lbl>
        <Inp value={password} onChange={v=>{setPassword(v);setError("");}} type="password" placeholder="Inserisci la password..."
          onKeyDown={e=>{if(e.key==="Enter")tryPassword();}}/>
        {error&&<div style={{color:"#DC2626",fontSize:12,marginTop:6}}>{error}</div>}
        <div style={{display:"flex",gap:10,marginTop:12}}>
          <Btn variant="secondary" onClick={()=>{setStep("role");setPassword("");setError("");}}>← Indietro</Btn>
          <Btn onClick={tryPassword} disabled={!password} style={{flex:1}}>Accedi</Btn>
        </div>
        {selectedRole==="architetto"&&<div style={{marginTop:8,fontSize:11,color:C.muted,textAlign:"center"}}>Password default: <strong>arch2024</strong></div>}
        {selectedRole==="committente"&&<div style={{marginTop:8,fontSize:11,color:C.muted,textAlign:"center"}}>Password default: <strong>client2024</strong></div>}
      </div>
    </div>
  );

  if(step==="supplier") return (
    <div style={{minHeight:"100vh",background:C.bg,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",fontFamily:"'DM Sans',sans-serif"}}>
      <link href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:ital,wght@0,400;0,600;1,400&family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet"/>
      <div style={{width:380,padding:"0 24px"}}>
        <div style={{textAlign:"center",marginBottom:32}}>
          <div style={{fontSize:36,marginBottom:8}}>📦</div>
          <h2 style={{fontFamily:"'Cormorant Garamond',serif",fontSize:32,fontWeight:600,color:C.text,margin:0}}>Accesso Fornitore</h2>
          <div style={{fontSize:13,color:C.muted,marginTop:6}}>{project.name}</div>
          <div style={{width:40,height:2,background:"#7A8C6E",margin:"12px auto 0",borderRadius:1}}/>
        </div>
        <Lbl>Ragione sociale / Nome azienda</Lbl>
        <Inp value={supplierName} onChange={setSupplierName} placeholder="Es. Ceramiche Siciliane srl"
          onKeyDown={e=>{if(e.key==="Enter"&&supplierName.trim())onAuth("fornitore",supplierName.trim());}}/>
        <div style={{fontSize:12,color:C.muted,marginTop:6}}>Vedrai solo le richieste di preventivo a te assegnate</div>
        <div style={{display:"flex",gap:10,marginTop:12}}>
          <Btn variant="secondary" onClick={()=>setStep("role")}>← Indietro</Btn>
          <Btn onClick={()=>onAuth("fornitore",supplierName.trim())} disabled={!supplierName.trim()} style={{flex:1}}>Accedi</Btn>
        </div>
      </div>
    </div>
  );

  return (
    <div style={{minHeight:"100vh",background:C.bg,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",fontFamily:"'DM Sans',sans-serif",backgroundImage:"radial-gradient(circle at 20% 50%,rgba(196,113,74,0.06),transparent 50%)"}}>
      <link href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:ital,wght@0,400;0,600;1,400&family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet"/>
      <div style={{textAlign:"center",marginBottom:40}}>
        <div style={{fontFamily:"'DM Mono',monospace",fontSize:11,letterSpacing:"0.2em",color:C.gold,marginBottom:6,textTransform:"uppercase"}}>{project.name}</div>
        <h2 style={{fontFamily:"'Cormorant Garamond',serif",fontSize:38,fontWeight:600,color:C.text,margin:0}}>Chi sei?</h2>
        <div style={{width:40,height:2,background:`linear-gradient(90deg,${C.accent},${C.gold})`,margin:"12px auto 0",borderRadius:1}}/>
      </div>
      <div style={{display:"flex",gap:18,flexWrap:"wrap",justifyContent:"center",padding:"0 24px"}}>
        {Object.entries(ROLES).map(([key,r])=>(
          <div key={key}
            onClick={()=>{setSelectedRole(key);if(r.needsPassword)setStep("password");else setStep("supplier");}}
            onMouseEnter={()=>setHov(key)} onMouseLeave={()=>setHov(null)}
            style={{width:200,padding:24,background:hov===key?"#FFF":"rgba(255,255,255,0.7)",borderRadius:14,cursor:"pointer",transition:"all .2s",textAlign:"center",
              border:hov===key?`2px solid ${r.color}`:"2px solid transparent",
              boxShadow:hov===key?`0 8px 28px ${r.color}22`:"0 2px 10px rgba(42,37,32,0.06)",
              transform:hov===key?"translateY(-4px)":"none"}}>
            <div style={{fontSize:32,marginBottom:10}}>{r.icon}</div>
            <div style={{fontFamily:"'Cormorant Garamond',serif",fontSize:20,fontWeight:600,color:hov===key?r.color:C.text,marginBottom:4}}>{r.label}</div>
            <div style={{fontSize:11,color:C.muted,lineHeight:1.5}}>{r.sub}</div>
          </div>
        ))}
      </div>
      <div style={{marginTop:32}}>
        <Btn variant="ghost" onClick={onBack} style={{fontSize:12}}>← Cambia progetto</Btn>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════
// SIDEBAR
// ═══════════════════════════════════════════════════════════════════════
function Sidebar({view,setView,role,supplierName,projectName,onSwitchRole,onSwitchProject,pendingOffers}){
  const r=ROLES[role];
  const nav=[
    {id:"dashboard",icon:"◈",label:"Dashboard"},
    {id:"abaco",    icon:"≡",label:"Abaco Forniture"},
    {id:"zone",     icon:"⬡",label:"Vista per Zone"},
    ...(role==="fornitore"?[{id:"portal",icon:"📋",label:"Il mio Portale"}]:[]),
    ...(r.canSettings?[{id:"settings",icon:"⚙",label:"Impostazioni"}]:[]),
  ];
  return (
    <div style={{width:220,background:C.sidebar,display:"flex",flexDirection:"column",flexShrink:0}}>
      <div style={{padding:"22px 18px 18px",borderBottom:"1px solid rgba(255,255,255,0.07)"}}>
        <div style={{fontFamily:"'DM Mono',monospace",fontSize:9,letterSpacing:"0.2em",color:C.gold,textTransform:"uppercase",marginBottom:3}}>Progetto</div>
        <div style={{fontFamily:"'Cormorant Garamond',serif",fontSize:16,color:"#F6F1EA",lineHeight:1.25,fontWeight:600,wordBreak:"break-word"}}>{projectName}</div>
        <div style={{marginTop:10,display:"inline-flex",alignItems:"center",gap:6,background:"rgba(196,113,74,0.15)",borderRadius:6,padding:"4px 8px"}}>
          <span style={{fontSize:11}}>{r.icon}</span>
          <div>
            <div style={{fontSize:11,color:C.accent,fontWeight:500,lineHeight:1.2}}>{r.label}</div>
            {supplierName&&<div style={{fontSize:10,color:"rgba(255,255,255,0.4)",marginTop:1,lineHeight:1.2}}>{supplierName}</div>}
          </div>
        </div>
      </div>
      <nav style={{flex:1,padding:"14px 10px"}}>
        {nav.map(n=>(
          <button key={n.id} onClick={()=>setView(n.id)}
            style={{width:"100%",display:"flex",alignItems:"center",justifyContent:"space-between",padding:"9px 10px",borderRadius:8,border:"none",cursor:"pointer",marginBottom:2,
              background:view===n.id?"rgba(196,113,74,0.15)":"transparent",
              color:view===n.id?C.accent:"rgba(255,255,255,0.5)",
              fontFamily:"'DM Sans',sans-serif",fontSize:13,fontWeight:view===n.id?600:400,transition:"all .15s",textAlign:"left"}}>
            <span style={{display:"flex",alignItems:"center",gap:9}}>
              <span style={{fontSize:15,width:17,textAlign:"center"}}>{n.icon}</span>{n.label}
            </span>
            {n.id==="abaco"&&pendingOffers>0&&<span style={{background:C.accent,color:"#FFF",borderRadius:99,padding:"1px 7px",fontSize:10,fontWeight:700}}>{pendingOffers}</span>}
          </button>
        ))}
      </nav>
      <div style={{padding:"12px 10px",borderTop:"1px solid rgba(255,255,255,0.07)",display:"flex",flexDirection:"column",gap:6}}>
        <button onClick={onSwitchRole} style={{width:"100%",padding:"8px 10px",borderRadius:8,border:"1px solid rgba(255,255,255,0.12)",background:"transparent",color:"rgba(255,255,255,0.45)",fontSize:12,fontFamily:"'DM Sans',sans-serif",cursor:"pointer",textAlign:"left"}}>
          ⇄ Cambia ruolo
        </button>
        <button onClick={onSwitchProject} style={{width:"100%",padding:"8px 10px",borderRadius:8,border:"1px solid rgba(255,255,255,0.08)",background:"transparent",color:"rgba(255,255,255,0.3)",fontSize:12,fontFamily:"'DM Sans',sans-serif",cursor:"pointer",textAlign:"left"}}>
          ← Cambia progetto
        </button>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════
// DASHBOARD
// ═══════════════════════════════════════════════════════════════════════
function Dashboard({project,role}){
  const {items,zones,categories}=project;
  const perm=ROLES[role];
  const totalVal=perm.canSeePrice?items.reduce((s,i)=>s+getTotal(i),0):null;
  const byStatus=Object.fromEntries(Object.keys(STATUS).map(k=>[k,items.filter(i=>i.status===k).length]));
  const byZone=zones.map(z=>({...z,count:items.filter(i=>i.zoneId===z.id).length,val:items.filter(i=>i.zoneId===z.id).reduce((s,i)=>s+getTotal(i),0)}));
  const maxVal=Math.max(...byZone.map(z=>z.val),1);
  const totalOffers=items.reduce((s,i)=>s+(i.offers?.length||0),0);
  const pendingOffers=items.reduce((s,i)=>s+(i.offers?.filter(o=>!o.isSelected).length||0),0);

  const kpis=[
    {label:"Voci totali",    value:items.length,           sub:"elementi",       color:C.accent},
    ...(perm.canSeePrice?[{label:"Budget totale",  value:fmt(totalVal),          sub:"valore complessivo", color:C.gold}]:[]),
    {label:"In valutazione", value:byStatus.in_valutazione||0, sub:"da approvare", color:"#3B82F6"},
    {label:"Ordinati",       value:byStatus.ordinato||0,   sub:"ordini attivi",  color:"#F59E0B"},
    {label:"Offerte",        value:totalOffers,             sub:`${pendingOffers} da valutare`, color:"#7A8C6E"},
    {label:"Installati",     value:byStatus.installato||0, sub:`su ${items.length}`, color:"#059669"},
  ];

  return (
    <div>
      <h2 style={{fontFamily:"'Cormorant Garamond',serif",fontSize:28,fontWeight:600,color:C.text,margin:"0 0 22px"}}>Panoramica Progetto</h2>
      <div style={{display:"grid",gridTemplateColumns:`repeat(${perm.canSeePrice?3:3},1fr)`,gap:14,marginBottom:24}}>
        {kpis.slice(0,6).map((k,i)=>(
          <Card key={i} style={{padding:"18px 20px",borderTop:`3px solid ${k.color}`}}>
            <div style={{fontSize:11,fontWeight:600,color:C.muted,textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:5}}>{k.label}</div>
            <div style={{fontFamily:"'Cormorant Garamond',serif",fontSize:30,fontWeight:600,color:k.color}}>{k.value}</div>
            <div style={{fontSize:11,color:"#9CA3AF",marginTop:2}}>{k.sub}</div>
          </Card>
        ))}
      </div>
      <div style={{display:"grid",gridTemplateColumns:perm.canSeePrice?"1fr 1fr":"1fr",gap:18}}>
        <Card style={{padding:"18px 22px"}}>
          <div style={{fontSize:13,fontWeight:600,color:C.text,marginBottom:14}}>Stato Forniture</div>
          {Object.entries(STATUS).map(([k,cfg])=>{
            const n=byStatus[k]||0; const pct=items.length?(n/items.length)*100:0;
            return <div key={k} style={{marginBottom:11}}>
              <div style={{display:"flex",justifyContent:"space-between",fontSize:12,marginBottom:3}}>
                <span style={{color:C.text,display:"flex",alignItems:"center",gap:6}}>
                  <span style={{width:8,height:8,borderRadius:"50%",background:cfg.dot,display:"inline-block"}}/>{cfg.label}
                </span>
                <span style={{color:C.muted,fontFamily:"'DM Mono',monospace"}}>{n}</span>
              </div>
              <div style={{height:3,background:"#F3F0EC",borderRadius:99}}><div style={{height:"100%",width:`${pct}%`,background:cfg.dot,borderRadius:99}}/></div>
            </div>;
          })}
        </Card>
        {perm.canSeePrice&&<Card style={{padding:"18px 22px"}}>
          <div style={{fontSize:13,fontWeight:600,color:C.text,marginBottom:14}}>Budget per Zona</div>
          {byZone.filter(z=>z.count>0).map(z=>(
            <div key={z.id} style={{marginBottom:11}}>
              <div style={{display:"flex",justifyContent:"space-between",fontSize:12,marginBottom:3}}>
                <span style={{color:C.text}}>{z.name}</span>
                <span style={{color:C.muted,fontFamily:"'DM Mono',monospace",fontSize:11}}>{fmt(z.val)}</span>
              </div>
              <div style={{height:3,background:"#F3F0EC",borderRadius:99}}><div style={{height:"100%",width:`${(z.val/maxVal)*100}%`,background:`linear-gradient(90deg,${C.accent},${C.gold})`,borderRadius:99}}/></div>
            </div>
          ))}
        </Card>}
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════
// ITEM MODAL — SCHEDA TAB
// ═══════════════════════════════════════════════════════════════════════
function SchedaTab({form,setForm,project,isNew,perm}){
  const {floors,zones,categories}=project;
  const set=(k,v)=>setForm(f=>({...f,[k]:v}));
  const imgRef=useRef();
  const filteredZones=zones; // show all zones regardless of floor
  const ambList=zones.find(z=>z.id===form.zoneId)?.ambienti||[];
  const readOnly=!perm.canEdit&&!isNew;

  useEffect(()=>{
    if(isNew&&form.zoneId&&form.catId) set("code",genCode(form.zoneId,form.catId,project.items));
  },[form.zoneId,form.catId,isNew]);

  const handleImg=async e=>{
    const f=e.target.files?.[0]; if(!f) return;
    try{const img=await resizeImage(f);set("image",img);}catch(err){alert(err.message);}
  };

  const catInfo=categories.find(c=>c.id===form.catId);

  return (
    <div style={{display:"flex",flexDirection:"column",gap:14}}>
      {/* Piano + Zona + Ambiente */}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10}}>
        <div><Lbl>Piano</Lbl>
          {readOnly?<div style={{fontSize:13,padding:"6px 0"}}>{floors.find(f=>f.id===form.floorId)?.name||"—"}</div>:
          <Sel value={form.floorId||""} onChange={v=>set("floorId",v)}>
            <option value="">— Piano —</option>
            {floors.map(f=><option key={f.id} value={f.id}>{f.name}</option>)}
          </Sel>}
        </div>
        <div><Lbl>Zona</Lbl>
          {readOnly?<div style={{fontSize:13,padding:"6px 0"}}>{zones.find(z=>z.id===form.zoneId)?.name||form.zoneId}</div>:
          <Sel value={form.zoneId} onChange={v=>{set("zoneId",v);set("ambiente","");}}>
            {filteredZones.map(z=><option key={z.id} value={z.id}>{z.name}</option>)}
          </Sel>}
        </div>
        <div><Lbl>Ambiente</Lbl>
          {readOnly?<div style={{fontSize:13,padding:"6px 0"}}>{form.ambiente||"—"}</div>:
          <Sel value={form.ambiente} onChange={v=>set("ambiente",v)}>
            <option value="">— Ambiente —</option>
            {ambList.map(a=><option key={a}>{a}</option>)}
          </Sel>}
        </div>
      </div>

      {/* Categoria */}
      <div><Lbl>Categoria</Lbl>
        {catInfo&&<div style={{fontSize:11,color:C.muted,marginBottom:6,fontStyle:"italic"}}>{catInfo.description}</div>}
        <div style={{display:"flex",flexWrap:"wrap",gap:5}}>
          {categories.map(c=>(
            <button key={c.id} onClick={()=>!readOnly&&set("catId",c.id)} disabled={readOnly} title={`${c.name} — ${c.description}`}
              style={{padding:"4px 9px",borderRadius:6,border:`1.5px solid ${form.catId===c.id?c.color:C.border}`,
                background:form.catId===c.id?c.color+"18":"#FFF",color:form.catId===c.id?c.color:C.muted,
                fontSize:11,fontWeight:600,cursor:readOnly?"default":"pointer",fontFamily:"'DM Mono',monospace",display:"flex",alignItems:"center",gap:4}}>
              <span style={{width:6,height:6,borderRadius:"50%",background:c.color,flexShrink:0}}/>
              {c.id}
            </button>
          ))}
        </div>
        {!readOnly&&form.catId&&<div style={{fontSize:11,color:C.muted,marginTop:5}}>
          <strong>{catInfo?.name}</strong> — {catInfo?.description}
        </div>}
      </div>

      {/* Descrizione */}
      <div><Lbl>Descrizione *</Lbl>
        {readOnly?<div style={{fontSize:13,fontWeight:500,color:C.text,padding:"6px 0"}}>{form.description}</div>:
        <Inp value={form.description} onChange={v=>set("description",v)} placeholder="Es. Pavimento in pietra lavica 60×60"/>}
      </div>

      {/* Codice */}
      <div><Lbl>Codice Abaco</Lbl>
        <Inp value={form.code||""} onChange={v=>set("code",v)} readOnly={readOnly}
          style={{fontFamily:"'DM Mono',monospace",background:"#F8F4EF",color:C.gold}}/>
      </div>

      {/* Fornitore */}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
        <div><Lbl>Fornitore di riferimento</Lbl>
          {readOnly?<div style={{fontSize:13,padding:"6px 0"}}>{form.referenceSupplier||"—"}</div>:
          <Inp value={form.referenceSupplier||""} onChange={v=>set("referenceSupplier",v)} placeholder="Es. Agape Design"/>}
        </div>
        <div><Lbl>Codice fornitore</Lbl>
          {readOnly?<div style={{fontSize:13,fontFamily:"'DM Mono',monospace",padding:"6px 0",color:C.muted}}>{form.supplierCode||"—"}</div>:
          <Inp value={form.supplierCode||""} onChange={v=>set("supplierCode",v)} placeholder="AGP-001" style={{fontFamily:"'DM Mono',monospace"}}/>}
        </div>
      </div>

      {/* Link */}
      <div><Lbl>Link scheda prodotto</Lbl>
        <div style={{display:"flex",gap:8,alignItems:"center"}}>
          {readOnly?<div style={{fontSize:13,padding:"6px 0",flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{form.webLink||"—"}</div>:
          <Inp value={form.webLink||""} onChange={v=>set("webLink",v)} placeholder="https://..." style={{flex:1}}/>}
          {form.webLink&&<a href={form.webLink} target="_blank" rel="noopener noreferrer"
            style={{flexShrink:0,padding:"8px 13px",borderRadius:8,background:"#EFF6FF",color:"#1D4ED8",fontSize:12,fontWeight:500,textDecoration:"none",whiteSpace:"nowrap"}}>↗ Apri</a>}
        </div>
      </div>

      {/* Q.tà + Prezzo (only if canSeePrice OR new item by architect) */}
      {perm.canSeePrice&&<div style={{display:"grid",gridTemplateColumns:"1fr 80px 1fr 1fr",gap:10}}>
        <div><Lbl>Quantità</Lbl>
          {readOnly?<div style={{fontSize:13,fontFamily:"'DM Mono',monospace",padding:"6px 0"}}>{form.qty} {form.unit}</div>:
          <Inp value={form.qty} onChange={v=>set("qty",v)} placeholder="0" style={{textAlign:"right"}}/>}
        </div>
        {!readOnly&&<div><Lbl>U.M.</Lbl><Sel value={form.unit} onChange={v=>set("unit",v)}>{UNITS.map(u=><option key={u}>{u}</option>)}</Sel></div>}
        <div><Lbl>Prezzo fornitura (€)</Lbl>
          {readOnly?<div style={{fontSize:13,fontFamily:"'DM Mono',monospace",padding:"6px 0"}}>{fmt(form.unitPrice)}</div>:
          <Inp value={form.unitPrice} onChange={v=>set("unitPrice",v)} placeholder="0" style={{textAlign:"right"}}/>}
        </div>
        <div><Lbl>Prezzo posa (€)</Lbl>
          {readOnly?<div style={{fontSize:13,fontFamily:"'DM Mono',monospace",padding:"6px 0"}}>{fmt(form.unitPriceInstall)}</div>:
          <Inp value={form.unitPriceInstall||""} onChange={v=>set("unitPriceInstall",v)} placeholder="0" style={{textAlign:"right"}}/>}
        </div>
      </div>}

      {/* Quantità senza prezzo per fornitore */}
      {!perm.canSeePrice&&<div style={{display:"grid",gridTemplateColumns:"1fr 80px",gap:10}}>
        <div><Lbl>Quantità richiesta</Lbl>
          <div style={{fontSize:13,fontFamily:"'DM Mono',monospace",padding:"8px 12px",background:"#F8F4EF",borderRadius:8,border:`1px solid ${C.border}`}}>{form.qty} {form.unit}</div>
        </div>
        <div/>
      </div>}

      {/* Totale */}
      {perm.canSeePrice&&!readOnly&&form.qty&&(Number(form.unitPrice)||Number(form.unitPriceInstall))?
        <div style={{background:"#FFF8F0",border:`1px solid #F0DCC8`,borderRadius:8,padding:"10px 14px",display:"flex",flexDirection:"column",gap:4}}>
          {Number(form.unitPrice)>0&&<div style={{display:"flex",justifyContent:"space-between",fontSize:12,color:C.muted}}><span>Fornitura ({form.qty} × {fmt(form.unitPrice)})</span><span style={{fontFamily:"'DM Mono',monospace"}}>{fmt(Number(form.qty)*Number(form.unitPrice))}</span></div>}
          {Number(form.unitPriceInstall)>0&&<div style={{display:"flex",justifyContent:"space-between",fontSize:12,color:C.muted}}><span>Posa ({form.qty} × {fmt(form.unitPriceInstall)})</span><span style={{fontFamily:"'DM Mono',monospace"}}>{fmt(Number(form.qty)*Number(form.unitPriceInstall))}</span></div>}
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",borderTop:`1px solid ${C.border}`,paddingTop:5,marginTop:2}}>
            <span style={{fontSize:12,color:C.muted}}>Totale stimato</span>
            <span style={{fontFamily:"'Cormorant Garamond',serif",fontSize:20,fontWeight:600,color:C.accent}}>{fmt(Number(form.qty)*((Number(form.unitPrice)||0)+(Number(form.unitPriceInstall)||0)))}</span>
          </div>
        </div>:null}

      {/* Campi personalizzati */}
      {form.customFields&&Object.keys(form.customFields).length>0&&<div>
        <Lbl>Campi personalizzati</Lbl>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
          {Object.entries(form.customFields).map(([k,v])=>(
            <div key={k} style={{display:"flex",flexDirection:"column",gap:3}}>
              <span style={{fontSize:11,color:C.muted,fontFamily:"'DM Mono',monospace"}}>★ {k}</span>
              {readOnly?<div style={{fontSize:13,padding:"4px 0"}}>{v||"—"}</div>:
              <Inp value={v||""} onChange={nv=>set("customFields",{...form.customFields,[k]:nv})}/>}
            </div>
          ))}
        </div>
      </div>}

      {/* Stato */}
      {perm.canEdit&&<div><Lbl>Stato</Lbl>
        <div style={{display:"flex",flexWrap:"wrap",gap:5}}>
          {Object.entries(STATUS).map(([k,cfg])=>(
            <button key={k} onClick={()=>set("status",k)}
              style={{padding:"4px 12px",borderRadius:99,border:`1.5px solid ${form.status===k?cfg.dot:C.border}`,
                background:form.status===k?cfg.bg:"#FFF",color:form.status===k?cfg.dot:C.muted,
                fontSize:11,fontWeight:500,cursor:"pointer",display:"flex",alignItems:"center",gap:4}}>
              <span style={{width:6,height:6,borderRadius:"50%",background:cfg.dot}}/>{cfg.label}
            </button>
          ))}
        </div>
      </div>}

      {/* Immagine */}
      <div><Lbl>Immagine di riferimento</Lbl>
        {form.image?<div style={{position:"relative",display:"inline-block",width:"100%"}}>
          <img src={form.image} alt="ref" style={{width:"100%",maxHeight:180,objectFit:"contain",borderRadius:10,border:`1px solid ${C.border}`}}/>
          {perm.canEdit&&<button onClick={()=>set("image",null)} style={{position:"absolute",top:6,right:6,background:"rgba(0,0,0,0.65)",color:"#FFF",border:"none",borderRadius:"50%",width:24,height:24,cursor:"pointer",fontSize:14,display:"flex",alignItems:"center",justifyContent:"center"}}>×</button>}
        </div>:
        perm.canEdit&&<div onClick={()=>imgRef.current?.click()}
          style={{border:`2px dashed ${C.border}`,borderRadius:10,padding:"20px",textAlign:"center",cursor:"pointer",color:C.muted,fontSize:12}}>
          <div style={{fontSize:22,marginBottom:4}}>🖼</div>Clicca per caricare un'immagine · JPG, PNG · Max 5MB
        </div>}
        <input ref={imgRef} type="file" accept="image/*" style={{display:"none"}} onChange={handleImg}/>
      </div>

      {/* Note */}
      <div><Lbl>Note tecniche</Lbl>
        {readOnly?<div style={{fontSize:13,color:C.text,lineHeight:1.5,whiteSpace:"pre-wrap"}}>{form.notes||"—"}</div>:
        <textarea value={form.notes||""} onChange={e=>set("notes",e.target.value)} rows={3} placeholder="Finiture, varianti, specifiche tecniche..."
          style={{width:"100%",border:`1.5px solid ${C.border}`,borderRadius:8,padding:"8px 12px",background:"#FFF",color:C.text,fontSize:13,fontFamily:"inherit",resize:"vertical",outline:"none",boxSizing:"border-box"}}/>}
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════
// OFFERTE TAB
// ═══════════════════════════════════════════════════════════════════════
function OfferteTab({item,perm,supplierName,onUpdate}){
  const [showForm,setShowForm]=useState(false);
  const [ofr,setOfr]=useState({supplierName:supplierName||"",unitPrice:"",totalNote:"",attName:null,attType:null,attData:null,attKey:null});
  const attRef=useRef();
  const myOffer=supplierName?item.offers?.find(o=>o.supplierName?.toLowerCase()===supplierName.toLowerCase()):null;
  const minP=item.offers?.length?Math.min(...item.offers.map(o=>o.unitPrice)):null;

  useEffect(()=>{if(myOffer){setOfr({...myOffer,attData:null,attName:myOffer.attachmentName,attType:myOffer.attachmentType,attKey:myOffer.attachmentKey});setShowForm(true);}}, [myOffer?.id]);

  const handleAtt=async e=>{
    const f=e.target.files?.[0];if(!f)return;
    try{const a=await readFileB64(f);setOfr(o=>({...o,attName:a.name,attType:a.type,attData:a.data}));}catch(err){alert(err.message);}
  };

  const submit=async()=>{
    if(!ofr.supplierName.trim()||!ofr.unitPrice)return;
    const id=myOffer?.id||uid();
    const aKey=ofr.attData?attKey(id):(myOffer?.attachmentKey||null);
    if(ofr.attData) await stSet(aKey,ofr.attData);
    const newOffer={id,supplierName:ofr.supplierName.trim(),unitPrice:Number(ofr.unitPrice),totalNote:ofr.totalNote,
      attachmentName:ofr.attData?ofr.attName:(myOffer?.attachmentName||null),
      attachmentType:ofr.attData?ofr.attType:(myOffer?.attachmentType||null),
      attachmentKey:aKey,submittedAt:new Date().toISOString(),isSelected:myOffer?.isSelected||false};
    const newOffers=myOffer?item.offers.map(o=>o.id===id?newOffer:o):[...(item.offers||[]),newOffer];
    onUpdate({...item,offers:newOffers});setShowForm(false);
  };

  const select=o=>{
    const newOffers=item.offers.map(x=>({...x,isSelected:x.id===o.id}));
    onUpdate({...item,offers:newOffers,unitPrice:o.unitPrice,referenceSupplier:o.supplierName});
  };
  const dlAtt=async o=>{
    if(!o.attachmentKey)return;
    const d=await stGet(o.attachmentKey);
    if(d)dlFile(o.attachmentName,o.attachmentType,d); else alert("Allegato non trovato.");
  };

  return (
    <div style={{display:"flex",flexDirection:"column",gap:14}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <div>
          <div style={{fontSize:13,fontWeight:600,color:C.text}}>{item.offers?.length||0} offerte ricevute</div>
          {item.offers?.length>1&&perm.canSeePrice&&<div style={{fontSize:11,color:C.muted}}>Prezzo migliore: <strong style={{color:"#059669"}}>{fmt(minP)}</strong></div>}
        </div>
        {(perm.canOffer||perm.canAdd)&&!showForm&&<Btn size="sm" onClick={()=>setShowForm(true)}>{myOffer?"✏ Modifica":"+ Invia preventivo"}</Btn>}
      </div>

      {showForm&&<div style={{background:"#F8F4EF",borderRadius:12,padding:16,border:`1px solid ${C.border}`}}>
        <div style={{fontSize:13,fontWeight:600,color:C.text,marginBottom:12}}>{myOffer?"Modifica il tuo preventivo":"Nuovo preventivo"}</div>
        <div style={{display:"flex",flexDirection:"column",gap:10}}>
          {!perm.canOffer&&<div><Lbl>Fornitore</Lbl><Inp value={ofr.supplierName} onChange={v=>setOfr(o=>({...o,supplierName:v}))} placeholder="Nome azienda"/></div>}
          <div><Lbl>Prezzo unitario (€) *</Lbl>
            <div style={{display:"flex",gap:8,alignItems:"center"}}>
              <Inp value={ofr.unitPrice} onChange={v=>setOfr(o=>({...o,unitPrice:v}))} placeholder="0" style={{textAlign:"right",maxWidth:160}}/>
              {ofr.unitPrice&&item.qty&&<div style={{fontSize:12,color:C.muted}}>Totale: <strong style={{color:C.accent}}>{fmt(Number(ofr.unitPrice)*item.qty)}</strong> per {item.qty} {item.unit}</div>}
            </div>
          </div>
          <div><Lbl>Note offerta</Lbl>
            <textarea value={ofr.totalNote||""} onChange={e=>setOfr(o=>({...o,totalNote:e.target.value}))} rows={2} placeholder="Validità, condizioni, tempi consegna..."
              style={{width:"100%",border:`1.5px solid ${C.border}`,borderRadius:8,padding:"8px 12px",background:"#FFF",color:C.text,fontSize:13,fontFamily:"inherit",resize:"vertical",outline:"none",boxSizing:"border-box"}}/>
          </div>
          <div><Lbl>Allegato {(ofr.attName)&&<span style={{fontWeight:400,textTransform:"none",fontSize:10}}>· {ofr.attName}</span>}</Lbl>
            <Btn variant="secondary" size="sm" onClick={()=>attRef.current?.click()}>📎 {ofr.attName?"Cambia file":"Allega PDF / Immagine"}</Btn>
            <input ref={attRef} type="file" accept=".pdf,.jpg,.jpeg,.png" style={{display:"none"}} onChange={handleAtt}/>
            <div style={{fontSize:10,color:C.muted,marginTop:3}}>PDF, JPG, PNG · Max 5MB</div>
          </div>
          <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
            <Btn variant="secondary" size="sm" onClick={()=>setShowForm(false)}>Annulla</Btn>
            <Btn size="sm" onClick={submit} disabled={!ofr.supplierName.trim()||!ofr.unitPrice}>{myOffer?"Aggiorna":"Invia offerta"}</Btn>
          </div>
        </div>
      </div>}

      {(item.offers||[]).length===0&&!showForm&&<div style={{textAlign:"center",padding:"32px 0",color:C.muted,fontSize:13}}><div style={{fontSize:28,marginBottom:6}}>📭</div>Nessuna offerta</div>}
      {(item.offers||[]).map(o=>{
        // Fornitore sees only their own offer
        if(perm.canOffer&&supplierName&&o.supplierName?.toLowerCase()!==supplierName?.toLowerCase()) return null;
        const isMin=perm.canSeePrice&&o.unitPrice===minP&&(item.offers?.length||0)>1;
        return (
          <div key={o.id} style={{borderRadius:12,border:`2px solid ${o.isSelected?"#10B981":C.border}`,background:o.isSelected?"#F0FDF4":"#FFF",padding:14,position:"relative"}}>
            {o.isSelected&&<span style={{position:"absolute",top:-1,right:12,background:"#10B981",color:"#FFF",fontSize:10,fontWeight:700,padding:"2px 10px",borderRadius:"0 0 8px 8px"}}>SELEZIONATA</span>}
            {isMin&&!o.isSelected&&<span style={{position:"absolute",top:-1,right:12,background:"#059669",color:"#FFF",fontSize:10,fontWeight:700,padding:"2px 10px",borderRadius:"0 0 8px 8px"}}>MIGLIOR PREZZO</span>}
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:8}}>
              <div>
                <div style={{fontWeight:600,fontSize:14,color:C.text}}>{o.supplierName}</div>
                <div style={{fontSize:11,color:C.muted,marginTop:1}}>{fmtDate(o.submittedAt)}</div>
              </div>
              {perm.canSeePrice&&<div style={{textAlign:"right"}}>
                <div style={{fontFamily:"'Cormorant Garamond',serif",fontSize:22,fontWeight:600,color:isMin?"#059669":C.text}}>{fmt(o.unitPrice)}</div>
                <div style={{fontSize:11,color:C.muted}}>× {item.qty} {item.unit} = {fmt(o.unitPrice*item.qty)}</div>
              </div>}
              {perm.canOffer&&<div style={{textAlign:"right"}}>
                <div style={{fontFamily:"'Cormorant Garamond',serif",fontSize:22,fontWeight:600,color:C.accent}}>{fmt(o.unitPrice)}</div>
                <div style={{fontSize:11,color:C.muted}}>tua offerta</div>
              </div>}
            </div>
            {o.totalNote&&<div style={{fontSize:12,color:C.muted,marginBottom:8,background:"#F8F4EF",padding:"7px 10px",borderRadius:6,lineHeight:1.4}}>{o.totalNote}</div>}
            <div style={{display:"flex",gap:6,alignItems:"center"}}>
              {o.attachmentKey&&<Btn variant="secondary" size="sm" onClick={()=>dlAtt(o)}>📎 {o.attachmentName}</Btn>}
              {perm.canApprove&&!o.isSelected&&<Btn variant="green" size="sm" onClick={()=>select(o)}>✓ Seleziona</Btn>}
              {o.isSelected&&perm.canApprove&&<Btn variant="ghost" size="sm" onClick={()=>onUpdate({...item,offers:item.offers.map(x=>({...x,isSelected:false}))})}>✕ Deseleziona</Btn>}
            </div>
          </div>
        );
      })}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════
// COMMENTI TAB
// ═══════════════════════════════════════════════════════════════════════
function CommentiTab({item,role,authorName,onUpdate}){
  const [text,setText]=useState("");
  const r=ROLES[role];
  const roleCol={architetto:C.accent,committente:C.gold,fornitore:"#7A8C6E"};
  const add=()=>{
    if(!text.trim())return;
    const c={id:uid(),author:authorName||r.label,role,text:text.trim(),createdAt:new Date().toISOString()};
    onUpdate({...item,comments:[...(item.comments||[]),c]});setText("");
  };
  return (
    <div style={{display:"flex",flexDirection:"column",gap:12}}>
      {(item.comments||[]).length===0&&<div style={{textAlign:"center",padding:"28px 0",color:C.muted,fontSize:13}}><div style={{fontSize:26,marginBottom:6}}>💬</div>Nessun commento</div>}
      {(item.comments||[]).map(c=>(
        <div key={c.id} style={{display:"flex",gap:10}}>
          <div style={{width:32,height:32,borderRadius:"50%",background:(roleCol[c.role]||C.accent)+"22",color:roleCol[c.role]||C.accent,display:"flex",alignItems:"center",justifyContent:"center",fontSize:13,flexShrink:0}}>{ROLES[c.role]?.icon||"👤"}</div>
          <div style={{flex:1,background:"#F8F4EF",borderRadius:"4px 12px 12px 12px",padding:"9px 13px"}}>
            <div style={{display:"flex",justifyContent:"space-between",marginBottom:3}}>
              <span style={{fontSize:12,fontWeight:600,color:roleCol[c.role]||C.accent}}>{c.author}</span>
              <span style={{fontSize:10,color:C.muted}}>{fmtDate(c.createdAt)}</span>
            </div>
            <div style={{fontSize:13,color:C.text,lineHeight:1.5,whiteSpace:"pre-wrap"}}>{c.text}</div>
          </div>
        </div>
      ))}
      <div style={{display:"flex",gap:8,alignItems:"flex-end",marginTop:4}}>
        <div style={{width:32,height:32,borderRadius:"50%",background:(roleCol[role]||C.accent)+"22",color:roleCol[role]||C.accent,display:"flex",alignItems:"center",justifyContent:"center",fontSize:13,flexShrink:0}}>{r.icon}</div>
        <div style={{flex:1,display:"flex",gap:8}}>
          <textarea value={text} onChange={e=>setText(e.target.value)} placeholder="Aggiungi un commento... (Ctrl+Enter per inviare)" rows={2}
            onKeyDown={e=>{if(e.key==="Enter"&&(e.ctrlKey||e.metaKey))add();}}
            style={{flex:1,border:`1.5px solid ${C.border}`,borderRadius:8,padding:"8px 12px",background:"#FFF",color:C.text,fontSize:13,fontFamily:"inherit",resize:"vertical",outline:"none",boxSizing:"border-box"}}/>
          <Btn onClick={add} disabled={!text.trim()}>Invia</Btn>
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════
// ITEM MODAL WRAPPER
// ═══════════════════════════════════════════════════════════════════════
function ItemModal({item,project,role,supplierName,onSave,onClose}){
  const isNew=!item.id;
  const perm=ROLES[role];
  const [activeTab,setActiveTab]=useState(role==="fornitore"&&!isNew?"offerte":"scheda");
  const [form,setForm]=useState({
    zoneId:project.zones[0]?.id||"",ambiente:"",catId:project.categories[0]?.id||"",
    description:"",referenceSupplier:"",supplierCode:"",webLink:"",qty:"",unit:"pz",
    unitPrice:"",unitPriceInstall:"",status:"da_definire",notes:"",image:null,floorId:"",offers:[],comments:[],customFields:{},...item
  });
  const handleSave=()=>{
    if(!form.description.trim())return;
    const code=form.code||genCode(form.zoneId,form.catId,project.items);
    onSave({...form,code,qty:Number(form.qty)||0,unitPrice:Number(form.unitPrice)||0,unitPriceInstall:Number(form.unitPriceInstall)||0});
  };
  const tabs=[
    {id:"scheda",  icon:"📋",label:"Scheda"},
    {id:"offerte", icon:"💰",label:"Offerte", badge:form.offers?.length||null},
    {id:"commenti",icon:"💬",label:"Commenti",badge:form.comments?.length||null},
  ];
  return (
    <div style={{position:"fixed",inset:0,zIndex:1000,display:"flex"}}>
      <div onClick={onClose} style={{flex:1,background:"rgba(28,25,22,0.5)",backdropFilter:"blur(2px)"}}/>
      <div style={{width:580,background:C.bg,display:"flex",flexDirection:"column",boxShadow:"-8px 0 40px rgba(0,0,0,0.2)",overflowY:"auto"}}>
        <div style={{padding:"18px 26px 0",background:"#FFF",borderBottom:`1px solid ${C.border}`}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:14}}>
            <div>
              <div style={{fontFamily:"'DM Mono',monospace",fontSize:11,color:C.gold,marginBottom:3}}>{isNew?"NUOVO ELEMENTO":form.code}</div>
              <h3 style={{fontFamily:"'Cormorant Garamond',serif",fontSize:20,margin:0,color:C.text,maxWidth:460,lineHeight:1.2}}>{isNew?"Aggiungi voce":form.description}</h3>
            </div>
            <button onClick={onClose} style={{background:"none",border:"none",cursor:"pointer",fontSize:22,color:C.muted,lineHeight:1}}>×</button>
          </div>
          <div style={{display:"flex"}}>
            {tabs.map(t=>(
              <button key={t.id} onClick={()=>setActiveTab(t.id)}
                style={{padding:"9px 16px",background:"none",border:"none",cursor:"pointer",fontSize:13,fontFamily:"inherit",
                  fontWeight:activeTab===t.id?600:400,color:activeTab===t.id?C.accent:C.muted,
                  borderBottom:activeTab===t.id?`2px solid ${C.accent}`:"2px solid transparent",marginBottom:-1,
                  display:"flex",alignItems:"center",gap:5,whiteSpace:"nowrap"}}>
                {t.icon} {t.label}
                {t.badge>0&&<span style={{background:C.accent,color:"#FFF",borderRadius:99,padding:"1px 6px",fontSize:10,fontWeight:700}}>{t.badge}</span>}
              </button>
            ))}
          </div>
        </div>
        <div style={{flex:1,padding:"20px 26px",overflowY:"auto"}}>
          {activeTab==="scheda"  &&<SchedaTab   form={form} setForm={setForm} project={project} isNew={isNew} perm={perm}/>}
          {activeTab==="offerte" &&<OfferteTab  item={form} perm={perm} supplierName={supplierName} onUpdate={setForm}/>}
          {activeTab==="commenti"&&<CommentiTab item={form} role={role} authorName={supplierName} onUpdate={setForm}/>}
        </div>
        <div style={{padding:"14px 26px",background:"#FFF",borderTop:`1px solid ${C.border}`,display:"flex",justifyContent:"flex-end",gap:8}}>
          <Btn variant="secondary" onClick={onClose}>Chiudi</Btn>
          {(perm.canEdit||isNew)&&activeTab==="scheda"&&<Btn onClick={handleSave} disabled={!form.description.trim()}>{isNew?"Aggiungi":"Salva"}</Btn>}
          {(activeTab==="offerte"||activeTab==="commenti")&&<Btn onClick={()=>onSave(form)}>Salva</Btn>}
        </div>
      </div>
    </div>
  );
}
// ═══════════════════════════════════════════════════════════════════════
// EXCEL IMPORT WIZARD
// ═══════════════════════════════════════════════════════════════════════
function ExcelImportWizard({project,onImport,onClose}){
  const [step,setStep]=useState("upload");
  const [rawRows,setRawRows]=useState([]);
  const [headers,setHeaders]=useState([]);
  const [mapping,setMapping]=useState({});
  const [unknownHeaders,setUnknownHeaders]=useState([]);
  const [preview,setPreview]=useState([]);
  const [importing,setImporting]=useState(false);
  const [result,setResult]=useState(null);
  const fileRef=useRef();

  const KNOWN_FIELDS=[
    {key:"description",label:"Descrizione"},{key:"referenceSupplier",label:"Fornitore"},
    {key:"supplierCode",label:"Codice fornitore"},{key:"webLink",label:"Link prodotto"},
    {key:"qty",label:"Quantità"},{key:"unit",label:"Unità misura"},
    {key:"unitPrice",label:"Prezzo fornitura"},{key:"unitPriceInstall",label:"Prezzo posa in opera"},
    {key:"notes",label:"Note"},{key:"status",label:"Stato"},{key:"zoneId",label:"Zona"},
    {key:"ambiente",label:"Ambiente"},{key:"floorId",label:"Piano"},
    {key:"catId",label:"Categoria"},{key:"code",label:"Codice abaco"},
  ];

  const handleFile=async e=>{
    const file=e.target.files?.[0];if(!file)return;
    const ab=await file.arrayBuffer();
    const wb=XLSX.read(ab,{type:"array"});
    const ws=wb.Sheets[wb.SheetNames[0]];
    const data=XLSX.utils.sheet_to_json(ws,{header:1,defval:""});
    if(!data.length)return;
    const hdrs=data[0].map(h=>String(h||"").trim()).filter(Boolean);
    const rows=data.slice(1).filter(r=>r.some(c=>c!==""));
    setHeaders(hdrs);setRawRows(rows);
    const autoMap={};const unknown=[];
    hdrs.forEach(h=>{
      const offerSupplier=detectOfferHeader(h);
      if(offerSupplier){
        autoMap[h]={action:"offer",customLabel:offerSupplier};
        unknown.push(h);
        return;
      }
      const match=autoMatchField(h);
      if(match){autoMap[h]={action:"map",target:match};return;}
      autoMap[h]={action:"custom",customLabel:h};
      unknown.push(h);
    });
    setMapping(autoMap);setUnknownHeaders(unknown);
    setStep(unknown.length>0?"mapping":"preview");
  };

  const buildPreview=useCallback(()=>rawRows.slice(0,4).map(row=>{
    const obj={};
    headers.forEach((h,i)=>{
      const m=mapping[h];if(!m||m.action==="skip")return;
      const val=String(row[i]||"").trim();
      if(m.action==="map"){
        const fld=KNOWN_FIELDS.find(f=>f.key===m.target);
        obj[fld?.label||m.target]=val;
      }
      else if(m.action==="offer"&&m.customLabel)obj[`💰 ${m.customLabel}`]=val;
      else if(m.action==="custom"&&m.customLabel)obj[`★ ${m.customLabel}`]=val;
    });
    return obj;
  }),[rawRows,headers,mapping]);

  useEffect(()=>{if(step==="preview")setPreview(buildPreview());},[step,buildPreview]);

  const doImport=async()=>{
    setImporting(true);
    const customLabels=new Set();
    const imported=rawRows.map(row=>{
      const obj={id:uid(),offers:[],comments:[],image:null,customFields:{},created:new Date().toISOString().slice(0,10)};
      headers.forEach((h,i)=>{
        const m=mapping[h];if(!m||m.action==="skip")return;
        const val=String(row[i]||"").trim();
        if(m.action==="map"){
          if(m.target==="qty"||m.target==="unitPrice"||m.target==="unitPriceInstall")obj[m.target]=parseFloat(val.replace(",","."))||0;
          else if(m.target==="status")obj[m.target]=statusFromStr(val);
          else obj[m.target]=val;
        }else if(m.action==="offer"&&m.customLabel){
          const price=parseFloat(val.replace(",","."));
          if(price>0){
            obj.offers.push({
              id:uid(),supplierName:m.customLabel,unitPrice:price,
              totalNote:"Importato da Excel",attachmentName:null,attachmentKey:null,attachmentType:null,
              submittedAt:new Date().toISOString(),isSelected:false,
            });
          }
        }else if(m.action==="custom"&&m.customLabel&&val){
          obj.customFields[m.customLabel]=val;
          customLabels.add(m.customLabel);
        }
      });
      if(!obj.description&&!obj.code)return null;
      if(!obj.status)obj.status="da_definire";
      if(!obj.unit)obj.unit="pz";
      if(!obj.code)obj.code=genCode(obj.zoneId||"XX",obj.catId||"XX",[]);
      if(Object.keys(obj.customFields).length===0)delete obj.customFields;
      return obj;
    }).filter(Boolean);
    setResult({count:imported.length,customFields:[...customLabels]});
    onImport(imported);
    setStep("done");setImporting(false);
  };

  const canProceed=!unknownHeaders.some(h=>{
    const m=mapping[h];
    return(m?.action==="map"&&!m?.target)||(m?.action==="custom"&&!m?.customLabel)||(m?.action==="offer"&&!m?.customLabel);
  });

  const stepLabels=["upload","mapping","preview","done"];
  const stepNames=["Carica","Mappa","Anteprima","Fine"];

  return(
    <div style={{position:"fixed",inset:0,zIndex:2000,display:"flex",alignItems:"center",justifyContent:"center",background:"rgba(28,25,22,0.55)",backdropFilter:"blur(3px)"}}>
      <div style={{background:C.bg,borderRadius:18,width:step==="mapping"?680:520,maxWidth:"96vw",maxHeight:"90vh",display:"flex",flexDirection:"column",boxShadow:"0 24px 64px rgba(0,0,0,0.25)",border:`1px solid ${C.border}`,overflow:"hidden"}}>
        {/* Header */}
        <div style={{padding:"18px 22px",borderBottom:`1px solid ${C.border}`,display:"flex",justifyContent:"space-between",alignItems:"center",flexShrink:0}}>
          <div>
            <div style={{fontSize:15,fontWeight:600,color:C.text}}>Importa da Excel</div>
            <div style={{fontSize:12,color:C.muted,marginTop:2}}>
              {step==="upload"&&"Carica il file Excel con il tuo abaco"}
              {step==="mapping"&&`${unknownHeaders.length} colonne non riconosciute — scegli cosa farne`}
              {step==="preview"&&`Anteprima — ${rawRows.length} righe trovate`}
              {step==="done"&&"Importazione completata"}
            </div>
          </div>
          <button onClick={onClose} style={{background:"none",border:"none",cursor:"pointer",fontSize:20,color:C.muted}}>×</button>
        </div>
        {/* Step bar */}
        <div style={{display:"flex",borderBottom:`1px solid ${C.border}`,flexShrink:0}}>
          {stepLabels.map((s,i)=>(
            <div key={s} style={{flex:1,padding:"9px 0",textAlign:"center",fontSize:11,fontWeight:500,
              color:step===s?C.accent:stepLabels.indexOf(step)>i?"#059669":C.muted,
              borderBottom:step===s?`2px solid ${C.accent}`:"2px solid transparent"}}>
              {stepNames[i]}
            </div>
          ))}
        </div>
        {/* Body */}
        <div style={{flex:1,overflowY:"auto",padding:"20px 22px"}}>

          {step==="upload"&&(
            <div style={{display:"flex",flexDirection:"column",gap:14}}>
              <div onClick={()=>fileRef.current?.click()}
                style={{border:`2px dashed ${C.border}`,borderRadius:12,padding:"40px 20px",textAlign:"center",cursor:"pointer",background:"#FFF"}}
                onMouseEnter={e=>e.currentTarget.style.borderColor=C.accent} onMouseLeave={e=>e.currentTarget.style.borderColor=C.border}>
                <div style={{fontSize:32,marginBottom:8}}>📊</div>
                <div style={{fontSize:14,fontWeight:500,color:C.text,marginBottom:4}}>Clicca per selezionare il file</div>
                <div style={{fontSize:12,color:C.muted}}>Excel .xlsx .xls oppure .csv</div>
              </div>
              <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={handleFile}/>
              <div style={{background:"#F8F4EF",borderRadius:8,padding:"12px 14px"}}>
                <div style={{fontSize:11,fontWeight:600,color:C.muted,marginBottom:6}}>COLONNE RICONOSCIUTE AUTOMATICAMENTE</div>
                <div style={{display:"flex",flexWrap:"wrap",gap:5}}>
                  {["descrizione","fornitore","quantità","prezzo unit","zona","ambiente","piano","categoria","stato","note","link","codice fornitore"].map(a=>(
                    <span key={a} style={{fontSize:11,padding:"2px 8px",borderRadius:99,background:C.border,color:C.muted}}>{a}</span>
                  ))}
                </div>
              </div>
            </div>
          )}

          {step==="mapping"&&(
            <div style={{display:"flex",flexDirection:"column",gap:10}}>
              {Object.entries(mapping).filter(([,m])=>m.action==="map").length>0&&(
                <div style={{background:"#F0FDF4",borderRadius:8,padding:"10px 14px",border:"1px solid #D1FAE5",marginBottom:4}}>
                  <div style={{fontSize:11,fontWeight:600,color:"#059669",marginBottom:4}}>✓ Colonne abbinate automaticamente</div>
                  <div style={{display:"flex",flexWrap:"wrap",gap:4}}>
                    {Object.entries(mapping).filter(([,m])=>m.action==="map").map(([h,m])=>(
                      <span key={h} style={{fontSize:11,padding:"2px 8px",borderRadius:99,background:"white",border:`1px solid ${C.border}`,color:C.text}}>
                        <span style={{color:C.muted}}>{h}</span> → <strong>{KNOWN_FIELDS.find(f=>f.key===m.target)?.label||m.target}</strong>
                      </span>
                    ))}
                  </div>
                </div>
              )}
              <div style={{fontSize:12,fontWeight:600,color:C.text}}>Colonne non riconosciute:</div>
              {unknownHeaders.map(h=>{
                const m=mapping[h]||{action:"skip"};
                const exVal=String(rawRows[0]?.[headers.indexOf(h)]||"").slice(0,30);
                return(
                  <div key={h} style={{border:`1px solid ${C.border}`,borderRadius:8,padding:"12px 14px",background:m.action==="skip"?"#F8F4EF":"#FFF"}}>
                    <div style={{display:"flex",alignItems:"flex-start",gap:10}}>
                      <div style={{flex:"0 0 auto"}}>
                        <div style={{fontSize:13,fontWeight:600,color:C.text,fontFamily:"'DM Mono',monospace"}}>"{h}"</div>
                        {exVal&&<div style={{fontSize:11,color:C.muted,marginTop:2}}>es: "{exVal}"</div>}
                      </div>
                      <div style={{flex:1,display:"flex",flexDirection:"column",gap:6}}>
                        <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
                          {[["map","Abbina a campo"],["offer","Offerta da fornitore"],["custom","Campo personalizzato"],["skip","Ignora"]].map(([act,lbl])=>(
                            <label key={act} style={{display:"flex",alignItems:"center",gap:4,cursor:"pointer",fontSize:12,padding:"4px 10px",borderRadius:99,border:`1px solid ${m.action===act?"#3B82F6":C.border}`,background:m.action===act?"#EFF6FF":"#F8F4EF",color:m.action===act?"#3B82F6":C.muted,fontWeight:m.action===act?600:400}}>
                              <input type="radio" checked={m.action===act} onChange={()=>setMapping(mp=>({...mp,[h]:{...m,action:act}}))} style={{display:"none"}}/>
                              {lbl}
                            </label>
                          ))}
                        </div>
                        {m.action==="map"&&(
                          <Sel value={m.target||""} onChange={v=>setMapping(mp=>({...mp,[h]:{...m,target:v}}))} style={{width:"auto",maxWidth:240}}>
                            <option value="">— Seleziona campo —</option>
                            {KNOWN_FIELDS.map(f=><option key={f.key} value={f.key}>{f.label}</option>)}
                          </Sel>
                        )}
                        {m.action==="offer"&&(
                          <Inp value={m.customLabel||""} onChange={v=>setMapping(mp=>({...mp,[h]:{...m,customLabel:v}}))} placeholder="Nome fornitore" style={{maxWidth:240}}/>
                        )}
                        {m.action==="custom"&&(
                          <Inp value={m.customLabel||h} onChange={v=>setMapping(mp=>({...mp,[h]:{...m,customLabel:v}}))} placeholder="Nome campo personalizzato" style={{maxWidth:240}}/>
                        )}
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>
          )}

          {step==="preview"&&(
            <div>
              <div style={{fontSize:12,color:C.muted,marginBottom:12}}>Prime 4 righe su {rawRows.length} totali</div>
              <div style={{overflowX:"auto",borderRadius:8,border:`1px solid ${C.border}`}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                  <thead>
                    <tr style={{background:"#F8F4EF"}}>
                      {Object.keys(preview[0]||{}).map(k=>(
                        <th key={k} style={{padding:"7px 10px",textAlign:"left",fontWeight:600,color:C.muted,whiteSpace:"nowrap",borderBottom:`1px solid ${C.border}`,fontSize:11}}>{k}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {preview.map((row,i)=>(
                      <tr key={i} style={{borderBottom:`1px solid ${C.border}`}}>
                        {Object.values(row).map((v,j)=>(
                          <td key={j} style={{padding:"6px 10px",color:C.text,maxWidth:160,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{String(v||"—")}</td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <div style={{marginTop:12,padding:"10px 14px",background:"#FFFBEB",borderRadius:8,fontSize:12,color:C.gold}}>
                ⚠ Verranno importate {rawRows.length} righe come nuove voci nel progetto.
              </div>
            </div>
          )}

          {step==="done"&&result&&(
            <div style={{textAlign:"center",padding:"24px 0"}}>
              <div style={{fontSize:40,marginBottom:12}}>✅</div>
              <div style={{fontSize:16,fontWeight:600,color:C.text,marginBottom:6}}>{result.count} voci importate</div>
              {result.customFields?.length>0&&(
                <div style={{fontSize:12,color:C.muted,marginTop:8}}>
                  Campi personalizzati creati: {result.customFields.map(c=>(<span key={c} style={{display:"inline-block",margin:"2px 3px",padding:"2px 8px",borderRadius:99,background:C.border,color:C.text,fontFamily:"'DM Mono',monospace",fontSize:11}}>★ {c}</span>))}
                </div>
              )}
              <div style={{fontSize:12,color:C.muted,marginTop:8}}>Vai all'Abaco per completare i campi mancanti</div>
            </div>
          )}
        </div>

        {/* Footer */}
        <div style={{padding:"14px 22px",borderTop:`1px solid ${C.border}`,display:"flex",justifyContent:"space-between",alignItems:"center",flexShrink:0}}>
          <div style={{fontSize:12,color:C.muted}}>
            {step==="mapping"&&`${rawRows.length} righe · ${headers.length} colonne`}
            {step==="preview"&&`${rawRows.length} righe pronte`}
          </div>
          <div style={{display:"flex",gap:8}}>
            <Btn variant="secondary" onClick={onClose}>Annulla</Btn>
            {step==="mapping"&&<Btn onClick={()=>setStep("preview")} disabled={!canProceed}>Continua →</Btn>}
            {step==="preview"&&<><Btn variant="secondary" onClick={()=>setStep("mapping")}>← Indietro</Btn><Btn onClick={doImport} disabled={importing}>{importing?"Importo...":"Importa adesso"}</Btn></>}
            {step==="done"&&<Btn onClick={onClose}>Chiudi</Btn>}
          </div>
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════
// ABACO TABLE
// ═══════════════════════════════════════════════════════════════════════
function AbacoTable({project,role,supplierName,onEdit,onDelete,onAdd,onExport,onImport,filters,setFilters}){
  const {items,zones,categories,floors}=project;
  const perm=ROLES[role];
  const [confirmDel,setConfirmDel]=useState(null);
  const [showImport,setShowImport]=useState(false);
  const [filterFloor,    setFilterFloor]    = useState("");
  const [filterAmbiente, setFilterAmbiente] = useState("");
  const [filterSupplier, setFilterSupplier] = useState("");
  const {search,filterZone,filterCat,filterStatus}=filters;

  const baseItems=perm.canOffer&&supplierName
    ? items.filter(i=>i.referenceSupplier?.toLowerCase()===supplierName.toLowerCase())
    : items;

  const filtered = useMemo(() => baseItems.filter(i => {
    if(filterFloor    && i.floorId            !== filterFloor)    return false;
    if(filterZone     && i.zoneId             !== filterZone)     return false;
    if(filterAmbiente && i.ambiente           !== filterAmbiente) return false;
    if(filterSupplier && i.referenceSupplier  !== filterSupplier) return false;
    if(filterCat      && i.catId              !== filterCat)      return false;
    if(filterStatus   && i.status             !== filterStatus)   return false;
    if(search){const q=search.toLowerCase();return(i.code||"").toLowerCase().includes(q)||(i.description||"").toLowerCase().includes(q)||(i.referenceSupplier||"").toLowerCase().includes(q)||(i.ambiente||"").toLowerCase().includes(q);}
    return true;
  }),[baseItems,filterFloor,filterZone,filterAmbiente,filterSupplier,filterCat,filterStatus,search]);

  const grandTotal=perm.canSeePrice?filtered.reduce((s,i)=>s+getTotal(i),0):null;
  const activeFilters=[filterZone,filterCat,filterStatus,filterFloor,filterAmbiente,filterSupplier].filter(Boolean).length;

  const resetFilters=()=>{
    setFilters({search:"",filterZone:"",filterCat:"",filterStatus:""});
    setFilterFloor(""); setFilterAmbiente(""); setFilterSupplier("");
  };

  return (
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:18}}>
        <div>
          <h2 style={{fontFamily:"'Cormorant Garamond',serif",fontSize:28,fontWeight:600,color:C.text,margin:0}}>Abaco Forniture</h2>
          <p style={{margin:"3px 0 0",color:C.muted,fontSize:13}}>
            {filtered.length} voci{perm.canSeePrice&&<> · Totale: <strong style={{color:C.accent}}>{fmt(grandTotal)}</strong></>}
          </p>
        </div>
        <div style={{display:"flex",gap:8}}>
         {perm.canSeePrice&&<Btn variant="secondary" onClick={onExport} size="sm">↓ Excel</Btn>}
          {perm.canAdd&&<Btn variant="secondary" onClick={()=>setShowImport(true)} size="sm">↑ Importa Excel</Btn>}
          {perm.canAdd&&<Btn onClick={onAdd} size="sm">+ Aggiungi voce</Btn>}
        </div>
      </div>

      <Card style={{padding:"12px 14px",marginBottom:14,display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
        <input value={search} onChange={e=>setFilters(f=>({...f,search:e.target.value}))} placeholder="🔍  Cerca codice, descrizione, fornitore..."
          style={{flex:1,minWidth:180,border:`1.5px solid ${C.border}`,borderRadius:8,padding:"7px 11px",background:"#F8F4EF",color:C.text,fontSize:13,fontFamily:"inherit",outline:"none"}}/>
        <Sel value={filterFloor} onChange={setFilterFloor} style={{width:"auto"}}>
          <option value="">Tutti i piani</option>
          {[...new Set(items.map(i=>i.floorId).filter(Boolean))].map(id=>{
            const f=floors.find(x=>x.id===id)||DEFAULT_FLOORS.find(x=>x.id===id);
            return <option key={id} value={id}>{f?.name||id}</option>;
          })}
        </Sel>
        <Sel value={filterZone} onChange={v=>setFilters(f=>({...f,filterZone:v}))} style={{width:"auto"}}>
          <option value="">Tutte le zone</option>
          {zones.map(z=><option key={z.id} value={z.id}>{z.name}</option>)}
        </Sel>
        <Sel value={filterAmbiente} onChange={setFilterAmbiente} style={{width:"auto"}}>
          <option value="">Tutti gli ambienti</option>
          {[...new Set(items.filter(i=>!filterZone||i.zoneId===filterZone).map(i=>i.ambiente).filter(Boolean))].map(a=>(
            <option key={a} value={a}>{a}</option>
          ))}
        </Sel>
        <Sel value={filterSupplier} onChange={setFilterSupplier} style={{width:"auto"}}>
          <option value="">Tutti i fornitori</option>
          {[...new Set(items.map(i=>i.referenceSupplier).filter(Boolean))].sort().map(s=>(
            <option key={s} value={s}>{s}</option>
          ))}
        </Sel>
        <Sel value={filterCat} onChange={v=>setFilters(f=>({...f,filterCat:v}))} style={{width:"auto"}}>
          <option value="">Tutte le categorie</option>
          {categories.map(c=><option key={c.id} value={c.id}>{c.id} – {c.name}</option>)}
        </Sel>
        <Sel value={filterStatus} onChange={v=>setFilters(f=>({...f,filterStatus:v}))} style={{width:"auto"}}>
          <option value="">Tutti gli stati</option>
          {Object.entries(STATUS).map(([k,v])=><option key={k} value={k}>{v.label}</option>)}
        </Sel>
        {activeFilters>0&&<Btn variant="ghost" size="sm" onClick={resetFilters}>✕ Pulisci ({activeFilters})</Btn>}
      </Card>

      <Card style={{overflow:"hidden"}}>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
            <thead>
              <tr style={{background:"#F8F4EF",borderBottom:`2px solid ${C.border}`}}>
                {["","Codice","Descrizione","Piano / Zona / Amb.","Cat.","Fornitore","Q.tà",
                  ...(perm.canSeePrice?["Totale"]:[""]),"Offerte","Stato",""].map((h,i)=>(
                  <th key={i} style={{padding:"9px 11px",textAlign:i>=6?"right":"left",fontFamily:"'DM Sans',sans-serif",fontSize:11,fontWeight:600,color:C.muted,textTransform:"uppercase",letterSpacing:"0.05em",whiteSpace:"nowrap"}}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.length===0?
                <tr><td colSpan={12} style={{padding:"40px",textAlign:"center",color:C.muted,fontSize:13}}>Nessuna voce trovata</td></tr>:
                filtered.map((item,idx)=>{
                  const zone=zones.find(z=>z.id===item.zoneId);
                  const floor=floors.find(f=>f.id===item.floorId);
                  const offCount=item.offers?.length||0;
                  const hasSel=item.offers?.some(o=>o.isSelected);
                  return (
                    <tr key={item.id} style={{borderBottom:`1px solid #F0EBE4`,background:idx%2===0?"#FFF":"#FDFAF7",cursor:"pointer",transition:"background .1s"}}
                      onClick={()=>onEdit(item)}>
                      <td style={{padding:"7px 9px",width:38}}>
                        {item.image?<img src={item.image} alt="" style={{width:30,height:30,borderRadius:5,objectFit:"cover",border:`1px solid ${C.border}`}}/>:
                        <div style={{width:30,height:30,borderRadius:5,background:"#F3F0EC",display:"flex",alignItems:"center",justifyContent:"center"}}>
                          <CatBadge catId={item.catId} cats={categories}/>
                        </div>}
                      </td>
                      <td style={{padding:"7px 11px",whiteSpace:"nowrap"}} onClick={e=>e.stopPropagation()}>
                        <div style={{display:"flex",alignItems:"center",gap:5}}>
                          <span style={{fontFamily:"'DM Mono',monospace",fontSize:11,color:C.gold}}>{item.code}</span>
                          {item.webLink&&<a href={item.webLink} target="_blank" rel="noopener noreferrer" onClick={e=>e.stopPropagation()} style={{fontSize:10,color:"#1D4ED8",textDecoration:"none",padding:"1px 4px",background:"#DBEAFE",borderRadius:3}}>↗</a>}
                        </div>
                      </td>
                      <td style={{padding:"7px 11px",maxWidth:200}}>
                        <div style={{fontWeight:500,color:C.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{item.description}</div>
                        {item.comments?.length>0&&<span style={{fontSize:10,color:C.muted}}>💬 {item.comments.length}</span>}
                      </td>
                      <td style={{padding:"7px 11px",whiteSpace:"nowrap"}}>
                        {floor&&<div style={{fontSize:10,color:C.gold,fontFamily:"'DM Mono',monospace"}}>{floor.name}</div>}
                        <div style={{fontSize:12,color:C.text}}>{zone?.name||item.zoneId}</div>
                        <div style={{fontSize:11,color:C.muted}}>{item.ambiente}</div>
                      </td>
                      <td style={{padding:"7px 11px"}}>
                        <span title={categories.find(c=>c.id===item.catId)?.name} style={{display:"inline-block",padding:"2px 7px",borderRadius:4,background:(categories.find(c=>c.id===item.catId)?.color||"#888")+"18",color:categories.find(c=>c.id===item.catId)?.color||"#888",fontSize:11,fontWeight:600,fontFamily:"'DM Mono',monospace",cursor:"help"}}>
                          {item.catId}
                        </span>
                      </td>
                      <td style={{padding:"7px 11px",whiteSpace:"nowrap"}}>
                        <div style={{fontSize:12,color:C.text}}>{item.referenceSupplier||"—"}</div>
                        {item.supplierCode&&<div style={{fontFamily:"'DM Mono',monospace",fontSize:10,color:C.muted}}>{item.supplierCode}</div>}
                      </td>
                      <td style={{padding:"7px 11px",textAlign:"right",fontFamily:"'DM Mono',monospace",whiteSpace:"nowrap",color:C.muted,fontSize:12}}>{item.qty} {item.unit}</td>
                      {perm.canSeePrice&&<td style={{padding:"7px 11px",textAlign:"right",fontFamily:"'Cormorant Garamond',serif",fontSize:16,fontWeight:600,color:C.text,whiteSpace:"nowrap"}}>{fmt(getTotal(item))}</td>}
                      {!perm.canSeePrice&&<td/>}
                      <td style={{padding:"7px 11px",textAlign:"right",whiteSpace:"nowrap"}}>
                        {offCount>0&&<span style={{display:"inline-flex",alignItems:"center",gap:3,fontSize:10,padding:"2px 7px",borderRadius:99,background:hasSel?"#D1FAE5":"#FEF3C7",color:hasSel?"#059669":"#B45309"}}>
                          {hasSel?"✓":"⏳"} {offCount}
                        </span>}
                      </td>
                      <td style={{padding:"7px 11px",whiteSpace:"nowrap"}}><StatusBadge status={item.status}/></td>
                      <td style={{padding:"7px 11px",textAlign:"right"}} onClick={e=>e.stopPropagation()}>
                        {perm.canDelete&&(confirmDel===item.id?
                          <span style={{display:"flex",gap:3}}>
                            <Btn variant="danger" size="xs" onClick={()=>{onDelete(item.id);setConfirmDel(null);}}>Sì</Btn>
                            <Btn variant="secondary" size="xs" onClick={()=>setConfirmDel(null)}>No</Btn>
                          </span>:
                          <Btn variant="ghost" size="sm" onClick={()=>setConfirmDel(item.id)} style={{color:"#DC2626"}}>✕</Btn>)}
                      </td>
                    </tr>
                  );
                })
              }
            </tbody>
            {filtered.length>0&&perm.canSeePrice&&<tfoot>
              <tr style={{background:"#F8F4EF",borderTop:`2px solid ${C.border}`}}>
                <td colSpan={7} style={{padding:"9px 11px",fontWeight:600,fontSize:12,color:C.muted}}>TOTALE ({filtered.length})</td>
                <td style={{padding:"9px 11px",textAlign:"right",fontFamily:"'Cormorant Garamond',serif",fontSize:20,fontWeight:600,color:C.accent}}>{fmt(grandTotal)}</td>
                <td colSpan={4}/>
              </tr>
            </tfoot>}
          </table>
        </div>
      </Card>
   {showImport&&<ExcelImportWizard project={project} onImport={imported=>{onImport(imported);setShowImport(false);}} onClose={()=>setShowImport(false)}/>}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════
// ZONE VIEW
// ═══════════════════════════════════════════════════════════════════════
function ZoneView({project,role}){
  const {items,zones,floors,categories}=project;
  const perm=ROLES[role];
  const [open,setOpen]=useState(null);
  return (
    <div>
      <h2 style={{fontFamily:"'Cormorant Garamond',serif",fontSize:28,fontWeight:600,color:C.text,margin:"0 0 22px"}}>Vista per Zone</h2>
      <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:14}}>
        {zones.map(zone=>{
          const zItems=items.filter(i=>i.zoneId===zone.id);
          const zVal=zItems.reduce((s,i)=>s+getTotal(i),0);
          const isOpen=open===zone.id;
          const byAmb=zone.ambienti.map(a=>({name:a,items:zItems.filter(i=>i.ambiente===a)})).filter(a=>a.items.length>0);
          return (
            <Card key={zone.id} style={{overflow:"hidden",cursor:"pointer"}}>
              <div style={{padding:"16px 20px",display:"flex",justifyContent:"space-between",alignItems:"center",borderBottom:isOpen?`1px solid #F0EBE4`:"none"}} onClick={()=>setOpen(isOpen?null:zone.id)}>
                <div>
                  <div style={{fontFamily:"'DM Mono',monospace",fontSize:10,color:C.gold,letterSpacing:"0.12em",marginBottom:3}}>{zone.id}</div>
                  <div style={{fontFamily:"'Cormorant Garamond',serif",fontSize:17,fontWeight:600,color:C.text}}>{zone.name}</div>
                  <div style={{fontSize:12,color:C.muted,marginTop:2}}>{zItems.length} voci{perm.canSeePrice&&<> · {fmt(zVal)}</>}</div>
                </div>
                <span style={{fontSize:18,color:C.accent,transition:"transform .2s",transform:isOpen?"rotate(180deg)":"none",display:"inline-block"}}>▾</span>
              </div>
              {isOpen&&<div style={{padding:"6px 0"}}>
                {byAmb.length===0?<div style={{padding:"10px 20px",color:C.muted,fontSize:12}}>Nessuna voce</div>:
                byAmb.map(a=>(
                  <div key={a.name} style={{padding:"9px 20px",borderBottom:`1px solid #F8F4EF`}}>
                    <div style={{display:"flex",justifyContent:"space-between",marginBottom:5}}>
                      <span style={{fontSize:13,fontWeight:500,color:C.text}}>{a.name}</span>
                      <span style={{fontFamily:"'DM Mono',monospace",fontSize:11,color:C.muted}}>{a.items.length} voci</span>
                    </div>
                    <div style={{display:"flex",flexWrap:"wrap",gap:3}}>
                      {a.items.map(i=><span key={i.id} title={i.description} style={{fontSize:10,padding:"2px 6px",borderRadius:4,background:"#F8F4EF",color:C.muted,fontFamily:"'DM Mono',monospace",cursor:"help"}}>{i.catId}</span>)}
                    </div>
                  </div>
                ))}
              </div>}
            </Card>
          );
        })}
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════
// PORTALE FORNITORE
// ═══════════════════════════════════════════════════════════════════════
function PortaleFornitore({project,supplierName,onEdit}){
  const {items,zones,categories}=project;
  const assigned=items.filter(i=>i.referenceSupplier?.toLowerCase()===supplierName?.toLowerCase());
  const myOfferItems=assigned; // items already filtered to supplier
  const myOffersCount=myOfferItems.reduce((s,i)=>s+(i.offers?.filter(o=>o.supplierName?.toLowerCase()===supplierName?.toLowerCase()).length||0),0);
  const selectedCount=myOfferItems.reduce((s,i)=>s+(i.offers?.filter(o=>o.isSelected&&o.supplierName?.toLowerCase()===supplierName?.toLowerCase()).length||0),0);

  return (
    <div>
      <div style={{marginBottom:24}}>
        <div style={{fontFamily:"'DM Mono',monospace",fontSize:11,letterSpacing:"0.15em",color:C.gold,textTransform:"uppercase",marginBottom:4}}>Portale Fornitore</div>
        <h2 style={{fontFamily:"'Cormorant Garamond',serif",fontSize:32,fontWeight:600,color:C.text,margin:0}}>{supplierName}</h2>
        <p style={{color:C.muted,fontSize:13,marginTop:4}}>{assigned.length} richieste di preventivo · {myOffersCount} offerte inviate</p>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:14,marginBottom:24}}>
        {[{label:"Richieste assegnate",val:assigned.length,c:C.accent},{label:"Offerte inviate",val:myOffersCount,c:"#7A8C6E"},{label:"Offerte selezionate",val:selectedCount,c:"#059669"}].map((k,i)=>(
          <Card key={i} style={{padding:"16px 20px",borderTop:`3px solid ${k.c}`}}>
            <div style={{fontSize:11,fontWeight:600,color:C.muted,textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:4}}>{k.label}</div>
            <div style={{fontFamily:"'Cormorant Garamond',serif",fontSize:34,fontWeight:600,color:k.c}}>{k.val}</div>
          </Card>
        ))}
      </div>
      <div style={{display:"flex",flexDirection:"column",gap:10}}>
        {myOfferItems.length===0&&<div style={{textAlign:"center",padding:"40px",color:C.muted,background:"#FFF",borderRadius:14,fontSize:13}}>
          <div style={{fontSize:32,marginBottom:8}}>📭</div>Nessuna richiesta assegnata
        </div>}
        {myOfferItems.map(item=>{
          const myOffer=item.offers?.find(o=>o.supplierName?.toLowerCase()===supplierName?.toLowerCase());
          const zone=zones.find(z=>z.id===item.zoneId);
          return (
            <Card key={item.id} style={{padding:16,cursor:"pointer",border:myOffer?.isSelected?`2px solid #10B981`:`2px solid transparent`}} onClick={()=>onEdit(item,"offerte")}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
                <div style={{flex:1}}>
                  <div style={{display:"flex",gap:7,alignItems:"center",marginBottom:5}}>
                    <span style={{fontFamily:"'DM Mono',monospace",fontSize:11,color:C.gold}}>{item.code}</span>
                    {myOffer?.isSelected&&<span style={{fontSize:10,background:"#D1FAE5",color:"#059669",padding:"2px 8px",borderRadius:99,fontWeight:600}}>✓ Selezionata</span>}
                  </div>
                  <div style={{fontSize:14,fontWeight:600,color:C.text,marginBottom:3}}>{item.description}</div>
                  <div style={{fontSize:12,color:C.muted}}>{zone?.name} · {item.ambiente} · <strong>{item.qty} {item.unit}</strong></div>
                  {item.notes&&<div style={{fontSize:11,color:C.muted,marginTop:4,fontStyle:"italic"}}>{item.notes}</div>}
                  {item.webLink&&<a href={item.webLink} target="_blank" rel="noopener noreferrer" onClick={e=>e.stopPropagation()} style={{fontSize:12,color:"#1D4ED8",textDecoration:"none",display:"inline-flex",alignItems:"center",gap:3,marginTop:4}}>↗ Scheda prodotto</a>}
                </div>
                <div style={{marginLeft:16,textAlign:"right",flexShrink:0}}>
                  {myOffer?<div>
                    <div style={{fontFamily:"'Cormorant Garamond',serif",fontSize:22,fontWeight:600,color:C.accent}}>{fmt(myOffer.unitPrice)}</div>
                    <div style={{fontSize:11,color:C.muted}}>tua offerta</div>
                  </div>:<Btn size="sm">+ Invia preventivo</Btn>}
                  <div style={{marginTop:6}}><StatusBadge status={item.status}/></div>
                </div>
              </div>
            </Card>
          );
        })}
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════
// SETTINGS VIEW
// ═══════════════════════════════════════════════════════════════════════
function SettingsView({project,onUpdate,onRenameIndex}){
  const [tab,setTab]=useState("progetto");
  const tabs=[{id:"progetto",label:"⚙ Progetto"},{id:"struttura",label:"🏗 Struttura"},{id:"categorie",label:"🏷 Categorie"}];

  return (
    <div>
      <h2 style={{fontFamily:"'Cormorant Garamond',serif",fontSize:28,fontWeight:600,color:C.text,margin:"0 0 20px"}}>Impostazioni</h2>
      <div style={{display:"flex",gap:0,marginBottom:20,borderBottom:`1px solid ${C.border}`}}>
        {tabs.map(t=><button key={t.id} onClick={()=>setTab(t.id)}
          style={{padding:"10px 20px",background:"none",border:"none",cursor:"pointer",fontSize:13,fontFamily:"inherit",
            fontWeight:tab===t.id?600:400,color:tab===t.id?C.accent:C.muted,
            borderBottom:tab===t.id?`2px solid ${C.accent}`:"2px solid transparent",marginBottom:-1}}>{t.label}</button>)}
      </div>
      {tab==="progetto"  &&<ProjectSettings    project={project} onUpdate={onUpdate} onRenameIndex={onRenameIndex}/>}
      {tab==="struttura" &&<StrutturaSettings  project={project} onUpdate={onUpdate}/>}
      {tab==="categorie" &&<CategorieSettings  project={project} onUpdate={onUpdate}/>}
    </div>
  );
}

function ProjectSettings({project,onUpdate,onRenameIndex}){
  const [name,setName]=useState(project.name);
  const [archPw,setArchPw]=useState(project.archPassword);
  const [clientPw,setClientPw]=useState(project.clientPassword);
  const [saved,setSaved]=useState(false);
  const save=()=>{
    const upd={...project,name:name.trim()||project.name,archPassword:archPw||project.archPassword,clientPassword:clientPw||project.clientPassword};
    onUpdate(upd); onRenameIndex(upd.id,upd.name); setSaved(true); setTimeout(()=>setSaved(false),2000);
  };
  return (
    <div style={{maxWidth:480,display:"flex",flexDirection:"column",gap:18}}>
      <Card style={{padding:22}}>
        <div style={{fontFamily:"'Cormorant Garamond',serif",fontSize:16,fontWeight:600,color:C.text,marginBottom:14}}>Nome progetto</div>
        <Inp value={name} onChange={setName} placeholder="Es. Hotel Taormina"/>
      </Card>
      <Card style={{padding:22}}>
        <div style={{fontFamily:"'Cormorant Garamond',serif",fontSize:16,fontWeight:600,color:C.text,marginBottom:14}}>Password di accesso</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
          <div><Lbl>Architetto</Lbl><Inp value={archPw} onChange={setArchPw} type="password" placeholder="Password architetto"/></div>
          <div><Lbl>Committente</Lbl><Inp value={clientPw} onChange={setClientPw} type="password" placeholder="Password committente"/></div>
        </div>
        <div style={{fontSize:11,color:C.muted,marginTop:8}}>I fornitori accedono con il nome azienda, senza password.</div>
      </Card>
      <div style={{display:"flex",gap:10,alignItems:"center"}}>
        <Btn onClick={save}>{saved?"✓ Salvato":"Salva impostazioni"}</Btn>
        {saved&&<span style={{fontSize:12,color:"#059669"}}>Modifiche salvate</span>}
      </div>
    </div>
  );
}

function StrutturaSettings({project,onUpdate}){
  const [tab,setTab]=useState("piani");
  // Floors
  const [newFloor,setNewFloor]=useState("");
  const [editFloor,setEditFloor]=useState(null);
  // Zones
  const [newZoneName,setNewZoneName]=useState("");
  const [newZoneId,setNewZoneId]=useState("");
  const [editZone,setEditZone]=useState(null);
  const [expandedZone,setExpandedZone]=useState(null);
  const [newAmb,setNewAmb]=useState("");

  const addFloor=()=>{
    if(!newFloor.trim())return;
    const id="F"+uid().slice(0,5).toUpperCase();
    onUpdate({...project,floors:[...project.floors,{id,name:newFloor.trim()}]});
    setNewFloor("");
  };
  const delFloor=id=>{ if(project.items.some(i=>i.floorId===id)){alert("Piano in uso da alcune voci. Rimuovilo prima dalle voci.");return;} onUpdate({...project,floors:project.floors.filter(f=>f.id!==id)}); };
  const saveFloor=()=>{ onUpdate({...project,floors:project.floors.map(f=>f.id===editFloor.id?{...f,name:editFloor.name}:f)}); setEditFloor(null); };

  const addZone=()=>{
    if(!newZoneName.trim()||!newZoneId.trim())return;
    const id=newZoneId.toUpperCase().slice(0,5);
    if(project.zones.find(z=>z.id===id)){alert("ID zona già esistente.");return;}
    onUpdate({...project,zones:[...project.zones,{id,name:newZoneName.trim(),ambienti:[]}]});
    setNewZoneName("");setNewZoneId("");
  };
  const delZone=id=>{ if(project.items.some(i=>i.zoneId===id)){alert("Zona in uso da alcune voci.");return;} onUpdate({...project,zones:project.zones.filter(z=>z.id!==id)}); };
  const saveZone=()=>{ onUpdate({...project,zones:project.zones.map(z=>z.id===editZone.id?{...z,name:editZone.name}:z)}); setEditZone(null); };

  const addAmb=(zoneId)=>{
    if(!newAmb.trim())return;
    onUpdate({...project,zones:project.zones.map(z=>z.id===zoneId?{...z,ambienti:[...z.ambienti,newAmb.trim()]}:z)});
    setNewAmb("");
  };
  const delAmb=(zoneId,a)=>{ if(project.items.some(i=>i.zoneId===zoneId&&i.ambiente===a)){alert("Ambiente in uso.");return;} onUpdate({...project,zones:project.zones.map(z=>z.id===zoneId?{...z,ambienti:z.ambienti.filter(x=>x!==a)}:z)}); };

  const stt=[{id:"piani",label:"Piani"},{id:"zone",label:"Zone & Ambienti"}];

  return (
    <div>
      <div style={{display:"flex",gap:10,marginBottom:18}}>
        {stt.map(t=><Btn key={t.id} variant={tab===t.id?"primary":"secondary"} size="sm" onClick={()=>setTab(t.id)}>{t.label}</Btn>)}
      </div>

      {tab==="piani"&&<div style={{maxWidth:480}}>
        <Card style={{padding:20,marginBottom:14}}>
          <div style={{fontSize:13,fontWeight:600,color:C.text,marginBottom:12}}>Piani del progetto</div>
          {project.floors.map(f=>(
            <div key={f.id} style={{display:"flex",alignItems:"center",gap:8,padding:"8px 0",borderBottom:`1px solid ${C.border}`}}>
              {editFloor?.id===f.id?
                <><Inp value={editFloor.name} onChange={v=>setEditFloor(e=>({...e,name:v}))} style={{flex:1}}/><Btn size="xs" onClick={saveFloor}>✓</Btn><Btn size="xs" variant="ghost" onClick={()=>setEditFloor(null)}>✕</Btn></>:
                <><span style={{flex:1,fontSize:13,color:C.text}}>{f.name}</span><span style={{fontFamily:"'DM Mono',monospace",fontSize:10,color:C.muted,padding:"2px 7px",background:"#F8F4EF",borderRadius:4}}>{f.id}</span>
                <Btn size="xs" variant="ghost" onClick={()=>setEditFloor({...f})}>✏</Btn>
                <Btn size="xs" variant="ghost" onClick={()=>delFloor(f.id)} style={{color:"#DC2626"}}>✕</Btn></>}
            </div>
          ))}
          <div style={{display:"flex",gap:8,marginTop:12}}>
            <Inp value={newFloor} onChange={setNewFloor} placeholder="Es. Quarto Piano" style={{flex:1}} onKeyDown={e=>{if(e.key==="Enter")addFloor();}}/>
            <Btn size="sm" onClick={addFloor} disabled={!newFloor.trim()}>+ Aggiungi</Btn>
          </div>
        </Card>
      </div>}

      {tab==="zone"&&<div>
        <Card style={{padding:20,marginBottom:14}}>
          <div style={{fontSize:13,fontWeight:600,color:C.text,marginBottom:12}}>Zone del progetto</div>
          {project.zones.map(z=>(
            <div key={z.id} style={{borderBottom:`1px solid ${C.border}`,marginBottom:4}}>
              <div style={{display:"flex",alignItems:"center",gap:8,padding:"9px 0"}}>
                {editZone?.id===z.id?
                  <><Inp value={editZone.name} onChange={v=>setEditZone(e=>({...e,name:v}))} style={{flex:1}}/><Btn size="xs" onClick={saveZone}>✓</Btn><Btn size="xs" variant="ghost" onClick={()=>setEditZone(null)}>✕</Btn></>:
                  <><button onClick={()=>setExpandedZone(expandedZone===z.id?null:z.id)} style={{flex:1,textAlign:"left",background:"none",border:"none",cursor:"pointer",fontSize:13,color:C.text,display:"flex",alignItems:"center",gap:6}}>
                    <span style={{color:C.accent,fontSize:12}}>{expandedZone===z.id?"▾":"▸"}</span>
                    <span style={{fontWeight:500}}>{z.name}</span>
                    <span style={{fontFamily:"'DM Mono',monospace",fontSize:10,color:C.muted,padding:"1px 6px",background:"#F8F4EF",borderRadius:3}}>{z.id}</span>
                    <span style={{fontSize:11,color:C.muted}}>({z.ambienti.length} ambienti)</span>
                  </button>
                  <Btn size="xs" variant="ghost" onClick={()=>setEditZone({...z})}>✏</Btn>
                  <Btn size="xs" variant="ghost" onClick={()=>delZone(z.id)} style={{color:"#DC2626"}}>✕</Btn></>}
              </div>
              {expandedZone===z.id&&<div style={{paddingLeft:20,paddingBottom:10}}>
                {z.ambienti.map(a=>(
                  <div key={a} style={{display:"flex",alignItems:"center",gap:6,padding:"4px 0",fontSize:13,color:C.muted}}>
                    <span style={{flex:1}}>· {a}</span>
                    <Btn size="xs" variant="ghost" onClick={()=>delAmb(z.id,a)} style={{color:"#DC2626"}}>✕</Btn>
                  </div>
                ))}
                <div style={{display:"flex",gap:6,marginTop:6}}>
                  <Inp value={newAmb} onChange={setNewAmb} placeholder="Nuovo ambiente..." style={{flex:1}} onKeyDown={e=>{if(e.key==="Enter")addAmb(z.id);}}/>
                  <Btn size="xs" onClick={()=>addAmb(z.id)} disabled={!newAmb.trim()}>+</Btn>
                </div>
              </div>}
            </div>
          ))}
        </Card>
        <Card style={{padding:20}}>
          <div style={{fontSize:13,fontWeight:600,color:C.text,marginBottom:12}}>Aggiungi nuova zona</div>
          <div style={{display:"grid",gridTemplateColumns:"100px 1fr",gap:10}}>
            <div><Lbl>Codice (max 5)</Lbl><Inp value={newZoneId} onChange={v=>setNewZoneId(v.slice(0,5).toUpperCase())} placeholder="Es. TER"/></div>
            <div><Lbl>Nome zona</Lbl><Inp value={newZoneName} onChange={setNewZoneName} placeholder="Es. Terrazza Sud"/></div>
          </div>
          <Btn size="sm" onClick={addZone} disabled={!newZoneName.trim()||!newZoneId.trim()} style={{marginTop:10}}>+ Aggiungi zona</Btn>
        </Card>
      </div>}
    </div>
  );
}

function CategorieSettings({project,onUpdate}){
  const [newCat,setNewCat]=useState({id:"",name:"",description:"",color:CAT_COLORS[0]});
  const [editCat,setEditCat]=useState(null);

  const addCat=()=>{
    if(!newCat.id.trim()||!newCat.name.trim())return;
    const id=newCat.id.toUpperCase().slice(0,5);
    if(project.categories.find(c=>c.id===id)){alert("Codice categoria già esistente.");return;}
    onUpdate({...project,categories:[...project.categories,{...newCat,id}]});
    setNewCat({id:"",name:"",description:"",color:CAT_COLORS[0]});
  };
  const delCat=id=>{ if(project.items.some(i=>i.catId===id)){alert("Categoria in uso da alcune voci.");return;} onUpdate({...project,categories:project.categories.filter(c=>c.id!==id)}); };
  const saveCat=()=>{ onUpdate({...project,categories:project.categories.map(c=>c.id===editCat.id?editCat:c)}); setEditCat(null); };

  return (
    <div>
      <div style={{display:"flex",flexDirection:"column",gap:10,marginBottom:18}}>
        {project.categories.map(c=>(
          <Card key={c.id} style={{padding:"12px 16px"}}>
            {editCat?.id===c.id?
              <div style={{display:"flex",flexDirection:"column",gap:10}}>
                <div style={{display:"grid",gridTemplateColumns:"90px 1fr",gap:10}}>
                  <div><Lbl>Codice</Lbl><Inp value={editCat.id} onChange={()=>{}} readOnly style={{fontFamily:"'DM Mono',monospace"}}/></div>
                  <div><Lbl>Nome</Lbl><Inp value={editCat.name} onChange={v=>setEditCat(e=>({...e,name:v}))}/></div>
                </div>
                <div><Lbl>Descrizione</Lbl><Inp value={editCat.description||""} onChange={v=>setEditCat(e=>({...e,description:v}))} placeholder="Descrizione per committente e fornitori"/></div>
                <div><Lbl>Colore</Lbl>
                  <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
                    {CAT_COLORS.map(col=><button key={col} onClick={()=>setEditCat(e=>({...e,color:col}))} style={{width:22,height:22,borderRadius:"50%",background:col,border:editCat.color===col?"3px solid #333":"2px solid transparent",cursor:"pointer"}}/>)}
                  </div>
                </div>
                <div style={{display:"flex",gap:8}}><Btn size="sm" onClick={saveCat}>✓ Salva</Btn><Btn size="sm" variant="ghost" onClick={()=>setEditCat(null)}>Annulla</Btn></div>
              </div>:
              <div style={{display:"flex",alignItems:"center",gap:12}}>
                <span style={{width:12,height:12,borderRadius:"50%",background:c.color,flexShrink:0}}/>
                <span style={{fontFamily:"'DM Mono',monospace",fontSize:12,fontWeight:700,color:c.color,background:c.color+"18",padding:"2px 8px",borderRadius:4,flexShrink:0}}>{c.id}</span>
                <div style={{flex:1}}>
                  <div style={{fontSize:13,fontWeight:600,color:C.text}}>{c.name}</div>
                  {c.description&&<div style={{fontSize:11,color:C.muted,marginTop:1}}>{c.description}</div>}
                </div>
                <Btn size="xs" variant="ghost" onClick={()=>setEditCat({...c})}>✏</Btn>
                <Btn size="xs" variant="ghost" onClick={()=>delCat(c.id)} style={{color:"#DC2626"}}>✕</Btn>
              </div>}
          </Card>
        ))}
      </div>
      <Card style={{padding:20}}>
        <div style={{fontSize:13,fontWeight:600,color:C.text,marginBottom:12}}>Aggiungi nuova categoria</div>
        <div style={{display:"flex",flexDirection:"column",gap:10}}>
          <div style={{display:"grid",gridTemplateColumns:"90px 1fr",gap:10}}>
            <div><Lbl>Codice (max 5)</Lbl><Inp value={newCat.id} onChange={v=>setNewCat(c=>({...c,id:v.slice(0,5).toUpperCase()}))} placeholder="COS" style={{fontFamily:"'DM Mono',monospace"}}/></div>
            <div><Lbl>Nome categoria</Lbl><Inp value={newCat.name} onChange={v=>setNewCat(c=>({...c,name:v}))} placeholder="Es. Cosmetici Bagno"/></div>
          </div>
          <div><Lbl>Descrizione</Lbl><Inp value={newCat.description} onChange={v=>setNewCat(c=>({...c,description:v}))} placeholder="Descrizione visibile a committente e fornitori"/></div>
          <div><Lbl>Colore</Lbl>
            <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
              {CAT_COLORS.map(col=><button key={col} onClick={()=>setNewCat(c=>({...c,color:col}))} style={{width:22,height:22,borderRadius:"50%",background:col,border:newCat.color===col?"3px solid #333":"2px solid transparent",cursor:"pointer"}}/>)}
            </div>
          </div>
          <Btn size="sm" onClick={addCat} disabled={!newCat.id.trim()||!newCat.name.trim()}>+ Aggiungi categoria</Btn>
        </div>
      </Card>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════════════
// MAIN APP
// ═══════════════════════════════════════════════════════════════════════
function App(){
  const [screen,setScreen]=useState("projects"); // "projects" | "auth" | "main"
  const [projectMeta,setProjectMeta]=useState(null); // {id,name} lightweight
  const [project,setProject]=useState(null);
  const [role,setRole]=useState(null);
  const [supplierName,setSupplierName]=useState("");
  const [view,setView]=useState("dashboard");
  const [modal,setModal]=useState(null);
  const [filters,setFilters]=useState({search:"",filterZone:"",filterCat:"",filterStatus:""});

  const updateProject=useCallback(async upd=>{
    setProject(upd); await saveProject(upd);
  },[]);

  const renameIndex=useCallback(async(id,name)=>{
    const idx=await loadIndex();
    const updated=idx.map(p=>p.id===id?{...p,name}:p);
    await saveIndex(updated);
  },[]);

  const updateItem=useCallback(item=>{
    const prev=project.items.find(i=>i.id===item.id);
    let updatedItems=project.items.map(i=>i.id===item.id?item:i);
    if(prev){
      const target=normDesc(item.description);
      if(target){
        const sync={};
        if(prev.catId!==item.catId)sync.catId=item.catId;
        if(Number(prev.unitPrice||0)!==Number(item.unitPrice||0))sync.unitPrice=item.unitPrice;
        if(Number(prev.unitPriceInstall||0)!==Number(item.unitPriceInstall||0))sync.unitPriceInstall=item.unitPriceInstall;
        const keys=Object.keys(sync);
        if(keys.length>0){
          const matches=updatedItems.filter(i=>i.id!==item.id&&normDesc(i.description)===target);
          if(matches.length>0){
            const labels={catId:"categoria",unitPrice:"prezzo fornitura",unitPriceInstall:"prezzo posa"};
            const fieldList=keys.map(k=>labels[k]).join(", ");
            if(window.confirm(`Trovate ${matches.length} altre voci con la stessa descrizione "${item.description}".\n\nAggiornare anche quelle (${fieldList})?`)){
              updatedItems=updatedItems.map(i=>(i.id!==item.id&&normDesc(i.description)===target)?{...i,...sync}:i);
            }
          }
        }
      }
    }
    updateProject({...project,items:updatedItems});
  },[project,updateProject]);
  const addItem=useCallback(item=>updateProject({...project,items:[...project.items,{...item,id:uid(),created:new Date().toISOString().slice(0,10)}]}),[project,updateProject]);
  const deleteItem=useCallback(id=>updateProject({...project,items:project.items.filter(i=>i.id!==id)}),[project,updateProject]);

  const exportExcel=useCallback(()=>{
    const perm=ROLES[role];
    const src=perm.canOffer&&supplierName?project.items.filter(i=>i.referenceSupplier?.toLowerCase()===supplierName.toLowerCase()):project.items;
    const customKeys=[...new Set(src.flatMap(i=>Object.keys(i.customFields||{})))];
    const rows=src.map(i=>({
      "Codice":i.code||"","Descrizione":i.description||"",
      "Piano":project.floors.find(f=>f.id===i.floorId)?.name||"",
      "Zona":project.zones.find(z=>z.id===i.zoneId)?.name||i.zoneId,
      "Ambiente":i.ambiente||"",
      "Categoria":project.categories.find(c=>c.id===i.catId)?.name||i.catId,
      "Fornitore ref.":i.referenceSupplier||"","Cod. Fornitore":i.supplierCode||"",
      "Link prodotto":i.webLink||"","Quantità":i.qty||0,"U.M.":i.unit||"",
      ...(perm.canSeePrice?{"Prezzo fornitura (€)":i.unitPrice||0,"Prezzo posa (€)":i.unitPriceInstall||0,"Totale (€)":getTotal(i)}:{}),
      "N. Offerte":i.offers?.length||0,
      ...(perm.canSeePrice?{"Offerta sel.":i.offers?.find(o=>o.isSelected)?.supplierName||"","Prezzo migliore":i.offers?.length?Math.min(...i.offers.map(o=>o.unitPrice)):""}:{}),
      "Stato":STATUS[i.status]?.label||i.status,"Note":i.notes||"","Creato":i.created||"",
      ...Object.fromEntries(customKeys.map(k=>[`★ ${k}`,i.customFields?.[k]||""])),
    }));
    const ws=XLSX.utils.json_to_sheet(rows);
    const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"Abaco");
    XLSX.writeFile(wb,`Abaco_${project.name.replace(/\s+/g,"_")}.xlsx`);
  },[project,role,supplierName]);

  const pendingOffers=useMemo(()=>project?project.items.reduce((s,i)=>s+(i.offers?.filter(o=>!o.isSelected).length||0),0):0,[project]);
  const perm=role?ROLES[role]:null;

  if(screen==="projects") return <ProjectSelector onSelect={proj=>{setProjectMeta({id:proj.id,name:proj.name});setProject(proj);setScreen("auth");}}/>;
  if(screen==="auth") return <AuthScreen project={project} onAuth={(r,sn)=>{setRole(r);setSupplierName(sn);setView(r==="fornitore"?"portal":"dashboard");setScreen("main");}} onBack={()=>setScreen("projects")}/>;

  return (
    <div style={{display:"flex",height:"100vh",fontFamily:"'DM Sans',sans-serif",background:C.bg,color:C.text,overflow:"hidden"}}>
      <link href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:ital,wght@0,400;0,600;1,400&family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet"/>

      <Sidebar view={view} setView={setView} role={role} supplierName={supplierName}
        projectName={project.name}
        onSwitchRole={()=>setScreen("auth")}
        onSwitchProject={()=>{setScreen("projects");setRole(null);setSupplierName("");}}
        pendingOffers={pendingOffers}/>

      <div style={{flex:1,overflow:"auto",display:"flex",flexDirection:"column"}}>
        {/* Topbar */}
        <div style={{background:"#FFF",borderBottom:`1px solid ${C.border}`,padding:"11px 28px",display:"flex",justifyContent:"space-between",alignItems:"center",flexShrink:0}}>
          <div style={{display:"flex",alignItems:"center",gap:8}}>
            <span style={{fontFamily:"'DM Mono',monospace",fontSize:11,color:C.gold}}>{project.name}</span>
            <span style={{color:C.border}}>·</span>
            <span style={{fontSize:13,color:C.muted}}>{{dashboard:"Panoramica",abaco:"Abaco Forniture",zone:"Vista per Zone",portal:"Portale Fornitore",settings:"Impostazioni"}[view]}</span>
          </div>
          <div style={{display:"flex",alignItems:"center",gap:10}}>
            {pendingOffers>0&&role!=="fornitore"&&<div style={{display:"flex",alignItems:"center",gap:5,background:"#FEF3C7",padding:"4px 10px",borderRadius:99}}>
              <span style={{width:6,height:6,borderRadius:"50%",background:"#F59E0B",display:"inline-block"}}/>
              <span style={{fontSize:12,color:"#B45309",fontWeight:500}}>{pendingOffers} offerte da valutare</span>
            </div>}
            <div style={{display:"flex",alignItems:"center",gap:5,background:perm.color+"15",padding:"4px 10px",borderRadius:99}}>
              <span style={{fontSize:12}}>{perm.icon}</span>
              <span style={{fontSize:12,color:perm.color,fontWeight:500}}>{supplierName||perm.label}</span>
            </div>
          </div>
        </div>

        {/* Main content */}
        <main style={{flex:1,padding:"24px 28px",overflow:"auto"}}>
          {view==="dashboard" &&<Dashboard project={project} role={role}/>}
         {view==="abaco"     &&<AbacoTable project={project} role={role} supplierName={supplierName}
            onEdit={item=>setModal(item)} onDelete={perm.canDelete?deleteItem:null}
            onAdd={perm.canAdd?()=>setModal({}):null} onExport={exportExcel}
            onImport={imported=>updateProject({...project,items:[...project.items,...imported]})}
            filters={filters} setFilters={setFilters}/>}
          {view==="zone"      &&<ZoneView project={project} role={role}/>}
          {view==="portal"    &&<PortaleFornitore project={project} supplierName={supplierName} onEdit={(item)=>setModal(item)}/>}
          {view==="settings"  &&<SettingsView project={project} onUpdate={updateProject} onRenameIndex={renameIndex}/>}
        </main>
      </div>

      {modal!==null&&<ItemModal item={modal} project={project} role={role} supplierName={supplierName}
        onSave={item=>{ if(item.id)updateItem(item); else addItem(item); setModal(null); }}
        onClose={()=>setModal(null)}/>}
    </div>
  );
}

ReactDOM.createRoot(document.getElementById('root')).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
)
