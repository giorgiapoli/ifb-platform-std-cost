// @ts-nocheck
import React, { useState, useMemo, useCallback, useRef, useEffect } from "react";
import * as XLSX from "xlsx";

// ─── GLOBAL CSS RESET ─────────────────────────────────────────────────────────
const injectCSS = () => {
  const id = "ifb-global";
  if (document.getElementById(id)) return;
  const s = document.createElement("style");
  s.id = id;
  s.textContent = `*,*::before,*::after{box-sizing:border-box}html,body{margin:0;padding:0;background:#0f1117}::-webkit-scrollbar{width:6px;height:6px}::-webkit-scrollbar-track{background:#1a1d27}::-webkit-scrollbar-thumb{background:#2a2d3e;border-radius:3px}`;
  document.head.appendChild(s);
};

// ─── THEME ────────────────────────────────────────────────────────────────────
const C = {
  bg:"#0f1117", card:"#1a1d27", border:"#2a2d3e", accent:"#4f8ef7",
  green:"#22c55e", red:"#ef4444", yellow:"#f59e0b", purple:"#a855f7",
  text:"#e2e8f0", muted:"#64748b", white:"#ffffff",
};
const inputStyle = (w="100%") => ({
  width:w, background:"#111827", border:`1px solid ${C.border}`,
  color:C.text, padding:"6px 10px", borderRadius:6, fontSize:13,
  outline:"none", boxSizing:"border-box",
});

// ─── LOCAL STORAGE ────────────────────────────────────────────────────────────
const LS = {
  get:(k,d)=>{try{const v=localStorage.getItem(k);return v?JSON.parse(v):d;}catch{return d;}},
  set:(k,v)=>{try{localStorage.setItem(k,JSON.stringify(v));}catch{}},
};

// ─── SEED DATA (vuoti — tutto dai file importati) ─────────────────────────────
const SEED_PRODUCTS = [];
const SEED_LOGISTIC = [];
const SEED_PRICES   = [];
const SEED_FX = [
  {branch:"HK",  month:"2026-06", rate:9.1437},
  {branch:"CAN", month:"2026-06", rate:1.0},
  {branch:"AUS", month:"2026-06", rate:1.6420},
];

// ─── FIELD ALIASES ────────────────────────────────────────────────────────────
const BC_FIELD_ALIASES = {
  nHK:         ["n hk","nhk","hk code","hk no","n_hk","codice hk","gc code","gc no","hong kong no"],
  code:        ["no_","no.","no","item no.","item no","ifb no","ifb n","codice","code"],
  description: ["description","descrizione","desc","item description"],
  vendorName:  ["vendor name 3","vendor name","vendor","fornitore","vendor name 2","vendor name 1"],
  category:    ["category","categoria","cat","item category","product group","gruppo"],
  uom:         ["uom","unit","unit of measure","unità","um","base unit"],
  qtyPerBox:   ["qty/box","qty per box","qtà/box","qty box","pcs/box","quantity per box","pz/box"],
  boxPerPallet:["box/plt","box per pallet","boxes/pallet","scatole/pallet","box per plt"],
  pltPerContainer:["plt/cont","plt per container","pallets/container","pallet/cont","plt cont"],
  netWeight:   ["net weight","peso netto","net wt","peso net","weight net"],
  grossWeight: ["gross weight","peso lordo","gross wt","peso gross"],
  volume:      ["volume","vol","cbm","cubic","m3"],
};

const PRICE_FIELD_ALIASES = {
  code:         ["item no","item no.","no_","codice","code","ifb no","ifb n","no"],
  nHK:          ["n hk","nhk","hk code","hong kong no","gc code"],
  description:  ["description","descrizione","desc"],
  branch:       ["branch","filiale","paese","country"],
  month:        ["month","mese","period","periodo","competenza"],
  priceType:    ["price type","tipo prezzo","type","tipo","price list type"],
  fcaGross:     ["fca gross","fca lordo","fca price","prezzo fca","fca"],
  discountPct:  ["discount %","sconto %","disc %","discount","sconto","disc"],
  fcaDiscounted:["fca net","fca netto","fca discounted","fca scontato","net fca"],
  dapGross:     ["dap gross","dap lordo","dap price","prezzo dap","dap"],
  dapDiscount:  ["dap discount","sconto dap","dap disc"],
  dapFinal:     ["dap final","dap netto","dap net","dap finale"],
  mtsPrice:     ["mts price","prezzo mts","mts","made to stock price"],
};

// ─── COST ENGINE ──────────────────────────────────────────────────────────────
const CERT_COST = 185;
const VGM_COST  = 22;
const HC_COST   = 110;

function selectPrice(p, ubicazione) {
  if (!p) return 0;
  if (ubicazione==="FOR") return p.fcaDiscounted||0;
  if (ubicazione==="MTO") return p.dapFinal||0;
  return p.mtsPrice&&p.mtsPrice!==0 ? p.mtsPrice : (p.dapFinal||0);
}

function calcHK({priceInput,ubicazione,product:p,logistic:l,eurToHkd}) {
  if (!priceInput||priceInput<=0) return null;
  if (!p.qtyPerBox||!p.boxPerPallet||!l.pltPerContainer) return null;
  const conv = l.convFactor||1;
  const unitsPerContainer = p.qtyPerBox * p.boxPerPallet * l.pltPerContainer;
  if (!unitsPerContainer) return null;
  const priceEur = priceInput / conv;
  const fob = ubicazione==="FOR" ? priceEur*0.015 : 0;
  const lic = l.hasCert ? CERT_COST/unitsPerContainer : 0;
  const vgm = VGM_COST / unitsPerContainer;
  const hc  = HC_COST  / unitsPerContainer;
  const plt = 0;
  const alc = l.hasAlcTax ? (l.alcTax||0)/unitsPerContainer : 0;
  const carriageUnit = (l.carriage||0) / unitsPerContainer;
  const step1Eur = priceEur + fob + lic + vgm + hc + plt + alc + carriageUnit;
  const step1Hkd = step1Eur * eurToHkd;
  const wh = step1Hkd * 0.005;
  const step2Eur = step1Eur;
  const step2Hkd = step1Hkd + wh;
  return {priceEur,fob,lic,vgm,hc,plt,alc,carriageUnit,step1Eur,step1Hkd,wh,step2Eur,step2Hkd,rate:eurToHkd};
}

// ─── HELPERS ─────────────────────────────────────────────────────────────────
function fmt2(n){if(n==null||isNaN(n))return "—";return Number(n).toFixed(2);}
function fmtPct(n){if(n==null||isNaN(n))return "—";return (Number(n)*100).toFixed(1)+"%";}
function monthLabel(m){
  if(!m)return "—";
  const [y,mo]=m.split("-");
  const names=["","Gen","Feb","Mar","Apr","Mag","Giu","Lug","Ago","Set","Ott","Nov","Dic"];
  return `${names[parseInt(mo)||0]} ${y}`;
}
function readXlsx(file){
  return new Promise((res,rej)=>{
    const r=new FileReader();
    r.onload=e=>{
      try{
        const wb=XLSX.read(e.target.result,{type:"array"});
        const ws=wb.Sheets[wb.SheetNames[0]];
        res(XLSX.utils.sheet_to_json(ws,{header:1,defval:""}));
      }catch(err){rej(err);}
    };
    r.onerror=rej;
    r.readAsArrayBuffer(file);
  });
}
function normalize(s){return String(s||"").toLowerCase().trim().replace(/\s+/g," ");}
function findAlias(headers,aliases){
  for(let i=0;i<headers.length;i++){if(aliases.includes(normalize(headers[i])))return i;}
  return -1;
}
function getCol(row,idx,fallback=""){return idx>=0&&idx<row.length?row[idx]:fallback;}

// ─── SHARED COMPONENTS ───────────────────────────────────────────────────────
function PageHeader({title,sub}){
  return(
    <div style={{marginBottom:20}}>
      <div style={{fontSize:22,fontWeight:700,color:C.white}}>{title}</div>
      {sub&&<div style={{fontSize:13,color:C.muted,marginTop:3}}>{sub}</div>}
    </div>
  );
}
function Section({title,children}){
  return(
    <div style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:10,padding:18,marginBottom:16}}>
      {title&&<div style={{fontSize:12,fontWeight:600,color:C.muted,marginBottom:12,textTransform:"uppercase",letterSpacing:1}}>{title}</div>}
      {children}
    </div>
  );
}
function KPI({label,value,color}){
  return(
    <div style={{background:"#111827",borderRadius:8,padding:"12px 16px",minWidth:120}}>
      <div style={{fontSize:11,color:C.muted,marginBottom:4}}>{label}</div>
      <div style={{fontSize:22,fontWeight:700,color:color||C.white}}>{value}</div>
    </div>
  );
}
function THead({cols}){
  return(
    <thead>
      <tr>{cols.map(c=>(
        <th key={c} style={{padding:"6px 10px",textAlign:"left",fontSize:11,color:C.muted,
          borderBottom:`1px solid ${C.border}`,whiteSpace:"nowrap"}}>{c}</th>
      ))}</tr>
    </thead>
  );
}
function TD({children,right,mono}){
  return(
    <td style={{padding:"5px 10px",fontSize:12,color:C.text,borderBottom:`1px solid #1e2130`,
      textAlign:right?"right":"left",fontFamily:mono?"monospace":"inherit",whiteSpace:"nowrap"}}>
      {children}
    </td>
  );
}
function Chip({label,color}){
  return <span style={{background:color||C.border,color:C.white,fontSize:10,padding:"2px 6px",borderRadius:4,fontWeight:600}}>{label}</span>;
}
function UbicChip({u}){
  const col=u==="MTO"?C.accent:u==="FOR"?C.green:u==="MTS"?C.yellow:C.muted;
  return <Chip label={u} color={col}/>;
}
function DeltaBadge({delta}){
  const pct=delta*100;
  const col=pct>0?C.green:pct<0?C.red:C.muted;
  return <span style={{color:col,fontSize:11,fontWeight:600}}>{pct>0?"+":""}{pct.toFixed(1)}%</span>;
}
function ActionBtn({label,onClick,color,small}){
  return(
    <button onClick={onClick} style={{
      background:color||C.accent,color:C.white,border:"none",borderRadius:6,
      padding:small?"4px 10px":"7px 16px",fontSize:small?11:13,fontWeight:600,
      cursor:"pointer",whiteSpace:"nowrap",
    }}>{label}</button>
  );
}
function MiniBtn({label,onClick,active}){
  return(
    <button onClick={onClick} style={{
      background:active?C.accent:"#1e2130",color:active?C.white:C.muted,
      border:`1px solid ${active?C.accent:C.border}`,borderRadius:5,
      padding:"3px 9px",fontSize:11,cursor:"pointer",
    }}>{label}</button>
  );
}
function SearchBar({value,onChange,placeholder}){
  return(
    <input value={value} onChange={e=>onChange(e.target.value)}
      placeholder={placeholder||"Cerca..."}
      style={{...inputStyle("260px"),height:32}}/>
  );
}
function DropZone({onFile,label}){
  const ref=useRef(null);
  return(
    <div onClick={()=>ref.current?.click()} style={{
      border:`2px dashed ${C.border}`,borderRadius:8,padding:"24px",
      textAlign:"center",cursor:"pointer",color:C.muted,fontSize:13,
    }}
    onMouseEnter={e=>e.currentTarget.style.borderColor=C.accent}
    onMouseLeave={e=>e.currentTarget.style.borderColor=C.border}>
      <div style={{fontSize:28,marginBottom:8}}>📂</div>
      <div>{label||"Clicca o trascina file Excel/CSV"}</div>
      <input ref={ref} type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}}
        onChange={e=>{if(e.target.files?.[0])onFile(e.target.files[0]);e.target.value="";}}/>
    </div>
  );
}
function EmptyState({message,btnLabel,onBtn}){
  return(
    <div style={{textAlign:"center",padding:"48px 24px",color:C.muted}}>
      <div style={{fontSize:40,marginBottom:12}}>📭</div>
      <div style={{marginBottom:16,fontSize:14}}>{message}</div>
      {btnLabel&&<ActionBtn label={btnLabel} onClick={onBtn}/>}
    </div>
  );
}
function StepBar({steps,current}){
  return(
    <div style={{display:"flex",gap:8,marginBottom:20,alignItems:"center"}}>
      {steps.map((s,i)=>(
        <React.Fragment key={s}>
          <div style={{
            padding:"4px 14px",borderRadius:20,fontSize:12,fontWeight:600,
            background:i===current?C.accent:i<current?"#1a3a1a":C.card,
            color:i<=current?C.white:C.muted,border:`1px solid ${i===current?C.accent:C.border}`
          }}>{i+1}. {s}</div>
          {i<steps.length-1&&<div style={{color:C.border,fontSize:16}}>›</div>}
        </React.Fragment>
      ))}
    </div>
  );
}
function FormField({label,children}){
  return(
    <div style={{marginBottom:12}}>
      <label style={{display:"block",fontSize:12,color:C.muted,marginBottom:4}}>{label}</label>
      {children}
    </div>
  );
}
function SelectField({label,value,onChange,options}){
  return(
    <FormField label={label}>
      <select value={value} onChange={e=>onChange(e.target.value)} style={inputStyle()}>
        {options.map(o=><option key={o.value} value={o.value}>{o.label}</option>)}
      </select>
    </FormField>
  );
}
function CheckBox({label,checked,onChange}){
  return(
    <label style={{display:"flex",alignItems:"center",gap:8,cursor:"pointer",fontSize:13,color:C.text}}>
      <input type="checkbox" checked={checked} onChange={e=>onChange(e.target.checked)}
        style={{accentColor:C.accent,width:15,height:15}}/>
      {label}
    </label>
  );
}

// ─── APP ROOT ─────────────────────────────────────────────────────────────────
const PAGES = ["Dashboard","Anagrafica","Logistica","Listini","FX","Costi","Mail","Fattura","Storico","XRef","Note","Importa BC","Importa Prezzi"];

export default function App() {
  useEffect(()=>injectCSS(),[]);

  const [page,  setPage]         = useState(()=>LS.get("page","Dashboard"));
  const [products,  setProductsRaw]  = useState(()=>LS.get("products",SEED_PRODUCTS));
  const [logistics, setLogisticsRaw] = useState(()=>LS.get("logistics",SEED_LOGISTIC));
  const [prices,    setPricesRaw]    = useState(()=>LS.get("prices",SEED_PRICES));
  const [fxRates,   setFxRatesRaw]   = useState(()=>LS.get("fxRates",SEED_FX));
  const [xrefs,     setXrefsRaw]     = useState(()=>LS.get("xrefs",[]));
  const [notes,     setNotesRaw]     = useState(()=>LS.get("notes",""));
  const [branch,    setBranch]       = useState(()=>LS.get("branch","HK"));
  const [month,     setMonth]        = useState(()=>LS.get("month","2026-06"));

  const setProducts  = useCallback(v=>{LS.set("products",v); setProductsRaw(v);},[]);
  const setLogistics = useCallback(v=>{LS.set("logistics",v);setLogisticsRaw(v);},[]);
  const setPrices    = useCallback(v=>{LS.set("prices",v);   setPricesRaw(v);},[]);
  const setFxRates   = useCallback(v=>{LS.set("fxRates",v);  setFxRatesRaw(v);},[]);
  const setXrefs     = useCallback(v=>{LS.set("xrefs",v);    setXrefsRaw(v);},[]);
  const setNotes     = useCallback(v=>{LS.set("notes",v);    setNotesRaw(v);},[]);
  const changeBranch = v=>{LS.set("branch",v);setBranch(v);};
  const changeMonth  = v=>{LS.set("month",v); setMonth(v);};
  const nav          = p=>{LS.set("page",p);  setPage(p);};

  const costRows = useMemo(()=>{
    const fxEntry = fxRates.find(f=>f.branch===branch&&f.month===month);
    const fxRate  = fxEntry?.rate||1;
    return products.map(prod=>{
      const log = logistics.find(l=>l.productId===prod.id&&l.branch===branch)||null;
      const pr  = prices.find(p=>p.productId===prod.id&&p.branch===branch&&p.month===month)||null;
      const effectiveLog = log
        ? {...log, category:prod.category}
        : {productId:prod.id,branch,area:"NORD",ubicazione:"MTO",
           pltPerContainer:prod.pltPerContainer||20,
           hasCert:false,hasAlcTax:false,alcTax:0,convFactor:1,carriage:0,
           category:prod.category};
      if(!pr) return {...prod,cost:null,prevCost:null,log:effectiveLog,pr,delta:null,
        skipReason:log?`NO PREZZO (${branch}/${month})`:`NO LOG+PREZZO`};
      const priceInput = selectPrice(pr,effectiveLog.ubicazione);
      const cost = calcHK({priceInput,ubicazione:effectiveLog.ubicazione,product:prod,logistic:effectiveLog,eurToHkd:fxRate});
      if(!cost) return {...prod,cost:null,prevCost:null,log:effectiveLog,pr,delta:null,
        skipReason:`CALC NULL (qty=${prod.qtyPerBox} box=${prod.boxPerPallet} plt=${effectiveLog.pltPerContainer})`};
      return {...prod,cost,prevCost:null,log:effectiveLog,pr,delta:null,skipReason:null};
    });
  },[products,logistics,prices,fxRates,branch,month]);

  const ctx={products,setProducts,logistics,setLogistics,prices,setPrices,
    fxRates,setFxRates,xrefs,setXrefs,notes,setNotes,
    branch,changeBranch,month,changeMonth,costRows,nav};

  return(
    <div style={{minHeight:"100vh",background:C.bg,color:C.text,fontFamily:"'Inter',system-ui,sans-serif"}}>
      <TopBar page={page} setPage={nav} branch={branch} changeBranch={changeBranch} month={month} changeMonth={changeMonth}/>
      <div style={{maxWidth:1400,margin:"0 auto",padding:"20px 16px"}}>
        {page==="Dashboard"     && <Dashboard     {...ctx}/>}
        {page==="Anagrafica"    && <ProductsPage  {...ctx}/>}
        {page==="Logistica"     && <LogisticsPage {...ctx}/>}
        {page==="Listini"       && <PricesPage    {...ctx}/>}
        {page==="FX"            && <FxRatesPage   {...ctx}/>}
        {page==="Costi"         && <CostTable     {...ctx}/>}
        {page==="Mail"          && <MailGen       {...ctx}/>}
        {page==="Fattura"       && <SalesInvoice  {...ctx}/>}
        {page==="Storico"       && <Storico       {...ctx}/>}
        {page==="XRef"          && <XRefPage      {...ctx}/>}
        {page==="Note"          && <NotesPage     {...ctx}/>}
        {page==="Importa BC"    && <ImportBC      {...ctx}/>}
        {page==="Importa Prezzi"&& <ImportPrices  {...ctx}/>}
      </div>
    </div>
  );
}

// ─── TOP BAR ─────────────────────────────────────────────────────────────────
function TopBar({page,setPage,branch,changeBranch,month,changeMonth}){
  const tabs=["Dashboard","Anagrafica","Logistica","Listini","FX","Costi","Mail","Fattura","Storico","XRef","Note","Importa BC","Importa Prezzi"];
  return(
    <div style={{background:C.card,borderBottom:`1px solid ${C.border}`,position:"sticky",top:0,zIndex:100}}>
      <div style={{maxWidth:1400,margin:"0 auto",padding:"0 16px"}}>
        <div style={{display:"flex",alignItems:"center",gap:12,height:50}}>
          <div style={{fontWeight:800,fontSize:14,color:C.white,whiteSpace:"nowrap"}}>IFB Cost Intelligence</div>
          <div style={{flex:1,overflowX:"auto",display:"flex",gap:2,scrollbarWidth:"none"}}>
            {tabs.map(t=>(
              <button key={t} onClick={()=>setPage(t)} style={{
                padding:"4px 11px",fontSize:11,borderRadius:5,border:"none",cursor:"pointer",whiteSpace:"nowrap",
                background:page===t?C.accent:"transparent",color:page===t?C.white:C.muted,
                fontWeight:page===t?700:400,
              }}>{t}</button>
            ))}
          </div>
          <div style={{display:"flex",gap:8,alignItems:"center",flexShrink:0}}>
            <select value={branch} onChange={e=>changeBranch(e.target.value)} style={{...inputStyle("80px"),height:30,fontSize:12}}>
              {["HK","CAN","AUS"].map(b=><option key={b} value={b}>{b}</option>)}
            </select>
            <input type="month" value={month} onChange={e=>changeMonth(e.target.value)}
              style={{...inputStyle("130px"),height:30,fontSize:12}}/>
          </div>
        </div>
      </div>
    </div>
  );
}

// ─── DASHBOARD ────────────────────────────────────────────────────────────────
function Dashboard({products,logistics,prices,fxRates,costRows,branch,month,nav}){
  const calculated=costRows.filter(r=>r.cost!==null).length;
  const skipped   =costRows.filter(r=>r.cost===null).length;
  const fxEntry   =fxRates.find(f=>f.branch===branch&&f.month===month);
  const logBranch =logistics.filter(l=>l.branch===branch).length;
  const priceBranch=prices.filter(p=>p.branch===branch&&p.month===month).length;
  const hasData   =products.length>0;
  return(
    <div>
      <PageHeader title={`Dashboard — ${branch} ${monthLabel(month)}`}/>
      {!hasData?(
        <Section>
          <EmptyState message="Nessun dato caricato. Inizia importando l'anagrafica prodotti da Business Central." btnLabel="📥 Importa Anagrafica BC" onBtn={()=>nav("Importa BC")}/>
        </Section>
      ):(
        <>
          <div style={{display:"flex",gap:12,flexWrap:"wrap",marginBottom:20}}>
            <KPI label="Prodotti" value={products.length}/>
            <KPI label="Calcolati" value={calculated} color={C.green}/>
            <KPI label="Saltati" value={skipped} color={skipped>0?C.red:C.muted}/>
            <KPI label={`Logistica ${branch}`} value={logBranch}/>
            <KPI label={`Prezzi ${branch}/${monthLabel(month)}`} value={priceBranch}/>
            <KPI label={`FX EUR→${branch}`} value={fxEntry?fxEntry.rate.toFixed(4):"N/D"} color={C.yellow}/>
          </div>
          <div style={{display:"flex",gap:12,flexWrap:"wrap",marginBottom:20}}>
            {[
              {label:"Importa Anagrafica BC",icon:"📥",page:"Importa BC"},
              {label:"Importa Prezzi",icon:"💰",page:"Importa Prezzi"},
              {label:"Logistica",icon:"🚢",page:"Logistica"},
              {label:"Visualizza Costi",icon:"📊",page:"Costi"},
              {label:"Genera Mail",icon:"✉️",page:"Mail"},
            ].map(x=>(
              <button key={x.page} onClick={()=>nav(x.page)} style={{
                background:C.card,border:`1px solid ${C.border}`,borderRadius:10,
                padding:"16px 20px",cursor:"pointer",color:C.text,fontSize:13,
                display:"flex",flexDirection:"column",alignItems:"center",gap:8,minWidth:120,
              }}>
                <span style={{fontSize:24}}>{x.icon}</span>{x.label}
              </button>
            ))}
          </div>
          {skipped>0&&(
            <Section title="Prodotti non calcolati (campione)">
              <table style={{width:"100%",borderCollapse:"collapse"}}>
                <THead cols={["Codice","Descrizione","Motivo"]}/>
                <tbody>
                  {costRows.filter(r=>r.cost===null).slice(0,8).map(r=>(
                    <tr key={r.id}><TD mono>{r.code}</TD><TD>{r.description}</TD><TD><Chip label={r.skipReason||"?"} color={C.red}/></TD></tr>
                  ))}
                </tbody>
              </table>
            </Section>
          )}
        </>
      )}
    </div>
  );
}

// ─── PRODUCTS PAGE ────────────────────────────────────────────────────────────
function ProductsPage({products,setProducts,nav}){
  const [search,setSearch]=useState("");
  const [onlyIFB,setOnlyIFB]=useState(false);
  const filtered=products.filter(p=>{
    if(onlyIFB&&!String(p.vendorName||"").toUpperCase().includes("INALCA FOOD"))return false;
    if(!search)return true;
    const q=search.toLowerCase();
    return p.code?.toLowerCase().includes(q)||p.description?.toLowerCase().includes(q)||p.nHK?.toLowerCase().includes(q);
  });
  if(products.length===0){
    return(
      <div><PageHeader title="Anagrafica Prodotti"/>
        <Section><EmptyState message="Nessun prodotto. Importa l'anagrafica da Business Central." btnLabel="📥 Importa Anagrafica BC" onBtn={()=>nav("Importa BC")}/></Section>
      </div>
    );
  }
  return(
    <div>
      <PageHeader title="Anagrafica Prodotti" sub={`${filtered.length} / ${products.length} prodotti`}/>
      <div style={{display:"flex",gap:10,marginBottom:14,flexWrap:"wrap",alignItems:"center"}}>
        <SearchBar value={search} onChange={setSearch} placeholder="Cerca codice, descrizione, HK..."/>
        <button onClick={()=>setOnlyIFB(v=>!v)} style={{
          background:onlyIFB?C.accent:C.card,color:onlyIFB?C.white:C.muted,
          border:`1px solid ${onlyIFB?C.accent:C.border}`,borderRadius:6,
          padding:"5px 14px",fontSize:12,cursor:"pointer",fontWeight:600,
        }}>{onlyIFB?"✓ Solo IFB":"Solo IFB"}</button>
        <ActionBtn label="Svuota" onClick={()=>{if(confirm("Eliminare tutti i prodotti?"))setProducts([]);}} color={C.red} small/>
      </div>
      <Section>
        <table style={{width:"100%",borderCollapse:"collapse"}}>
          <THead cols={["Codice IFB","N HK","Descrizione","Fornitore","Cat","UOM","Qty/Box","Box/Plt","Plt/Cont"]}/>
          <tbody>
            {filtered.slice(0,300).map(p=>(
              <tr key={p.id}>
                <TD mono>{p.code}</TD><TD mono>{p.nHK||"—"}</TD><TD>{p.description}</TD>
                <TD>{p.vendorName||"—"}</TD><TD>{p.category||"—"}</TD><TD>{p.uom||"—"}</TD>
                <TD right>{p.qtyPerBox||"—"}</TD><TD right>{p.boxPerPallet||"—"}</TD><TD right>{p.pltPerContainer||"—"}</TD>
              </tr>
            ))}
          </tbody>
        </table>
        {filtered.length>300&&<div style={{color:C.muted,fontSize:12,marginTop:8}}>Mostrati 300/{filtered.length}</div>}
      </Section>
    </div>
  );
}

// ─── LOGISTICS PAGE ───────────────────────────────────────────────────────────
function LogisticsPage({products,logistics,setLogistics,branch,nav}){
  const [mapStep,setMapStep]=useState("idle");
  const [logRawRows,setLogRawRows]=useState([]);
  const [importMsg,setImportMsg]=useState("");
  const [editId,setEditId]=useState(null);
  const [form,setForm]=useState({});
  const branchLog=logistics.filter(l=>l.branch===branch);

  async function parseLogFile(file){
    try{
      const rows=await readXlsx(file);
      if(rows.length<2){setImportMsg("File vuoto o non valido.");return;}
      let hi=0;
      for(let i=0;i<Math.min(10,rows.length);i++){
        const r=rows[i].map(normalize);
        if(r.some(c=>c.includes("ubicazione")||c.includes("n hk")||c.includes("nhk"))){hi=i;break;}
      }
      const headers=rows[hi].map(c=>String(c||"").trim());
      const hn=headers.map(normalize);
      window._logIdx={
        iNHK:  hn.findIndex(h=>["n hk","nhk","hk code","gc code","hong kong no"].includes(h)),
        iIFB:  hn.findIndex(h=>["no_","no.","no","item no.","item no","ifb no","codice","code"].includes(h)),
        iUb:   hn.findIndex(h=>["ubicazione","location","ubic","loc"].includes(h)),
        iArea: hn.findIndex(h=>["area","zona","zone"].includes(h)),
        iPlt:  hn.findIndex(h=>["plt/cont","plt per container","pallet/cont","pallets/container","plt cont"].includes(h)),
        iCert: hn.findIndex(h=>["cert","certificate","certificato","has cert","hascert"].includes(h)),
        iCarr: hn.findIndex(h=>["carriage","trasporto","transport","nolo","freight"].includes(h)),
      };
      setLogRawRows(rows.slice(hi+1));
      setMapStep("ready");
      setImportMsg(`Trovate ${rows.length-hi-1} righe. Pronto per importare in branch: ${branch}`);
    }catch(e){setImportMsg("Errore lettura file: "+String(e));}
  }

  function applyLogFile(){
    const idx=window._logIdx||{};
    const g=(row,i)=>i>=0?row[i]:"";
    let matched=0,skipped=0;
    const updated=[...logistics];
    for(const row of logRawRows){
      if(!row.some(c=>String(c||"").trim()))continue;
      const nhk =String(g(row,idx.iNHK)||"").trim();
      const ifb =String(g(row,idx.iIFB)||"").trim();
      const prod=products.find(p=>(nhk&&(p.nHK===nhk||p.code===nhk))||(ifb&&p.code===ifb));
      if(!prod){skipped++;continue;}
      const ub  =String(g(row,idx.iUb)||"MTO").trim().toUpperCase()||"MTO";
      const area=String(g(row,idx.iArea)||"NORD").trim().toUpperCase()||"NORD";
      const plt =parseFloat(String(g(row,idx.iPlt)||"20"))||20;
      const certRaw=String(g(row,idx.iCert)||"").toLowerCase();
      const hasCert=certRaw==="si"||certRaw==="yes"||certRaw==="1"||certRaw==="true";
      const carriage=parseFloat(String(g(row,idx.iCarr)||"0"))||0;
      const ei=updated.findIndex(l=>l.productId===prod.id&&l.branch===branch);
      const entry={productId:prod.id,branch,area,ubicazione:ub,pltPerContainer:plt,hasCert,hasAlcTax:false,alcTax:0,convFactor:1,carriage};
      if(ei>=0)updated[ei]=entry;else updated.push(entry);
      matched++;
    }
    setLogistics(updated);
    setImportMsg(`✓ Importati: ${matched} | Non trovati: ${skipped}`);
    setMapStep("idle");
  }

  function startEdit(l){setEditId(l.productId);setForm({...l});}
  function saveEdit(){
    if(!editId)return;
    const updated=logistics.filter(l=>!(l.productId===editId&&l.branch===branch));
    updated.push({...form,productId:editId,branch});
    setLogistics(updated);setEditId(null);
  }
  function delLog(pid){
    if(confirm("Eliminare logistica per questo prodotto?"))
      setLogistics(logistics.filter(l=>!(l.productId===pid&&l.branch===branch)));
  }
  const f=k=>({value:String(form[k]??""),onChange:e=>setForm(v=>({...v,[k]:e.target.value}))});

  if(products.length===0){
    return(<div><PageHeader title={`Logistica — ${branch}`}/>
      <Section><EmptyState message="Nessun prodotto in anagrafica. Importa prima i prodotti." btnLabel="📥 Importa Anagrafica BC" onBtn={()=>nav("Importa BC")}/></Section>
    </div>);
  }

  return(
    <div>
      <PageHeader title={`Logistica — ${branch}`} sub={`${branchLog.length} prodotti configurati`}/>
      <Section title="Importa da file">
        {mapStep==="idle"&&(
          <DropZone onFile={parseLogFile} label="Carica file logistica con colonne: N HK / IFB, Ubicazione, Area, Plt/Cont, Cert, Carriage"/>
        )}
        {mapStep==="ready"&&(
          <div>
            <div style={{color:C.green,marginBottom:12,fontSize:13}}>{importMsg}</div>
            <div style={{display:"flex",gap:10}}>
              <ActionBtn label="✓ Applica importazione" onClick={applyLogFile} color={C.green}/>
              <ActionBtn label="Annulla" onClick={()=>{setMapStep("idle");setImportMsg("");}} color={C.red} small/>
            </div>
          </div>
        )}
        {importMsg&&mapStep==="idle"&&<div style={{color:C.muted,fontSize:12,marginTop:8}}>{importMsg}</div>}
      </Section>

      {editId&&(
        <Section title="Modifica logistica">
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(180px,1fr))",gap:12}}>
            <SelectField label="Ubicazione" value={form.ubicazione||"MTO"} onChange={v=>setForm(x=>({...x,ubicazione:v}))}
              options={[{value:"MTO",label:"MTO"},{value:"FOR",label:"FOR"},{value:"MTS",label:"MTS"}]}/>
            <SelectField label="Area" value={form.area||"NORD"} onChange={v=>setForm(x=>({...x,area:v}))}
              options={[{value:"NORD",label:"NORD"},{value:"SUD",label:"SUD"},{value:"OVEST",label:"OVEST"}]}/>
            <FormField label="Plt/Container"><input type="number" {...f("pltPerContainer")} style={inputStyle()}/></FormField>
            <FormField label="Carriage (€)"><input type="number" {...f("carriage")} style={inputStyle()}/></FormField>
            <FormField label="Conv Factor"><input type="number" {...f("convFactor")} style={inputStyle()}/></FormField>
            <FormField label="Alc Tax (€)"><input type="number" {...f("alcTax")} style={inputStyle()}/></FormField>
          </div>
          <div style={{display:"flex",gap:12,marginTop:8}}>
            <CheckBox label="Certificato" checked={!!form.hasCert} onChange={v=>setForm(x=>({...x,hasCert:v}))}/>
            <CheckBox label="Tassa Alcolici" checked={!!form.hasAlcTax} onChange={v=>setForm(x=>({...x,hasAlcTax:v}))}/>
          </div>
          <div style={{display:"flex",gap:8,marginTop:12}}>
            <ActionBtn label="Salva" onClick={saveEdit} color={C.green}/>
            <ActionBtn label="Annulla" onClick={()=>setEditId(null)} color={C.red} small/>
          </div>
        </Section>
      )}

      <Section>
        <table style={{width:"100%",borderCollapse:"collapse"}}>
          <THead cols={["Codice","Descrizione","Ubic","Area","Plt/Cont","Cert","Carriage","Azioni"]}/>
          <tbody>
            {products.map(p=>{
              const l=logistics.find(x=>x.productId===p.id&&x.branch===branch);
              return(
                <tr key={p.id}>
                  <TD mono>{p.code}</TD><TD>{p.description}</TD>
                  <TD>{l?<UbicChip u={l.ubicazione}/>:<Chip label="—" color={C.muted}/>}</TD>
                  <TD>{l?.area||"—"}</TD><TD right>{l?.pltPerContainer||"—"}</TD>
                  <TD>{l?.hasCert?"✓":"—"}</TD><TD right>{l?.carriage?fmt2(l.carriage):"—"}</TD>
                  <TD>
                    <div style={{display:"flex",gap:4}}>
                      <MiniBtn label={l?"✏":"+"} onClick={()=>startEdit(l||{productId:p.id,branch,area:"NORD",ubicazione:"MTO",pltPerContainer:p.pltPerContainer||20,hasCert:false,hasAlcTax:false,alcTax:0,convFactor:1,carriage:0})}/>
                      {l&&<MiniBtn label="✕" onClick={()=>delLog(p.id)}/>}
                    </div>
                  </TD>
                </tr>
              );
            })}
          </tbody>
        </table>
      </Section>
    </div>
  );
}

// ─── PRICES PAGE ──────────────────────────────────────────────────────────────
function PricesPage({products,prices,setPrices,branch,month,nav}){
  const [search,setSearch]=useState("");
  const branchPrices=prices.filter(p=>p.branch===branch&&p.month===month);
  const filtered=branchPrices.filter(p=>{
    if(!search)return true;
    const q=search.toLowerCase();
    const prod=products.find(x=>x.id===p.productId);
    return p.productId.toLowerCase().includes(q)||(prod?.description||"").toLowerCase().includes(q);
  });
  if(branchPrices.length===0){
    return(<div><PageHeader title={`Listini — ${branch} ${monthLabel(month)}`}/>
      <Section><EmptyState message={`Nessun prezzo per ${branch} / ${monthLabel(month)}. Importa il listino.`} btnLabel="📥 Importa Prezzi" onBtn={()=>nav("Importa Prezzi")}/></Section>
    </div>);
  }
  return(
    <div>
      <PageHeader title={`Listini — ${branch} ${monthLabel(month)}`} sub={`${filtered.length} prezzi`}/>
      <div style={{display:"flex",gap:10,marginBottom:14,alignItems:"center"}}>
        <SearchBar value={search} onChange={setSearch}/>
        <ActionBtn label="Svuota branch/mese" onClick={()=>{if(confirm(`Eliminare prezzi ${branch}/${month}?`))setPrices(prices.filter(p=>!(p.branch===branch&&p.month===month)));}} color={C.red} small/>
      </div>
      <Section>
        <table style={{width:"100%",borderCollapse:"collapse"}}>
          <THead cols={["Codice","Descrizione","Tipo","FCA Lordo","Sc%","FCA Netto","DAP","DAP Net","MTS"]}/>
          <tbody>
            {filtered.slice(0,300).map((p,i)=>{
              const prod=products.find(x=>x.id===p.productId);
              return(
                <tr key={i}>
                  <TD mono>{p.productId}</TD><TD>{prod?.description||"—"}</TD>
                  <TD><Chip label={p.priceType||"—"}/></TD>
                  <TD right>{fmt2(p.fcaGross)}</TD><TD right>{fmt2(p.discountPct)}%</TD>
                  <TD right>{fmt2(p.fcaDiscounted)}</TD><TD right>{fmt2(p.dapGross)}</TD>
                  <TD right>{fmt2(p.dapFinal)}</TD><TD right>{fmt2(p.mtsPrice)}</TD>
                </tr>
              );
            })}
          </tbody>
        </table>
        {filtered.length>300&&<div style={{color:C.muted,fontSize:12,marginTop:8}}>Mostrati 300/{filtered.length}</div>}
      </Section>
    </div>
  );
}

// ─── FX RATES ─────────────────────────────────────────────────────────────────
function FxRatesPage({fxRates,setFxRates}){
  const [editIdx,setEditIdx]=useState(null);
  const [form,setForm]=useState({branch:"",month:"",rate:1});
  const [newRow,setNewRow]=useState({branch:"HK",month:"2026-06",rate:1});
  function saveEdit(){if(editIdx===null)return;const u=[...fxRates];u[editIdx]=form;setFxRates(u);setEditIdx(null);}
  function addRow(){setFxRates([...fxRates,{...newRow,rate:parseFloat(String(newRow.rate))||1}]);setNewRow({branch:"HK",month:"2026-06",rate:1});}
  function del(i){const u=[...fxRates];u.splice(i,1);setFxRates(u);}
  return(
    <div>
      <PageHeader title="Tassi di Cambio FX" sub="EUR → valuta branch"/>
      <Section>
        <table style={{width:"100%",borderCollapse:"collapse"}}>
          <THead cols={["Branch","Mese","Tasso EUR→","Azioni"]}/>
          <tbody>
            {fxRates.map((fx,i)=>(
              <tr key={i}>
                {editIdx===i?(
                  <>
                    <TD><input value={form.branch} onChange={e=>setForm(x=>({...x,branch:e.target.value}))} style={inputStyle("70px")}/></TD>
                    <TD><input type="month" value={form.month} onChange={e=>setForm(x=>({...x,month:e.target.value}))} style={inputStyle("130px")}/></TD>
                    <TD><input type="number" step="0.0001" value={form.rate} onChange={e=>setForm(x=>({...x,rate:parseFloat(e.target.value)||1}))} style={inputStyle("100px")}/></TD>
                    <TD><ActionBtn label="Salva" onClick={saveEdit} small/></TD>
                  </>
                ):(
                  <>
                    <TD>{fx.branch}</TD><TD>{fx.month}</TD><TD right>{fx.rate.toFixed(4)}</TD>
                    <TD><div style={{display:"flex",gap:4}}><MiniBtn label="✏" onClick={()=>{setEditIdx(i);setForm({...fx});}}/><MiniBtn label="✕" onClick={()=>del(i)}/></div></TD>
                  </>
                )}
              </tr>
            ))}
            <tr>
              <TD><input value={newRow.branch} onChange={e=>setNewRow(x=>({...x,branch:e.target.value}))} style={inputStyle("70px")}/></TD>
              <TD><input type="month" value={newRow.month} onChange={e=>setNewRow(x=>({...x,month:e.target.value}))} style={inputStyle("130px")}/></TD>
              <TD><input type="number" step="0.0001" value={newRow.rate} onChange={e=>setNewRow(x=>({...x,rate:parseFloat(e.target.value)||1}))} style={inputStyle("100px")}/></TD>
              <TD><ActionBtn label="+ Aggiungi" onClick={addRow} small/></TD>
            </tr>
          </tbody>
        </table>
      </Section>
    </div>
  );
}

// ─── COST TABLE ───────────────────────────────────────────────────────────────
function CostTable({costRows,branch,month,nav}){
  const [search,setSearch]=useState("");
  const [showSkipped,setShowSkipped]=useState(false);
  const [detail,setDetail]=useState(null);
  const rows=costRows.filter(r=>{
    if(!showSkipped&&!r.cost)return false;
    if(!search)return true;
    const q=search.toLowerCase();
    return r.code?.toLowerCase().includes(q)||r.description?.toLowerCase().includes(q);
  });
  function exportXlsx(){
    const data=costRows.filter(r=>r.cost).map(r=>({
      "Codice IFB":r.code,"N HK":r.nHK||"","Descrizione":r.description,
      "Branch":branch,"Mese":month,"Ubicazione":r.log?.ubicazione||"",
      "Prezzo EUR":fmt2(r.cost?.priceEur),"FOB":fmt2(r.cost?.fob),
      "VGM":fmt2(r.cost?.vgm),"HC":fmt2(r.cost?.hc),"Alc":fmt2(r.cost?.alc),
      "Carriage Unit":fmt2(r.cost?.carriageUnit),
      "Step1 EUR":fmt2(r.cost?.step1Eur),"Step1 HKD":fmt2(r.cost?.step1Hkd),
      "WH":fmt2(r.cost?.wh),"Step2 HKD":fmt2(r.cost?.step2Hkd),
    }));
    const ws=XLSX.utils.json_to_sheet(data);
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,"Costi");
    XLSX.writeFile(wb,`SC_${branch}_${month}_${new Date().toISOString().slice(0,10)}.xlsx`);
  }
  if(costRows.length===0){
    return(<div><PageHeader title={`Costi Standard — ${branch} ${monthLabel(month)}`}/>
      <Section><EmptyState message="Nessun dato. Importa anagrafica e prezzi per calcolare i costi." btnLabel="📥 Importa Anagrafica BC" onBtn={()=>nav("Importa BC")}/></Section>
    </div>);
  }
  return(
    <div>
      <PageHeader title={`Costi Standard — ${branch} ${monthLabel(month)}`}
        sub={`${costRows.filter(r=>r.cost).length} calcolati / ${costRows.length} totali`}/>
      <div style={{display:"flex",gap:10,marginBottom:14,flexWrap:"wrap",alignItems:"center"}}>
        <SearchBar value={search} onChange={setSearch}/>
        <MiniBtn label={showSkipped?"Nascondi saltati":"Mostra saltati"} onClick={()=>setShowSkipped(v=>!v)} active={showSkipped}/>
        <ActionBtn label="📥 Esporta Excel" onClick={exportXlsx} small/>
      </div>
      <Section>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse",minWidth:900}}>
            <THead cols={["Codice","N HK","Descrizione","Ubic","Prezzo€","FOB","VGM+HC","Alc","Carr","Step1€","Step1","WH","Step2",""]}/>
            <tbody>
              {rows.slice(0,300).map(r=>(
                <tr key={r.id} style={{opacity:r.cost?1:0.5}}>
                  <TD mono>{r.code}</TD><TD mono>{r.nHK||"—"}</TD><TD>{r.description}</TD>
                  <TD>{r.log?<UbicChip u={r.log.ubicazione}/>:<Chip label="—" color={C.muted}/>}</TD>
                  {r.cost?(
                    <>
                      <TD right>{fmt2(r.cost.priceEur)}</TD>
                      <TD right>{fmt2(r.cost.fob)}</TD>
                      <TD right>{fmt2(r.cost.vgm+r.cost.hc)}</TD>
                      <TD right>{fmt2(r.cost.alc)}</TD>
                      <TD right>{fmt2(r.cost.carriageUnit)}</TD>
                      <TD right>{fmt2(r.cost.step1Eur)}</TD>
                      <TD right><b>{fmt2(r.cost.step1Hkd)}</b></TD>
                      <TD right>{fmt2(r.cost.wh)}</TD>
                      <TD right><b style={{color:C.green}}>{fmt2(r.cost.step2Hkd)}</b></TD>
                    </>
                  ):(
                    <td colSpan={9} style={{padding:"5px 10px",fontSize:11,color:C.red}}>{r.skipReason}</td>
                  )}
                  <TD><MiniBtn label="+" onClick={()=>setDetail(r)}/></TD>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Section>
      {detail&&<CostBreakdown row={detail} onClose={()=>setDetail(null)} branch={branch} month={month}/>}
    </div>
  );
}

function CostBreakdown({row,onClose,branch,month}){
  const c=row.cost;
  if(!c)return null;
  const lines=[
    ["Prezzo acquisto EUR",fmt2(c.priceEur)],
    ["FOB surcharge",fmt2(c.fob)],
    ["Licenza certificato",fmt2(c.lic)],
    ["VGM",fmt2(c.vgm)],
    ["HC (handling/customs)",fmt2(c.hc)],
    ["PLT cost",fmt2(c.plt)],
    ["Tassa alcolici",fmt2(c.alc)],
    ["Carriage (unit)",fmt2(c.carriageUnit)],
    ["── STEP 1 EUR",fmt2(c.step1Eur)],
    [`── STEP 1 ${branch}`,fmt2(c.step1Hkd)],
    ["WH surcharge 0.5%",fmt2(c.wh)],
    ["── STEP 2 EUR",fmt2(c.step2Eur)],
    [`── STEP 2 ${branch}`,fmt2(c.step2Hkd)],
    ["Tasso FX",c.rate.toFixed(4)],
  ];
  return(
    <div style={{position:"fixed",top:0,left:0,right:0,bottom:0,background:"rgba(0,0,0,.7)",zIndex:200,display:"flex",alignItems:"center",justifyContent:"center"}}
      onClick={onClose}>
      <div style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:12,padding:24,minWidth:360,maxWidth:480,maxHeight:"90vh",overflowY:"auto"}}
        onClick={e=>e.stopPropagation()}>
        <div style={{fontWeight:700,fontSize:16,marginBottom:4,color:C.white}}>{row.code}</div>
        <div style={{fontSize:12,color:C.muted,marginBottom:16}}>{row.description} — {branch} {monthLabel(month)}</div>
        {lines.map(([k,v])=>(
          <div key={k} style={{display:"flex",justifyContent:"space-between",padding:"4px 0",borderBottom:`1px solid ${C.border}`,fontSize:13}}>
            <span style={{color:k.startsWith("──")?C.white:C.muted}}>{k}</span>
            <span style={{color:k.startsWith("──")?C.green:C.text,fontWeight:k.startsWith("──")?700:400}}>{v}</span>
          </div>
        ))}
        <button onClick={onClose} style={{marginTop:16,background:C.border,color:C.white,border:"none",borderRadius:6,padding:"6px 16px",cursor:"pointer"}}>Chiudi</button>
      </div>
    </div>
  );
}

// ─── MAIL GEN ─────────────────────────────────────────────────────────────────
function MailGen({costRows,branch,month,nav}){
  const [selected,setSelected]=useState(()=>new Set());
  const [intro,setIntro]=useState("Please find below the updated standard costs for your reference.");
  const [copied,setCopied]=useState(false);
  const calculated=costRows.filter(r=>r.cost!==null);
  function toggleAll(){if(selected.size===calculated.length)setSelected(new Set());else setSelected(new Set(calculated.map(r=>r.id)));}
  function toggle(id){setSelected(s=>{const n=new Set(s);n.has(id)?n.delete(id):n.add(id);return n;});}
  const selRows=calculated.filter(r=>selected.has(r.id));
  const mailBody=`Dear Team,\n\n${intro}\n\nStandard Cost Update — ${branch} ${monthLabel(month)}\n\n`+
    selRows.map(r=>`${r.code} | ${r.description} | ${fmt2(r.cost?.step2Hkd)} ${branch}`).join("\n")+
    "\n\nBest regards,\nIFB Cost Team";
  function copy(){navigator.clipboard.writeText(mailBody).then(()=>{setCopied(true);setTimeout(()=>setCopied(false),2000);});}
  if(calculated.length===0){
    return(<div><PageHeader title={`Genera Mail — ${branch} ${monthLabel(month)}`}/>
      <Section><EmptyState message="Nessun costo calcolato. Completa l'importazione di anagrafica e prezzi." btnLabel="📊 Vai a Costi" onBtn={()=>nav("Costi")}/></Section>
    </div>);
  }
  return(
    <div>
      <PageHeader title={`Genera Mail — ${branch} ${monthLabel(month)}`}/>
      <div style={{display:"flex",gap:16,flexWrap:"wrap"}}>
        <div style={{flex:"1 1 300px"}}>
          <Section title="Selezione prodotti">
            <div style={{marginBottom:8,display:"flex",gap:8,alignItems:"center"}}>
              <MiniBtn label={selected.size===calculated.length?"Deseleziona tutti":"Seleziona tutti"} onClick={toggleAll}/>
              <span style={{color:C.muted,fontSize:12}}>{selected.size} selezionati</span>
            </div>
            <div style={{maxHeight:400,overflowY:"auto"}}>
              {calculated.map(r=>(
                <label key={r.id} style={{display:"flex",alignItems:"center",gap:8,padding:"4px 0",cursor:"pointer",fontSize:12,color:C.text}}>
                  <input type="checkbox" checked={selected.has(r.id)} onChange={()=>toggle(r.id)} style={{accentColor:C.accent}}/>
                  <span style={{color:C.muted,minWidth:90,fontFamily:"monospace"}}>{r.code}</span>
                  <span style={{flex:1}}>{r.description}</span>
                  <span style={{color:C.green,minWidth:65,textAlign:"right"}}>{fmt2(r.cost?.step2Hkd)}</span>
                </label>
              ))}
            </div>
          </Section>
        </div>
        <div style={{flex:"1 1 300px"}}>
          <Section title="Anteprima mail">
            <FormField label="Testo introduttivo">
              <textarea value={intro} onChange={e=>setIntro(e.target.value)} style={{...inputStyle(),height:60,resize:"vertical"}}/>
            </FormField>
            <pre style={{background:"#111827",border:`1px solid ${C.border}`,borderRadius:6,padding:12,fontSize:11,color:C.text,whiteSpace:"pre-wrap",maxHeight:280,overflowY:"auto"}}>{mailBody}</pre>
            <div style={{marginTop:8}}>
              <ActionBtn label={copied?"✓ Copiato!":"📋 Copia testo"} onClick={copy} color={copied?C.green:C.accent}/>
            </div>
          </Section>
        </div>
      </div>
    </div>
  );
}

// ─── SALES INVOICE ────────────────────────────────────────────────────────────
function SalesInvoice({costRows,branch,month,nav}){
  // Persisted per branch so navigating away doesn't lose data
  const lsKey=`invoice_lines_${branch}`;
  const [lines,setLinesRaw]=useState(()=>LS.get(lsKey,[]));
  const [search,setSearch]=useState("");
  const setLines=v=>{LS.set(lsKey,v);setLinesRaw(v);};
  const calculated=costRows.filter(r=>r.cost!==null);
  const filtered=calculated.filter(r=>!search||(r.code+r.description).toLowerCase().includes(search.toLowerCase()));
  function addLine(r){if(lines.find(l=>l.id===r.id))return;setLines([...lines,{id:r.id,qty:1,markup:1.3}]);}
  function removeLine(id){setLines(lines.filter(l=>l.id!==id));}
  function updateLine(id,key,val){setLines(lines.map(l=>l.id===id?{...l,[key]:val}:l));}
  const total=lines.reduce((sum,l)=>{
    const row=calculated.find(r=>r.id===l.id);
    return sum+(row?.cost?.step2Hkd||0)*l.qty*l.markup;
  },0);
  function exportXlsx(){
    const data=lines.map(l=>{
      const row=calculated.find(r=>r.id===l.id);
      const unitCost=row?.cost?.step2Hkd||0;
      const unitPrice=unitCost*l.markup;
      return {"Codice":row?.code,"Descrizione":row?.description,"Qty":l.qty,"Costo Unit":fmt2(unitCost),"Markup":l.markup,"Prezzo Unit":fmt2(unitPrice),"Totale":fmt2(unitPrice*l.qty)};
    });
    const ws=XLSX.utils.json_to_sheet(data);
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,"Fattura");
    XLSX.writeFile(wb,`Fattura_${branch}_${month}.xlsx`);
  }
  if(calculated.length===0){
    return(<div><PageHeader title={`Fattura di Vendita — ${branch} ${monthLabel(month)}`}/>
      <Section><EmptyState message="Nessun costo calcolato per questo branch/mese." btnLabel="📊 Vai a Costi" onBtn={()=>nav("Costi")}/></Section>
    </div>);
  }
  return(
    <div>
      <PageHeader title={`Fattura di Vendita — ${branch} ${monthLabel(month)}`}/>
      <div style={{display:"flex",gap:16,flexWrap:"wrap"}}>
        <div style={{flex:"1 1 280px"}}>
          <Section title="Aggiungi prodotti">
            <SearchBar value={search} onChange={setSearch} placeholder="Cerca prodotto..."/>
            <div style={{maxHeight:380,overflowY:"auto",marginTop:10}}>
              {filtered.slice(0,100).map(r=>(
                <div key={r.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"4px 0",borderBottom:`1px solid ${C.border}`}}>
                  <div>
                    <div style={{fontSize:12,fontFamily:"monospace",color:C.muted}}>{r.code}</div>
                    <div style={{fontSize:11,color:C.text}}>{r.description}</div>
                  </div>
                  <MiniBtn label={lines.find(l=>l.id===r.id)?"✓":"+"} onClick={()=>addLine(r)} active={!!lines.find(l=>l.id===r.id)}/>
                </div>
              ))}
            </div>
          </Section>
        </div>
        <div style={{flex:"2 1 380px"}}>
          <Section title={`Righe fattura (${lines.length})`}>
            {lines.length===0&&<div style={{color:C.muted,fontSize:13,padding:"16px 0"}}>Nessuna riga. Aggiungi prodotti dalla lista.</div>}
            {lines.length>0&&(
              <>
                <table style={{width:"100%",borderCollapse:"collapse"}}>
                  <THead cols={["Codice","Descrizione","Costo","Qty","Markup","Totale",""]}/>
                  <tbody>
                    {lines.map(l=>{
                      const row=calculated.find(r=>r.id===l.id);
                      const unitCost=row?.cost?.step2Hkd||0;
                      return(
                        <tr key={l.id}>
                          <TD mono>{row?.code}</TD><TD>{row?.description}</TD>
                          <TD right>{fmt2(unitCost)}</TD>
                          <TD><input type="number" min={1} value={l.qty} onChange={e=>updateLine(l.id,"qty",parseFloat(e.target.value)||1)} style={{...inputStyle("55px")}}/></TD>
                          <TD><input type="number" step=".01" min={1} value={l.markup} onChange={e=>updateLine(l.id,"markup",parseFloat(e.target.value)||1)} style={{...inputStyle("65px")}}/></TD>
                          <TD right><b>{fmt2(unitCost*l.qty*l.markup)}</b></TD>
                          <TD><MiniBtn label="✕" onClick={()=>removeLine(l.id)}/></TD>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
                <div style={{marginTop:12,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                  <div style={{fontSize:18,fontWeight:700,color:C.green}}>Totale: {fmt2(total)} {branch}</div>
                  <div style={{display:"flex",gap:8}}>
                    <ActionBtn label="📥 Esporta Excel" onClick={exportXlsx} small/>
                    <ActionBtn label="Svuota" onClick={()=>setLines([])} color={C.red} small/>
                  </div>
                </div>
              </>
            )}
          </Section>
        </div>
      </div>
    </div>
  );
}

// ─── STORICO ─────────────────────────────────────────────────────────────────
function Storico({costRows,products,prices,logistics,fxRates,branch,month,nav}){
  const [compareMonth,setCompareMonth]=useState("");
  const prevPrices=prices.filter(p=>p.branch===branch&&p.month===compareMonth);
  const hasPrev=compareMonth&&prevPrices.length>0;
  const rows=costRows.filter(r=>r.cost!==null).map(r=>{
    let prevCost=null;
    if(hasPrev){
      const pp=prevPrices.find(p=>p.productId===r.id);
      if(pp){
        const fxE=fxRates.find(f=>f.branch===branch&&f.month===compareMonth);
        const fxRate=fxE?.rate||1;
        const log=logistics.find(l=>l.productId===r.id&&l.branch===branch);
        const effectiveLog=log||{productId:r.id,branch,area:"NORD",ubicazione:"MTO",pltPerContainer:20,hasCert:false,hasAlcTax:false,alcTax:0,convFactor:1,carriage:0};
        const priceInput=selectPrice(pp,effectiveLog.ubicazione);
        prevCost=calcHK({priceInput,ubicazione:effectiveLog.ubicazione,product:r,logistic:{...effectiveLog,category:r.category},eurToHkd:fxRate});
      }
    }
    const delta=prevCost&&r.cost?(r.cost.step2Hkd-prevCost.step2Hkd)/prevCost.step2Hkd:null;
    return {...r,prevCost,delta};
  });
  if(costRows.filter(r=>r.cost).length===0){
    return(<div><PageHeader title={`Storico Costi — ${branch}`}/>
      <Section><EmptyState message={`Nessun costo calcolato per ${branch}. Importa prezzi e anagrafica.`} btnLabel="📥 Importa Prezzi" onBtn={()=>nav("Importa Prezzi")}/></Section>
    </div>);
  }
  return(
    <div>
      <PageHeader title={`Storico Costi — ${branch}`}/>
      <Section title="Confronto periodi">
        <div style={{display:"flex",gap:16,alignItems:"center",flexWrap:"wrap"}}>
          <div>
            <div style={{fontSize:12,color:C.muted,marginBottom:4}}>Mese corrente</div>
            <div style={{fontSize:15,fontWeight:600,color:C.white}}>{monthLabel(month)}</div>
          </div>
          <div style={{color:C.muted,fontSize:20}}>vs</div>
          <FormField label="Mese precedente da confrontare">
            <input type="month" value={compareMonth} onChange={e=>setCompareMonth(e.target.value)} style={inputStyle("140px")}/>
          </FormField>
          {compareMonth&&!hasPrev&&<div style={{color:C.yellow,fontSize:12}}>Nessun prezzo trovato per {branch}/{compareMonth}</div>}
        </div>
      </Section>
      <Section>
        <table style={{width:"100%",borderCollapse:"collapse"}}>
          <THead cols={["Codice","Descrizione","Costo corrente","Costo precedente","Variazione"]}/>
          <tbody>
            {rows.map(r=>(
              <tr key={r.id}>
                <TD mono>{r.code}</TD><TD>{r.description}</TD>
                <TD right>{fmt2(r.cost?.step2Hkd)}</TD>
                <TD right>{r.prevCost?fmt2(r.prevCost.step2Hkd):"—"}</TD>
                <TD>{r.delta!==null?<DeltaBadge delta={r.delta}/>:<span style={{color:C.muted,fontSize:11}}>{hasPrev?"N/D":"—"}</span>}</TD>
              </tr>
            ))}
          </tbody>
        </table>
      </Section>
    </div>
  );
}

// ─── XREF PAGE ────────────────────────────────────────────────────────────────
function XRefPage({xrefs,setXrefs}){
  const [newRow,setNewRow]=useState({hkCode:"",ifbCode:"",note:""});
  function add(){if(!newRow.hkCode||!newRow.ifbCode)return;setXrefs([...xrefs,{...newRow}]);setNewRow({hkCode:"",ifbCode:"",note:""});}
  function del(i){const u=[...xrefs];u.splice(i,1);setXrefs(u);}
  return(
    <div>
      <PageHeader title="Tabella XRef" sub="Mappatura codici HK → IFB (globale, tutti i branch)"/>
      <Section>
        {xrefs.length===0&&(
          <div style={{color:C.muted,fontSize:13,marginBottom:16}}>
            Nessuna mappatura. Aggiungi sotto se un codice HK non corrisponde al codice IFB.
          </div>
        )}
        <table style={{width:"100%",borderCollapse:"collapse"}}>
          <THead cols={["Codice HK","Codice IFB","Nota",""]}/>
          <tbody>
            {xrefs.map((x,i)=>(
              <tr key={i}>
                <TD mono>{x.hkCode}</TD><TD mono>{x.ifbCode}</TD><TD>{x.note}</TD>
                <TD><MiniBtn label="✕" onClick={()=>del(i)}/></TD>
              </tr>
            ))}
            <tr>
              <TD><input value={newRow.hkCode} onChange={e=>setNewRow(x=>({...x,hkCode:e.target.value}))} placeholder="HK Code" style={inputStyle()}/></TD>
              <TD><input value={newRow.ifbCode} onChange={e=>setNewRow(x=>({...x,ifbCode:e.target.value}))} placeholder="IFB Code" style={inputStyle()}/></TD>
              <TD><input value={newRow.note} onChange={e=>setNewRow(x=>({...x,note:e.target.value}))} placeholder="Nota" style={inputStyle()}/></TD>
              <TD><ActionBtn label="+ Aggiungi" onClick={add} small/></TD>
            </tr>
          </tbody>
        </table>
      </Section>
    </div>
  );
}

// ─── NOTES PAGE ───────────────────────────────────────────────────────────────
function NotesPage({notes,setNotes}){
  return(
    <div>
      <PageHeader title="Note" sub="Appunti liberi — salvati automaticamente"/>
      <Section>
        <textarea value={notes} onChange={e=>setNotes(e.target.value)}
          style={{...inputStyle(),height:400,resize:"vertical",lineHeight:1.6}}
          placeholder="Inserisci note, commenti, decisioni di calcolo..."/>
      </Section>
    </div>
  );
}

// ─── IMPORT BC ────────────────────────────────────────────────────────────────
function ImportBC({products,setProducts}){
  const [step,setStep]=useState("idle");
  const [preview,setPreview]=useState([]);
  const [msg,setMsg]=useState("");
  async function onFile(file){
    try{
      const rows=await readXlsx(file);
      if(rows.length<2){setMsg("File vuoto.");return;}
      let hi=0;
      for(let i=0;i<Math.min(10,rows.length);i++){
        const r=rows[i].map(normalize);
        if(r.some(c=>["no_","no.","no","item no","item no."].includes(c)||c.includes("ifb"))){hi=i;break;}
      }
      const headers=rows[hi].map(c=>String(c||"").trim());
      const hn=headers.map(normalize);
      const idx=aliases=>findAlias(hn,aliases);
      const iCode=idx(BC_FIELD_ALIASES.code);
      const iNHK=idx(BC_FIELD_ALIASES.nHK);
      const iDesc=idx(BC_FIELD_ALIASES.description);
      const iVend=idx(BC_FIELD_ALIASES.vendorName);
      const iCat=idx(BC_FIELD_ALIASES.category);
      const iUom=idx(BC_FIELD_ALIASES.uom);
      const iQty=idx(BC_FIELD_ALIASES.qtyPerBox);
      const iBpp=idx(BC_FIELD_ALIASES.boxPerPallet);
      const iPlt=idx(BC_FIELD_ALIASES.pltPerContainer);
      const iNW=idx(BC_FIELD_ALIASES.netWeight);
      const iGW=idx(BC_FIELD_ALIASES.grossWeight);
      const iVol=idx(BC_FIELD_ALIASES.volume);
      const result=[];
      for(let i=hi+1;i<rows.length;i++){
        const row=rows[i];
        const code=String(getCol(row,iCode,"")||"").trim();
        if(!code)continue;
        if(/^P_BC_/i.test(code))continue;
        result.push({
          id:code, nHK:String(getCol(row,iNHK,"")||"").trim(), code,
          description:String(getCol(row,iDesc,"")||"").trim(),
          vendorName:String(getCol(row,iVend,"")||"").trim(),
          category:String(getCol(row,iCat,"")||"").trim(),
          uom:String(getCol(row,iUom,"")||"").trim(),
          qtyPerBox:parseFloat(String(getCol(row,iQty,0)))||0,
          boxPerPallet:parseFloat(String(getCol(row,iBpp,0)))||0,
          pltPerContainer:parseFloat(String(getCol(row,iPlt,0)))||0,
          netWeight:parseFloat(String(getCol(row,iNW,0)))||0,
          grossWeight:parseFloat(String(getCol(row,iGW,0)))||0,
          volume:parseFloat(String(getCol(row,iVol,0)))||0,
        });
      }
      setPreview(result);setStep("preview");
      setMsg(`Trovati ${result.length} prodotti da importare.`);
    }catch(e){setMsg("Errore: "+String(e));}
  }
  function executeImport(mode){
    if(mode==="replace"){setProducts(preview);}
    else{const map=new Map(products.map(p=>[p.id,p]));preview.forEach(p=>map.set(p.id,p));setProducts(Array.from(map.values()));}
    setMsg(`✓ Importati ${preview.length} prodotti (${mode}).`);
    setStep("idle");setPreview([]);
  }
  return(
    <div>
      <PageHeader title="Importa Anagrafica BC" sub="Importa prodotti da Business Central export"/>
      <StepBar steps={["Carica file","Anteprima","Importa"]} current={step==="idle"?0:1}/>
      {step==="idle"&&<DropZone onFile={onFile} label="Carica export BC (Excel/CSV) — colonne: No., Description, Vendor Name, Qty/Box, Box/Plt..."/>}
      {msg&&<div style={{color:step==="idle"?C.muted:C.green,marginTop:10,fontSize:13}}>{msg}</div>}
      {step==="preview"&&(
        <div>
          <div style={{display:"flex",gap:10,marginBottom:14,flexWrap:"wrap"}}>
            <ActionBtn label="✓ Sostituisci tutto" onClick={()=>executeImport("replace")} color={C.accent}/>
            <ActionBtn label="Merge (aggiorna esistenti)" onClick={()=>executeImport("merge")} color={C.purple}/>
            <ActionBtn label="Annulla" onClick={()=>{setStep("idle");setPreview([]);setMsg("");}} color={C.red} small/>
          </div>
          <Section>
            <table style={{width:"100%",borderCollapse:"collapse"}}>
              <THead cols={["Codice IFB","N HK","Descrizione","Fornitore","Cat","Qty/Box","Box/Plt"]}/>
              <tbody>
                {preview.slice(0,50).map((p,i)=>(
                  <tr key={i}>
                    <TD mono>{p.code}</TD><TD mono>{p.nHK||"—"}</TD><TD>{p.description}</TD>
                    <TD>{p.vendorName||"—"}</TD><TD>{p.category||"—"}</TD>
                    <TD right>{p.qtyPerBox||"—"}</TD><TD right>{p.boxPerPallet||"—"}</TD>
                  </tr>
                ))}
              </tbody>
            </table>
            {preview.length>50&&<div style={{color:C.muted,fontSize:12,marginTop:8}}>Mostrati 50/{preview.length} — verranno importati tutti</div>}
          </Section>
        </div>
      )}
    </div>
  );
}

// ─── IMPORT PRICES ────────────────────────────────────────────────────────────
function ImportPrices({products,prices,setPrices,xrefs,branch,month}){
  const [step,setStep]=useState("idle");
  const [preview,setPreview]=useState([]);
  const [msg,setMsg]=useState("");
  async function onFile(file){
    try{
      const rows=await readXlsx(file);
      if(rows.length<2){setMsg("File vuoto.");return;}
      let hi=0;
      for(let i=0;i<Math.min(10,rows.length);i++){
        const r=rows[i].map(normalize);
        if(r.some(c=>["no_","no.","no","item no","code","codice"].includes(c)||c.includes("fca"))){hi=i;break;}
      }
      const headers=rows[hi].map(c=>String(c||"").trim());
      const hn=headers.map(normalize);
      const idx=aliases=>findAlias(hn,aliases);
      const iCode=idx(PRICE_FIELD_ALIASES.code);
      const iNHK=idx(PRICE_FIELD_ALIASES.nHK);
      const iBranch=idx(PRICE_FIELD_ALIASES.branch);
      const iMonth=idx(PRICE_FIELD_ALIASES.month);
      const iType=idx(PRICE_FIELD_ALIASES.priceType);
      const iFcaG=idx(PRICE_FIELD_ALIASES.fcaGross);
      const iDisc=idx(PRICE_FIELD_ALIASES.discountPct);
      const iFcaN=idx(PRICE_FIELD_ALIASES.fcaDiscounted);
      const iDapG=idx(PRICE_FIELD_ALIASES.dapGross);
      const iDapD=idx(PRICE_FIELD_ALIASES.dapDiscount);
      const iDapF=idx(PRICE_FIELD_ALIASES.dapFinal);
      const iMts=idx(PRICE_FIELD_ALIASES.mtsPrice);
      const result=[];
      for(let i=hi+1;i<rows.length;i++){
        const row=rows[i];
        const rawCode=String(getCol(row,iCode,"")||"").trim();
        if(!rawCode)continue;
        if(/^P_BC_/i.test(rawCode))continue;
        if(/^P_\d/.test(rawCode)&&rawCode.length>8)continue;
        const rowBranch=String(getCol(row,iBranch,branch)||"").trim()||branch;
        const rowMonth=String(getCol(row,iMonth,month)||"").trim()||month;
        const discountPct=parseFloat(String(getCol(row,iDisc,0)))||0;
        const fcaGross=parseFloat(String(getCol(row,iFcaG,0)))||0;
        const fcaDiscounted=parseFloat(String(getCol(row,iFcaN,0)))||(fcaGross*(1-discountPct/100));
        const dapGross=parseFloat(String(getCol(row,iDapG,0)))||0;
        const dapFinal=parseFloat(String(getCol(row,iDapF,0)))||dapGross;
        const mtsPrice=parseFloat(String(getCol(row,iMts,0)))||0;
        const xref=xrefs.find(x=>x.hkCode===rawCode);
        const prod=products.find(p=>p.code===rawCode||p.nHK===rawCode)||(xref&&products.find(p=>p.code===xref.ifbCode));
        const productId=prod?.id||rawCode;
        result.push({
          productId,branch:rowBranch,month:rowMonth,
          priceType:String(getCol(row,iType,"")||"").trim(),
          fcaGross,discountPct,fcaDiscounted,dapGross,
          dapDiscount:parseFloat(String(getCol(row,iDapD,0)))||0,
          dapFinal,mtsPrice,
        });
      }
      setPreview(result);setStep("preview");
      setMsg(`Trovati ${result.length} prezzi da importare.`);
    }catch(e){setMsg("Errore: "+String(e));}
  }
  function executeImport(mode){
    if(mode==="replace"){
      const other=prices.filter(p=>!(p.branch===branch&&p.month===month));
      setPrices([...other,...preview]);
    }else{
      const map=new Map(prices.map(p=>[`${p.productId}|${p.branch}|${p.month}`,p]));
      preview.forEach(p=>map.set(`${p.productId}|${p.branch}|${p.month}`,p));
      setPrices(Array.from(map.values()));
    }
    setMsg(`✓ Importati ${preview.length} prezzi (${mode}).`);
    setStep("idle");setPreview([]);
  }
  return(
    <div>
      <PageHeader title="Importa Prezzi / Listini" sub={`Branch: ${branch} — Mese: ${monthLabel(month)}`}/>
      <StepBar steps={["Carica file","Anteprima","Importa"]} current={step==="idle"?0:1}/>
      {step==="idle"&&<DropZone onFile={onFile} label="Carica listino prezzi (Excel/CSV) — colonne: codice, FCA, DAP, MTS..."/>}
      {msg&&<div style={{color:step==="idle"?C.muted:C.green,marginTop:10,fontSize:13}}>{msg}</div>}
      {step==="preview"&&(
        <div>
          <div style={{display:"flex",gap:10,marginBottom:14,flexWrap:"wrap"}}>
            <ActionBtn label={`✓ Sostituisci ${branch}/${month}`} onClick={()=>executeImport("replace")} color={C.accent}/>
            <ActionBtn label="Merge globale" onClick={()=>executeImport("merge")} color={C.purple}/>
            <ActionBtn label="Annulla" onClick={()=>{setStep("idle");setPreview([]);setMsg("");}} color={C.red} small/>
          </div>
          <Section>
            <table style={{width:"100%",borderCollapse:"collapse"}}>
              <THead cols={["ID Prodotto","Branch","Mese","Tipo","FCA Netto","DAP Final","MTS"]}/>
              <tbody>
                {preview.slice(0,50).map((p,i)=>(
                  <tr key={i}>
                    <TD mono>{p.productId}</TD><TD>{p.branch}</TD><TD>{p.month}</TD>
                    <TD><Chip label={p.priceType||"—"}/></TD>
                    <TD right>{fmt2(p.fcaDiscounted)}</TD>
                    <TD right>{fmt2(p.dapFinal)}</TD>
                    <TD right>{fmt2(p.mtsPrice)}</TD>
                  </tr>
                ))}
              </tbody>
            </table>
            {preview.length>50&&<div style={{color:C.muted,fontSize:12,marginTop:8}}>Mostrati 50/{preview.length} — verranno importati tutti</div>}
          </Section>
        </div>
      )}
    </div>
  );
}