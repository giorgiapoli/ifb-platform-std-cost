import { useState, useMemo } from "react";
import * as XLSX from "xlsx";

const T = {
  bg:"#0B0F14", surface:"#111720", panel:"#111720", card:"#161E28",
  border:"rgba(255,255,255,0.07)", borderHi:"rgba(255,255,255,0.14)",
  text:"#E2D9CC", muted:"rgba(226,217,204,0.45)", dim:"rgba(226,217,204,0.22)",
  gold:"#C9A84C", goldDim:"rgba(201,168,76,0.18)",
  blue:"#4A8FB5", green:"#4BA87A", red:"#B5534A", orange:"#C47A3B", purple:"#7B5AC4",
};
const BRANCH_CFG = {
  HK:  { label:"Hong Kong", flag:"🇭🇰", color:T.gold,   currency:"HKD", defaultRate:9.1437, active:true  },
  MAC: { label:"Macao",     flag:"🇲🇴", color:T.green,  currency:"MOP", defaultRate:9.08,   active:false },
  CAN: { label:"Canarie",   flag:"🇮🇨", color:T.blue,   currency:"EUR", defaultRate:1,      active:false },
  AUS: { label:"Australia", flag:"🇦🇺", color:T.orange, currency:"AUD", defaultRate:1.6420, active:false },
};
const IFB_VENDOR = "INALCA FOOD & BEVERAGE";

const NOW = () => new Date().toISOString().slice(0,7);
const roundN = (n, d=2) => Math.round((n||0)*Math.pow(10,d))/Math.pow(10,d);
const EXCLUDED_INVOICE_DESC = ["health certificate costs","late payment interest","interest on intercompany","handling costs","freight cost"];
const isExcludedDesc = d => EXCLUDED_INVOICE_DESC.some(ex=>String(d||"").toLowerCase().includes(ex));
const AIR_TYPES = ["air","ch air","fr air","dr air","chilled air","frozen air","dry air"];
const isAirTransport = t => AIR_TYPES.some(a=>String(t||"").toLowerCase().trim()===a);
const isIFBVendor = v => String(v||"").toUpperCase().includes("INALCA FOOD");

function findProduct(code, products, xrefs=[]) {
  if(!code) return null;
  const c = String(code).trim();
  let p = products.find(pr=>pr.code===c); if(p) return p;
  p = products.find(pr=>pr.nHK&&pr.nHK===c); if(p) return p;
  const xr = xrefs.find(x=>x.nHK===c); if(xr){ p=products.find(pr=>pr.code===xr.ifbNo); if(p) return p; }
  const xr2 = xrefs.find(x=>x.ifbNo===c); if(xr2){ p=products.find(pr=>pr.nHK===xr2.nHK); if(p) return p; }
  return null;
}

// ─── COST ENGINE ─────────────────────────────────────────────────────────────
// All costs confirmed from 05_Modello_Standard_Cost.xlsx COSTS (LOG) sheet
const COSTS = {
  FOB:{ DRY:{NORD:2000,CENTRO:0,SUD:1108.55}, FRESH:{NORD:3500,CENTRO:3500,SUD:0}, FROZEN:{NORD:4000,CENTRO:0,SUD:0} },
  LIC_HKD: 4100+3800,   // LOCAL IMPORT CHARGES in HKD: C MEDIO CON PALLET + C TERMINAL
  VGM: 100,             // EUR per container
  HC:  80,              // EUR per container (Health Certificate)
  PLT: 30,              // EUR per pallet
  // MTO: In+Stock+Out per pallet (EUR)
  MTO:{ DRY:8.16, FRESH:10.2, FROZEN:12.24 },
  // MTS breakdown per pallet
  MTS_D:{ DRY:14.4228, FRESH:16.4832, FROZEN:24.7248 },  // Deposito
  MTS_I:{ DRY:2.5755,  FRESH:3.6057,  FROZEN:3.6057  },  // Inbound
  MTS_P:{ DRY:0.303,   FRESH:0.3434,  FROZEN:0.3535  },  // Picking (per collo)
  // Wine/spirits carriage: EUR per pallet
  CARRIAGE_WINE: 60,
  VENDOR_CARRIAGE:{
    "ALICO SRL":70,"ANTICO PASTIFICIO MORELLI SRL":80,"AZ. AGRICOLA MANCINI SRL AGRICOLA":70,
    "BONOMI SPA":30,"CAPURSO AZIENDA CASEARIA S.R.L.":90,"CECCHINI DARIO SRL":20,
    "CONSERVAS ANGELACHU S.L.":200,"DELIZIA 2000 SRL":75,"GRA-COM S.R.L.":90,
    "GREENS FOOD SPA":30,"INALCA S.P.A. A SOCIO UNICO":40,"ITALPIZZA S.R.L.":30,
    "OILALA' SRL":100,"QUANTOBASTA S.R.L.":140,"VALLE FINE FOODS ITALIA S.R.L.S.":150,
  },
};

/**
 * Main cost calculation for HK (sea).
 * Returns null if totalUnits = 0 or missing data.
 *
 * FORMULA (per unit):
 *   Step1 = purchasePrice + FOB + LIC + VGM + HC + PLT + alcTax + carriageUnit
 *   Step2 = Step1 + warehouseCost
 */
function calcHK({ priceInput, ubicazione, product, logistic, eurToHkd }) {
  const { uom, qtyPerBox, boxPerPallet, kgPerBox, temperature } = product;
  const { area, pltPerContainer, hasCert, hasAlcTax, alcTax, convFactor=1, carriage=0, vendorName="", category="" } = logistic;

  // Units per pallet
  let unitsPerPlt;
  if(uom==="BOX")      unitsPerPlt = boxPerPallet;
  else if(uom==="KG")  unitsPerPlt = (kgPerBox||qtyPerBox)*boxPerPallet;
  else                 unitsPerPlt = qtyPerBox*boxPerPallet;

  // "Divisore collo" = units per box/parcel for MTS picking cost
  const divisoreCollo = uom==="BOX" ? 1 : uom==="KG" ? (kgPerBox||qtyPerBox) : qtyPerBox;

  const totalUnits = unitsPerPlt * pltPerContainer;
  if(!totalUnits) return null;

  const priceEur = (priceInput||0) * convFactor;

  // FOB per unit (EUR)
  const fob = (COSTS.FOB[temperature]?.[area] ?? 0) / totalUnits;

  // LOCAL IMPORT CHARGES: (4100+3800) HKD converted to EUR, per unit
  const lic = (COSTS.LIC_HKD / eurToHkd) / totalUnits;

  // VGM per unit (EUR)
  const vgm = COSTS.VGM / totalUnits;

  // Health Certificate per unit (EUR)
  const hc = hasCert ? COSTS.HC / totalUnits : 0;

  // Pallet cost per unit (EUR)
  const plt = COSTS.PLT / unitsPerPlt;

  // Alcohol tax per unit (EUR) — alcTax is already a per-unit value from Work_tab TASSA ALCOLICA
  const alc = hasAlcTax ? (alcTax||0) : 0;

  // Carriage per unit (EUR)
  const isWineSpirits = category==="WINE" || category==="SPIRITS";
  const carriagePlt = carriage>0 ? carriage : isWineSpirits ? COSTS.CARRIAGE_WINE : (COSTS.VENDOR_CARRIAGE[vendorName]||0);
  const carriageUnit = carriagePlt > 0 ? carriagePlt / unitsPerPlt : 0;

  const step1Eur = priceEur + fob + lic + vgm + hc + plt + alc + carriageUnit;

  // Warehouse cost per unit (EUR)
  let wh = 0;
  if(ubicazione==="MTO") {
    wh = COSTS.MTO[temperature] / unitsPerPlt;
  } else if(ubicazione==="MTS") {
    wh = COSTS.MTS_D[temperature] / unitsPerPlt
       + COSTS.MTS_I[temperature] / unitsPerPlt
       + COSTS.MTS_P[temperature] / divisoreCollo;
  }
  // FOR: no warehouse cost (supplier delivers directly)

  const step2Eur = step1Eur + wh;

  return {
    priceEur, fob, lic, vgm, hc, plt, alc, carriageUnit,
    step1Eur, step1Hkd: step1Eur * eurToHkd,
    wh, step2Eur, step2Hkd: Math.round(step2Eur * eurToHkd * 100)/100,
    rate: eurToHkd,
  };
}

function selectPrice(pr, ubicazione) {
  if(!pr) return 0;
  if(ubicazione==="FOR") return pr.fcaDiscounted || pr.dapFinal || 0;
  if(ubicazione==="MTO") return pr.dapFinal || 0;
  if(ubicazione==="MTS") { const m=pr.mtsPrice||0; return m!==0?m:(pr.dapFinal||0); }
  return pr.dapFinal || 0;
}

// ─── FIELD ALIASES ────────────────────────────────────────────────────────────
const BC_FIELD_ALIASES = {
  nHK:         ["n hk","nhk","hk code","hk no","n_hk","gc code","gc no","hong kong no","no"],
  code:        ["no_","no.","item no.","item no","ifb no","ifb n","ifb item","codice","code"],
  description: ["description","descrizione","desc","item description"],
  category:    ["sectiondescription","section description","section desc","section","item category code","item category","category","categoria"],
  uom:         ["salesunitofmeasure","sales unit of measure","base unit of measure","uom","unit of measure","base uom","unit"],
  qtyPerBox:   ["quantityxpackaging","quantity x packaging","units per parcel","qty per box","qty/box","pz per cartone"],
  boxPerPallet:["packagingxpallet","packaging x pallet","parcels per pallet","box per pallet","cartoni per pallet"],
  kgPerBox:    ["kgperbox","kg per box","net weight","peso netto","kg per cartone","netweight"],
  temperature: ["producttype","product type","product type rettificato","product type - anagrafica","item tracking code","temperatura","temperature","storage"],
  active:      ["blocked","bloccato","active","attivo"],
  vendorName:  ["vendorname","vendor name","vendor name 2","vendor","fornitore","vendor name2"],
};

const PRICE_FIELD_ALIASES = {
  code:          ["no_","no.","no","item no.","codice","code","n hk","ifb item","ifb no","ifb n"],
  vendorName:    ["vendor name 3","vendor name","vendor","fornitore","vendor name 2"],
  section:       ["section description","sectiondescription","section desc","section","sezione","categoria","category"],
  mtsPrice:      ["mts price","mts","mts price (eur)"],
  fcaPrice:      ["fca price","fca"],
  fcaDiscount:   ["fca discount","fca disc","fca discount %"],
  fcaDiscounted: ["fca discounted","fca disc.","fca final"],
  dapPrice:      ["dap price","dap"],
  dapDiscount:   ["dap discount","dap disc"],
  dapDiscounted: ["dap discounted","dap final discounted"],
  dapFinalDirect:["dap final","dap final price","final price","prezzo acquisto"],
};

const FOR_VENDORS = new Set(["ALICO SRL","ANTICO PASTIFICIO MORELLI SRL","AZ. AGRICOLA MANCINI SRL AGRICOLA","BONOMI SPA","CAPURSO AZIENDA CASEARIA S.R.L.","CECCHINI DARIO SRL","CONSERVAS ANGELACHU S.L.","DELIZIA 2000 SRL","GRA-COM S.R.L.","GREENS FOOD SPA","INALCA S.P.A. A SOCIO UNICO","ITALPIZZA S.R.L.","OILALA' SRL","QUANTOBASTA S.R.L.","VALLE FINE FOODS ITALIA S.R.L.S"]);

function calcDAPFinal({ dapDiscounted, fcaPrice, fcaDiscounted, vendorName, section, products, code }) {
  const prod = products.find(p=>p.code===code);
  let unitsPerPlt = 1;
  if(prod) {
    const{uom,qtyPerBox,boxPerPallet,kgPerBox}=prod;
    if(uom==="BOX")     unitsPerPlt=boxPerPallet;
    else if(uom==="KG") unitsPerPlt=(kgPerBox||qtyPerBox)*boxPerPallet;
    else                unitsPerPlt=qtyPerBox*boxPerPallet;
  }
  const sec = (section||"").toUpperCase();
  const isWine = sec==="WINE"||sec==="SPIRITS";
  const isX = isWine || FOR_VENDORS.has(vendorName||"");
  const pltCost = isWine ? 60 : (COSTS.VENDOR_CARRIAGE[vendorName]||0);
  const cu = unitsPerPlt>0 ? pltCost/unitsPerPlt : 0;
  const dd=dapDiscounted||0, fp=fcaPrice||0, fd=fcaDiscounted||0;
  if(dd!==0) return { dapFinal:dd, carriageUnit:cu, note:"DAP Disc." };
  if(!isX)   return { dapFinal:0,  carriageUnit:0,  note:"non-X" };
  if(isWine) return { dapFinal:fp!==0?fp+cu:0, carriageUnit:cu, note:"Wine FCA+C" };
  return { dapFinal:fd!==0?fd+cu:0, carriageUnit:cu, note:"FCA Disc+C" };
}

const LS = {
  get: (k,def) => { try{ const v=localStorage.getItem(k); return v?JSON.parse(v):def; }catch{ return def; } },
  set: (k,v)   => { try{ localStorage.setItem(k,JSON.stringify(v)); }catch{} },
};

// Seed data (minimal)
const SEED_PRODUCTS = [];
const SEED_LOGISTIC = [];
const SEED_PRICES   = [];
const SEED_FX = [
  {branch:"HK", month:"2026-05",rate:9.1200},{branch:"HK", month:"2026-06",rate:9.1437},
  {branch:"MAC",month:"2026-05",rate:9.08},  {branch:"MAC",month:"2026-06",rate:9.08},
  {branch:"CAN",month:"2026-05",rate:1.0},   {branch:"CAN",month:"2026-06",rate:1.0},
  {branch:"AUS",month:"2026-05",rate:1.6200},{branch:"AUS",month:"2026-06",rate:1.6420},
];

// ─── APP ─────────────────────────────────────────────────────────────────────
export default function App() {
  const[products,setProducts]   = useState(()=>LS.get("ifb_products",SEED_PRODUCTS));
  const[logistics,setLogistics] = useState(()=>LS.get("ifb_logistics",SEED_LOGISTIC));
  const[prices,setPrices]       = useState(()=>LS.get("ifb_prices",SEED_PRICES));
  const[fx,setFx]               = useState(()=>LS.get("ifb_fx",SEED_FX));
  const[xrefs,setXrefs]         = useState(()=>LS.get("ifb_xrefs",[]));
  const[airList,setAirList]     = useState(()=>LS.get("ifb_airlist",[]));
  const[salesRows,setSalesRows] = useState(()=>LS.get("ifb_sales_invoice",[]));
  const[importLogs,setImportLogs] = useState(()=>LS.get("ifb_importlogs",[]));
  const[snapshots,setSnapshots]   = useState(()=>LS.get("ifb_snapshots",[]));
  const[costHistory,setCostHistory] = useState(()=>LS.get("ifb_costhistory",[]));
  const[lastImportTs,setLastImportTs] = useState(()=>LS.get("ifb_last_import_ts",0));
  const[lastCalcTs,setLastCalcTs]     = useState(()=>LS.get("ifb_last_calc_ts",0));
  const[page,setPage]     = useState("branchSelect");
  const[branch,setBranch] = useState("");
  const[month,setMonth]   = useState(NOW());
  const[toast,setToast]   = useState(null);

  const showToast = (msg,color=T.green) => { setToast({msg,color}); setTimeout(()=>setToast(null),3500); };
  const bumpImportTs = () => { const ts=Date.now(); setLastImportTs(ts); LS.set("ifb_last_import_ts",ts); return ts; };

  // Pallet/container default formula: =SE(temp="DRY";25;SE(temp="FRESH"||"FROZEN";23;20))
  // Matches Excel formula in STDC sheet column "Numero Pallet per Container"
  const pltDefault = (temp) => {
    const t = (temp||"DRY").toUpperCase();
    if(t==="DRY")   return 25;
    if(t==="FRESH"||t==="FROZEN") return 23;
    return 20;
  };

  // ── Cost rows: only IFB vendor, only SEA (non-AIR)
  const costRows = useMemo(()=>{
    if(!branch) return [];
    const fxRate = fx.find(f=>f.branch===branch&&f.month===month)?.rate || BRANCH_CFG[branch]?.defaultRate || 9.1437;
    const [yr,mo] = month.split("-").map(Number);
    const prevM = mo===1 ? `${yr-1}-12` : `${yr}-${String(mo-1).padStart(2,"0")}`;
    const eligible = products.filter(p => p.active && isIFBVendor(p.vendorName));

    return eligible.map(prod => {
      const airEntry = airList.find(a=>a.productId===prod.id);
      if(airEntry && isAirTransport(airEntry.transportation))
        return { ...prod, cost:null, prevCost:null, priceInput:null, isAir:true, skipReason:"AIR" };

      const logRaw = logistics.find(l=>l.productId===prod.id&&l.branch===branch);
      if(!logRaw) return { ...prod, cost:null, prevCost:null, priceInput:null, skipReason:"NO LOGISTICA" };

      // Apply pltPerContainer default formula if value from Work_tab is 0 or missing
      // Formula: =SE(ProductType="DRY";25;SE(OR(ProductType="FRESH";ProductType="FROZEN");23;0))
      const pltFromFile = logRaw.pltPerContainer || 0;
      const plt = pltFromFile > 0 ? pltFromFile : pltDefault(prod.temperature);
      const log = { ...logRaw, pltPerContainer: plt };

      const pr     = prices.find(p=>p.productId===prod.id&&p.branch===branch&&p.month===month);
      const prPrev = prices.find(p=>p.productId===prod.id&&p.branch===branch&&p.month===prevM);

      if(!pr) return { ...prod, cost:null, prevCost:null, priceInput:null, skipReason:`NO PREZZO (${branch}/${month})` };

      const ub  = log.ubicazione;
      const pi  = selectPrice(pr, ub);
      const piP = prPrev ? selectPrice(prPrev, ub) : null;

      const cost = calcHK({ priceInput:pi, ubicazione:ub, product:prod, logistic:{...log,category:prod.category}, eurToHkd:fxRate });
      if(!cost) return { ...prod, cost:null, prevCost:null, priceInput:pi,
        skipReason:`CALC=0 (qty=${prod.qtyPerBox} box/plt=${prod.boxPerPallet} plt=${plt} uom=${prod.uom})` };

      const prevCost = piP!=null ? calcHK({ priceInput:piP, ubicazione:ub, product:prod, logistic:{...log,category:prod.category}, eurToHkd:fxRate }) : null;
      const delta    = cost&&prevCost ? (cost.step2Hkd-prevCost.step2Hkd)/prevCost.step2Hkd*100 : null;
      return { ...prod, cost, prevCost, delta, priceInput:pi, isNew:!prPrev,
               flagged: delta!==null && Math.abs(delta)>=3, ubicazione:ub, pltUsed:plt };
    });
  }, [products,logistics,prices,fx,airList,branch,month]);

  const NAV = [
    {id:"dashboard",  icon:"⬡", label:"Dashboard"},
    {id:"products",   icon:"◈", label:"Anagrafica"},
    {id:"importAnag", icon:"⇪", label:"Import Anagrafica", badge:"BC"},
    {id:"xref",       icon:"⇄", label:"XRef N HK / IFB"},
    {id:"logistics",  icon:"◎", label:"Logistica"},
    {id:"prices",     icon:"◉", label:"Listini"},
    {id:"importPrice",icon:"💶", label:"Import Listini",  badge:"BC"},
    {id:"fx",         icon:"◌", label:"Cambi"},
    {id:"air",        icon:"✈", label:"AIR Transport"},
    {id:"costs",      icon:"◆", label:"Standard Cost"},
    {id:"sales",      icon:"📋", label:"Sales Invoice"},
    {id:"storico",    icon:"⧖", label:"Storico & Diff"},
    {id:"mail",       icon:"◻", label:"Mail Mensile"},
    {id:"notes",      icon:"📝", label:"Note & Ambiguità"},
  ];

  // ── Page: branch selection splash ──────────────────────────────────────────
  if(page==="branchSelect") return (
    <div style={{display:"flex",height:"100vh",width:"100vw",background:T.bg,alignItems:"center",justifyContent:"center",fontFamily:"'Palatino Linotype','Book Antiqua',Palatino,serif"}}>
      <div style={{textAlign:"center",maxWidth:"600px",padding:"40px"}}>
        <div style={{fontSize:"10px",letterSpacing:"4px",color:T.gold,textTransform:"uppercase",marginBottom:"8px"}}>IFB Platform</div>
        <h1 style={{color:T.text,margin:"0 0 8px",fontSize:"28px",fontWeight:"bold"}}>Cost Intelligence</h1>
        <div style={{color:T.muted,fontSize:"13px",marginBottom:"48px"}}>Seleziona la filiale con cui vuoi lavorare</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"16px"}}>
          {Object.entries(BRANCH_CFG).map(([key,cfg])=>(
            <button key={key} onClick={()=>{ setBranch(key); setPage("dashboard"); }}
              style={{padding:"28px 20px",background:T.card,border:`2px solid ${cfg.color}44`,
                borderRadius:"16px",cursor:"pointer",fontFamily:"inherit",transition:"all 0.2s",
                display:"flex",flexDirection:"column",alignItems:"center",gap:"10px",
                opacity: cfg.active ? 1 : 0.55,
              }}
              onMouseEnter={e=>{ e.currentTarget.style.borderColor=cfg.color; e.currentTarget.style.background=`${cfg.color}11`; }}
              onMouseLeave={e=>{ e.currentTarget.style.borderColor=`${cfg.color}44`; e.currentTarget.style.background=T.card; }}>
              <span style={{fontSize:"36px"}}>{cfg.flag}</span>
              <span style={{fontSize:"16px",fontWeight:"bold",color:cfg.color}}>{cfg.label}</span>
              <span style={{fontSize:"11px",color:T.muted}}>{cfg.currency}</span>
              {!cfg.active && <span style={{fontSize:"9px",color:cfg.color,background:`${cfg.color}22`,padding:"2px 8px",borderRadius:"4px",letterSpacing:"1px"}}>COMING SOON</span>}
            </button>
          ))}
        </div>
      </div>
    </div>
  );

  const needsRecalc = lastImportTs > lastCalcTs;
  const cfg = BRANCH_CFG[branch] || BRANCH_CFG.HK;

  const pages = {
    dashboard:   <Dashboard costRows={costRows} branch={branch} month={month} setPage={setPage}/>,
    products:    <Products products={products}/>,
    importAnag:  <ImportBC products={products} setProducts={setProducts} importLogs={importLogs} setImportLogs={setImportLogs} snapshots={snapshots} setSnapshots={setSnapshots} showToast={showToast} bumpImportTs={bumpImportTs}/>,
    xref:        <XRefPage xrefs={xrefs} setXrefs={setXrefs} snapshots={snapshots} setSnapshots={setSnapshots} importLogs={importLogs} setImportLogs={setImportLogs} showToast={showToast} bumpImportTs={bumpImportTs}/>,
    logistics:   <Logistics logistics={logistics} setLogistics={setLogistics} products={products} branch={branch} showToast={showToast} bumpImportTs={bumpImportTs}/>,
    prices:      <Prices prices={prices} products={products} branch={branch} month={month}/>,
    importPrice: <ImportPrices prices={prices} setPrices={setPrices} products={products} xrefs={xrefs} branch={branch} month={month} importLogs={importLogs} setImportLogs={setImportLogs} snapshots={snapshots} setSnapshots={setSnapshots} showToast={showToast} bumpImportTs={bumpImportTs}/>,
    fx:          <FxRates fx={fx} setFx={setFx} branch={branch} month={month}/>,
    air:         <AirListPage airList={airList} setAirList={setAirList} products={products} xrefs={xrefs} snapshots={snapshots} setSnapshots={setSnapshots} importLogs={importLogs} setImportLogs={setImportLogs} showToast={showToast} bumpImportTs={bumpImportTs}/>,
    costs:       <CostTable costRows={costRows} branch={branch} month={month} logistics={logistics} lastImportTs={lastImportTs} lastCalcTs={lastCalcTs} setLastCalcTs={setLastCalcTs} setCostHistory={setCostHistory}/>,
    sales:       <SalesInvoice rows={salesRows} setRows={setSalesRows} airList={airList} products={products} xrefs={xrefs} snapshots={snapshots} setSnapshots={setSnapshots} importLogs={importLogs} setImportLogs={setImportLogs} showToast={showToast} bumpImportTs={bumpImportTs}/>,
    storico:     <Storico snapshots={snapshots} setSnapshots={setSnapshots} costHistory={costHistory} branch={branch}/>,
    mail:        <MailGen costRows={costRows} branch={branch} month={month}/>,
    notes:       <NotesPage/>,
  };

  return (
    <div style={{display:"flex",height:"100vh",width:"100vw",background:T.bg,fontFamily:"'Palatino Linotype','Book Antiqua',Palatino,serif",color:T.text,overflow:"hidden"}}>
      {/* Sidebar */}
      <div style={{width:"200px",flexShrink:0,background:T.surface,borderRight:`1px solid ${T.border}`,display:"flex",flexDirection:"column",height:"100vh",overflow:"hidden"}}>
        <div style={{padding:"18px 16px 14px",borderBottom:`1px solid ${T.border}`}}>
          <div style={{fontSize:"9px",letterSpacing:"3px",color:T.gold,textTransform:"uppercase",marginBottom:"3px"}}>IFB Platform</div>
          <div style={{fontSize:"14px",fontWeight:"bold",lineHeight:1.2}}>Cost Intelligence</div>
        </div>
        {/* Branch switcher */}
        <div style={{padding:"10px 12px",borderBottom:`1px solid ${T.border}`}}>
          <div style={{fontSize:"9px",letterSpacing:"2px",color:T.dim,textTransform:"uppercase",marginBottom:"6px"}}>Filiale</div>
          <div style={{display:"flex",flexDirection:"column",gap:"3px"}}>
            {Object.entries(BRANCH_CFG).map(([key,c])=>(
              <button key={key} onClick={()=>setBranch(key)}
                style={{padding:"5px 8px",background:branch===key?`${c.color}20`:"transparent",
                  border:`1px solid ${branch===key?c.color:"transparent"}`,borderRadius:"6px",
                  color:branch===key?c.color:T.muted,cursor:"pointer",fontFamily:"inherit",
                  fontSize:"11px",textAlign:"left",display:"flex",alignItems:"center",gap:"6px"}}>
                <span>{c.flag}</span>{c.label}
                {!c.active&&<span style={{fontSize:"7px",color:c.color,marginLeft:"auto",background:`${c.color}22`,padding:"1px 4px",borderRadius:"3px"}}>SOON</span>}
              </button>
            ))}
          </div>
        </div>
        {/* Month */}
        <div style={{padding:"10px 12px",borderBottom:`1px solid ${T.border}`}}>
          <div style={{fontSize:"9px",letterSpacing:"2px",color:T.dim,textTransform:"uppercase",marginBottom:"5px"}}>Mese</div>
          <input type="month" value={month} onChange={e=>setMonth(e.target.value)}
            style={{width:"100%",padding:"5px 7px",background:"rgba(255,255,255,0.05)",border:`1px solid ${T.border}`,borderRadius:"6px",color:T.text,fontFamily:"inherit",fontSize:"11px",outline:"none",boxSizing:"border-box"}}/>
        </div>
        <nav style={{flex:1,padding:"8px",display:"flex",flexDirection:"column",gap:"1px",overflowY:"auto"}}>
          {NAV.map(n=>(
            <button key={n.id} onClick={()=>setPage(n.id)}
              style={{padding:"7px 10px",background:page===n.id?T.goldDim:"transparent",
                border:`1px solid ${page===n.id?T.gold+"44":"transparent"}`,borderRadius:"6px",
                color:page===n.id?T.gold:T.muted,cursor:"pointer",fontFamily:"inherit",
                fontSize:"11px",textAlign:"left",display:"flex",alignItems:"center",gap:"7px"}}>
              <span style={{fontSize:"10px",opacity:0.8}}>{n.icon}</span>{n.label}
              {n.badge&&<span style={{marginLeft:"auto",fontSize:"7px",background:`${T.blue}33`,color:T.blue,padding:"1px 4px",borderRadius:"3px"}}>{n.badge}</span>}
            </button>
          ))}
          <button onClick={()=>setPage("branchSelect")}
            style={{padding:"7px 10px",background:"transparent",border:`1px solid transparent`,
              borderRadius:"6px",color:T.dim,cursor:"pointer",fontFamily:"inherit",
              fontSize:"11px",textAlign:"left",display:"flex",alignItems:"center",gap:"7px",marginTop:"8px"}}>
            <span style={{fontSize:"10px"}}>⇦</span>Cambia filiale
          </button>
        </nav>
        <div style={{padding:"10px 12px",borderTop:`1px solid ${T.border}`,fontSize:"9px",color:T.dim,lineHeight:1.5}}>
          Inalca Food & Beverage<br/>© 2026 · v4.0
        </div>
      </div>
      {/* Main content */}
      <div style={{flex:1,display:"flex",flexDirection:"column",minWidth:0,overflow:"hidden"}}>
        <div style={{padding:"10px 24px",borderBottom:`1px solid ${T.border}`,background:T.surface,
          display:"flex",alignItems:"center",gap:"10px",flexShrink:0,zIndex:10,flexWrap:"wrap"}}>
          <span style={{fontSize:"16px"}}>{cfg.flag}</span>
          <span style={{fontSize:"13px",fontWeight:"bold",color:cfg.color}}>{cfg.label}</span>
          <span style={{color:T.dim}}>·</span>
          <span style={{fontSize:"12px",color:T.muted}}>{NAV.find(n=>n.id===page)?.label}</span>
          <span style={{color:T.dim}}>·</span>
          <span style={{fontSize:"11px",color:T.gold}}>{month}</span>
          {needsRecalc&&<span style={{padding:"2px 10px",background:`${T.orange}20`,color:T.orange,borderRadius:"10px",fontSize:"11px"}}>⚠ Nuovi dati — ricalcola Standard Cost</span>}
          <div style={{marginLeft:"auto",display:"flex",gap:"6px"}}>
            <button onClick={()=>setPage("importAnag")} style={{padding:"5px 12px",background:`${T.blue}15`,border:`1px solid ${T.blue}44`,borderRadius:"5px",color:T.blue,cursor:"pointer",fontFamily:"inherit",fontSize:"10px"}}>⇪ Anagrafica</button>
            <button onClick={()=>setPage("importPrice")} style={{padding:"5px 12px",background:`${T.purple}15`,border:`1px solid ${T.purple}44`,borderRadius:"5px",color:T.purple,cursor:"pointer",fontFamily:"inherit",fontSize:"10px"}}>💶 Listini</button>
            <button onClick={()=>setPage("costs")} style={{padding:"5px 12px",background:"rgba(255,255,255,0.05)",border:`1px solid ${T.border}`,borderRadius:"5px",color:T.muted,cursor:"pointer",fontFamily:"inherit",fontSize:"10px"}}>◆ Costi</button>
            <button onClick={()=>setPage("mail")} style={{padding:"5px 12px",background:T.gold,border:"none",borderRadius:"5px",color:T.bg,cursor:"pointer",fontFamily:"inherit",fontSize:"10px",fontWeight:"bold"}}>✉ Mail</button>
          </div>
        </div>
        <div style={{flex:1,padding:"20px 28px",overflow:"auto",width:"calc(100% - 200px)",boxSizing:"border-box"}}>{pages[page]}</div>
      </div>
      {toast&&<div style={{position:"fixed",bottom:"24px",right:"24px",padding:"10px 18px",background:toast.color,borderRadius:"8px",color:"#fff",fontSize:"12px",fontWeight:"bold",boxShadow:"0 8px 24px rgba(0,0,0,0.4)",zIndex:1000}}>{toast.msg}</div>}
    </div>
  );
}

// ─── XREF PAGE ────────────────────────────────────────────────────────────────
function XRefPage({xrefs,setXrefs,snapshots,setSnapshots,importLogs,setImportLogs,showToast,bumpImportTs}) {
  const[step,setStep]=useState("main");
  const[rawRows,setRawRows]=useState([]);
  const[headers,setHeaders]=useState([]);
  const[fileName,setFileName]=useState("");
  const[colNHK,setColNHK]=useState("");
  const[colIFB,setColIFB]=useState("");
  const[preview,setPreview]=useState([]);
  const[search,setSearch]=useState("");

  function parseFile(file) {
    setFileName(file.name);
    const reader=new FileReader();
    reader.onload=e=>{
      try{
        const wb=XLSX.read(e.target.result,{type:"binary"});
        const ws=wb.Sheets[wb.SheetNames[0]];
        const data=XLSX.utils.sheet_to_json(ws,{header:1,defval:""});
        if(data.length<2){showToast("File vuoto",T.red);return;}
        const hdrs=data[0].map(h=>String(h).trim()).filter(h=>h);
        const rows=data.slice(1).filter(r=>r.some(c=>c!==""));
        setHeaders(hdrs);setRawRows(rows);
        const nhkA=["n hk","nhk","hk","n_hk","gc code","gc no","hk code","hk no","hong kong"];
        const ifbA=["ifb n","ifb no","ifb no.","no_","no.","no","codice","code","item no"];
        setColNHK(hdrs.find(h=>nhkA.some(a=>h.toLowerCase().includes(a)))||"");
        setColIFB(hdrs.find(h=>ifbA.some(a=>h.toLowerCase()===a||h.toLowerCase().includes(a)))||"");
        setStep("map");
      }catch(err){showToast("Errore: "+err.message,T.red);}
    };
    reader.readAsBinaryString(file);
  }

  function buildPreview() {
    const iN=headers.indexOf(colNHK),iI=headers.indexOf(colIFB);
    const mapped=rawRows.map((row,idx)=>{
      const nHK=String(row[iN]||"").trim(),ifbNo=String(row[iI]||"").trim();
      if(!nHK&&!ifbNo) return null;
      const ex=xrefs.find(x=>x.nHK===nHK);
      return{_idx:idx,nHK,ifbNo,_isNew:!ex,_changed:ex&&ex.ifbNo!==ifbNo,_oldIFB:ex?.ifbNo};
    }).filter(Boolean);
    setPreview(mapped);setStep("preview");
  }

  function executeImport() {
    const id=Date.now();
    const incoming=preview.filter(r=>r.nHK&&r.ifbNo);
    const diffs=incoming.map(r=>({nHK:r.nHK,ifbNo:r.ifbNo,isNew:r._isNew,changed:r._changed,oldIFB:r._oldIFB}));
    const kept=xrefs.filter(x=>!incoming.find(i=>i.nHK===x.nHK));
    const next=[...incoming.map(r=>({nHK:r.nHK,ifbNo:r.ifbNo})),...kept];
    setXrefs(next);LS.set("ifb_xrefs",next);
    const log={id,type:"xref",fileName,date:new Date(id).toISOString(),count:incoming.length,diffs,branch:"ALL"};
    const newLogs=[log,...importLogs];setImportLogs(newLogs);LS.set("ifb_importlogs",newLogs);
    const newSnaps=[log,...snapshots].slice(0,50);setSnapshots(newSnaps);LS.set("ifb_snapshots",newSnaps);
    bumpImportTs();showToast(`XRef: ${incoming.length} voci · ${diffs.filter(d=>d.isNew).length} nuove ✓`,T.gold);
    setStep("main");setPreview([]);setRawRows([]);setHeaders([]);
  }

  const displayed=xrefs.filter(x=>!search||x.nHK?.toLowerCase().includes(search.toLowerCase())||x.ifbNo?.toLowerCase().includes(search.toLowerCase()));

  return(
    <div>
      <PageHeader title="⇄ Cross Reference N HK ↔ IFB N" sub="Tabella di corrispondenza per il matching automatico dei listini"/>
      {step==="map"&&(
        <Section title={`Mappatura — ${fileName} · ${rawRows.length} righe`}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"16px",marginBottom:"16px"}}>
            {[["Colonna N HK *",colNHK,setColNHK],["Colonna IFB N *",colIFB,setColIFB]].map(([lbl,val,setter])=>(
              <div key={lbl}>
                <label style={{display:"block",fontSize:"11px",color:T.gold,marginBottom:"5px"}}>{lbl}</label>
                <select value={val} onChange={e=>setter(e.target.value)} style={{...inputStyle(),cursor:"pointer"}}>
                  <option value="">— seleziona —</option>
                  {headers.map(h=><option key={h} value={h}>{h}</option>)}
                </select>
              </div>
            ))}
          </div>
          <div style={{display:"flex",gap:"10px"}}>
            <ActionBtn label="← Ricarica" onClick={()=>setStep("main")}/>
            <ActionBtn label="Preview →" onClick={buildPreview} primary disabled={!colNHK||!colIFB}/>
          </div>
        </Section>
      )}
      {step==="preview"&&(
        <div>
          <div style={{display:"flex",gap:"12px",marginBottom:"16px"}}>
            {[[preview.filter(r=>r._isNew).length,"Nuove",T.gold],[preview.filter(r=>r._changed).length,"Modificate",T.orange],[preview.filter(r=>!r._isNew&&!r._changed).length,"Invariate",T.dim],[preview.length,"Totale",T.text]].map(([n,l,c])=>(
              <div key={l} style={{padding:"10px 16px",background:T.card,border:`1px solid ${T.border}`,borderRadius:"8px"}}>
                <div style={{fontSize:"18px",fontWeight:"bold",color:c}}>{n}</div>
                <div style={{fontSize:"10px",color:T.dim,marginTop:"2px"}}>{l}</div>
              </div>
            ))}
          </div>
          <div style={{display:"flex",gap:"10px",marginBottom:"16px"}}>
            <ActionBtn label="← Torna" onClick={()=>setStep("map")}/>
            <ActionBtn label={`✓ Aggiorna XRef (${preview.filter(r=>r.nHK&&r.ifbNo).length} voci)`} onClick={executeImport} primary/>
          </div>
          <Section title="Preview (prime 50)">
            <table style={{width:"100%",borderCollapse:"collapse"}}>
              <THead cols={["N HK","IFB N","Stato"]}/>
              <tbody>{preview.slice(0,50).map(r=>(
                <tr key={r._idx} style={{borderBottom:`1px solid ${T.border}`,background:r._isNew?`${T.gold}07`:r._changed?`${T.orange}07`:""}}>
                  <TD mono><span style={{color:T.gold}}>{r.nHK}</span></TD>
                  <TD mono>{r.ifbNo}</TD>
                  <TD>{r._isNew?<Chip label="NUOVO" color={T.gold}/>:r._changed?<><Chip label="MODIF." color={T.orange}/><span style={{fontSize:"10px",color:T.dim,marginLeft:"6px"}}>{r._oldIFB}→{r.ifbNo}</span></>:<span style={{color:T.dim}}>=</span>}</TD>
                </tr>
              ))}</tbody>
            </table>
          </Section>
        </div>
      )}
      {step==="main"&&(
        <>
          <div style={{border:`2px dashed ${T.borderHi}`,borderRadius:"10px",padding:"20px 28px",textAlign:"center",cursor:"pointer",marginBottom:"20px"}}
            onClick={()=>document.getElementById("_xref_in")?.click()}>
            <div style={{fontSize:"24px",marginBottom:"6px"}}>⇄</div>
            <div style={{fontSize:"13px",color:T.text,marginBottom:"4px"}}>Carica file XRef (Excel/CSV)</div>
            <div style={{fontSize:"11px",color:T.muted}}>Due colonne: N HK · IFB N</div>
            <input id="_xref_in" type="file" accept=".xlsx,.xls,.csv"
              onChange={e=>{const f=e.target.files?.[0];if(f)parseFile(f);e.target.value="";}} style={{display:"none"}}/>
          </div>
          <SearchBar value={search} onChange={setSearch} placeholder="🔍 Cerca per N HK o IFB N…"/>
          <Section title={`${displayed.length} / ${xrefs.length} corrispondenze`}>
            {xrefs.length===0?<div style={{padding:"24px",textAlign:"center",color:T.dim,fontSize:"13px"}}>Nessuna XRef caricata.</div>:(
              <table style={{width:"100%",borderCollapse:"collapse"}}>
                <THead cols={["N HK","IFB N","Azioni"]}/>
                <tbody>{displayed.map((x,i)=>(
                  <tr key={x.nHK+i} style={{borderBottom:`1px solid ${T.border}`}}>
                    <TD mono><span style={{color:T.gold}}>{x.nHK}</span></TD>
                    <TD mono>{x.ifbNo}</TD>
                    <TD><MiniBtn label="✕" onClick={()=>{const n=xrefs.filter((_,j)=>j!==xrefs.indexOf(x));setXrefs(n);LS.set("ifb_xrefs",n);}} color={T.red}/></TD>
                  </tr>
                ))}</tbody>
              </table>
            )}
          </Section>
        </>
      )}
    </div>
  );
}

// ─── NOTES PAGE ───────────────────────────────────────────────────────────────
function NotesPage() {
  const sections=[
    {title:"🔴 Problemi noti",color:T.red,items:[
      "Import Listini: il codice nel file PBI deve essere N HK o IFB N — usa XRef per il matching automatico.",
      "Colonna 'No_' nel file CURRENT PRICELIST (ACQ) contiene a volte il codice HK (es. WCAN01-NV) — la mappatura automatica lo rileva.",
      "I codici P_BC_xxx vengono filtrati automaticamente (ID interni Power BI).",
      "Canarie (CAN), Macao (MAC) e Australia (AUS): calcoli costo non ancora attivati.",
    ]},
    {title:"🟡 Ambiguità",color:T.orange,items:[
      "Formula DAP Final: SE(DAP Disc≠0, DAP Disc, SE(isX, SE(WINE, FCA+carriage, FCADisc+carriage), 0)). isX dipende da FOR_VENDORS hardcodato.",
      "TASSA ALCOLICA: letta dalla colonna 'TASSA ALCOLICA' del Work_tab. Solo prodotti con GRADI ≥30° (SPIRITS).",
      "LIC = (4100+3800) HKD / tasso cambio / totalUnits — confermare se cambia ogni mese.",
      "Health Certificate: €80/container. Incluso se HEALTH CERTIFICATE = SI nel Work_tab.",
      "Standard Cost calcolato SOLO per articoli Vendor = INALCA FOOD & BEVERAGE, trasporto SEA (non AIR).",
    ]},
    {title:"🟢 Future",color:T.green,items:[
      "Export Standard Cost a Excel con struttura originale.",
      "Grafico andamento prezzi per prodotto.",
      "Gestione multi-container.",
      "Backup/restore completo.",
    ]},
  ];
  return(
    <div>
      <PageHeader title="📝 Note & Ambiguità" sub="Stato del progetto · regole di calcolo"/>
      {sections.map(s=>(
        <Section key={s.title} title={s.title} accent={s.color}>
          <ul style={{margin:0,padding:"0 0 0 18px"}}>
            {s.items.map((item,i)=><li key={i} style={{fontSize:"12px",color:T.muted,lineHeight:"1.8",marginBottom:"4px"}}>{item}</li>)}
          </ul>
        </Section>
      ))}
    </div>
  );
}

// ─── IMPORT LISTINI ───────────────────────────────────────────────────────────
function ImportPrices({prices,setPrices,products,xrefs,branch,month,importLogs,setImportLogs,snapshots,setSnapshots,showToast,bumpImportTs}) {
  const[step,setStep]=useState("upload");
  const[rawRows,setRawRows]=useState([]);
  const[headers,setHeaders]=useState([]);
  const[fileName,setFileName]=useState("");
  const[mapping,setMapping]=useState({});
  const[preview,setPreview]=useState([]);
  const[importMonth,setImportMonth]=useState(month);
  const[doneInfo,setDoneInfo]=useState(null);

  function parseFile(file) {
    setFileName(file.name);
    const reader=new FileReader();
    reader.onload=e=>{
      try{
        const wb=XLSX.read(e.target.result,{type:"binary"});
        const ws=wb.Sheets[wb.SheetNames[0]];
        const data=XLSX.utils.sheet_to_json(ws,{header:1,defval:""});
        if(data.length<2){showToast("File vuoto",T.red);return;}
        const hdrs=data[0].map(h=>String(h).trim()).filter(h=>h);
        const rows=data.slice(1).filter(r=>r.some(c=>c!==""));
        setHeaders(hdrs);setRawRows(rows);
        const am={};
        Object.keys(PRICE_FIELD_ALIASES).forEach(field=>{
          const aliases=PRICE_FIELD_ALIASES[field];
          for(const h of hdrs){const hl=h.toLowerCase().trim();if(aliases.some(a=>hl===a)){am[field]=h;break;}}
          if(!am[field]) for(const h of hdrs){const hl=h.toLowerCase().trim();if(aliases.some(a=>hl.includes(a)&&a.length>3)){am[field]=h;break;}}
        });
        setMapping(am);setStep("map");
      }catch(err){showToast("Errore: "+err.message,T.red);}
    };
    reader.readAsBinaryString(file);
  }

  function buildPreview() {
    const get=(row,field)=>{const col=mapping[field];if(!col)return null;const i=headers.indexOf(col);return i>=0?row[i]:null;};
    let skipped=0;
    const mapped=rawRows.map((row,idx)=>{
      const rawCode=String(get(row,"code")||"").trim();
      if(!rawCode) return null;
      // Filter internal Power BI IDs: P_BC_xxx, P_xxx, or pure long numeric strings (>8 digits)
      if(/^P_/i.test(rawCode) || /^\d{7,}$/.test(rawCode.replace(/[^0-9]/g,""))){skipped++;return null;}
      const prod=findProduct(rawCode,products,xrefs);
      if(!prod){skipped++;return null;}
      

      const mtsPrice=parseFloat(get(row,"mtsPrice"))||0;
      const fcaPrice=parseFloat(get(row,"fcaPrice"))||0;
      const fcaDiscount=parseFloat(get(row,"fcaDiscount"))||0;
      const fcaDiscounted=parseFloat(get(row,"fcaDiscounted"))||(fcaPrice-fcaDiscount*fcaPrice/100)||0;
      const dapPrice=parseFloat(get(row,"dapPrice"))||0;
      const dapDiscount=parseFloat(get(row,"dapDiscount"))||0;
      const dapDiscounted=parseFloat(get(row,"dapDiscounted"))||(dapPrice-dapDiscount*dapPrice/100)||0;
      const vendorName=String(get(row,"vendorName")||"").trim();
      const section=String(get(row,"section")||"").trim();
      const dapFinalDirect=parseFloat(get(row,"dapFinalDirect"))||0;
      let dapFinal=0,dapNote="";
      if(dapFinalDirect!==0){dapFinal=dapFinalDirect;dapNote="da file";}
      else{const calc=calcDAPFinal({dapDiscounted,fcaPrice,fcaDiscounted,vendorName,section,products,code:prod.code});dapFinal=calc.dapFinal;dapNote=calc.note;}
      const existing=prices.find(p=>p.productId===prod.id&&p.branch===branch&&p.month===importMonth);
      return{_idx:idx,rawCode,productId:prod.id,nHK:prod.nHK||"—",ifbNo:prod.code,description:prod.description,
        dapFinal:roundN(dapFinal),mtsPrice:roundN(mtsPrice),fcaDiscounted:roundN(fcaDiscounted),
        dapPrice:roundN(dapPrice),fcaPrice:roundN(fcaPrice),dapNote,_hasProduct:true,_existing:!!existing};
    }).filter(Boolean);
    setPreview(mapped);
    window._priceSkipped=skipped;
    setStep("preview");
  }

  function executeImport() {
    const snId=Date.now();
    const updated=[...prices];
    const diffs=[];
    let count=0,newCount=0,changed=0;
    preview.filter(r=>r._hasProduct).forEach(r=>{
      const idx=updated.findIndex(p=>p.productId===r.productId&&p.branch===branch&&p.month===importMonth);
      const entry={productId:r.productId,branch,month:importMonth,dapFinal:r.dapFinal,mtsPrice:r.mtsPrice,fcaDiscounted:r.fcaDiscounted,dapPrice:r.dapPrice,fcaPrice:r.fcaPrice};
      const prev=idx>=0?updated[idx]:null;
      const diffFields=[];
      ["dapFinal","mtsPrice","fcaDiscounted","dapPrice","fcaPrice"].forEach(f=>{
        const oldR=roundN(prev?.[f]||0),newR=roundN(entry[f]||0);
        if(Math.abs(oldR-newR)>=0.005) diffFields.push({field:f,old:oldR,new:newR,delta:oldR>0?((newR-oldR)/oldR*100):null});
      });
      if(!prev) newCount++; else if(diffFields.length>0) changed++;
      if(diffFields.length>0||!prev) diffs.push({productId:r.productId,nHK:r.nHK,ifbNo:r.ifbNo,description:r.description,isNew:!prev,fields:diffFields});
      if(idx>=0) updated[idx]=entry; else updated.push(entry);
      count++;
    });
    setPrices(updated);LS.set("ifb_prices",updated);
    const log={id:snId,type:"prices",fileName,branch,month:importMonth,date:new Date(snId).toISOString(),count,newCount,updateCount:changed,diffs};
    const newLogs=[log,...importLogs];setImportLogs(newLogs);LS.set("ifb_importlogs",newLogs);
    const newSnaps=[log,...snapshots].slice(0,50);setSnapshots(newSnaps);LS.set("ifb_snapshots",newSnaps);
    setDoneInfo({count,newCount,changed,unchanged:count-newCount-changed});
    bumpImportTs();setStep("done");
  }

  const reset=()=>{setStep("upload");setRawRows([]);setHeaders([]);setFileName("");setMapping({});setPreview([]);setDoneInfo(null);};

  if(step==="done"&&doneInfo) return(
    <div>
      <PageHeader title="✓ Import Listini completato" sub={fileName}/>
      <div style={{padding:"20px",background:`${T.green}11`,border:`1px solid ${T.green}33`,borderRadius:"8px",marginBottom:"16px",fontSize:"13px",color:T.muted,lineHeight:"2"}}>
        Mese: <strong style={{color:T.gold}}>{importMonth}</strong> · Filiale: <strong style={{color:T.text}}>{branch}</strong><br/>
        Prezzi totali: <strong style={{color:T.text}}>{doneInfo.count}</strong> &nbsp;·&nbsp;
        <span style={{color:T.green}}>🆕 {doneInfo.newCount} nuovi</span> &nbsp;·&nbsp;
        <span style={{color:T.orange}}>✏️ {doneInfo.changed} modificati</span> &nbsp;·&nbsp;
        <span style={{color:T.dim}}>{doneInfo.unchanged} invariati</span>
      </div>
      <ActionBtn label="💶 Nuovo import" onClick={reset} primary/>
    </div>
  );

  return(
    <div>
      <PageHeader title="💶 Import Listini" sub={`${branch} · filtra automaticamente Vendor = INALCA F&B`}/>
      <StepBar steps={["upload","map","preview","done"]} current={step}/>
      {step==="upload"&&(
        <div>
          <Section title="Mese di riferimento">
            <div style={{display:"flex",alignItems:"center",gap:"12px"}}>
              <span style={{fontSize:"12px",color:T.muted}}>Mese:</span>
              <input type="month" value={importMonth} onChange={e=>setImportMonth(e.target.value)} style={{...inputStyle(),width:"160px"}}/>
            </div>
          </Section>
          <Section title="Carica file export PBI / CURRENT PRICELIST">
            <DropZone onFile={parseFile}/>
          </Section>
        </div>
      )}
      {step==="map"&&(
        <Section title={`Mappatura — ${fileName} · ${rawRows.length} righe`}>
          <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:"12px",marginBottom:"18px"}}>
            {Object.keys(PRICE_FIELD_ALIASES).map(field=>{
              const labels={code:"Codice * (N HK o IFB N)",vendorName:"Vendor Name",section:"Section",mtsPrice:"MTS Price",fcaPrice:"FCA Price",fcaDiscount:"FCA Discount %",fcaDiscounted:"FCA Discounted",dapPrice:"DAP Price",dapDiscount:"DAP Discount %",dapDiscounted:"DAP Discounted",dapFinalDirect:"DAP Final (già calcolato)"};
              return(
                <div key={field}>
                  <label style={{display:"block",fontSize:"11px",color:field==="code"?T.gold:T.muted,marginBottom:"5px"}}>{labels[field]}</label>
                  <select value={mapping[field]||""} onChange={e=>setMapping(m=>({...m,[field]:e.target.value||null}))} style={{...inputStyle(),cursor:"pointer",borderColor:!mapping[field]&&field==="code"?T.red+"88":T.border}}>
                    <option value="">— non mappato —</option>
                    {headers.map(h=><option key={h} value={h}>{h}</option>)}
                  </select>
                </div>
              );
            })}
          </div>
          <div style={{display:"flex",gap:"10px"}}>
            <ActionBtn label="← Ricarica" onClick={reset}/>
            <ActionBtn label="Preview →" onClick={buildPreview} primary disabled={!mapping["code"]}/>
          </div>
        </Section>
      )}
      {step==="preview"&&(
        <div>
          {step === "preview" && (
  <div>
    <div style={{ display: "flex", gap: "12px", marginBottom: "16px", flexWrap: "wrap" }}>
      {[
        [preview.length, "Trovati e importabili", T.green],
        [window._priceSkipped || 0, "Ignorati (non trovati / ID interni Power BI)", T.muted],
        [preview.filter(r => r._existing).length, "Aggiornamenti", T.orange]
      ].map(([n, l, c]) => (
        <div key={l} style={{ padding: "10px 16px", background: T.card, border: `1px solid ${T.border}`, borderRadius: "8px" }}>
          <div style={{ fontSize: "20px", fontWeight: "bold", color: c }}>{n}</div>
          <div style={{ fontSize: "10px", color: T.dim, marginTop: "2px" }}>{l}</div>
        </div>
      ))}
    </div>
    {/* Resto del codice uguale... */}
  </div>
)}
          <div style={{display:"flex",gap:"10px",marginBottom:"16px"}}>
            <ActionBtn label="← Torna" onClick={()=>setStep("map")}/>
            <ActionBtn label={`✓ Importa ${preview.filter(r=>r._hasProduct).length} prezzi per ${importMonth}`} onClick={executeImport} primary/>
          </div>
          <Section title={`Preview · ${importMonth} · ${branch}`}>
            <div style={{overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse"}}>
                <THead cols={["N HK","IFB N","Descrizione","DAP Final","MTS Price","FCA Disc.","Stato"]}/>
                <tbody>{preview.map(r=>(
                  <tr key={r._idx} style={{borderBottom:`1px solid ${T.border}`,opacity:r._hasProduct?1:0.4,background:r._existing?`${T.orange}08`:""}}>
                    <TD mono><span style={{color:T.gold}}>{r.nHK||"—"}</span></TD>
                    <TD mono>{r.ifbNo}</TD>
                    <TD>{r.description}</TD>
                    <TD mono><span style={{color:T.gold}}>{r.dapFinal>0?`€ ${r.dapFinal.toFixed(2)}`:"—"}</span>{r.dapNote&&<span style={{marginLeft:"4px",fontSize:"9px",color:T.dim}}>({r.dapNote})</span>}</TD>
                    <TD mono><span style={{color:T.blue}}>{r.mtsPrice>0?`€ ${r.mtsPrice.toFixed(2)}`:"—"}</span></TD>
                    <TD mono><span style={{color:T.muted}}>{r.fcaDiscounted>0?`€ ${r.fcaDiscounted.toFixed(2)}`:"—"}</span></TD>
                    <TD>{!r._hasProduct?<Chip label="NOT FOUND" color={T.red}/>:r._existing?<Chip label="AGGIORNA" color={T.orange}/>:<Chip label="NUOVO" color={T.green}/>}</TD>
                  </tr>
                ))}</tbody>
              </table>
            </div>
          </Section>
        </div>
      )}
    </div>
  );
}

// ─── IMPORT ANAGRAFICA BC ─────────────────────────────────────────────────────
function ImportBC({products,setProducts,importLogs,setImportLogs,snapshots,setSnapshots,showToast,bumpImportTs}) {
  const[step,setStep]=useState("upload");
  const[headers,setHeaders]=useState([]);
  const[rows,setRows]=useState([]);
  const[map,setMap]=useState({});
  const[preview,setPreview]=useState([]);
  const[doneInfo,setDoneInfo]=useState(null);
  const[fileName,setFileName]=useState("");

  // Added vendorName AND vendorName2
  const FIELDS=["nHK","code","description","category","uom","qtyPerBox","boxPerPallet","kgPerBox","temperature","active","vendorName","vendorName2"];
  const FLABELS={nHK:"N HK (No_)",code:"IFB Item *",description:"Descrizione *",category:"Section",uom:"UOM",qtyPerBox:"Qty/Cartone",boxPerPallet:"Cartoni/Pallet",kgPerBox:"Kg per Cartone / Net Weight",temperature:"Product Type",active:"Bloccato",vendorName:"Vendor Name (es. INALCA FOOD & BEVERAGE SRL)",vendorName2:"Vendor Name 2 (fornitore reale)"};

  const LOCAL_ALIASES = {
    nHK:         ["no","no_"],          // Anagrafica 'no' column = N HK
    code:        ["ifbitem","ifb item","ifb no","ifb n"],
    description: ["description"],
    category:    ["sectiondescription","section description","section"],
    uom:         ["salesunitofmeasure","sales unit of measure"],
    qtyPerBox:   ["quantityxpackaging","quantity x packaging"],
    boxPerPallet:["packagingxpallet","packaging x pallet"],
    kgPerBox:    ["netweight","net weight"],
    temperature: ["producttype","product type","product type rettificato","product type - anagrafica"],
    active:      ["blocked"],
    vendorName:  ["vendorname","vendor name"],    // exact match: 'vendorname' col in Anagrafica
    vendorName2: ["vendorname2","vendor name 2"],
  };

  function autoMap(hdrs) {
    const m: any={};
    for(const field of FIELDS){
      const aliases = LOCAL_ALIASES[field]||[];
      // Try exact match first (normalized)
      const h = hdrs.find(h=>aliases.some((a: string)=>h.toLowerCase().replace(/[\s_]/g,"")=== a.replace(/[\s_]/g,"")));
      if(h){ m[field]=h; continue; }
      // Then partial match
      const h2 = hdrs.find(h=>aliases.some((a: string)=>h.toLowerCase().replace(/[\s_]/g,"").includes(a.replace(/[\s_]/g,""))&&a.length>3));
      if(h2) m[field]=h2;
    }
    return m;
  }

  function parseFile(e) {
    const file=e.target.files?.[0]; if(!file) return;
    setFileName(file.name);
    const reader=new FileReader();
    reader.onload=ev=>{
      try{
        const wb=XLSX.read((ev.target as any).result,{type:"binary"});
        const ws=wb.Sheets[wb.SheetNames[0]];
        const data: any[][]=XLSX.utils.sheet_to_json(ws,{header:1,defval:""});
        if(!data.length){showToast("File vuoto",T.red);return;}
        // Find the actual header row: look for row containing 'no', 'ifbitem', 'vendorname'
        let hi=0;
        for(let i=0;i<Math.min(5,data.length);i++){
          const rNorm = data[i].map((c:any)=>String(c||"").toLowerCase().replace(/[\s_]/g,""));
          if(rNorm.some((c:string)=>["no","ifbitem","vendorname","description"].includes(c))){hi=i;break;}
        }
        const hdrs=data[hi].map((h:any)=>String(h||"").trim());
        setHeaders(hdrs);setRows(data.slice(hi+1).filter((r:any[])=>r.some((c:any)=>c!=="")));setMap(autoMap(hdrs));setStep("map");
      }catch(err:any){showToast("Errore lettura file",T.red);}
    };
    reader.readAsBinaryString(file);
  }

  function buildPreview() {
    const mapped=rows.map(r=>{
      const obj={};
      for(const field of FIELDS){const col=map[field];if(col){const idx=headers.indexOf(col);obj[field]=idx>=0?String(r[idx]||"").trim():""}else obj[field]="";}
      return obj;
    }).filter(r=>r.code||r.nHK);
    setPreview(mapped);setStep("preview");
  }

  function executeImport() {
    const now=Date.now();
    const newProds=preview.map((r)=>({
      id:r.code||r.nHK, code:r.code, nHK:r.nHK, description:r.description,
      category:mapBCVal("category",r.category), uom:mapBCVal("uom",r.uom),
      qtyPerBox:parseFloat(r.qtyPerBox)||0, boxPerPallet:parseFloat(r.boxPerPallet)||0,
      kgPerBox:parseFloat(r.kgPerBox)||0, temperature:mapBCVal("temperature",r.temperature),
      active:!["true","1","yes"].includes(String(r.active||"").toLowerCase()),
      vendorName: r.vendorName || "",
      vendorName2: r.vendorName2 || "",
    }));
    const prevMap=Object.fromEntries(products.map(p=>[p.id,p]));
    const diffs=[];
    for(const p of newProds){
      const old=prevMap[p.id];
      if(!old) diffs.push({id:p.id,isNew:true,description:p.description,fields:[]});
      else{
        const fields=[];
        for(const k of["description","category","uom","qtyPerBox","boxPerPallet","kgPerBox","temperature","active","vendorName"]){
          if(String(old[k])!==String(p[k])) fields.push({field:k,old:old[k],new:p[k]});
        }
        if(fields.length) diffs.push({id:p.id,isNew:false,description:p.description,fields});
      }
    }
    const snap={id:now,type:"anagrafica",date:new Date(now).toISOString(),count:newProds.length,diffs,branch:"ALL"};
    const newSnaps=[snap,...snapshots].slice(0,50);setSnapshots(newSnaps);LS.set("ifb_snapshots",newSnaps);
    setProducts(newProds);LS.set("ifb_products",newProds);
    const log={id:now,type:"anagrafica",date:new Date(now).toISOString(),msg:`Importati ${newProds.length} articoli`};
    const newLogs=[log,...importLogs];setImportLogs(newLogs);LS.set("ifb_importlogs",newLogs);
    const newCount=diffs.filter(d=>d.isNew).length,changed=diffs.filter(d=>!d.isNew&&d.fields.length>0).length;
    setDoneInfo({count:newProds.length,newCount,changed,unchanged:newProds.length-newCount-changed});
    bumpImportTs();setStep("done");
    showToast(`Importati ${newProds.length} articoli`,T.gold);
  }

  const mapBCVal=(field,raw)=>{
    const maps={
      category:{"food":"FOOD","alimenti":"FOOD","beverage":"WINE","wine":"WINE","spirits":"SPIRITS","vino":"WINE","meat":"MEAT","carni":"MEAT","salumi":"MEAT","milk and dairy products":"FOOD","cow cheese":"FOOD","sheep cheese":"FOOD","stretched-curd cheese":"FOOD","pork meat":"MEAT","ham":"MEAT","other cured meats":"MEAT","poultry and rabbit meat":"MEAT","egg products":"FOOD","eggs":"FOOD","flour and groats":"FOOD","preserved fish":"MEAT","fish processing":"MEAT","shellfish":"MEAT","molluscs and mussels":"MEAT","soft drinks":"FOOD","oil and fats":"FOOD","pasta and rice":"FOOD","condiments":"FOOD","meat processing":"MEAT"},
      uom:{"pcs":"PCS","pz":"PCS","piece":"PCS","pezzi":"PCS","box":"BOX","ctn":"BOX","cartone":"BOX","collo":"BOX","kg":"KG","kgs":"KG","kilogram":"KG"},
      temperature:{"dry":"DRY","secco":"DRY","ambient":"DRY","amb":"DRY","fresh":"FRESH","fresco":"FRESH","chilled":"FRESH","refrigerated":"FRESH","frozen":"FROZEN","surgelato":"FROZEN","congelato":"FROZEN"},
    };
    if(!maps[field]) return raw;
    return maps[field][String(raw||"").toLowerCase().trim()]||raw;
  };

  if(step==="done"&&doneInfo) return(
    <div>
      <PageHeader title="✓ Anagrafica importata" sub={fileName}/>
      <div style={{padding:"20px",background:`${T.green}11`,border:`1px solid ${T.green}33`,borderRadius:"8px",marginBottom:"16px",fontSize:"13px",lineHeight:"2"}}>
        Articoli totali: <strong style={{color:T.text}}>{doneInfo.count}</strong> &nbsp;·&nbsp;
        <span style={{color:T.green}}>🆕 {doneInfo.newCount} nuovi</span> &nbsp;·&nbsp;
        <span style={{color:T.orange}}>✏️ {doneInfo.changed} modificati</span> &nbsp;·&nbsp;
        <span style={{color:T.dim}}>{doneInfo.unchanged} invariati</span>
      </div>
      <ActionBtn label="Nuova importazione" onClick={()=>{setStep("upload");setDoneInfo(null);}}/>
    </div>
  );

  if(step==="preview") return(
    <div>
      <PageHeader title={`Preview Anagrafica · ${fileName}`} sub={`${preview.length} articoli`}/>
      <div style={{display:"flex",gap:"10px",marginBottom:"16px"}}>
        <ActionBtn label="← Mappa" onClick={()=>setStep("map")}/>
        <ActionBtn label={`✓ Importa ${preview.length} articoli`} onClick={executeImport} primary/>
      </div>
      <div style={{overflowX:"auto"}}>
        <table style={{width:"100%",borderCollapse:"collapse"}}>
          <THead cols={["N HK","IFB No","Descrizione","Vendor","Categoria","UOM","Qty/Box","Box/Plt","Kg/Box","Temp","Attivo"]}/>
          <tbody>{preview.slice(0,200).map((r,i)=>(
            <tr key={i} style={{borderBottom:`1px solid ${T.border}`}}>
              <TD mono><span style={{color:T.muted}}>{r.nHK||"—"}</span></TD>
              <TD mono><span style={{color:T.gold}}>{r.code}</span></TD>
              <TD>{r.description}</TD>
              <TD><span style={{fontSize:"11px",color:isIFBVendor(r.vendorName)?T.gold:T.muted}}>{r.vendorName||"—"}</span></TD>
              <TD>{r.category}</TD><TD>{r.uom}</TD>
              <TD mono>{r.qtyPerBox}</TD><TD mono>{r.boxPerPallet}</TD><TD mono>{r.kgPerBox}</TD>
              <TD>{r.temperature}</TD><TD>{r.active}</TD>
            </tr>
          ))}</tbody>
        </table>
      </div>
    </div>
  );

  if(step==="map") return(
    <div>
      <PageHeader title={`Mappatura Anagrafica · ${fileName}`} sub={`${rows.length} righe · mappatura auto da export BC`}/>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"12px",maxWidth:"700px",marginBottom:"20px"}}>
        {FIELDS.map(f=>(
          <div key={f}>
            <label style={{display:"block",fontSize:"11px",color:f==="code"||f==="description"?T.gold:T.muted,marginBottom:"5px"}}>{FLABELS[f]}</label>
            <select value={map[f]||""} onChange={e=>setMap(m=>({...m,[f]:e.target.value}))} style={{...inputStyle(),borderColor:map[f]?T.gold:T.border}}>
              <option value="">— non mappato —</option>
              {headers.map(h=><option key={h} value={h}>{h}</option>)}
            </select>
          </div>
        ))}
      </div>
      <div style={{display:"flex",gap:"10px"}}>
        <ActionBtn label="← Ricarica" onClick={()=>setStep("upload")}/>
        <ActionBtn label="Preview →" onClick={buildPreview} primary disabled={!map["code"]&&!map["description"]}/>
      </div>
    </div>
  );

  return(
    <div>
      <PageHeader title="⇪ Import Anagrafica BC" sub="Carica l'export da Business Central. Rilevamento automatico colonne + campo Vendor."/>
      <div style={{marginTop:"12px"}}>
        <label style={{display:"inline-block",padding:"10px 20px",background:T.gold,color:"#000",borderRadius:"6px",cursor:"pointer",fontWeight:"bold"}}>
          📂 Scegli file
          <input type="file" accept=".xlsx,.xls,.csv" onChange={parseFile} style={{display:"none"}}/>
        </label>
      </div>
    </div>
  );
}

// ─── AIR LIST PAGE ────────────────────────────────────────────────────────────
function AirListPage({airList,setAirList,products,xrefs,snapshots,setSnapshots,importLogs,setImportLogs,showToast,bumpImportTs}) {
  const[step,setStep]=useState("main");
  const[headers,setHeaders]=useState([]);
  const[rawRows,setRawRows]=useState([]);
  const[colCode,setColCode]=useState("");
  const[colTransport,setColTransport]=useState("");
  const[preview,setPreview]=useState([]);
  const[fileName,setFileName]=useState("");
  const[search,setSearch]=useState("");

  function parseFile(file) {
    setFileName(file.name);
    const reader=new FileReader();
    reader.onload=e=>{
      try{
        const wb=XLSX.read(e.target.result,{type:"binary"});
        const ws=wb.Sheets[wb.SheetNames[0]];
        const data=XLSX.utils.sheet_to_json(ws,{header:1,defval:""});
        if(data.length<2){showToast("File vuoto",T.red);return;}
        const hdrs=data[0].map(h=>String(h).trim()).filter(h=>h);
        const rows=data.slice(1).filter(r=>r.some(c=>c!==""));
        setHeaders(hdrs);setRawRows(rows);
        const codeA=["n hk","nhk","no_","no.","no","item no","ifb no","ifb n","code","codice"];
        const transA=["transportation","transport","trasporto","tipo trasporto","shipping method","shipment method","transp"];
        setColCode(hdrs.find(h=>codeA.some(a=>h.toLowerCase()===a||h.toLowerCase().includes(a)))||"");
        setColTransport(hdrs.find(h=>transA.some(a=>h.toLowerCase().includes(a)))||"");
        setStep("map");
      }catch(err){showToast("Errore: "+err.message,T.red);}
    };
    reader.readAsBinaryString(file);
  }

  function buildPreview() {
    const iC=headers.indexOf(colCode),iT=headers.indexOf(colTransport);
    const mapped=rawRows.map(row=>{
      const code=String(row[iC]||"").trim();
      const transportation=String(row[iT]||"").trim();
      if(!code) return null;
      const prod=findProduct(code,products,xrefs);
      return{code,transportation,productId:prod?.id||null,description:prod?.description||code,nHK:prod?.nHK||"",isAir:isAirTransport(transportation),_hasProduct:!!prod};
    }).filter(Boolean);
    setPreview(mapped);setStep("preview");
  }

  function executeImport() {
    const airOnly=preview.filter(r=>r.isAir&&r._hasProduct);
    const kept=airList.filter(a=>!airOnly.find(r=>r.productId===a.productId));
    const next=[...kept,...airOnly.map(r=>({productId:r.productId,code:r.code,nHK:r.nHK,description:r.description,transportation:r.transportation}))];
    setAirList(next);LS.set("ifb_airlist",next);
    const now=Date.now();
    const log={id:now,type:"air",date:new Date(now).toISOString(),count:airOnly.length,diffs:[],branch:"HK"};
    const newLogs=[log,...importLogs];setImportLogs(newLogs);LS.set("ifb_importlogs",newLogs);
    const newSnaps=[log,...snapshots].slice(0,50);setSnapshots(newSnaps);LS.set("ifb_snapshots",newSnaps);
    bumpImportTs();showToast(`AIR: ${airOnly.length} articoli marcati ✓`,T.gold);
    setStep("main");setPreview([]);setRawRows([]);setHeaders([]);
  }

  const displayed=airList.filter(a=>!search||a.description?.toLowerCase().includes(search.toLowerCase())||a.code?.includes(search));

  return(
    <div>
      <PageHeader title="✈ AIR Transport" sub="Articoli trasportati via aerea — esclusi da Standard Cost (calcolo solo SEA)"/>
      {step==="map"&&(
        <Section title={`Mappatura — ${fileName}`}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"16px",marginBottom:"16px"}}>
            {[["Colonna Codice *",colCode,setColCode],["Colonna Transportation *",colTransport,setColTransport]].map(([lbl,val,setter])=>(
              <div key={lbl}>
                <label style={{display:"block",fontSize:"11px",color:T.gold,marginBottom:"5px"}}>{lbl}</label>
                <select value={val} onChange={e=>setter(e.target.value)} style={{...inputStyle(),cursor:"pointer"}}>
                  <option value="">— seleziona —</option>
                  {headers.map(h=><option key={h} value={h}>{h}</option>)}
                </select>
              </div>
            ))}
          </div>
          <div style={{display:"flex",gap:"10px"}}>
            <ActionBtn label="← Ricarica" onClick={()=>setStep("main")}/>
            <ActionBtn label="Preview →" onClick={buildPreview} primary disabled={!colCode||!colTransport}/>
          </div>
        </Section>
      )}
      {step==="preview"&&(
        <div>
          <div style={{display:"flex",gap:"12px",marginBottom:"16px"}}>
            {[[preview.filter(r=>r.isAir).length,"AIR",T.orange],[preview.filter(r=>!r.isAir).length,"Non AIR",T.dim],[preview.filter(r=>!r._hasProduct).length,"Non trovati",T.red],[preview.length,"Totale",T.text]].map(([n,l,c])=>(
              <div key={l} style={{padding:"10px 16px",background:T.card,border:`1px solid ${T.border}`,borderRadius:"8px"}}>
                <div style={{fontSize:"18px",fontWeight:"bold",color:c}}>{n}</div>
                <div style={{fontSize:"10px",color:T.dim,marginTop:"2px"}}>{l}</div>
              </div>
            ))}
          </div>
          <div style={{display:"flex",gap:"10px",marginBottom:"16px"}}>
            <ActionBtn label="← Torna" onClick={()=>setStep("map")}/>
            <ActionBtn label={`✓ Salva ${preview.filter(r=>r.isAir&&r._hasProduct).length} articoli AIR`} onClick={executeImport} primary/>
          </div>
          <Section title="Preview (solo AIR)">
            <table style={{width:"100%",borderCollapse:"collapse"}}>
              <THead cols={["Codice","N HK","Descrizione","Transportation","AIR?","Trovato"]}/>
              <tbody>{preview.filter(r=>r.isAir).map((r,i)=>(
                <tr key={i} style={{borderBottom:`1px solid ${T.border}`,opacity:r._hasProduct?1:0.5}}>
                  <TD mono><span style={{color:T.gold}}>{r.code}</span></TD>
                  <TD mono><span style={{color:T.muted}}>{r.nHK||"—"}</span></TD>
                  <TD>{r.description}</TD>
                  <TD><Chip label={r.transportation} color={T.orange}/></TD>
                  <TD><Chip label="✈ AIR" color={T.orange}/></TD>
                  <TD>{r._hasProduct?<Chip label="OK" color={T.green}/>:<Chip label="NOT FOUND" color={T.red}/>}</TD>
                </tr>
              ))}</tbody>
            </table>
          </Section>
        </div>
      )}
      {step==="main"&&(
        <>
          <div style={{marginBottom:"20px"}}>
            <label style={{display:"inline-block",padding:"10px 20px",background:T.gold,color:"#000",borderRadius:"6px",cursor:"pointer",fontWeight:"bold"}}>
              📂 Carica report Transportation
              <input type="file" accept=".xlsx,.xls,.csv"
                onChange={e=>{const f=e.target.files?.[0];if(f)parseFile(f);e.target.value="";}} style={{display:"none"}}/>
            </label>
            <span style={{marginLeft:"12px",fontSize:"12px",color:T.muted}}>Colonne richieste: codice articolo + Transportation</span>
          </div>
          {airList.length>0&&(
            <>
              <SearchBar value={search} onChange={setSearch} placeholder="🔍 Cerca articolo AIR…"/>
              <Section title={`${displayed.length} articoli AIR (esclusi da Standard Cost)`}>
                <table style={{width:"100%",borderCollapse:"collapse"}}>
                  <THead cols={["Codice","N HK","Descrizione","Transportation","Azioni"]}/>
                  <tbody>{displayed.map((a,i)=>(
                    <tr key={a.productId||i} style={{borderBottom:`1px solid ${T.border}`}}>
                      <TD mono><span style={{color:T.gold}}>{a.code}</span></TD>
                      <TD mono><span style={{color:T.muted}}>{a.nHK||"—"}</span></TD>
                      <TD>{a.description}</TD>
                      <TD><Chip label={a.transportation} color={T.orange}/></TD>
                      <TD><MiniBtn label="✕ Rimuovi" onClick={()=>{const n=airList.filter((_,j)=>j!==airList.indexOf(a));setAirList(n);LS.set("ifb_airlist",n);}} color={T.red}/></TD>
                    </tr>
                  ))}</tbody>
                </table>
              </Section>
            </>
          )}
        </>
      )}
    </div>
  );
}

// ─── DASHBOARD ────────────────────────────────────────────────────────────────
function Dashboard({costRows,branch,month,setPage}) {
  const flagged = costRows.filter(r=>r.cost?.step2Hkd!=null&&r.prevCost?.step2Hkd!=null&&r.prevCost.step2Hkd>0&&Math.abs((r.cost.step2Hkd-r.prevCost.step2Hkd)/r.prevCost.step2Hkd)>=0.03);
  const air      = costRows.filter(r=>r.isAir);
  const noPrice  = costRows.filter(r=>!r.cost&&!r.isAir&&r.skipReason?.includes("NO PREZZO"));
  const noLog    = costRows.filter(r=>!r.cost&&!r.isAir&&r.skipReason==="NO LOGISTICA");
  const calcZero = costRows.filter(r=>!r.cost&&!r.isAir&&r.skipReason?.includes("CALC=0"));

  return(
    <div>
      <PageHeader title={`Dashboard · ${branch} · ${month}`} sub="Solo articoli INALCA FOOD &amp; BEVERAGE · SEA"/>
      <div style={{display:"flex",gap:"12px",marginBottom:"20px",flexWrap:"wrap"}}>
        {[[costRows.filter(r=>r.cost?.step2Hkd!=null).length,"Costi calcolati",T.green],[flagged.length,"Variazioni ≥3%",T.orange],[air.length,"AIR (esclusi)",T.blue],[noPrice.length,"Senza prezzo",T.red],[noLog.length,"No logistica",T.red],[calcZero.length,"Calc=0 (check UOM/qty)",T.orange]].map(([n,l,c])=>(
          <div key={l} style={{padding:"12px 20px",background:T.card,border:`1px solid ${T.border}`,borderRadius:"8px",minWidth:"120px"}}>
            <div style={{fontSize:"22px",fontWeight:"bold",color:c}}>{n}</div>
            <div style={{fontSize:"11px",color:T.dim,marginTop:"2px"}}>{l}</div>
          </div>
        ))}
      </div>
      {flagged.length>0&&(
        <Section title={`${flagged.length} articoli con variazione ≥ ±3%`} accent={T.orange}>
          <table style={{width:"100%",borderCollapse:"collapse"}}>
            <THead cols={["N HK","IFB No","Descrizione","Prec. HKD","Nuovo HKD","Δ%"]}/>
            <tbody>{flagged.map((r,i)=>{
              const pct=(r.cost.step2Hkd-r.prevCost.step2Hkd)/r.prevCost.step2Hkd*100;
              return<tr key={r.id} style={{borderBottom:`1px solid ${T.border}`,background:i%2===0?T.bg:T.surface}}>
                <TD mono><span style={{color:T.muted}}>{r.nHK||"—"}</span></TD>
                <TD mono><span style={{color:T.gold}}>{r.code}</span></TD>
                <TD>{r.description}</TD>
                <TD mono>{r.prevCost.step2Hkd.toFixed(2)}</TD>
                <TD mono><span style={{color:T.gold,fontWeight:"bold"}}>{r.cost.step2Hkd.toFixed(2)}</span></TD>
                <TD><span style={{color:pct>0?T.red:T.green,fontWeight:"bold"}}>{pct>0?"+":""}{pct.toFixed(1)}%</span></TD>
              </tr>;
            })}</tbody>
          </table>
        </Section>
      )}
      {flagged.length===0&&<div style={{padding:"32px",textAlign:"center",color:T.muted,fontSize:"13px"}}>Nessuna variazione ≥ ±3% questo mese.</div>}
    </div>
  );
}

// ─── PRODUCTS ─────────────────────────────────────────────────────────────────
function Products({products}) {
  const[search,setSearch]=useState("");
  const[onlyIFB,setOnlyIFB]=useState(true);
  const base = onlyIFB ? products.filter(p=>isIFBVendor(p.vendorName)) : products;
  const filtered = base.filter(p=>!search||p.description?.toLowerCase().includes(search.toLowerCase())||p.code?.includes(search)||p.nHK?.includes(search));
  return(
    <div>
      <PageHeader title="Anagrafica Articoli" sub={`${products.length} articoli totali · ${products.filter(p=>isIFBVendor(p.vendorName)).length} INALCA F&B`}/>
      <div style={{display:"flex",gap:"10px",marginBottom:"12px",alignItems:"center"}}>
        <SearchBar value={search} onChange={setSearch} placeholder="🔍 Cerca per codice o descrizione…"/>
        <button onClick={()=>setOnlyIFB(v=>!v)} style={{padding:"6px 14px",background:onlyIFB?T.gold:T.surface,color:onlyIFB?"#000":T.gold,border:`1px solid ${T.gold}`,borderRadius:"6px",cursor:"pointer",fontSize:"12px",whiteSpace:"nowrap"}}>
          {onlyIFB?"Solo IFB":"Tutti i vendor"}
        </button>
      </div>
      <Section title={`${filtered.length} articoli${onlyIFB?" (INALCA F&B)":""}`}>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse"}}>
            <THead cols={["N HK","IFB No","Descrizione","Vendor","Categoria","UOM","Qty/Box","Box/Plt","Kg/Box","Temp","Attivo"]}/>
            <tbody>{filtered.map((p,i)=>(
              <tr key={p.id} style={{borderBottom:`1px solid ${T.border}`,background:i%2===0?T.bg:T.surface}}>
                <TD mono><span style={{color:T.muted}}>{p.nHK||"—"}</span></TD>
                <TD mono><span style={{color:T.gold}}>{p.code}</span></TD>
                <TD>{p.description}</TD>
                <TD><span style={{fontSize:"10px",color:isIFBVendor(p.vendorName)?T.gold:T.dim}}>{p.vendorName||"—"}</span></TD>
                <TD><Chip label={p.category||"—"} color={p.category==="WINE"?T.purple:p.category==="MEAT"?T.red:p.category==="SPIRITS"?T.orange:T.blue}/></TD>
                <TD>{p.uom}</TD><TD mono>{p.qtyPerBox}</TD><TD mono>{p.boxPerPallet}</TD><TD mono>{p.kgPerBox||"—"}</TD>
                <TD><Chip label={p.temperature||"—"} color={p.temperature==="FROZEN"?T.blue:p.temperature==="FRESH"?T.green:T.muted}/></TD>
                <TD><Chip label={p.active?"Sì":"No"} color={p.active?T.green:T.red}/></TD>
              </tr>
            ))}</tbody>
          </table>
        </div>
      </Section>
    </div>
  );
}

// ─── LOGISTICS ────────────────────────────────────────────────────────────────
function Logistics({logistics,setLogistics,products,branch,showToast,bumpImportTs}) {
  const[search,setSearch]=useState("");
  const[showOnlyMissing,setShowOnlyMissing]=useState(false);
  const[mapStep,setMapStep]=useState("idle");
  const[logHeaders,setLogHeaders]=useState([]);
  const[logRawRows,setLogRawRows]=useState([]);
  // Store column indices in React state (NOT window) to avoid stale reference bug
  const[colIdx,setColIdx]=useState({});

  // Only IFB vendor products for logistics
  const allIFBProducts = products.filter(p=>isIFBVendor(p.vendorName));

  function getLog(productId) { return logistics.find(l=>l.productId===productId&&l.branch===branch)||null; }
  function getOrDefault(productId) {
    return getLog(productId)||{productId,branch,area:"NORD",ubicazione:"MTO",pltPerContainer:20,hasCert:false,hasAlcTax:false,alcTax:0,convFactor:1,carriage:0};
  }
  function update(productId,field,rawVal) {
    const val=["ubicazione","area"].includes(field)?rawVal:["hasCert","hasAlcTax"].includes(field)?rawVal==="true":parseFloat(rawVal)||0;
    const existing=getLog(productId);
    let next;
    if(existing){next=logistics.map(l=>l.productId===productId&&l.branch===branch?{...l,[field]:val}:l);}
    else{next=[...logistics,{...getOrDefault(productId),[field]:val}];}
    setLogistics(next);LS.set("ifb_logistics",next);
  }

  function parseLogFile(e) {
    const file=e.target.files?.[0]; if(!file) return;
    const reader=new FileReader();
    reader.onload=ev=>{
      try{
        const wb=XLSX.read(ev.target.result,{type:"binary"});
        const ws=wb.Sheets[wb.SheetNames[0]];
        const raw=XLSX.utils.sheet_to_json(ws,{header:1,defval:""});
        let headerRowIdx=raw.findIndex(r=>r.some(c=>String(c||"").toLowerCase().includes("ubicazione")));
        if(headerRowIdx<0) headerRowIdx=0;
        const hdrs=raw[headerRowIdx].map(h=>String(h||"").trim());
        const dataRows=raw.slice(headerRowIdx+1).filter(r=>r.some(c=>c!==""));
        const fi=aliases=>hdrs.findIndex(h=>aliases.some(a=>h.toLowerCase().replace(/[\s_°]/g,"").includes(a.replace(/[\s_°]/g,""))));
        const idx={
          iNHK:    fi(["nhk","n hk","gc"]),
          iIFB:    fi(["no_(ifb)","noifb","ifb","no_"]),
          iUb:     fi(["ubicazione","location","wh"]),
          iArea:   fi(["area"]),
          iPlt:    fi(["npltxcontainer","pltxcontainer","plt x container","nplt","pltpercontainer","n plt"]),
          iCert:   fi(["healthcertificate","health certificate","cert"]),
          iTemp:   fi(["rettificata","temperature","temp","trettificata"]),
          iCarriage:fi(["pltcostmedio","plt cost medio","pltcost","carriage"]),
          iAirSea: fi(["air/sea","airsea","air","sea"]),
          iAlcTax: fi(["tassa alcolica","tassaalcolica","alcolica","alctax","alc tax"]),
        };
        setColIdx(idx);
        setLogHeaders(hdrs);setLogRawRows(dataRows);
        setMapStep("ready");
        showToast(`File caricato: ${dataRows.length} righe`,T.gold);
      }catch(err){showToast("Errore: "+err.message,T.red);}
    };
    reader.readAsBinaryString(file);
    e.target.value="";
  }

  function applyLogFile() {
    const idx = colIdx;
    const { iNHK, iIFB, iUb, iArea, iPlt, iCert, iTemp, iCarriage, iAirSea, iAlcTax } = idx;
    let next = [...logistics];
    let countLog = 0;
    let countAir = 0;
  
    logRawRows.forEach(row => {
      const nHK = String(row[iNHK >= 0 ? iNHK : 99] || "").trim();
      const ifbNo = String(row[iIFB >= 0 ? iIFB : 99] || "").trim();
      
      // SKIP: non sono prodotti IFB (filtro per branch)
      if (!nHK && !ifbNo) return;
      
      const prod = products.find(p => 
        (nHK && (p.nHK === nHK || p.code === nHK)) || 
        (ifbNo && (p.code === ifbNo || p.id === ifbNo))
      );
      if (!prod) return;
      
      // SKIP: solo prodotti con vendor INALCA FOOD & BEVERAGE
      if (!isIFBVendor(prod.vendorName)) return;
  
      const ub = String(row[iUb >= 0 ? iUb : 99] || "").trim().toUpperCase();
      const area = String(row[iArea >= 0 ? iArea : 99] || "NORD").trim().toUpperCase();
      
      // Temperatura da colonna "T° RETTIFICATA"
      const tempRaw = String(row[iTemp >= 0 ? iTemp : 99] || "").trim().toUpperCase();
      const temp = tempRaw === "FRESH" ? "FRESH" : tempRaw === "FROZEN" ? "FROZEN" : "DRY";
      
      // PLT per container: se 0 usa formula
      const pltRaw = parseFloat(row[iPlt >= 0 ? iPlt : 99]) || 0;
      const plt = pltRaw > 0 ? pltRaw : (temp === "DRY" ? 25 : (temp === "FRESH" || temp === "FROZEN") ? 23 : 20);
      
      const certRaw = String(row[iCert >= 0 ? iCert : 99] || "").trim().toUpperCase();
      const hasCert = certRaw === "SI" || certRaw === "YES" || certRaw === "TRUE";
      const carriage = parseFloat(row[iCarriage >= 0 ? iCarriage : 99]) || 0;
      const airSea = String(row[iAirSea >= 0 ? iAirSea : 99] || "").trim().toUpperCase();
      const alcRaw = String(row[iAlcTax >= 0 ? iAlcTax : 99] || "").trim().toUpperCase();
      const hasAlcTax = alcRaw === "SI" || alcRaw === "YES" || alcRaw === "TRUE";
      
      if (airSea === "AIR") { countAir++; }
      if (!["MTO", "MTS", "FOR"].includes(ub)) return;
      
      const areaFixed = ["NORD", "CENTRO", "SUD"].includes(area) ? area : "NORD";
      
      // IMPORTANTE: filtro per branch corrente
      const existIdx = next.findIndex(l => l.productId === prod.id && l.branch === branch);
      const entry = {
        productId: prod.id,
        branch,  // <-- assegna il branch corrente
        area: areaFixed,
        ubicazione: ub,
        pltPerContainer: plt,
        hasCert,
        hasAlcTax,
        alcTax: 0,
        convFactor: 1,
        carriage
      };
      
      if (existIdx >= 0) {
        next[existIdx] = { ...next[existIdx], ...entry };
      } else {
        next.push(entry);
      }
      countLog++;
    });
    

    function applyLogFile() {
      const idx = colIdx;
      const { iNHK, iIFB, iUb, iArea, iPlt, iCert, iTemp, iCarriage, iAirSea, iAlcTax } = idx;
      let next = [...logistics];
      let countLog = 0;
      let countAir = 0;
      
      // DEBUG: controlla branch
      console.log("=== DEBUG applyLogFile ===");
      console.log("Branch corrente:", branch);
      console.log("Prodotti IFB disponibili:", products.filter(p => isIFBVendor(p.vendorName)).length);
      console.log("Righe file:", logRawRows.length);
    
      logRawRows.forEach((row, idxRow) => {
        const nHK = String(row[iNHK >= 0 ? iNHK : 99] || "").trim();
        const ifbNo = String(row[iIFB >= 0 ? iIFB : 99] || "").trim();
        
        console.log(`Riga ${idxRow}: nHK="${nHK}", ifbNo="${ifbNo}"`);
        
        if (!nHK && !ifbNo) {
          console.log(`  → SKIP: nessun codice`);
          return;
        }
        
        const prod = products.find(p => 
          (nHK && (p.nHK === nHK || p.code === nHK)) || 
          (ifbNo && (p.code === ifbNo || p.id === ifbNo))
        );
        
        if (!prod) {
          console.log(`  → SKIP: prodotto non trovato per ${nHK || ifbNo}`);
          return;
        }
        
        console.log(`  → TROVATO: ${prod.code} - ${prod.description}`);
        
        // SKIP: solo prodotti con vendor INALCA FOOD & BEVERAGE
        if (!isIFBVendor(prod.vendorName)) {
          console.log(`  → SKIP: vendor non IFB (${prod.vendorName})`);
          return;
        }
    
        const ub = String(row[iUb >= 0 ? iUb : 99] || "").trim().toUpperCase();
        console.log(`  → Ubicazione: ${ub}`);
        
        if (!["MTO", "MTS", "FOR"].includes(ub)) {
          console.log(`  → SKIP: ubicazione non valida (${ub})`);
          return;
        }
        
        const area = String(row[iArea >= 0 ? iArea : 99] || "NORD").trim().toUpperCase();
        
        // Temperatura da colonna "T° RETTIFICATA"
        const tempRaw = String(row[iTemp >= 0 ? iTemp : 99] || "").trim().toUpperCase();
        const temp = tempRaw === "FRESH" ? "FRESH" : tempRaw === "FROZEN" ? "FROZEN" : "DRY";
        
        // PLT per container: se 0 usa formula
        const pltRaw = parseFloat(row[iPlt >= 0 ? iPlt : 99]) || 0;
        const plt = pltRaw > 0 ? pltRaw : (temp === "DRY" ? 25 : (temp === "FRESH" || temp === "FROZEN") ? 23 : 20);
        
        const certRaw = String(row[iCert >= 0 ? iCert : 99] || "").trim().toUpperCase();
        const hasCert = certRaw === "SI" || certRaw === "YES" || certRaw === "TRUE";
        const carriage = parseFloat(row[iCarriage >= 0 ? iCarriage : 99]) || 0;
        const airSea = String(row[iAirSea >= 0 ? iAirSea : 99] || "").trim().toUpperCase();
        const alcRaw = String(row[iAlcTax >= 0 ? iAlcTax : 99] || "").trim().toUpperCase();
        const hasAlcTax = alcRaw === "SI" || alcRaw === "YES" || alcRaw === "TRUE";
        
        if (airSea === "AIR") { 
          countAir++; 
        }
        
        const areaFixed = ["NORD", "CENTRO", "SUD"].includes(area) ? area : "NORD";
        
        const existIdx = next.findIndex(l => l.productId === prod.id && l.branch === branch);
        console.log(`  → existIdx: ${existIdx}, branch: ${branch}`);
        
        const entry = {
          productId: prod.id,
          branch,
          area: areaFixed,
          ubicazione: ub,
          pltPerContainer: plt,
          hasCert,
          hasAlcTax,
          alcTax: 0,
          convFactor: 1,
          carriage
        };
        
        console.log(`  → ENTRY CREATA:`, entry);
        
        if (existIdx >= 0) {
          next[existIdx] = { ...next[existIdx], ...entry };
        } else {
          next.push(entry);
        }
        countLog++;
      });
      
      console.log("=== RISULTATO FINALE ===");
      console.log("countLog:", countLog);
      console.log("next length:", next.length);
      console.log("next per questo branch:", next.filter(l => l.branch === branch).length);
      
      setLogistics(next);
      LS.set("ifb_logistics", next);
      if (countAir > 0) showToast(`⚠ ${countAir} articoli AIR rilevati — gestiscili da ✈ AIR Transport`, T.orange);
      bumpImportTs();
      showToast(`Logistica aggiornata: ${countLog} prodotti ✓`, T.gold);
      setMapStep("idle");
      setLogHeaders([]);
      setLogRawRows([]);
    }
    setLogistics(next);
    LS.set("ifb_logistics", next);
    if (countAir > 0) showToast(`⚠ ${countAir} articoli AIR rilevati — gestiscili da ✈ AIR Transport`, T.orange);
    bumpImportTs();
    showToast(`Logistica aggiornata: ${countLog} prodotti ✓`, T.gold);
    setMapStep("idle");
    setLogHeaders([]);
    setLogRawRows([]);
  }

  const allProds=allIFBProducts.filter(p=>!search||p.description?.toLowerCase().includes(search.toLowerCase())||p.code?.includes(search));
  // MOSTRA solo prodotti che hanno logistica per QUESTO branch (quando showOnlyMissing è false)
  const displayed=showOnlyMissing 
    ? allProds.filter(p=>!getLog(p.id))   // solo quelli SENZA logistica
    : allProds.filter(p=>getLog(p.id) !== null);  // solo quelli CON logistica per questo branch
  const missingCount=allIFBProducts.filter(p=>!getLog(p.id)).length;
  const withCount=allIFBProducts.filter(p=>getLog(p.id) !== null).length;

  return(
    <div>
      <PageHeader title={`Logistica · ${branch}`} sub={`${withCount} con logistica · ${missingCount} senza logistica (totale ${allIFBProducts.length} IFB)`}/>
      {mapStep==="idle"?(
        <div style={{marginBottom:"16px",display:"flex",gap:"10px",alignItems:"center"}}>
          <label style={{display:"inline-block",padding:"8px 16px",background:T.surface,border:`1px solid ${T.border}`,borderRadius:"6px",cursor:"pointer",fontSize:"12px",color:T.text}}>
            📂 Carica Work_tab (08_Work_Tab.xlsx)
            <input type="file" accept=".xlsx,.xls,.csv" onChange={parseLogFile} style={{display:"none"}}/>
          </label>
          <span style={{fontSize:"11px",color:T.muted}}>Colonne: N HK / No_(IFB) / Ubicazione / Area / Cert / Carriage / TASSA ALCOLICA / AIR/SEA</span>
        </div>
      ):mapStep==="ready"?(
        <div style={{background:T.card,border:`1px solid ${T.green}`,borderRadius:"8px",padding:"16px",marginBottom:"16px"}}>
          <div style={{color:T.green,fontWeight:"bold",fontSize:"13px",marginBottom:"8px"}}>✓ File rilevato · {logRawRows.length} righe</div>
          <div style={{fontSize:"12px",color:T.muted,marginBottom:"12px",lineHeight:"1.8"}}>
            Verranno importati: <strong style={{color:T.text}}>Ubicazione, Area, Plt/Container, Health Certificate, Carriage, Tassa Alcolica</strong>
          </div>
          <div style={{display:"flex",gap:"10px"}}>
            <ActionBtn label="← Annulla" onClick={()=>setMapStep("idle")}/>
            <ActionBtn label={`✓ Importa logistica (${logRawRows.length} righe)`} onClick={applyLogFile} primary/>
          </div>
        </div>
      ):null}

      <div style={{display:"flex",gap:"10px",marginBottom:"12px",alignItems:"center"}}>
        <SearchBar value={search} onChange={setSearch} placeholder="🔍 Cerca prodotto IFB…"/>
        <button onClick={()=>setShowOnlyMissing(v=>!v)}
          style={{padding:"6px 14px",background:showOnlyMissing?T.orange:T.surface,color:showOnlyMissing?"#000":T.orange,border:`1px solid ${T.orange}`,borderRadius:"6px",cursor:"pointer",fontSize:"12px",whiteSpace:"nowrap"}}>
          ⚠ Solo senza logistica ({missingCount})
        </button>
      </div>

      {missingCount>0&&<div style={{background:`${T.orange}15`,border:`1px solid ${T.orange}44`,borderRadius:"6px",padding:"10px 14px",marginBottom:"14px",fontSize:"12px",color:T.orange}}>
        ⚠ {missingCount} prodotti IFB senza parametri logistici → Standard Cost non calcolabile. Carica Work_tab o imposta manualmente.
      </div>}

      <div style={{overflowX:"auto"}}>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:"12px"}}>
          <THead cols={["IFB No","N HK","Descrizione","Ubicaz.","Area","Plt/Cont","Cert.","Alcol >30°","Carriage","Conv."]}/>
          <tbody>{displayed.map((prod,i)=>{
            const l=getOrDefault(prod.id);
            const hasEntry=!!getLog(prod.id);
            return<tr key={prod.id} style={{borderBottom:`1px solid ${T.border}`,background:!hasEntry?`${T.orange}08`:i%2===0?T.bg:T.surface}}>
              <TD mono><span style={{color:T.gold}}>{prod.code}</span></TD>
              <TD mono><span style={{color:T.muted}}>{prod.nHK||"—"}</span></TD>
              <TD>{prod.description}{!hasEntry&&<span style={{marginLeft:"6px",fontSize:"9px",color:T.orange}}>⚠ nuovo</span>}</TD>
              <td style={{padding:"4px 8px",borderBottom:`1px solid ${T.border}`}}>
                <select value={l.ubicazione||"MTO"} onChange={e=>update(prod.id,"ubicazione",e.target.value)}
                  style={{background:T.card,color:T.gold,border:`1px solid ${T.border}`,borderRadius:"4px",padding:"3px 6px",fontSize:"11px"}}>
                  {["MTO","MTS","FOR"].map(v=><option key={v} value={v}>{v}</option>)}
                </select>
              </td>
              <td style={{padding:"4px 8px",borderBottom:`1px solid ${T.border}`}}>
                <select value={l.area||"NORD"} onChange={e=>update(prod.id,"area",e.target.value)}
                  style={{background:T.card,color:T.text,border:`1px solid ${T.border}`,borderRadius:"4px",padding:"3px 6px",fontSize:"11px"}}>
                  {["NORD","CENTRO","SUD"].map(v=><option key={v} value={v}>{v}</option>)}
                </select>
              </td>
              {[["pltPerContainer",l.pltPerContainer],["hasCert",l.hasCert,"bool"],["hasAlcTax",l.hasAlcTax,"bool"],["carriage",l.carriage],["convFactor",l.convFactor]].map(([field,val,type])=>(
                <td key={field} style={{padding:"4px 8px",borderBottom:`1px solid ${T.border}`}}>
                  {type==="bool"?(
                    <select value={String(val||false)} onChange={e=>update(prod.id,field,e.target.value)}
                      style={{background:T.card,color:T.text,border:`1px solid ${T.border}`,borderRadius:"4px",padding:"3px 6px",fontSize:"11px"}}>
                      <option value="false">No</option><option value="true">Sì</option>
                    </select>
                  ):(
                    <input type="number" defaultValue={val||0} key={prod.id+field+(val||0)}
                      onBlur={e=>update(prod.id,field,e.target.value)}
                      style={{width:"60px",background:"transparent",color:T.gold,border:"none",textAlign:"right",fontSize:"12px"}}/>
                  )}
                </td>
              ))}
            </tr>;
          })}</tbody>
        </table>
      </div>
    </div>
  );
}

// ─── PRICES ───────────────────────────────────────────────────────────────────
function Prices({prices,products,branch,month}) {
  const[search,setSearch]=useState("");
  const filtered=prices.filter(p=>p.branch===branch&&p.month===month);
  const displayed=filtered.filter(p=>{
    if(!search) return true;
    const prod=products.find(pr=>pr.id===p.productId);
    return prod?.description?.toLowerCase().includes(search.toLowerCase())||prod?.code?.includes(search)||prod?.nHK?.includes(search);
  });
  return(
    <div>
      <PageHeader title={`Listini · ${branch} · ${month}`} sub={`${filtered.length} prezzi caricati`}/>
      {filtered.length===0?<div style={{padding:"32px",textAlign:"center",color:T.muted,fontSize:"13px"}}>Nessun prezzo per {branch} · {month}. Usa "Import Listini".</div>:(
        <>
          <SearchBar value={search} onChange={setSearch} placeholder="🔍 Cerca prodotto…"/>
          <Section title={`${displayed.length} / ${filtered.length} prezzi`}>
            <div style={{overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse"}}>
                <THead cols={["N HK","IFB No","Descrizione","FCA Price","FCA Disc.","DAP Price","DAP Disc.","MTS Price","DAP Final"]}/>
                <tbody>{displayed.map((p,i)=>{
                  const prod=products.find(pr=>pr.id===p.productId);
                  return<tr key={p.productId} style={{borderBottom:`1px solid ${T.border}`,background:i%2===0?T.bg:T.surface}}>
                    <TD mono><span style={{color:T.muted}}>{prod?.nHK||"—"}</span></TD>
                    <TD mono><span style={{color:T.gold}}>{prod?.code||p.productId}</span></TD>
                    <TD>{prod?.description||p.productId}</TD>
                    {["fcaPrice","fcaDiscounted","dapPrice","dapDiscounted","mtsPrice","dapFinal"].map(f=>(
                      <TD key={f} mono><span style={{color:(p[f]||0)>0?T.text:T.dim}}>{(p[f]||0)>0?`€ ${roundN(p[f]).toFixed(2)}`:"—"}</span></TD>
                    ))}
                  </tr>;
                })}</tbody>
              </table>
            </div>
          </Section>
        </>
      )}
    </div>
  );
}

// ─── FX RATES ─────────────────────────────────────────────────────────────────
function FxRates({fx,setFx,branch,month}) {
  const branchRates=fx.filter(f=>f.branch===branch).sort((a,b)=>b.month.localeCompare(a.month));
  const current=fx.find(f=>f.branch===branch&&f.month===month)?.rate||BRANCH_CFG[branch]?.defaultRate;
  function update(targetMonth,val){
    const rate=parseFloat(val)||0;
    const exists=fx.find(f=>f.branch===branch&&f.month===targetMonth);
    const next=exists?fx.map(f=>f.branch===branch&&f.month===targetMonth?{...f,rate}:f):[...fx,{branch,month:targetMonth,rate}];
    setFx(next);LS.set("ifb_fx",next);
  }
  return(
    <div>
      <PageHeader title={`Tassi di cambio · ${branch}`} sub={`Tasso corrente (${month}): ${current} ${BRANCH_CFG[branch]?.currency}/EUR`}/>
      <Section title={`Tasso ${month} — EUR / ${BRANCH_CFG[branch]?.currency}`} accent={T.gold}>
        <div style={{display:"flex",alignItems:"center",gap:"12px"}}>
          <input type="number" step="0.0001" defaultValue={current} key={branch+month+current}
            onBlur={e=>update(month,e.target.value)}
            style={{width:"160px",padding:"8px 12px",background:T.card,color:T.gold,border:`1px solid ${T.gold}`,borderRadius:"6px",fontSize:"14px",fontWeight:"bold",textAlign:"right"}}/>
          <span style={{color:T.muted,fontSize:"13px"}}>{BRANCH_CFG[branch]?.currency} per 1 EUR</span>
        </div>
      </Section>
      {branchRates.length>0&&(
        <Section title="Storico tassi">
          <table style={{borderCollapse:"collapse"}}>
            <THead cols={["Mese","Tasso"]}/>
            <tbody>{branchRates.map((f,i)=>(
              <tr key={f.month} style={{borderBottom:`1px solid ${T.border}`,background:i%2===0?T.bg:T.surface}}>
                <TD mono><span style={{color:f.month===month?T.gold:T.text}}>{f.month}{f.month===month?" ✓":""}</span></TD>
                <TD mono><input type="number" step="0.0001" defaultValue={f.rate} key={f.branch+f.month+f.rate}
                  onBlur={e=>update(f.month,e.target.value)}
                  style={{width:"120px",background:"transparent",color:T.gold,border:"none",textAlign:"right",fontSize:"13px"}}/></TD>
              </tr>
            ))}</tbody>
          </table>
        </Section>
      )}
    </div>
  );
}

// ─── COST TABLE ───────────────────────────────────────────────────────────────
function CostTable({costRows,branch,month,logistics,lastImportTs,lastCalcTs,setLastCalcTs,setCostHistory}) {
  const[search,setSearch]=useState("");
  const[showDetail,setShowDetail]=useState(null);
  const needsRecalc=lastImportTs>lastCalcTs;

  function saveSnapshot(){
    const ts=Date.now();
    const snap={ts,date:new Date(ts).toISOString(),branch,month,rows:costRows.map(r=>({id:r.id,code:r.code,nHK:r.nHK,description:r.description,cost:r.cost?.step2Hkd??null,costEur:r.cost?.step2Eur??null,skipReason:r.skipReason||null}))};
    setCostHistory(prev=>{const n=[snap,...(prev||[])].slice(0,60);LS.set("ifb_costhistory",n);return n;});
    setLastCalcTs(ts);LS.set("ifb_last_calc_ts",ts);
  }

  const filtered=costRows.filter(r=>!search||r.description?.toLowerCase().includes(search.toLowerCase())||r.code?.includes(search)||r.nHK?.includes(search));
  const calc=filtered.filter(r=>r.cost?.step2Hkd!=null);
  const noPrice=filtered.filter(r=>!r.cost&&!r.isAir&&(r.skipReason?.includes("NO PREZZO")||r.skipReason?.includes("PREZZO=0")));
  const noLog  =filtered.filter(r=>!r.cost&&!r.isAir&&r.skipReason==="NO LOGISTICA");
  const noCalc =filtered.filter(r=>!r.cost&&!r.isAir&&r.skipReason&&!r.skipReason.includes("NO PREZZO")&&!r.skipReason.includes("PREZZO=0")&&r.skipReason!=="NO LOGISTICA"&&r.skipReason!=="AIR");

  return(
    <div>
      <div style={{display:"flex",alignItems:"center",gap:"14px",marginBottom:"16px",flexWrap:"wrap"}}>
        <PageHeader title={`Standard Cost · ${branch} · ${month}`} sub={`${calc.length} calcolati · solo INALCA F&B · SEA`}/>
        <button onClick={saveSnapshot} disabled={!needsRecalc}
          style={{padding:"9px 20px",background:needsRecalc?T.gold:"#333",color:needsRecalc?"#000":T.muted,border:"none",borderRadius:"6px",fontWeight:"bold",cursor:needsRecalc?"pointer":"not-allowed",fontSize:"13px",marginTop:"-8px"}}>
          {needsRecalc?"⟳ Ricalcola & Salva snapshot":"✓ Costi aggiornati"}
        </button>
        {needsRecalc&&<span style={{color:T.orange,fontSize:"12px",marginTop:"-8px"}}>⚠ Nuovi dati importati — clicca per salvare</span>}
      </div>
      {(noPrice.length>0||noCalc.length>0||noLog.length>0)&&(
        <div style={{background:`${T.orange}15`,border:`1px solid ${T.orange}44`,borderRadius:"6px",padding:"10px 14px",marginBottom:"14px",fontSize:"12px",color:T.orange}}>
          {noLog.length>0&&<div>⚠ {noLog.length} senza logistica — importa prima Work_tab dalla pagina "Logistica"</div>}
          {noPrice.length>0&&<div>⚠ {noPrice.length} senza prezzo ({month}) — vai su "Import Listini"</div>}
          {noCalc.length>0&&<div>⚠ {noCalc.length} altri errori — vedi colonna "Costo HKD" per dettaglio</div>}
        </div>
      )}
      <SearchBar value={search} onChange={setSearch} placeholder="🔍 Cerca articolo…"/>
      <div style={{overflowX:"auto"}}>
        <table style={{width:"100%",borderCollapse:"collapse"}}>
          <THead cols={["N HK","IFB No","Descrizione","Prezzo €","Ubicaz.","Step1 HKD","Costo HKD","Δ% vs prec",""]}/>
          <tbody>{filtered.map((r,i)=>{
            const hkd=r.cost?.step2Hkd??null;
            const prevHkd=r.prevCost?.step2Hkd??null;
            const pct=hkd!=null&&prevHkd!=null&&prevHkd>0?(hkd-prevHkd)/prevHkd*100:null;
            return<tr key={r.id} style={{borderBottom:`1px solid ${T.border}`,background:i%2===0?T.bg:T.surface,opacity:r.isAir?0.4:1}}>
              <TD mono><span style={{color:T.muted}}>{r.nHK||"—"}</span></TD>
              <TD mono><span style={{color:T.gold}}>{r.code}</span></TD>
              <TD>{r.description}{r.isAir&&<span style={{marginLeft:"6px",color:T.orange,fontSize:"10px"}}>✈ AIR</span>}</TD>
              <TD mono>{r.priceInput!=null?`€ ${roundN(r.priceInput).toFixed(2)}`:"—"}</TD>
              <TD><Chip label={r.ubicazione||"—"} color={r.ubicazione==="FOR"?T.purple:r.ubicazione==="MTS"?T.blue:T.green}/></TD>
              <TD mono>{r.cost?.step1Hkd!=null?roundN(r.cost.step1Hkd).toFixed(2):"—"}</TD>
              <TD mono><span style={{color:hkd!=null?T.gold:T.dim,fontWeight:"bold"}}>{hkd!=null?hkd.toFixed(2):r.skipReason||"—"}</span></TD>
              <TD><span style={{color:pct==null?T.dim:pct>3?T.red:pct<-3?T.green:T.text}}>{pct!=null?(pct>0?"+":"")+pct.toFixed(1)+"%":"—"}</span></TD>
              <TD>{r.cost&&<MiniBtn label="+" onClick={()=>setShowDetail(showDetail===r.id?null:r.id)}/>}</TD>
            </tr>;
          })}</tbody>
        </table>
      </div>
      {showDetail&&(()=>{
        const r=costRows.find(x=>x.id===showDetail);
        if(!r?.cost) return null;
        const c=r.cost;
        const logRaw=logistics.find(l=>l.productId===r.id&&l.branch===branch);
        const pltFromFile=logRaw?.pltPerContainer||0;
        const pltLabel=`Plt/Container: ${r.pltUsed||"—"}${pltFromFile<=0?" (formula: "+r.temperature+")":" (da Work_tab)"}`;
        return(
          <Section title={`Dettaglio calcolo: ${r.description}`} accent={T.gold} mt="16px">
            <div style={{fontSize:"11px",color:T.muted,marginBottom:"8px",padding:"5px 10px",background:`${T.gold}11`,borderRadius:"5px",border:`1px solid ${T.gold}33`}}>
              {pltLabel} · Units/plt: {r.cost?.unitsPerPlt||"—"} · Temp: {r.temperature} · Area: {logRaw?.area||"—"}
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(200px,1fr))",gap:"6px"}}>
              {[["Prezzo acquisto EUR",c.priceEur],["FOB + Nolo",c.fob],["LIC (4100+3800 HKD)",c.lic],["VGM",c.vgm],["Health Cert.",c.hc],["Pallet (€30/plt)",c.plt],["Tassa Alcolica",c.alc],["Carriage",c.carriageUnit],["= Step 1 EUR",c.step1Eur],["= Step 1 HKD",c.step1Hkd],["Magazzino (MTO/MTS)",c.wh],["= Step 2 EUR",c.step2Eur],["= Step 2 HKD ✓",c.step2Hkd],["Tasso EUR/HKD",c.rate]].map(([k,v])=>(
                <div key={k} style={{display:"flex",justifyContent:"space-between",padding:"5px 10px",background:T.card,borderRadius:"5px",fontSize:"12px"}}>
                  <span style={{color:T.muted}}>{k}</span>
                  <span style={{color:String(k).includes("Step")||String(k).includes("✓")?T.gold:T.text,fontWeight:String(k).includes("✓")?"bold":"normal"}}>{typeof v==="number"?v.toFixed(4):"—"}</span>
                </div>
              ))}
            </div>
          </Section>
        );
      })()}
    </div>
  );
}

// ─── MAIL GEN ─────────────────────────────────────────────────────────────────
// Only shows items with |delta| > 3% (point 7)
function MailGen({costRows,branch,month}) {
  const[copied,setCopied]=useState(false);
  // Only items where |delta| > 3%
  const changed=costRows.filter(r=>r.cost?.step2Hkd!=null&&r.prevCost?.step2Hkd!=null&&r.prevCost.step2Hkd>0&&Math.abs((r.cost.step2Hkd-r.prevCost.step2Hkd)/r.prevCost.step2Hkd*100)>3);
  const up=changed.filter(r=>(r.cost.step2Hkd-r.prevCost.step2Hkd)/r.prevCost.step2Hkd*100>3);
  const dn=changed.filter(r=>(r.cost.step2Hkd-r.prevCost.step2Hkd)/r.prevCost.step2Hkd*100<-3);
  const body=`Gentili colleghi,\n\ndi seguito le variazioni di Standard Cost (solo articoli con Δ > ±3%) per ${branch} — ${month}:\n\n`
    +(up.length?`📈 AUMENTI (${up.length}):\n`+up.map(r=>{const pct=(r.cost.step2Hkd-r.prevCost.step2Hkd)/r.prevCost.step2Hkd*100;return`• ${r.nHK||r.code}  ${r.description}: ${r.prevCost.step2Hkd.toFixed(2)} → ${r.cost.step2Hkd.toFixed(2)} HKD (+${pct.toFixed(1)}%)`;}).join("\n"):"")
    +(dn.length?`\n\n📉 RIDUZIONI (${dn.length}):\n`+dn.map(r=>{const pct=(r.cost.step2Hkd-r.prevCost.step2Hkd)/r.prevCost.step2Hkd*100;return`• ${r.nHK||r.code}  ${r.description}: ${r.prevCost.step2Hkd.toFixed(2)} → ${r.cost.step2Hkd.toFixed(2)} HKD (${pct.toFixed(1)}%)`;}).join("\n"):"")
    +"\n\nCordiali saluti,\nIFB Cost Intelligence";
  return(
    <div>
      <PageHeader title={`Mail Mensile · ${branch} · ${month}`} sub={`${changed.length} articoli con Δ > ±3% (${up.length} aumenti · ${dn.length} riduzioni)`}/>
      {changed.length===0?<div style={{padding:"32px",textAlign:"center",color:T.muted}}>Nessuna variazione &gt; ±3% questo mese.</div>:(
        <Section title="Testo mail">
          <ActionBtn label={copied?"✓ Copiato!":"📋 Copia testo"} onClick={()=>{navigator.clipboard.writeText(body);setCopied(true);setTimeout(()=>setCopied(false),2500);}} primary/>
          <pre style={{marginTop:"12px",background:T.card,color:T.text,padding:"16px",borderRadius:"8px",fontSize:"12px",whiteSpace:"pre-wrap",border:`1px solid ${T.border}`,lineHeight:"1.7"}}>{body}</pre>
        </Section>
      )}
    </div>
  );
}

// ─── SALES INVOICE ────────────────────────────────────────────────────────────
function SalesInvoice({rows,setRows,airList,products,xrefs,snapshots,setSnapshots,importLogs,setImportLogs,showToast,bumpImportTs}) {
  const[step,setStep]=useState(rows.length?"view":"upload");
  const[preview,setPreview]=useState([]);
  const[filter,setFilter]=useState("all");

  function parseFile(e) {
    const file=e.target.files?.[0]; if(!file) return;
    const reader=new FileReader();
    reader.onload=ev=>{
      try{
        const wb=XLSX.read(ev.target.result,{type:"binary"});
        const ws=wb.Sheets[wb.SheetNames[0]];
        const data=XLSX.utils.sheet_to_json(ws,{defval:""});
        if(!data.length){showToast("File vuoto",T.red);return;}
        const norm=s=>String(s||"").toLowerCase().replace(/[\s_]/g,"");
        const firstRow=data[0];
        const km={};
        for(const k of Object.keys(firstRow)){
          const n=norm(k);
          if(n.includes("description")||n.includes("descrizione")) km.description=k;
          else if(n.includes("amount")||n.includes("importo"))     km.amount=k;
          else if(n.includes("location")||n.includes("ubicazione"))km.location=k;
          else if((n.includes("no")||n.includes("code")||n.includes("item"))&&!km.itemCode) km.itemCode=k;
          else if(n.includes("date")||n.includes("data"))          km.date=k;
        }
        const parsed=data.map(r=>{
          const code=String(r[km.itemCode]||"");
          const prod=findProduct(code,products,xrefs);
          // Lookup N HK via xref or product
          const nHK=prod?.nHK||(xrefs.find(x=>x.ifbNo===code)?.nHK)||"";
          // Lookup AIR transport from airList
          const airEntry=prod?airList.find(a=>a.productId===prod.id):null;
          const transport=airEntry?airEntry.transportation:"";
          return{
            description:String(r[km.description]||""),
            amount:parseFloat(r[km.amount])||0,
            location:String(r[km.location]||""),
            itemCode:code,
            nHK,
            transport,
            date:String(r[km.date]||""),
          };
        }).filter(r=>!isExcludedDesc(r.description));
        setPreview(parsed);setStep("preview");
      }catch(err){showToast("Errore: "+err.message,T.red);}
    };
    reader.readAsBinaryString(file);
    e.target.value="";
  }

  function executeImport(){
    const now=Date.now();
    setRows(preview);LS.set("ifb_sales_invoice",preview);
    const log={id:now,type:"sales",date:new Date(now).toISOString(),count:preview.length,diffs:[],branch:"HK"};
    const newLogs=[log,...importLogs];setImportLogs(newLogs);LS.set("ifb_importlogs",newLogs);
    const newSnaps=[log,...snapshots].slice(0,50);setSnapshots(newSnaps);LS.set("ifb_snapshots",newSnaps);
    bumpImportTs();showToast(`Fattura: ${preview.length} righe importate ✓`,T.gold);setStep("view");
  }

  const activeRows=step==="view"?rows:preview;
  const airMismatches=activeRows.filter(r=>{
    const prod=findProduct(r.itemCode,products,xrefs);
    if(!prod) return false;
    const ae=airList.find(a=>a.productId===prod.id);
    if(!ae||!isAirTransport(ae.transportation)) return false;
    return !String(r.location||"").toUpperCase().includes("NCJ");
  });
  const displayedRows=filter==="air"?airMismatches:activeRows;

  const TableRows=({data})=>(
    <div style={{overflowX:"auto"}}>
      <table style={{width:"100%",borderCollapse:"collapse"}}>
        <THead cols={["Descrizione","Importo","Location","Codice","N HK","Transport","Data"]}/>
        <tbody>{data.slice(0,500).map((r,i)=>{
          const isMismatch=airMismatches.includes(r);
          return<tr key={i} style={{borderBottom:`1px solid ${T.border}`,background:isMismatch?"#2a1000":i%2===0?T.bg:T.surface}}>
            <TD>{r.description}</TD>
            <TD mono><span style={{color:isMismatch?T.orange:T.text}}>{r.amount?.toFixed(2)}</span></TD>
            <TD mono><span style={{color:isMismatch?T.orange:T.text}}>{r.location}</span></TD>
            <TD mono><span style={{color:T.gold}}>{r.itemCode}</span></TD>
            <TD mono><span style={{color:T.muted}}>{r.nHK||"—"}</span></TD>
            <TD>{r.transport?<Chip label={r.transport} color={isAirTransport(r.transport)?T.orange:T.blue}/>:<span style={{color:T.dim}}>—</span>}</TD>
            <TD mono>{r.date}</TD>
          </tr>;
        })}</tbody>
      </table>
    </div>
  );

  if(step==="preview") return(
    <div>
      <PageHeader title="Preview Sales Invoice" sub={`${preview.length} righe`}/>
      <div style={{display:"flex",gap:"10px",marginBottom:"16px"}}>
        <ActionBtn label="← Annulla" onClick={()=>setStep(rows.length?"view":"upload")}/>
        <ActionBtn label={`✓ Importa ${preview.length} righe`} onClick={executeImport} primary/>
      </div>
      <TableRows data={preview}/>
    </div>
  );

  return(
    <div>
      <PageHeader title="Sales Invoice" sub={rows.length?`${rows.length} righe caricate`:"Nessun file caricato"}/>
      {airMismatches.length>0&&(
        <div style={{background:"#2a1000",border:`1px solid ${T.orange}`,borderRadius:"6px",padding:"12px 16px",marginBottom:"16px",display:"flex",alignItems:"center",gap:"12px"}}>
          <span style={{color:T.orange,fontWeight:"bold"}}>⚠ {airMismatches.length} righe: trasporto AIR ma location ≠ NCJ</span>
          <button onClick={()=>setFilter(f=>f==="air"?"all":"air")}
            style={{padding:"4px 12px",background:T.orange,color:"#000",border:"none",borderRadius:"4px",cursor:"pointer",fontSize:"12px",fontWeight:"bold"}}>
            {filter==="air"?"Mostra tutte":"Mostra AIR mismatch"}
          </button>
        </div>
      )}
      <div style={{marginBottom:"16px"}}>
        <label style={{display:"inline-block",padding:"8px 16px",background:T.gold,color:"#000",borderRadius:"6px",cursor:"pointer",fontWeight:"bold",fontSize:"12px"}}>
          📂 {rows.length?"Ricarica fattura":"Carica fattura"}
          <input type="file" accept=".xlsx,.xls,.csv" onChange={parseFile} style={{display:"none"}}/>
        </label>
      </div>
      {rows.length>0&&<Section title={`${displayedRows.length} righe${filter==="air"?" (AIR mismatch)":""}`}><TableRows data={displayedRows}/></Section>}
    </div>
  );
}

// ─── STORICO ──────────────────────────────────────────────────────────────────
function Storico({snapshots,setSnapshots,costHistory,branch}) {
  const[sel,setSel]=useState(null);
  const[sortDir,setSortDir]=useState("asc");
  const[deltaFilter,setDeltaFilter]=useState("all");
  const[showModified,setShowModified]=useState(false);
  const[showNew,setShowNew]=useState(false);
  const[selCostSnap,setSelCostSnap]=useState(null);
  const costSnaps=(costHistory||[]).filter(s=>!branch||s.branch===branch);

  // Helper: get import date string from snapshot
  const snapDate = s => new Date(s.date||s.id).toLocaleDateString("it-IT",{day:"2-digit",month:"2-digit",year:"numeric"});

  function deleteSnap(id){
    const next=snapshots.filter(s=>s.id!==id);
    setSnapshots(next);LS.set("ifb_snapshots",next);
    if(sel?.id===id) setSel(null);
  }

  return(
    <div>
      <PageHeader title="Storico & Diff" sub="Snapshot import e Standard Cost"/>
      {/* COST HISTORY */}
      {costSnaps.length>0&&(
        <div style={{marginBottom:"32px"}}>
          <div style={{fontWeight:"bold",color:T.gold,fontSize:"11px",letterSpacing:"0.08em",textTransform:"uppercase",marginBottom:"10px",borderBottom:`1px solid ${T.gold}44`,paddingBottom:"6px"}}>📊 Storico Standard Cost · {branch}</div>
          <div style={{fontSize:"12px",color:T.muted,marginBottom:"10px"}}>Clicca una data per vedere i costi salvati in quel momento</div>
          <div style={{display:"flex",gap:"8px",flexWrap:"wrap",marginBottom:"12px"}}>
            {costSnaps.map(s=>(
              <button key={s.ts} onClick={()=>setSelCostSnap(selCostSnap?.ts===s.ts?null:s)}
                style={{padding:"6px 12px",background:selCostSnap?.ts===s.ts?T.gold:T.card,color:selCostSnap?.ts===s.ts?"#000":T.text,border:`1px solid ${selCostSnap?.ts===s.ts?T.gold:T.border}`,borderRadius:"6px",cursor:"pointer",fontSize:"12px"}}>
                {s.month||"?"} · {new Date(s.ts).toLocaleDateString("it-IT")}
              </button>
            ))}
          </div>
          {selCostSnap&&(
            <div style={{overflowX:"auto",marginBottom:"16px"}}>
              <table style={{width:"100%",borderCollapse:"collapse"}}>
                <THead cols={["N HK","IFB No","Descrizione","Costo HKD","Note"]}/>
                <tbody>{(selCostSnap.rows||[]).map((r,i)=>(
                  <tr key={r.id||i} style={{borderBottom:`1px solid ${T.border}`,background:i%2===0?T.bg:T.surface}}>
                    <TD mono><span style={{color:T.muted}}>{r.nHK||"—"}</span></TD>
                    <TD mono><span style={{color:T.gold}}>{r.code}</span></TD>
                    <TD>{r.description}</TD>
                    <TD mono><span style={{color:T.gold,fontWeight:"bold"}}>{r.cost!=null?roundN(r.cost).toFixed(2):"—"}</span></TD>
                    <TD><span style={{color:T.dim,fontSize:"11px"}}>{r.skipReason||""}</span></TD>
                  </tr>
                ))}</tbody>
              </table>
            </div>
          )}
        </div>
      )}

      {/* IMPORT SNAPSHOTS */}
      <div style={{fontWeight:"bold",color:T.gold,fontSize:"11px",letterSpacing:"0.08em",textTransform:"uppercase",marginBottom:"10px",borderBottom:`1px solid ${T.border}`,paddingBottom:"6px"}}>📥 Storico Import</div>
      {snapshots.length===0?<div style={{padding:"24px",textAlign:"center",color:T.dim,fontSize:"13px"}}>Nessuno snapshot ancora.</div>:(
        <div style={{display:"flex",flexDirection:"column",gap:"6px",marginBottom:"16px",maxHeight:"300px",overflowY:"auto"}}>
          {snapshots.map(s=>(
            <div key={s.id} style={{display:"flex",alignItems:"center",gap:"8px",padding:"8px 12px",background:sel?.id===s.id?`${T.gold}15`:T.card,border:`1px solid ${sel?.id===s.id?T.gold:T.border}`,borderRadius:"6px",cursor:"pointer"}}
              onClick={()=>{setSel(sel?.id===s.id?null:s);setShowModified(false);setShowNew(false);}}>
              <span style={{fontSize:"16px"}}>{s.type==="prices"?"💶":s.type==="anagrafica"?"◈":s.type==="xref"?"⇄":s.type==="air"?"✈":s.type==="sales"?"📋":"📥"}</span>
              <div style={{flex:1,minWidth:0}}>
                <div style={{fontSize:"12px",color:sel?.id===s.id?T.gold:T.text,fontWeight:"bold"}}>
                  {s.type==="prices"?"Import Listini":s.type==="anagrafica"?"Import Anagrafica":s.type==="xref"?"Import XRef":s.type==="air"?"Import AIR":s.type==="sales"?"Sales Invoice":"Import"}
                  {s.branch&&s.branch!=="ALL"&&<span style={{marginLeft:"6px",color:T.muted,fontWeight:"normal"}}>· {s.branch}</span>}
                  {s.month&&<span style={{marginLeft:"6px",color:T.gold,fontWeight:"normal",fontSize:"11px"}}>· {s.month}</span>}
                </div>
                <div style={{fontSize:"11px",color:T.muted,marginTop:"2px"}}>
                  {snapDate(s)} alle {new Date(s.date||s.id).toLocaleTimeString("it-IT",{hour:"2-digit",minute:"2-digit"})}
                  {" · "}{s.count} voci
                  {s.diffs&&s.diffs.length>0&&<span style={{color:T.orange}}> · {s.diffs.filter(d=>d.isNew).length} nuovi · {s.diffs.filter(d=>!d.isNew&&d.fields?.length>0).length} modif.</span>}
                </div>
              </div>
              <button onClick={e=>{e.stopPropagation();deleteSnap(s.id);}}
                style={{background:"none",border:"none",color:T.muted,cursor:"pointer",fontSize:"13px",padding:"2px 6px"}} title="Elimina">✕</button>
            </div>
          ))}
        </div>
      )}

      {sel&&(()=>{
        const diffs=sel.diffs||[];
        const newItems=diffs.filter(d=>d.isNew);
        const allModified=diffs.filter(d=>!d.isNew&&d.fields&&d.fields.length>0);
        const priceFields=["fcaPrice","fcaDiscounted","dapPrice","dapDiscounted","mtsPrice","dapFinal"];

        const realModified=allModified.map(d=>({
          ...d,
          fields:d.fields.filter(f=>{
            if(!priceFields.includes(f.field)) return true;
            return Math.abs(roundN(f.new||0)-roundN(f.old||0))>=0.005;
          })
        })).filter(d=>d.fields.length>0);

        const getPct=d=>{
          const pf=d.fields.find(f=>f.field==="dapFinal")||d.fields[0];
          if(!pf||!pf.old||pf.old===0) return 0;
          return(roundN(pf.new)-roundN(pf.old))/Math.abs(pf.old)*100;
        };

        let shownDiffs=realModified;
        if(deltaFilter==="minus") shownDiffs=realModified.filter(d=>getPct(d)<-3);
        else if(deltaFilter==="plus") shownDiffs=realModified.filter(d=>getPct(d)>3);
        shownDiffs=[...shownDiffs].sort((a,b)=>sortDir==="asc"?getPct(a)-getPct(b):getPct(b)-getPct(a));

        // Dates for "vecchio" and "nuovo" labels
        const thisDate = snapDate(sel);
        // Find previous snapshot of same type+branch+month
        const prevSnap = snapshots.find(s=>s.id!==sel.id&&s.type===sel.type&&s.branch===sel.branch&&s.month===sel.month);
        const prevDate = prevSnap ? snapDate(prevSnap) : "—";

        return(
          <div style={{background:T.card,borderRadius:"8px",padding:"16px",border:`1px solid ${T.border}`}}>
            <h3 style={{color:T.gold,marginTop:0,marginBottom:"12px"}}>{sel.type} · {thisDate} · {sel.branch||"ALL"} · {sel.count} voci</h3>
            <div style={{display:"flex",gap:"8px",flexWrap:"wrap",alignItems:"center",marginBottom:"14px"}}>
              <button onClick={()=>{setShowNew(v=>!v);setShowModified(false);}}
                style={{padding:"6px 14px",background:showNew?T.green:T.surface,color:showNew?"#000":T.green,border:`1px solid ${T.green}`,borderRadius:"6px",cursor:"pointer",fontSize:"12px",fontWeight:"bold"}}>
                🆕 {newItems.length} nuovi
              </button>
              <button onClick={()=>{setShowModified(v=>!v);setShowNew(false);}}
                style={{padding:"6px 14px",background:showModified?T.orange:T.surface,color:showModified?"#000":T.orange,border:`1px solid ${T.orange}`,borderRadius:"6px",cursor:"pointer",fontSize:"12px",fontWeight:"bold"}}>
                ✏️ {realModified.length} modificati (reali)
              </button>
              {showModified&&(
                <>
                  <button onClick={()=>setDeltaFilter(f=>f==="minus"?"all":"minus")}
                    style={{padding:"4px 10px",background:deltaFilter==="minus"?T.red:T.surface,color:deltaFilter==="minus"?"#fff":T.red,border:`1px solid ${T.red}`,borderRadius:"4px",cursor:"pointer",fontSize:"11px"}}>{"< -3%"}</button>
                  <button onClick={()=>setDeltaFilter(f=>f==="plus"?"all":"plus")}
                    style={{padding:"4px 10px",background:deltaFilter==="plus"?T.green:T.surface,color:deltaFilter==="plus"?"#fff":T.green,border:`1px solid ${T.green}`,borderRadius:"4px",cursor:"pointer",fontSize:"11px"}}>{">"} +3%</button>
                  <button onClick={()=>setSortDir(d=>d==="asc"?"desc":"asc")}
                    style={{padding:"4px 10px",background:T.surface,color:T.muted,border:`1px solid ${T.border}`,borderRadius:"4px",cursor:"pointer",fontSize:"11px"}}>
                    Δ {sortDir==="asc"?"↑ crescente":"↓ decrescente"}
                  </button>
                  {deltaFilter!=="all"&&<span style={{fontSize:"11px",color:T.muted}}>({shownDiffs.length} su {realModified.length})</span>}
                </>
              )}
            </div>

            {showNew&&(
              <div style={{marginBottom:"14px"}}>
                <div style={{color:T.green,fontSize:"12px",fontWeight:"bold",marginBottom:"8px"}}>Articoli nuovi ({newItems.length})</div>
                {newItems.length===0?<div style={{color:T.dim,fontSize:"12px"}}>Nessun articolo nuovo.</div>:(
                  <table style={{width:"100%",borderCollapse:"collapse"}}>
                    <THead cols={["IFB No","Descrizione"]}/>
                    <tbody>{newItems.map((d,i)=>(
                      <tr key={i} style={{borderBottom:`1px solid ${T.border}`}}>
                        <TD mono><span style={{color:T.gold}}>{d.id||d.productId}</span></TD>
                        <TD>{d.description}</TD>
                      </tr>
                    ))}</tbody>
                  </table>
                )}
              </div>
            )}

            {showModified&&(
              <div>
                <div style={{color:T.orange,fontSize:"12px",fontWeight:"bold",marginBottom:"8px"}}>Modifiche reali ({shownDiffs.length})</div>
                {shownDiffs.length===0?<div style={{color:T.dim,fontSize:"12px"}}>Nessuna variazione reale.</div>:(
                  <table style={{width:"100%",borderCollapse:"collapse"}}>
                    {/* Show date of previous vs current import in header */}
                    <THead cols={["IFB No / N HK","Descrizione","Campo",`Vecchio (${prevDate})`,`Nuovo (${thisDate})`,"Δ%"]}/>
                    <tbody>{shownDiffs.map((d,i)=>d.fields.map((f,j)=>{
                      const oldR=roundN(f.old||0),newR=roundN(f.new||0);
                      const pct=oldR!==0?(newR-oldR)/Math.abs(oldR)*100:null;
                      return<tr key={`${i}-${j}`} style={{borderBottom:j===d.fields.length-1?`1px solid ${T.border}`:`1px solid ${T.border}44`,background:i%2===0?T.bg:T.surface}}>
                        {j===0&&<>
                          <td rowSpan={d.fields.length} style={{padding:"6px 12px",borderBottom:`1px solid ${T.border}`,verticalAlign:"top",fontFamily:"monospace",fontSize:"12px",color:T.gold}}>{d.ifbNo||d.id}<br/><span style={{color:T.muted,fontSize:"10px"}}>{d.nHK||""}</span></td>
                          <td rowSpan={d.fields.length} style={{padding:"6px 12px",borderBottom:`1px solid ${T.border}`,verticalAlign:"top",fontSize:"12px",color:T.text}}>{d.description}</td>
                        </>}
                        <TD><span style={{color:T.muted,fontSize:"11px"}}>{f.field}</span></TD>
                        <TD mono>{oldR.toFixed(2)}</TD>
                        <TD mono>{newR.toFixed(2)}</TD>
                        <TD><span style={{color:pct==null?T.dim:pct>0?T.red:T.green,fontWeight:"bold"}}>{pct!=null?(pct>0?"+":"")+pct.toFixed(1)+"%":"—"}</span></TD>
                      </tr>;
                    }))}</tbody>
                  </table>
                )}
              </div>
            )}
          </div>
        );
      })()}
    </div>
  );
}

// ─── SHARED HELPERS ───────────────────────────────────────────────────────────
const inputStyle = () => ({
  width:"100%",padding:"7px 10px",background:"rgba(255,255,255,0.05)",
  border:`1px solid ${T.border}`,borderRadius:"6px",color:T.text,
  fontFamily:"inherit",fontSize:"12px",outline:"none",boxSizing:"border-box",
});

function PageHeader({title,sub}){
  return(
    <div style={{marginBottom:"20px"}}>
      <h2 style={{color:T.gold,margin:"0 0 4px",fontSize:"18px"}}>{title}</h2>
      {sub&&<div style={{fontSize:"12px",color:T.muted}}>{sub}</div>}
    </div>
  );
}
function Section({title,children,accent,mb,mt}){
  return(
    <div style={{marginBottom:mb||"24px",marginTop:mt||"0"}}>
      <div style={{fontWeight:"bold",color:accent||T.muted,fontSize:"11px",letterSpacing:"0.08em",textTransform:"uppercase",marginBottom:"10px",borderBottom:`1px solid ${accent?accent+"44":T.border}`,paddingBottom:"6px"}}>{title}</div>
      {children}
    </div>
  );
}
function ActionBtn({label,onClick,primary=false,disabled=false}){
  return<button onClick={onClick} disabled={disabled}
    style={{padding:"8px 18px",background:disabled?"#333":primary?T.gold:T.surface,color:disabled?T.muted:primary?"#000":T.text,border:primary?"none":`1px solid ${T.border}`,borderRadius:"6px",cursor:disabled?"not-allowed":"pointer",fontWeight:primary?"bold":"normal",fontSize:"13px"}}>{label}</button>;
}
function StepBar({steps,current}){
  const idx=steps.indexOf(current);
  return(
    <div style={{display:"flex",gap:"4px",marginBottom:"24px",alignItems:"center", width: "100%", marginLeft: 0, paddingLeft: 0 }}>
      {steps.map((s,i)=>(
        <span key={s} style={{display:"flex",alignItems:"center",gap:"4px"}}>
          <span style={{padding:"3px 10px",borderRadius:"10px",fontSize:"11px",background:i<=idx?T.gold:T.surface,color:i<=idx?"#000":T.muted,border:`1px solid ${i<=idx?T.gold:T.border}`}}>{s}</span>
          {i<steps.length-1&&<span style={{width:"16px",height:"1px",display:"inline-block",background:T.border}}/>}
        </span>
      ))}
    </div>
  );
}
function DropZone({onFile,label}){
  return(
    <div style={{border:`2px dashed ${T.borderHi}`,borderRadius:"10px",padding:"28px",textAlign:"center",cursor:"pointer"}}
      onClick={()=>document.getElementById("_dz_in")?.click()}>
      <div style={{fontSize:"28px",marginBottom:"8px"}}>📂</div>
      <div style={{fontSize:"13px",color:T.text,marginBottom:"4px"}}>{label||"Trascina o clicca per caricare"}</div>
      <div style={{fontSize:"11px",color:T.muted}}>Excel (.xlsx, .xls) o CSV</div>
      <input id="_dz_in" type="file" accept=".xlsx,.xls,.csv"
        onChange={e=>{const f=e.target.files?.[0];if(f)onFile(f);e.target.value="";}} style={{display:"none"}}/>
    </div>
  );
}
function SearchBar({value,onChange,placeholder}){
  return<input value={value} onChange={e=>onChange(e.target.value)} placeholder={placeholder||"Cerca..."}
    style={{...inputStyle(),maxWidth:"320px",marginBottom:"14px"}}/>;
}
function THead({cols}){
  return<thead><tr>{cols.map(c=><th key={c} style={{padding:"7px 12px",background:T.card,color:T.muted,textAlign:"left",borderBottom:`1px solid ${T.border}`,fontSize:"11px",fontWeight:"normal",letterSpacing:"0.05em",whiteSpace:"nowrap"}}>{c}</th>)}</tr></thead>;
}
function TD({children,mono=false}){
  return<td style={{padding:"7px 12px",borderBottom:`1px solid ${T.border}`,fontSize:"12px",fontFamily:mono?"monospace":"inherit",verticalAlign:"middle"}}>{children}</td>;
}
function Chip({label,color}){
  return<span style={{padding:"2px 7px",background:`${color}22`,color,borderRadius:"4px",fontSize:"10px",fontWeight:"bold",letterSpacing:"0.04em"}}>{label}</span>;
}
function MiniBtn({label,onClick,color}){
  return<button onClick={onClick} style={{padding:"3px 8px",background:"none",border:`1px solid ${color||T.border}`,borderRadius:"4px",color:color||T.muted,cursor:"pointer",fontSize:"11px"}}>{label}</button>;
}