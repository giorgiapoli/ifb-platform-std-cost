import { useState, useMemo, useEffect, useRef } from "react";
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
  CAN: { label:"Canarie",   flag:"🇮🇨", color:T.blue,   currency:"EUR", defaultRate:1,      active:true },
  AUS: { label:"Australia", flag:"🇦🇺", color:T.orange, currency:"AUD", defaultRate:1.6420, active:false },
};
const IFB_VENDOR = "INALCA FOOD & BEVERAGE";

const NOW = () => new Date().toISOString().slice(0,7);
const roundN = (n, d=2) => Math.round((n||0)*Math.pow(10,d))/Math.pow(10,d);
const EXCLUDED_INVOICE_DESC = [
  "freight","health certificate","handling costs","freight cost",
  "interest on intercompany","pallets","vendita prodotti finiti",
  "late payment interest","unifreddo costs",
];
const isExcludedDesc = (d: string) =>
  EXCLUDED_INVOICE_DESC.some(ex => String(d||"").toLowerCase().includes(ex));

// Codici contabili tipo 51.9020.25 (cifre con punti)
const isAccountingCode = (c: string) => /^\d+\.\d+(\.\d+)+$/.test(String(c||"").trim());
const branchN = (branch: string) => branch === "CAN" ? "N COMIT" : "N HK";

const AIR_TYPES = ["air","sea"];
const isAirTransport = t => {
  const val = String(t||"").toLowerCase().trim();
  return val === "air";
};
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


function exportXLSX(rows: any[], sheetName: string, fileName: string) {
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, fileName);
}
/**
 * Main cost calculation for HK (sea).
 * Returns null if totalUnits = 0 or missing data.
 *
 * FORMULA (per unit):
 *   Step1 = purchasePrice + FOB + LIC + VGM + HC + PLT + alcTax + carriageUnit
 *   Step2 = Step1 + warehouseCost
 */
function calcHK({ priceInput, ubicazione, product, logistic, eurToHkd }: any) {
  const { uom, qtyPerBox, boxPerPallet, kgPerBox, kgxplt, temperature } = product;
  const { pltPerContainer, area, hasCert, hasAlcTax, alcTax, convFactor } = logistic || {};

  // ── Units per pallet ── (formula modello Excel)
  // PCS: qtyPerBox × boxPerPallet
  // BOX: boxPerPallet
  // KG:  kgxplt (kg per pallet = KgPerBox × qtyPerBox × boxPerPallet)
  let unitsPerPlt: number;
  if (uom==="BOX") unitsPerPlt = Number(boxPerPallet);
  else if (uom==="KG") {
    unitsPerPlt = Number(kgxplt) > 0 ? Number(kgxplt) : 300;
  }
  else unitsPerPlt = Number(qtyPerBox) * Number(boxPerPallet); // PCS

  // ── Divisore collo per MTS picking ──
  const divisoreCollo =
    uom==="BOX" ? 1 :
    uom==="KG"  ? Number(kgPerBox||qtyPerBox) :
                  Number(qtyPerBox);

  // ── Total units per container ──
  const totalUnits = unitsPerPlt * Number(pltPerContainer);
  if (!totalUnits) return null;

  const priceEur = Number(priceInput||0) * Number(convFactor);

  // ✅ SE IL PREZZO È ZERO O NON VALIDO, NON CALCOLARE IL COSTO
  if (priceEur === 0 || !priceInput) {
    return null;
  }

  // ── FOB ──
  const fobContainer = COSTS.FOB[temperature]?.[area] ?? 0;
  const fob = (fobContainer / pltPerContainer) / unitsPerPlt;

  // ── LIC = (4100+3800 HKD) / rate / totalUnits ──
  const lic = (COSTS.LIC_HKD / eurToHkd) / totalUnits;

  // ── VGM = 100 / totalUnits ──
  const vgm = COSTS.VGM / totalUnits;

  // ── HC = 80 / totalUnits (solo se certificato) ──
  const hc = hasCert ? COSTS.HC / totalUnits : 0;

  // ── Pallet = 30 / unitsPerPlt ──
  const plt = COSTS.PLT / unitsPerPlt;

  // ── Tassa alcolica ──
  const alc = hasAlcTax ? (Number(alcTax)||0) : 0;

  // ── Step 1 ──
  const step1Eur = priceEur + fob + lic + vgm + hc + plt + alc;

  // ── Warehouse ──
  let wh = 0;
  if (ubicazione==="MTO") {
    wh = (COSTS.MTO[temperature] ?? 0) / unitsPerPlt;
  } else if (ubicazione==="MTS") {
    wh = (COSTS.MTS_D[temperature] ?? 0) / unitsPerPlt
       + (COSTS.MTS_I[temperature] ?? 0) / unitsPerPlt
       + (COSTS.MTS_P[temperature] ?? 0) / divisoreCollo;
  }

  const step2Eur = step1Eur + wh;

  return {
    priceEur, fob, lic, vgm, hc, plt, alc,
    step1Eur,
    step1Hkd: step1Eur * eurToHkd,
    wh,
    step2Eur,
    step2Hkd: step2Eur * eurToHkd,
    rate: eurToHkd,
    unitsPerPlt,
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
  kgxplt: ["kgxplt","kg x pallet","kg per pallet","kgperpallet"],
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
  set: (k,v) => { try{ localStorage.setItem(k,JSON.stringify(v)); return true; }catch{ return false; } },
};

// Seed data (minimal)
// ─── IndexedDB per dati grandi ──────────────────────────────────────────────
const IDB = (() => {
  let dbP: Promise<IDBDatabase>|null = null;
  const open = () => {
    if(!dbP) dbP = new Promise((res,rej) => {
      const r = indexedDB.open("ifb_store",1);
      r.onupgradeneeded = e => (e.target as IDBOpenDBRequest).result.createObjectStore("store");
      r.onsuccess = e => res((e.target as IDBOpenDBRequest).result);
      r.onerror = () => { dbP=null; rej(); };
    });
    return dbP;
  };
  return {
    set: async (key:string, val:any) => {
      try { const db=await open(); await new Promise<void>((res,rej)=>{ const tx=db.transaction("store","readwrite"); tx.objectStore("store").put(val,key); tx.oncomplete=()=>res(); tx.onerror=rej; }); return true; } catch { return false; }
    },
    get: async (key:string, def:any=null) => {
      try { const db=await open(); return await new Promise(res=>{ const tx=db.transaction("store","readonly"); const r=tx.objectStore("store").get(key); r.onsuccess=()=>res(r.result??def); r.onerror=()=>res(def); }); } catch { return def; }
    },
    del: async (key:string) => {
      try { const db=await open(); await new Promise<void>((res)=>{ const tx=db.transaction("store","readwrite"); tx.objectStore("store").delete(key); tx.oncomplete=()=>res(); tx.onerror=()=>res(); }); } catch {}
    }
  };
})();

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
  const[products,setProducts]   = useState<any[]>([]);
  const[logistics,setLogistics] = useState(()=>LS.get("ifb_logistics",SEED_LOGISTIC));
  const[prices,setPrices]       = useState(()=>LS.get("ifb_prices",SEED_PRICES));
  const[fx,setFx]               = useState(()=>LS.get("ifb_fx",SEED_FX));
  const[xrefs,setXrefs]         = useState<any[]>([]);
  const[airList,setAirList]     = useState<any[]>([]);
  const[salesRows,setSalesRows] = useState<any[]>([]);
  const[importLogs,setImportLogs] = useState(()=>LS.get("ifb_importlogs",[]));
  const[snapshots,setSnapshots]   = useState(()=>LS.get("ifb_snapshots",[]));
  const[costHistory,setCostHistory] = useState(()=>LS.get("ifb_costhistory",[]));
  const[lastImportTs,setLastImportTs] = useState(()=>LS.get("ifb_last_import_ts",0));
  const[lastCalcTs,setLastCalcTs]     = useState(()=>LS.get("ifb_last_calc_ts",0));
  const[page,setPage]     = useState(()=>LS.get("ifb_branch","")?"dashboard":"branchSelect");
  const[branch,setBranch] = useState(()=>LS.get("ifb_branch",""));
  const[month,setMonth]   = useState(NOW());
  const[toast,setToast]   = useState(null);
  const[pageFilter, setPageFilter] = useState(null);
  const [meatPrices, setMeatPrices] = useState<any[]>(() => LS.get("ifb_meatprices", []));

  const navigate = (pageName, filter=null) => { setPageFilter(filter); setPage(pageName); };

  const branchRef = useRef(branch);
  useEffect(()=>{ branchRef.current = branch; },[branch]);
  useEffect(()=>{ if(branchRef.current&&products.length) IDB.set(`ifb_products_${branchRef.current}`, products); },[products]);
  useEffect(()=>{ if(logistics.length) LS.set("ifb_logistics",      logistics); }, [logistics]);
  useEffect(()=>{ if(branch) LS.set(`ifb_airlist_${branch}`, airList); },[airList,branch]);
  useEffect(()=>{ if(branch) IDB.set(`ifb_sales_invoice_${branch}`, salesRows); },[salesRows,branch]);
  useEffect(()=>{ if(prices.length)    LS.set("ifb_prices",         prices);    }, [prices]);
  useEffect(()=>{ if(branch) LS.set("ifb_branch",branch); },[branch]);
  useEffect(()=>{ if(meatPrices.length) LS.set("ifb_meatprices", meatPrices); }, [meatPrices]);
  // Ricarica dati branch-specifici ad ogni cambio filiale
  useEffect(()=>{
    if(!branch) return;
    (async()=>{
      setProducts(await IDB.get(`ifb_products_${branch}`,[]));
      setXrefs(LS.get(`ifb_xrefs_${branch}`,[]));
      setAirList(LS.get(`ifb_airlist_${branch}`,[]));
      setSalesRows(await IDB.get(`ifb_sales_invoice_${branch}`,[]));
    })();
  },[branch]);

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
      const airEntry = airList.find((a:any)=>
          a.productId === prod.id ||
          (a.code && a.code === prod.code) ||
          (a.nHK && prod.nHK && a.nHK === prod.nHK)
        );
      if(airEntry && isAirTransport(airEntry.transportation))
        return { ...prod, cost:null, prevCost:null, priceInput:null, isAir:true, skipReason:"AIR" };

      const logRaw = logistics.find(l=>l.productId===prod.id&&l.branch===branch);
      if(!logRaw) return { ...prod, cost:null, prevCost:null, priceInput:null, skipReason:"NO LOGISTICA" };

      // Apply pltPerContainer default formula if value from Work_tab is 0 or missing
      // Formula: =SE(ProductType="DRY";25;SE(OR(ProductType="FRESH";ProductType="FROZEN");23;0))
      const pltFromFile = logRaw.pltPerContainer || 0;
      const effectiveTemp = logRaw.temperatureOverride || prod.temperature;
      const plt = pltFromFile > 0 ? pltFromFile : pltDefault(effectiveTemp);
      const log = { ...logRaw, pltPerContainer: plt };

      const pr     = prices.find(p=>p.productId===prod.id&&p.branch===branch&&p.month===month);
      const prPrev = prices.find(p=>p.productId===prod.id&&p.branch===branch&&p.month===prevM);

      const ub = log.ubicazione;

      if(!pr) {
        // Fallback: listino carne (€/kg → €/unit)
        const meat = meatPrices.find((m:any) =>
          m.code === prod.code ||
          m.code === String(prod.id) ||
          (prod.nHK && m.code === prod.nHK)
        );
        if(!meat) return { ...prod, cost:null, prevCost:null, priceInput:null, ubicazione:ub, skipReason:`NO PREZZO (${branch}/${month})` };
        const kgPerUnit =
          prod.uom === "KG"  ? 1 :
          prod.uom === "BOX" ? (Number(prod.kgPerBox)||0) :
          (Number(prod.kgPerBox)||0) / Math.max(Number(prod.qtyPerBox)||1, 1);
        const pi = meat.pricePerKg * kgPerUnit;
        const effectiveProd2 = log.temperatureOverride ? { ...prod, temperature: log.temperatureOverride } : prod;
        const cost2 = calcHK({ priceInput:pi, ubicazione:ub, product:effectiveProd2, logistic:{...log,category:prod.category}, eurToHkd:fxRate });
        return { ...prod, cost:cost2, prevCost:null, delta:null, priceInput:pi, isNew:true,
          flagged:false, ubicazione:ub, pltUsed:plt, temperatureOverride:log.temperatureOverride||null,
          skipReason: cost2 ? undefined : "CALC=0", _fromMeatList:true };
      }

            const pi  = selectPrice(pr, ub);
      const piP = prPrev ? selectPrice(prPrev, ub) : null;

      // Se il prezzo da listino è zero, prova il fallback listino carne
      if (!pi || pi === 0) {
        const meat = meatPrices.find((m:any) =>
          m.code === prod.code ||
          m.code === String(prod.id) ||
          (prod.nHK && m.code === prod.nHK)
        );
        if (meat) {
          const kgPerUnit =
            prod.uom === "KG"  ? 1 :
            prod.uom === "BOX" ? (Number(prod.kgPerBox)||0) :
            (Number(prod.kgPerBox)||0) / Math.max(Number(prod.qtyPerBox)||1, 1);
          const piMeat = meat.pricePerKg * kgPerUnit;
          const effectiveProdM = log.temperatureOverride ? { ...prod, temperature: log.temperatureOverride } : prod;
          const costM = calcHK({ priceInput:piMeat, ubicazione:ub, product:effectiveProdM, logistic:{...log,category:prod.category}, eurToHkd:fxRate });
          return { ...prod, cost:costM, prevCost:null, delta:null, priceInput:piMeat, isNew:true,
            flagged:false, ubicazione:ub, pltUsed:plt, temperatureOverride:log.temperatureOverride||null,
            skipReason: costM ? undefined : "CALC=0", _fromMeatList:true };
        }
      }

            // Temperatura rettificata dal Work_tab ha precedenza su quella dell'anagrafica
      const effectiveProd = log.temperatureOverride ? { ...prod, temperature: log.temperatureOverride } : prod;
      const cost = calcHK({ priceInput:pi, ubicazione:ub, product:effectiveProd, logistic:{...log,category:prod.category}, eurToHkd:fxRate });
      if(!cost) return { ...prod, cost:null, prevCost:null, priceInput:pi,
        skipReason: !pi || pi === 0 ? "PREZZO ZERO" : `CALC=0 (qty=${prod.qtyPerBox} box/plt=${prod.boxPerPallet} plt=${plt} uom=${prod.uom})` };

      const prevCost = piP!=null ? calcHK({ priceInput:piP, ubicazione:ub, product:effectiveProd, logistic:{...log,category:prod.category}, eurToHkd:fxRate }) : null;
      
      const delta    = cost&&prevCost ? (cost.step2Hkd-prevCost.step2Hkd)/prevCost.step2Hkd*100 : null;
      return { ...prod, cost, prevCost, delta, priceInput:pi, isNew:!prPrev,
        flagged: delta!==null && Math.abs(delta)>=3, ubicazione:ub, pltUsed:plt,
        temperatureOverride: log.temperatureOverride || null };
    });
  }, [products,logistics,prices,fx,airList,meatPrices,branch,month]);

  const isCAN = branch === "CAN";

  const NAV = [
    {id:"dashboard",  icon:"⬡", label:"Dashboard"},
    {id:"products",   icon:"◈", label:"Anagrafica", badge:"⇪"},
    {id:"xref",       icon:"⇄", label:isCAN?"XRef N COMIT / IFB":"XRef N / IFB"},
    {id:"logistics",  icon:"◎", label:isCAN?"Work Tab (Logistica)":"Logistica"},
    {id:"prices",     icon:"◉", label:"Listini", badge:"💶"},
    {id:"meatlist",   icon:"🥩", label:"Listino Carne"},
    ...(!isCAN ? [{id:"fx",  icon:"◌", label:"Cambi"}] : []),
    ...(!isCAN ? [{id:"air", icon:"✈", label:"AIR Transport"}] : []),
    {id:"costs",      icon:"◆", label:"Standard Cost"},
    {id:"invoice",    icon:"📋", label:"Fatture & Costi", badge:"⇪"},
    {id:"storico",    icon:"⧖", label:"Storico & Diff"},
    {id:"mail",       icon:"◻", label:"Mail Mensile"},
    {id:"notes",      icon:"📝", label:"Guida & Istruzioni"},
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
  const costSnaps = snapshots.filter((s:any) => s.type === "cost");

  const pages = {
    dashboard:   <Dashboard costRows={costRows} branch={branch} month={month} navigate={navigate}/>,
    products: <Products 
  products={products} 
  setProducts={setProducts} 
  branch={branch} 
  importLogs={importLogs}
  setImportLogs={setImportLogs}
  snapshots={snapshots}
  setSnapshots={setSnapshots}
  showToast={showToast}
  bumpImportTs={bumpImportTs}
/>,
  
    xref:        <XRefPage 
      xrefs={xrefs} 
      setXrefs={setXrefs} 
      branch={branch} 
      snapshots={snapshots}
      setSnapshots={setSnapshots}
      importLogs={importLogs}
      setImportLogs={setImportLogs}
      showToast={showToast}
      bumpImportTs={bumpImportTs}
    />,
    logistics: <Logistics 
  logistics={logistics} 
  setLogistics={setLogistics} 
  products={products} 
  branch={branch} 
  showToast={showToast} 
  bumpImportTs={bumpImportTs} 
  initFilter={pageFilter} 
  importLogs={importLogs} 
  setImportLogs={setImportLogs}
  xrefs={xrefs}  
/>,
    prices: <Prices 
  prices={prices} 
  setPrices={setPrices} 
  products={products} 
  branch={branch} 
  month={month}
  salesRows={salesRows} 
  xrefs={xrefs}
  importLogs={importLogs}
  setImportLogs={setImportLogs}
  snapshots={snapshots}
  setSnapshots={setSnapshots}
  showToast={showToast}
  bumpImportTs={bumpImportTs}
/>,


    
    fx:          <FxRates fx={fx} setFx={setFx} branch={branch} month={month}/>,
    air:         <AirListPage airList={airList} setAirList={setAirList} products={products} xrefs={xrefs} branch={branch} snapshots={snapshots} setSnapshots={setSnapshots} importLogs={importLogs} setImportLogs={setImportLogs} showToast={showToast} bumpImportTs={bumpImportTs}/>,
    meatlist: <MeatPriceListPage meatPrices={meatPrices} setMeatPrices={setMeatPrices} products={products} xrefs={xrefs} importLogs={importLogs} setImportLogs={setImportLogs} snapshots={snapshots} setSnapshots={setSnapshots} showToast={showToast} bumpImportTs={bumpImportTs}/>,
    costs:       <CostTable costRows={costRows} branch={branch} month={month} logistics={logistics} lastImportTs={lastImportTs} lastCalcTs={lastCalcTs} setLastCalcTs={setLastCalcTs} setCostHistory={setCostHistory} initFilter={pageFilter} salesRows={salesRows} products={products} xrefs={xrefs}/>,
    invoice: <InvoiceAndCosts rows={salesRows} setRows={setSalesRows} branch={branch} airList={airList} products={products} xrefs={xrefs} costRows={costRows} snapshots={snapshots} setSnapshots={setSnapshots} importLogs={importLogs} setImportLogs={setImportLogs} showToast={showToast} bumpImportTs={bumpImportTs}/>,
    storico: <Storico 
      snapshots={snapshots}
      setSnapshots={setSnapshots}
      costHistory={costHistory}
      setCostHistory={setCostHistory}
      branch={branch}
      showToast={showToast}
    />,
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
              <button key={key} onClick={()=>{ setBranch(key); setPage("dashboard"); }}
                style={{padding:"5px 8px",background:branch===key?`${c.color}20`:"transparent",
                  border:`1px solid ${branch===key?c.color:"transparent"}`,borderRadius:"6px",
                  color:branch===key?c.color:T.muted,fontFamily:"inherit",
                  fontSize:"11px",textAlign:"left",display:"flex",alignItems:"center",gap:"6px",
                  opacity:c.active?1:0.45,cursor:c.active?"pointer":"default"}}>
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
        <div style={{flex:1,paddingTop:"20px",paddingLeft:"28px",paddingBottom:"20px",paddingRight:0,overflow:"auto",width:"calc(100% - 200px)",boxSizing:"border-box"}}>
  {isCAN && (page==="air" || page==="fx") ? pages["dashboard"] : pages[page]}
</div>
      </div>
      {toast&&<div style={{position:"fixed",bottom:"24px",right:"24px",padding:"10px 18px",background:toast.color,borderRadius:"8px",color:"#fff",fontSize:"12px",fontWeight:"bold",boxShadow:"0 8px 24px rgba(0,0,0,0.4)",zIndex:1000}}>{toast.msg}</div>}
    </div>
  );
}

// ─── XREF PAGE ────────────────────────────────────────────────────────────────
function XRefPage({xrefs,setXrefs,branch,snapshots,setSnapshots,importLogs,setImportLogs,showToast,bumpImportTs}) {
  const branchCode = branch === "CAN" ? "N COMIT" : "N HK";
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
    setXrefs(next);LS.set(`ifb_xrefs_${branch}`,next);
    const log={id,type:"xref",fileName,date:new Date(id).toISOString(),count:incoming.length,diffs,branch:"ALL"};
    const newLogs=[log,...importLogs];setImportLogs(newLogs);LS.set("ifb_importlogs",newLogs);
    const newSnaps=[log,...snapshots].slice(0,50);setSnapshots(newSnaps);LS.set("ifb_snapshots",newSnaps);
    bumpImportTs();showToast(`XRef: ${incoming.length} voci · ${diffs.filter(d=>d.isNew).length} nuove ✓`,T.gold);
    setStep("main");setPreview([]);setRawRows([]);setHeaders([]);
  }

  const displayed=xrefs.filter(x=>!search||x.nHK?.toLowerCase().includes(search.toLowerCase())||x.ifbNo?.toLowerCase().includes(search.toLowerCase()));

  return(
    <div>
      <PageHeader title={`⇄ XRef ${branchCode} / IFB N · ${branch}`} sub="Codici filiale ↔ IFB N — ogni filiale ha la propria tabella"/>
      {step==="map"&&(
        <Section title={`Mappatura — ${fileName} · ${rawRows.length} righe`}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"16px",marginBottom:"16px"}}>
            {[[`Colonna ${branchCode} *`,colNHK,setColNHK],["Colonna IFB N *",colIFB,setColIFB]].map(([lbl,val,setter])=>(
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
              <THead cols={[branchCode,"IFB N","Stato"]} sticky />
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
            <div style={{fontSize:"11px",color:T.muted}}>Due colonne: {branchCode} · IFB N</div>
            <input id="_xref_in" type="file" accept=".xlsx,.xls,.csv"
              onChange={e=>{const f=e.target.files?.[0];if(f)parseFile(f);e.target.value="";}} style={{display:"none"}}/>
          </div>
          <SearchBar value={search} onChange={setSearch} placeholder={`🔍 Cerca per ${branchCode} o IFB N…`}/>
          {xrefs.length>0&&(
            <div style={{marginBottom:"10px",display:"flex",justifyContent:"flex-end"}}>
              <button onClick={()=>{if(window.confirm(`Eliminare tutte le ${xrefs.length} XRef di ${branch}?`)){setXrefs([]);LS.set(`ifb_xrefs_${branch}`,[]);}}}
                style={{padding:"5px 14px",background:"none",border:`1px solid ${T.red}44`,borderRadius:"6px",color:T.red,cursor:"pointer",fontSize:"11px"}}>
                ✕ Svuota lista ({xrefs.length})
              </button>
            </div>
          )}
          <Section title={`${displayed.length} / ${xrefs.length} corrispondenze`}>
            {xrefs.length===0?<div style={{padding:"24px",textAlign:"center",color:T.dim,fontSize:"13px"}}>Nessuna XRef caricata.</div>:(
              <table style={{width:"100%",borderCollapse:"collapse"}}>
                <THead cols={[branchCode,"IFB N","Azioni"]} sticky />
                <tbody>{displayed.map((x,i)=>(
                  <tr key={x.nHK+i} style={{borderBottom:`1px solid ${T.border}`}}>
                    <TD mono><span style={{color:T.gold}}>{x.nHK}</span></TD>
                    <TD mono>{x.ifbNo}</TD>
                    <TD><MiniBtn label="✕" onClick={()=>{const n=xrefs.filter((_,j)=>j!==xrefs.indexOf(x));setXrefs(n);LS.set(`ifb_xrefs_${branch}`,n);}} color={T.red}/></TD>
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

  // ✅ NUOVA SEZIONE ISTRUZIONI
  const instructions = [
    {
      title: "🏁 1. Selezione Filiale",
      steps: [
        "All'avvio, seleziona la filiale desiderata (Hong Kong è attualmente l'unica attiva)",
        "Il mese viene impostato automaticamente sul mese corrente",
        "Puoi cambiare filiale o mese in qualsiasi momento dalla sidebar sinistra"
      ]
    },
    {
      title: "📦 2. Importazione Anagrafica",
      steps: [
        "Vai su 'Anagrafica' → clicca '📂 Carica anagrafica (BC export)'",
        "Seleziona il file Excel esportato da Business Central",
        "Mappa le colonne se necessario (il sistema prova a farlo automaticamente)",
        "Clicca 'Preview' per verificare i dati, poi 'Importa'",
        "I dati vengono salvati automaticamente e rimangono disponibili al prossimo accesso"
      ]
    },
    {
      title: "💶 3. Importazione Listini",
      steps: [
        "Vai su 'Listini' → clicca '📂 Carica listini'",
        "Seleziona il file PBI o CURRENT PRICELIST",
        "Seleziona il mese di riferimento (se diverso da quello corrente)",
        "Verifica il mapping automatico del codice articolo",
        "Clicca 'Preview' per vedere i prezzi trovati, poi 'Importa'",
        "I prezzi vengono salvati per mese e filiale"
      ]
    },
    {
      title: "🚚 4. Logistica (Work_tab)",
      steps: [
        "Vai su 'Logistica' → clicca '📂 Carica Work_tab'",
        "Seleziona il file 08_Work_Tab.xlsx",
        "Il sistema rileva automaticamente colonne: Ubicazione, Area, Plt/Container, Health Certificate, Carriage, Tassa Alcolica",
        "Clicca 'Importa' per salvare i parametri logistici",
        "⚠️ Le righe importate sono in SOLA LETTURA (colore dorato)",
        "Le righe senza logistica sono modificabili (colore arancione)",
        "Puoi svuotare tutti i dati con '🗑 Svuota tutti i dati'"
      ]
    },
    {
      title: "✈️ 5. Lista AIR Transport",
      steps: [
        "Vai su 'AIR Transport' → clicca '📂 Carica lista AIR'",
        "Carica un file Excel con una colonna di codici articolo (N HK o IFB N)",
        "Il sistema mostra quali codici sono trovati/non trovati in anagrafica",
        "Clicca 'Salva' per marcare questi articoli come trasporto aereo",
        "Questi articoli verranno ESCLUSI dal calcolo Standard Cost",
        "Nota: anche la location 'NCJ' in fattura viene considerata AIR automaticamente"
      ]
    },
    {
      title: "📊 6. Calcolo Standard Cost",
      steps: [
        "Vai su 'Standard Cost' → clicca '⟳ Ricalcola & Salva'",
        "Il sistema calcola il costo per tutti gli articoli INALCA F&B con trasporto SEA",
        "Vedi la tabella con breakdown dettagliato: Prezzo, FOB, LIC, VGM, HC, Pallet, Tassa Alcolica, Magazzino",
        "Clicca su una riga per vedere il dettaglio completo del calcolo",
        "Le variazioni ≥ ±3% sono evidenziate",
        "I costi vengono salvati come snapshot per confronti futuri"
      ]
    },
    {
      title: "📋 7. Sales Invoice (Fatture)",
      steps: [
        "Vai su 'Sales Invoice' → clicca '📂 Ricarica'",
        "Carica il file fattura (Excel/CSV)",
        "Mappa le colonne: Codice articolo, Descrizione, Data, Quantità, Prezzo, Location",
        "Il trasporto AIR/SEA viene determinato automaticamente dalla lista AIR o dalla location 'NCJ'",
        "Clicca 'Importa' per salvare i dati",
        "Puoi filtrare per 'Solo AIR' o 'Solo SEA'"
      ]
    },
    {
      title: "📈 8. Dashboard & Monitoraggio",
      steps: [
        "La Dashboard mostra statistiche riassuntive: costi calcolati, variazioni, AIR esclusi, mancanti",
        "Clicca su qualsiasi card per vedere la lista dettagliata degli articoli in quella categoria",
        "Usa i filtri e la ricerca per trovare articoli specifici"
      ]
    },
    {
      title: "⏱️ 9. Storico & Snapshot",
      steps: [
        "Vai su 'Storico & Diff' per vedere tutti gli import effettuati",
        "Clicca su uno snapshot per vedere le differenze rispetto alla versione precedente",
        "Gli snapshot Standard Cost vengono creati ogni volta che clicchi 'Ricalcola & Salva'",
        "Puoi eliminare snapshot singoli o tutti"
      ]
    },
    {
      title: "✉️ 10. Mail Mensile",
      steps: [
        "Vai su 'Mail Mensile' → il sistema prepara automaticamente il testo delle variazioni > ±3%",
        "Clicca 'Copia testo' per copiare negli appunti e incollare nella mail"
      ]
    }
  ];

  return(
    <div>
      <PageHeader title="📝 Guida & Istruzioni" sub="Manuale d'uso · regole di calcolo · note tecniche"/>

      {/* ✅ NUOVA SEZIONE ISTRUZIONI */}
      <Section title="" accent={T.gold}>
        <div style={{display:"flex",flexDirection:"column",gap:"16px"}}>
          {instructions.map((inst, idx) => (
            <div key={idx} style={{
              background:T.card,
              borderLeft: `4px solid ${T.gold}`,
              borderRadius:"8px",
              padding:"12px 16px"
            }}>
              <div style={{
                fontSize:"13px",
                fontWeight:"bold",
                color:T.gold,
                marginBottom:"8px",
                display:"flex",
                alignItems:"center",
                gap:"8px"
              }}>
                <span style={{fontSize:"16px"}}>📌</span>
                {inst.title}
              </div>
              <ul style={{
                margin:0,
                paddingLeft:"20px",
                fontSize:"12px",
                color:T.muted,
                lineHeight:"1.8"
              }}>
                {inst.steps.map((step, i) => (
                  <li key={i}>{step}</li>
                ))}
              </ul>
            </div>
          ))}
        </div>
      </Section>

      {/* Sezioni esistenti */}
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

  // Funzione per verificare se un codice è valido (NON Power BI)
  function isValidCode(code) {
    if(!code) return false;
    const str = String(code).trim();
    if(/^P_/i.test(str)) return false;
    if(/^\d{7,}$/.test(str.replace(/[^0-9]/g, ""))) return false;
    if(str.includes("P_BC_")) return false;
    return true;
  }

  function parseFile(file) {
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, {type:"binary"});
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(ws, {header:1, defval:""});
        if(data.length < 2) { showToast("File vuoto", T.red); return; }
        
        const hdrs = data[0].map(h => String(h||"").trim());
        const rows = data.slice(1).filter(r => r.some(c => c !== ""));
        setHeaders(hdrs);
        setRawRows(rows);
        
        // Auto-mapping dei campi (solo quelli necessari)
        const am = {};
        
        // Mappa il codice (obbligatorio)
        const codeAliases = ["no_", "no.", "no", "item no.", "codice", "code", "n hk", "ifb item", "ifb no", "ifb n"];
        for(const h of hdrs) {
          const hl = h.toLowerCase().trim();
          if(codeAliases.some(a => hl === a || hl.includes(a))) {
            am["code"] = h;
            break;
          }
        }
        
        // Mappa la descrizione (opzionale)
        const descAliases = ["description", "descrizione", "desc", "item description"];
        for(const h of hdrs) {
          const hl = h.toLowerCase().trim();
          if(descAliases.some(a => hl === a || hl.includes(a))) {
            am["description"] = h;
            break;
          }
        }
        
        // Mappa i prezzi
        const priceFields = ["mtsPrice", "fcaPrice", "fcaDiscount", "fcaDiscounted", "dapPrice", "dapDiscount", "dapDiscounted", "dapFinalDirect"];
        priceFields.forEach(field => {
          const aliases = PRICE_FIELD_ALIASES[field] || [];
          for(const h of hdrs) {
            const hl = h.toLowerCase().trim();
            if(aliases.some(a => hl === a || (a.length > 3 && hl.includes(a)))) {
              am[field] = h;
              break;
            }
          }
        });
        
        setMapping(am);
        setStep("map");
      } catch(err) { 
        showToast("Errore: "+err.message, T.red); 
      }
    };
    reader.readAsBinaryString(file);
  }

  function buildPreview() {
    const get = (row, field) => {
      const col = mapping[field];
      if(!col) return null;
      const i = headers.indexOf(col);
      return i >= 0 ? row[i] : null;
    };
    
    let skipped = 0;
    let notFound = 0;
    
    const mapped = rawRows.map((row, idx) => {
      const rawCode = String(get(row, "code") || "").trim();
      const rawDescription = String(get(row, "description") || get(row, "code") || "").trim();
      
      if(!rawCode) { skipped++; return null; }
      if(!isValidCode(rawCode)) { skipped++; return null; }
      
      const prod = findProduct(rawCode, products, xrefs);
      
      const mtsPrice = parseFloat(get(row, "mtsPrice")) || 0;
      const fcaPrice = parseFloat(get(row, "fcaPrice")) || 0;
      const fcaDiscount = parseFloat(get(row, "fcaDiscount")) || 0;
      const fcaDiscounted = parseFloat(get(row, "fcaDiscounted")) || (fcaPrice - (fcaDiscount * fcaPrice / 100)) || 0;
      const dapPrice = parseFloat(get(row, "dapPrice")) || 0;
      const dapDiscount = parseFloat(get(row, "dapDiscount")) || 0;
      const dapDiscounted = parseFloat(get(row, "dapDiscounted")) || (dapPrice - (dapDiscount * dapPrice / 100)) || 0;
      const dapFinalDirect = parseFloat(get(row, "dapFinalDirect")) || 0;
      
      let dapFinal = 0;
      let dapNote = "";
      if(dapFinalDirect !== 0) {
        dapFinal = dapFinalDirect;
        dapNote = "da file";
      } else if(prod) {
        dapFinal = dapDiscounted || 0;
        dapNote = dapDiscounted ? "da DAP Disc." : "";
      }
      
      const existing = prod ? prices.find(p => p.productId === prod.id && p.branch === branch && p.month === importMonth) : null;
      
      return {
        _idx: idx,
        rawCode,
        ifbNo_from_file: rawCode,
        description_from_file: rawDescription,
        productId: prod?.id || null,
        nHK_from_anag: prod?.nHK || "",
        ifbNo_from_anag: prod?.code || "",
        description_from_anag: prod?.description || "",
        dapFinal: roundN(dapFinal),
        mtsPrice: roundN(mtsPrice),
        fcaDiscounted: roundN(fcaDiscounted),
        dapPrice: roundN(dapPrice),
        fcaPrice: roundN(fcaPrice),
        dapNote,
        _hasProduct: !!prod,
        _existing: !!existing
      };
    }).filter(Boolean);
    
    setPreview(mapped);
    setStep("preview");
  }

  function executeImport() {
    const snId = Date.now();
    const updated = [...prices];
    const diffs = [];
    let count = 0, newCount = 0, changed = 0;
    
    preview.forEach(r => {
      if(!r._hasProduct) return;
      
      const idx = updated.findIndex(p => p.productId === r.productId && p.branch === branch && p.month === importMonth);
      const entry = {
        productId: r.productId,
        branch,
        month: importMonth,
        dapFinal: r.dapFinal,
        mtsPrice: r.mtsPrice,
        fcaDiscounted: r.fcaDiscounted,
        dapPrice: r.dapPrice,
        fcaPrice: r.fcaPrice
      };
      const prev = idx >= 0 ? updated[idx] : null;
      const diffFields = [];
      
      ["dapFinal","mtsPrice","fcaDiscounted","dapPrice","fcaPrice"].forEach(f => {
        const oldR = roundN(prev?.[f] || 0);
        const newR = roundN(entry[f] || 0);
        if(Math.abs(oldR - newR) >= 0.005) {
          diffFields.push({field: f, old: oldR, new: newR, delta: oldR > 0 ? ((newR - oldR) / oldR * 100) : null});
        }
      });
      
      if(!prev) newCount++;
      else if(diffFields.length > 0) changed++;
      
      if(diffFields.length > 0 || !prev) {
        diffs.push({
          productId: r.productId,
          nHK: r.nHK_from_anag,
          ifbNo: r.ifbNo_from_anag,
          description: r.description_from_anag,
          isNew: !prev,
          fields: diffFields
        });
      }
      
      if(idx >= 0) updated[idx] = entry;
      else updated.push(entry);
      count++;
    });
    
    setPrices(updated);
    LS.set("ifb_prices", updated);
    
    const log = {
      id: snId,
      type: "prices",
      fileName,
      branch,
      month: importMonth,
      date: new Date(snId).toISOString(),
      count,
      newCount,
      updateCount: changed,
      diffs
    };
    
    const newLogs = [log, ...importLogs];
    setImportLogs(newLogs);
    LS.set("ifb_importlogs", newLogs);
    
    const newSnaps = [log, ...snapshots].slice(0, 50);
    setSnapshots(newSnaps);
    LS.set("ifb_snapshots", newSnaps);
    
    setDoneInfo({ count, newCount, changed, unchanged: count - newCount - changed });
    bumpImportTs();
    setStep("done");
  }

  const reset = () => {
    setStep("upload");
    setRawRows([]);
    setHeaders([]);
    setFileName("");
    setMapping({});
    setPreview([]);
    setDoneInfo(null);
  };

  if(step === "done" && doneInfo) {
    return (
      <div>
        <PageHeader title="✓ Import Listini completato" sub={fileName}/>
        <div style={{padding:"20px", background:`${T.green}11`, border:`1px solid ${T.green}33`, borderRadius:"8px", marginBottom:"16px", fontSize:"13px", color:T.muted, lineHeight:"2"}}>
          Mese: <strong style={{color:T.gold}}>{importMonth}</strong> · Filiale: <strong style={{color:T.text}}>{branch}</strong><br/>
          Prezzi totali: <strong style={{color:T.text}}>{doneInfo.count}</strong> &nbsp;·&nbsp;
          <span style={{color:T.green}}>🆕 {doneInfo.newCount} nuovi</span> &nbsp;·&nbsp;
          <span style={{color:T.orange}}>✏️ {doneInfo.changed} modificati</span> &nbsp;·&nbsp;
          <span style={{color:T.dim}}>{doneInfo.unchanged} invariati</span>
        </div>
        <ActionBtn label="💶 Nuovo import" onClick={reset} primary/>
      </div>
    );
  }

  return (
    <div>
      <PageHeader title="💶 Import Listini" sub={`${branch} · importa da file PBI / CURRENT PRICELIST`}/>
      <StepBar steps={["upload","map","preview","done"]} current={step}/>
      
      {step === "upload" && (
        <div>
          <Section title="Mese di riferimento">
            <div style={{display:"flex", alignItems:"center", gap:"12px"}}>
              <span style={{fontSize:"12px", color:T.muted}}>Mese:</span>
              <input type="month" value={importMonth} onChange={e => setImportMonth(e.target.value)} style={{...inputStyle(), width:"160px"}}/>
            </div>
          </Section>
          <Section title="Carica file export PBI / CURRENT PRICELIST">
            <DropZone onFile={parseFile}/>
          </Section>
        </div>
      )}
      
      {step === "map" && (
        <Section title={`Mappatura — ${fileName} · ${rawRows.length} righe`}>
          <div style={{display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:"12px", marginBottom:"18px"}}>
            <div>
              <label style={{display:"block", fontSize:"11px", color:T.gold, marginBottom:"5px"}}>📌 Codice * (N HK o IFB N)</label>
              <select 
                value={mapping["code"] || ""} 
                onChange={e => setMapping(m => ({...m, code: e.target.value || null}))} 
                style={{...inputStyle(), cursor:"pointer", borderColor:!mapping["code"] ? T.red+"88" : T.border}}
              >
                <option value="">— seleziona colonna —</option>
                {headers.map(h => <option key={h} value={h}>{h}</option>)}
              </select>
            </div>
            
            <div>
              <label style={{display:"block", fontSize:"11px", color:T.muted, marginBottom:"5px"}}>📝 Descrizione</label>
              <select 
                value={mapping["description"] || ""} 
                onChange={e => setMapping(m => ({...m, description: e.target.value || null}))} 
                style={{...inputStyle(), cursor:"pointer"}}
              >
                <option value="">— non mappato —</option>
                {headers.map(h => <option key={h} value={h}>{h}</option>)}
              </select>
            </div>
          </div>
          
          <div style={{marginTop:"8px", padding:"8px", background:`${T.gold}08`, borderRadius:"6px", fontSize:"11px", color:T.muted}}>
            ⚡ I campi prezzi (MTS Price, FCA Price, DAP Price, etc.) vengono rilevati automaticamente.
          </div>
          
          <div style={{display:"flex", gap:"10px", marginTop:"16px"}}>
            <ActionBtn label="← Ricarica" onClick={reset}/>
            <ActionBtn label="Preview →" onClick={buildPreview} primary disabled={!mapping["code"]}/>
          </div>
        </Section>
      )}
      
      {step === "preview" && (
        <div>
          <div style={{display:"flex", gap:"12px", marginBottom:"16px", flexWrap:"wrap"}}>
            {[
              [preview.filter(r => r._hasProduct).length, "✅ Trovati in anagrafica", T.green],
              [preview.filter(r => !r._hasProduct).length, "❌ NON trovati in anagrafica", T.red],
              [preview.filter(r => r._existing).length, "✏️ Aggiornamenti", T.orange],
              [preview.filter(r => !r._existing && r._hasProduct).length, "🆕 Nuovi", T.gold]
            ].map(([n, l, c]) => (
              <div key={l} style={{padding:"10px 16px", background:T.card, border:`1px solid ${T.border}`, borderRadius:"8px"}}>
                <div style={{fontSize:"20px", fontWeight:"bold", color:c}}>{n}</div>
                <div style={{fontSize:"10px", color:T.dim, marginTop:"2px"}}>{l}</div>
              </div>
            ))}
          </div>
          
          <div style={{display:"flex", gap:"10px", marginBottom:"16px"}}>
            <ActionBtn label="← Torna" onClick={() => setStep("map")}/>
            <ActionBtn label={`✓ Importa ${preview.filter(r => r._hasProduct).length} prezzi per ${importMonth}`} onClick={executeImport} primary/>
          </div>
          
          <Section title={`Preview · ${importMonth} · ${branch}`}>
            <div style={{overflowX:"auto"}}>
              <table style={{width:"100%", borderCollapse:"collapse"}}>
                <thead>
                  <tr>
                  {["Codice (dal file)","Descrizione (dal file)","Match Anagrafica","DAP Final","MTS Price","FCA Disc.","Stato"].map(c=>(
                    <th key={c} style={{padding:"7px 12px",background:T.card,color:T.muted,textAlign:"left",borderBottom:`1px solid ${T.border}`,fontSize:"11px",position:"sticky",top:0,zIndex:10}}>{c}</th>
                  ))}
                   </tr>
                </thead>
                <tbody>
                  {preview.slice(0, 100).map(r => (
                    <tr key={r._idx} style={{borderBottom:`1px solid ${T.border}`, background: r._hasProduct ? (r._existing ? `${T.orange}08` : T.bg) : `${T.red}08`}}>
                      <td style={{padding:"7px 12px", fontSize:"12px", fontFamily:"monospace"}}><span style={{color: T.gold}}>{r.ifbNo_from_file}</span></td>
                      <td style={{padding:"7px 12px", fontSize:"12px"}}>{r.description_from_file || "—"}</td>
                      <td style={{padding:"7px 12px", fontSize:"12px"}}>{r._hasProduct ? <span style={{color: T.green}}>✓ {r.ifbNo_from_anag}</span> : <span style={{color: T.red}}>✗ non trovato</span>}</td>
                      <td style={{padding:"7px 12px", fontSize:"12px", fontFamily:"monospace"}}><span style={{color: T.gold}}>{r.dapFinal > 0 ? `€ ${r.dapFinal.toFixed(2)}` : "—"}</span>{r.dapNote && <span style={{marginLeft:"4px", fontSize:"9px", color:T.dim}}>({r.dapNote})</span>}</td>
                      <td style={{padding:"7px 12px", fontSize:"12px", fontFamily:"monospace"}}><span style={{color: T.blue}}>{r.mtsPrice > 0 ? `€ ${r.mtsPrice.toFixed(2)}` : "—"}</span></td>
                      <td style={{padding:"7px 12px", fontSize:"12px", fontFamily:"monospace"}}><span style={{color: T.muted}}>{r.fcaDiscounted > 0 ? `€ ${r.fcaDiscounted.toFixed(2)}` : "—"}</span></td>
                      <td style={{padding:"7px 12px", fontSize:"12px"}}>
                        {!r._hasProduct ? <span style={{color:T.red, fontSize:"10px"}}>❌ NON IN ANAGRAFICA</span> : r._existing ? <span style={{color:T.orange, fontSize:"10px"}}>✏️ AGGIORNAMENTO</span> : <span style={{color:T.green, fontSize:"10px"}}>🆕 NUOVO</span>}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            {preview.length > 100 && (
              <div style={{padding:"12px", textAlign:"center", color:T.muted, fontSize:"11px"}}>
                Mostrati primi 100 su {preview.length} risultati
              </div>
            )}
          </Section>
        </div>
      )}
    </div>
  );
}

// ─── IMPORT ANAGRAFICA BC ─────────────────────────────────────────────────────
function ImportBC({products,setProducts,branch,importLogs,setImportLogs,snapshots,setSnapshots,showToast,bumpImportTs}) {
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
    const m = {};
    for(const field of FIELDS){
      const aliases = LOCAL_ALIASES[field]||[];
      const h = hdrs.find(h=>aliases.some(a=>h.toLowerCase().replace(/[\s_]/g,"")=== a.replace(/[\s_]/g,"")));
      if(h){ m[field]=h; continue; }
      const h2 = hdrs.find(h=>aliases.some(a=>h.toLowerCase().replace(/[\s_]/g,"").includes(a.replace(/[\s_]/g,""))&&a.length>3));
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
      kgxplt: parseFloat(r.kgxplt) > 0
    ? parseFloat(r.kgxplt)
    : roundN((parseFloat(r.kgPerBox)||0) * (parseFloat(r.qtyPerBox)||1) * (parseFloat(r.boxPerPallet)||0)),
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
    const snap={id:now,type:"anagrafica",date:new Date(now).toISOString(),count:newProds.length,diffs,products:newProds,branch:"ALL"};
    const newSnaps=[snap,...snapshots].slice(0,50);setSnapshots(newSnaps);LS.set("ifb_snapshots",newSnaps);
    setProducts(newProds);
    const savedProd = LS.set(`ifb_products_${branch}`, newProds);
    if (!savedProd) showToast("⚠ LocalStorage piena: anagrafica NON salvata. Esporta i dati.", T.red);
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
          <THead cols={[branchN(branch),"IFB No","Descrizione","Vendor","Categoria","UOM","Qty/Box","Box/Plt","Kg/Box","Temp","Attivo"]}sticky/>
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
function AirListPage({airList,setAirList,products,xrefs,branch,snapshots,setSnapshots,importLogs,setImportLogs,showToast,bumpImportTs}) {
  const[step,setStep]=useState("main");
  const[headers,setHeaders]=useState([]);
  const[rawRows,setRawRows]=useState([]);
  const[colCode,setColCode]=useState("");
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
        setColCode(hdrs.find(h=>codeA.some(a=>h.toLowerCase()===a||h.toLowerCase().includes(a)))||"");
        setStep("map");
      }catch(err){showToast("Errore: "+err.message,T.red);}
    };
    reader.readAsBinaryString(file);
  }

  function buildPreview() {
    const iC=headers.indexOf(colCode);
    const mapped=rawRows.map(row=>{
      const code=String(row[iC]||"").trim();
      if(!code) return null;
      const prod=findProduct(code,products,xrefs);
      return{code,productId:prod?.id||null,description:prod?.description||code,nHK:prod?.nHK||"",_hasProduct:!!prod};
    }).filter(Boolean);
    setPreview(mapped);setStep("preview");
  }

  function executeImport() {
    const seen = new Set();
    const valid = preview
      .filter(r => r._hasProduct)
      .filter(r => { if(seen.has(r.productId)) return false; seen.add(r.productId); return true; });

    // Mantieni le altre filiali, sostituisce solo quella corrente
    const next = valid.map(r => ({
      productId: r.productId, code: r.code, nHK: r.nHK,
      description: r.description, transportation: "AIR", branch
    }));
    setAirList(next); LS.set(`ifb_airlist_${branch}`, next);
    const now = Date.now();
    IDB.set(`ifb_air_data_${now}`, next);
    const log = {id:now,type:"air",date:new Date(now).toISOString(),count:valid.length,diffs:[],branch};
    const newLogs = [log,...importLogs]; setImportLogs(newLogs); LS.set("ifb_importlogs",newLogs);
    const newSnaps = [log,...snapshots].slice(0,50); setSnapshots(newSnaps); LS.set("ifb_snapshots",newSnaps);
    bumpImportTs(); showToast(`AIR ${branch}: lista sostituita con ${valid.length} articoli ✓`, T.gold);
    setStep("main"); setPreview([]); setRawRows([]); setHeaders([]);
  }

  const branchAir = airList;
  const _sq=search.toLowerCase();
  const displayed=branchAir.filter((a:any)=>!search
    ||a.description?.toLowerCase().includes(_sq)
    ||a.code?.toLowerCase().includes(_sq)
    ||a.nHK?.toLowerCase().includes(_sq));

  return(
    <div>
      <PageHeader title="✈ AIR Transport" sub="Articoli trasportati via aerea — esclusi da Standard Cost (calcolo solo SEA)"/>

      {step==="map"&&(
        <Section title={`Mappatura — ${fileName}`}>
          <div style={{marginBottom:"16px",maxWidth:"320px"}}>
            <label style={{display:"block",fontSize:"11px",color:T.gold,marginBottom:"5px"}}>Colonna Codice * (N HK o IFB N)</label>
            <select value={colCode} onChange={e=>setColCode(e.target.value)} style={{...inputStyle(),cursor:"pointer"}}>
              <option value="">— seleziona —</option>
              {headers.map(h=><option key={h} value={h}>{h}</option>)}
            </select>
          </div>
          <div style={{fontSize:"11px",color:T.muted,marginBottom:"16px",padding:"8px 12px",background:`${T.gold}08`,borderRadius:"6px"}}>
            Tutti i prodotti di questo file verranno marcati come <strong style={{color:T.orange}}>✈ AIR</strong> — nessuna colonna Transportation richiesta.
          </div>
          <div style={{display:"flex",gap:"10px"}}>
            <ActionBtn label="← Ricarica" onClick={()=>setStep("main")}/>
            <ActionBtn label="Preview →" onClick={buildPreview} primary disabled={!colCode}/>
          </div>
        </Section>
      )}

{step==="preview"&&(
  <div>
    <div style={{display:"flex",gap:"12px",marginBottom:"16px",flexWrap:"wrap"}}>
      {[[preview.filter(r=>r._hasProduct).length,"✅ Trovati in anagrafica",T.green],
        [preview.filter(r=>!r._hasProduct).length,"❌ NON trovati in anagrafica",T.red],
        [preview.length,"📊 Totale",T.text]].map(([n,l,c])=>(
        <div key={l as string} style={{padding:"10px 16px",background:T.card,border:`1px solid ${c}44`,borderRadius:"8px"}}>
          <div style={{fontSize:"20px",fontWeight:"bold",color:c as string}}>{n as number}</div>
          <div style={{fontSize:"10px",color:T.dim,marginTop:"2px"}}>{l as string}</div>
        </div>
      ))}
    </div>
    
    <div style={{display:"flex",gap:"10px",marginBottom:"16px",flexWrap:"wrap"}}>
      <ActionBtn label="← Torna" onClick={()=>setStep("map")}/>
      <ActionBtn label={`✓ Salva ${preview.filter(r=>r._hasProduct).length} articoli AIR`} onClick={executeImport} primary/>
    </div>
    
    {/* ✅ SEZIONE NON TROVATI - visibile solo se ce ne sono */}
    {preview.filter(r=>!r._hasProduct).length > 0 && (
      <Section title={`❌ ${preview.filter(r=>!r._hasProduct).length} codici NON trovati in anagrafica`} accent={T.red}>
        <div style={{overflowX:"auto", marginBottom:"20px"}}>
          <table style={{width:"100%",borderCollapse:"collapse"}}>
            <THead cols={["Codice dal file","Descrizione dal file"]}sticky/>
            <tbody>
              {preview.filter(r=>!r._hasProduct).map((r,i)=>(
                <tr key={i} style={{borderBottom:`1px solid ${T.border}`,background:`${T.red}08`}}>
                  <TD mono><span style={{color:T.red, fontWeight:"bold"}}>{r.code}</span></TD>
                  <TD style={{color:T.muted}}>{r.description || "—"}</TD>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <div style={{fontSize:"11px", color:T.muted, padding:"8px 12px", background:`${T.red}10`, borderRadius:"6px"}}>
          💡 Questi codici non esistono nell'anagrafica. Verranno <strong>ignorati</strong> nell'import. 
          Verifica se sono digitati correttamente o aggiungili all'anagrafica.
        </div>
      </Section>
    )}
    
    {/* ✅ SEZIONE TROVATI (solo AIR) */}
    {preview.filter(r=>r._hasProduct).length > 0 && (
      <Section title={`✈ ${preview.filter(r=>r._hasProduct).length} articoli AIR trovati (verranno importati)`} accent={T.green}>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse"}}>
            <THead cols={["Codice",branchN(branch),"Descrizione"]}sticky/>
            <tbody>
              {preview.filter(r=>r._hasProduct).slice(0,100).map((r,i)=>(
                <tr key={i} style={{borderBottom:`1px solid ${T.border}`,background:i%2===0?T.bg:T.surface}}>
                  <TD mono><span style={{color:T.gold}}>{r.code}</span></TD>
                  <TD mono><span style={{color:T.muted}}>{r.nHK||"—"}</span></TD>
                  <TD>{r.description}</TD>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        {preview.filter(r=>r._hasProduct).length > 100 && (
          <div style={{padding:"8px", textAlign:"center", color:T.muted, fontSize:"11px"}}>
            Mostrati primi 100 su {preview.filter(r=>r._hasProduct).length}
          </div>
        )}
      </Section>
    )}
  </div>
)}

      {step==="main"&&(
         <>
         <div style={{marginBottom:"20px",display:"flex",gap:"10px",alignItems:"center",flexWrap:"wrap"}}>
         <label style={{display:"inline-block",padding:"10px 20px",background:T.gold,color:"#000",borderRadius:"6px",cursor:"pointer",fontWeight:"bold"}}>
          📂 Carica lista AIR
          <input type="file" accept=".xlsx,.xls,.csv"
            onChange={e=>{const f=e.target.files?.[0];if(f)parseFile(f);e.target.value="";}} style={{display:"none"}}/>
        </label>
        {importLogs.filter((l:any)=>l.type==="air"&&l.branch===branch).length>0&&(
          <select onChange={async e=>{
            if(!e.target.value) return;
            const snap=importLogs.find((l:any)=>String(l.id)===e.target.value);
            if(!snap) return;
            if(window.confirm(`Ripristinare la lista AIR del ${new Date(snap.id).toLocaleDateString("it-IT")} (${snap.count} articoli)?`)){
              const next=await IDB.get(`ifb_air_data_${snap.id}`, null);
              if(!next?.length){ showToast("Snapshot non disponibile — reimporta il file", T.orange); return; }
              setAirList(next);LS.set(`ifb_airlist_${branch}`,next);
              showToast(`Lista AIR ripristinata: ${snap.count} articoli ✓`,T.gold);
            }
            e.target.value="";
          }} style={{...inputStyle(),width:"auto",fontSize:"12px"}} defaultValue="">
            <option value="">📜 Carica da storico ({importLogs.filter((l:any)=>l.type==="air"&&l.branch===branch).length})</option>
            {importLogs.filter((l:any)=>l.type==="air"&&l.branch===branch).map((s:any)=>(
                  <option key={s.id} value={String(s.id)}>{new Date(s.id).toLocaleDateString("it-IT")} · {s.count} articoli</option>
            ))}
          </select>
        )}
 
{/* Bottone Mostra lista - stile coerente con il tema */}
<button
  onClick={() => {
    setSearch("");
    setStep("main");
  }}
  style={{
    padding: "8px 16px",
    background: branchAir.length > 0 ? `${T.blue}20` : T.surface,
    border: `1px solid ${branchAir.length > 0 ? T.blue : T.border}`,
    borderRadius: "6px",
    color: branchAir.length > 0 ? T.blue : T.muted,
    cursor: "pointer",
    fontSize: "12px",
    display: "flex",
    alignItems: "center",
    gap: "6px",
    transition: "all 0.2s"
  }}
  onMouseEnter={e => {
    if (branchAir.length > 0) {
      e.currentTarget.style.background = `${T.blue}30`;
      e.currentTarget.style.borderColor = T.blue;
    }
  }}
  onMouseLeave={e => {
    if (branchAir.length > 0) {
      e.currentTarget.style.background = `${T.blue}20`;
      e.currentTarget.style.borderColor = T.blue;
    } else {
      e.currentTarget.style.background = T.surface;
      e.currentTarget.style.borderColor = T.border;
    }
  }}
>
  <span style={{fontSize: "12px"}}>✈️</span>
  Lista AIR ({branchAir.length})
</button>
 {airList.length>0&&(
   <button
   onClick={()=>{if(window.confirm(`Eliminare i ${branchAir.length} articoli AIR di ${branch}?`)){
    setAirList([]);LS.set(`ifb_airlist_${branch}`,[]);
  }}}
     style={{padding:"8px 16px",background:"none",border:`1px solid ${T.red}44`,borderRadius:"6px",color:T.red,cursor:"pointer",fontSize:"12px"}}>
     ✕ Svuota lista ({branchAir.length})
   </button>
 )}
 <span style={{fontSize:"11px",color:T.muted}}>Colonna richiesta: N HK o IFB N · ogni import sostituisce la lista precedente</span>
</div>
          {airList.length>0&&(
            <>
              <SearchBar value={search} onChange={setSearch} placeholder="🔍 Cerca articolo AIR…"/>
              <Section title={`${displayed.length} articoli AIR (esclusi da Standard Cost)`}>
                <table style={{width:"100%",borderCollapse:"collapse"}}>
                  <THead cols={["Codice",branchN(branch),"Descrizione","Azioni"]}sticky/>
                  <tbody>{displayed.map((a,i)=>(
                    <tr key={a.productId||i} style={{borderBottom:`1px solid ${T.border}`}}>
                      <TD mono><span style={{color:T.gold}}>{a.code}</span></TD>
                      <TD mono><span style={{color:T.muted}}>{a.nHK||"—"}</span></TD>
                      <TD>{a.description}</TD>
                      <TD><MiniBtn label="✕ Rimuovi" onClick={()=>{const n=airList.filter((_,j)=>j!==airList.indexOf(a));setAirList(n);LS.set(`ifb_airlist_${branch}`,n);}} color={T.red}/></TD>
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
function Dashboard({costRows, branch, month, navigate}) {
  const [activePanel, setActivePanel] = useState<string|null>(null);

  const calcOk   = costRows.filter((r:any)=>r.cost?.step2Hkd!=null);
  const flagged  = costRows.filter((r:any)=>r.cost?.step2Hkd!=null&&r.prevCost?.step2Hkd!=null&&r.prevCost.step2Hkd>0&&Math.abs((r.cost.step2Hkd-r.prevCost.step2Hkd)/r.prevCost.step2Hkd)>=0.03);
  const air      = costRows.filter((r:any)=>r.isAir);
  const noPrice  = costRows.filter((r:any)=>!r.cost&&!r.isAir&&r.skipReason?.includes("NO PREZZO"));
  const noLog    = costRows.filter((r:any)=>!r.cost&&!r.isAir&&r.skipReason==="NO LOGISTICA");
  const calcZero = costRows.filter((r:any)=>!r.cost&&!r.isAir&&r.skipReason?.includes("CALC=0"));

  const STATS = [
    { id:"ok",      n:calcOk.length,   label:"Costi calcolati",       color:T.green,  rows:calcOk   },
    { id:"flagged", n:flagged.length,  label:"Variazioni ≥3%",        color:T.orange, rows:flagged  },
    { id:"air",     n:air.length,      label:"AIR (esclusi)",          color:T.blue,   rows:air      },
    { id:"noPrice", n:noPrice.length,  label:"Senza prezzo",           color:T.red,    rows:noPrice  },
    { id:"noLog",   n:noLog.length,    label:"No logistica",           color:T.red,    rows:noLog    },
    { id:"calc0",   n:calcZero.length, label:"Calc=0 (UOM/qty)",       color:T.orange, rows:calcZero },
  ];

  const panel = STATS.find(s=>s.id===activePanel);

  function renderPanel() {
    if(!panel||panel.rows.length===0)
      return <div style={{padding:"24px",textAlign:"center",color:T.dim,fontSize:"13px"}}>Nessun articolo in questa categoria.</div>;

    if(activePanel==="ok"||activePanel==="flagged") return (
      <table style={{width:"100%",borderCollapse:"collapse"}}>
        <THead cols={[branchN(branch),"IFB No","Descrizione","Ubicaz.","Step2 HKD","Prec. HKD","Δ%"]}sticky/>
        <tbody>{panel.rows.map((r:any,i:number)=>{
          const pct = r.cost&&r.prevCost&&r.prevCost.step2Hkd>0
            ? (r.cost.step2Hkd-r.prevCost.step2Hkd)/r.prevCost.step2Hkd*100 : null;
          return(
            <tr key={r.id} style={{borderBottom:`1px solid ${T.border}`,background:i%2===0?T.bg:T.surface}}>
              <TD mono><span style={{color:T.muted}}>{r.nHK||"—"}</span></TD>
              <TD mono><span style={{color:T.gold}}>{r.code}</span></TD>
              <TD>{r.description}</TD>
              <TD><Chip label={r.ubicazione||"—"} color={r.ubicazione==="FOR"?T.purple:r.ubicazione==="MTS"?T.blue:T.green}/></TD>
              <TD mono><span style={{color:T.gold,fontWeight:"bold"}}>{r.cost?.step2Hkd?.toFixed(2)||"—"}</span></TD>
              <TD mono><span style={{color:T.muted}}>{r.prevCost?.step2Hkd?.toFixed(2)||"—"}</span></TD>
              <TD>{pct!=null
                ? <span style={{color:Math.abs(pct)>=3?(pct>0?T.red:T.green):T.text,fontWeight:Math.abs(pct)>=3?"bold":"normal"}}>
                    {pct>0?"+":""}{pct.toFixed(1)}%{Math.abs(pct)>=3?" ⚡":""}
                  </span>
                : <span style={{color:T.dim}}>—</span>}
              </TD>
            </tr>
          );
        })}</tbody>
      </table>
    );

    if(activePanel==="air") return (
      <table style={{width:"100%",borderCollapse:"collapse"}}>
        <THead cols={[branchN(branch),"IFB No","Descrizione"]}sticky/>
        <tbody>{panel.rows.map((r:any,i:number)=>(
          <tr key={r.id} style={{borderBottom:`1px solid ${T.border}`,background:i%2===0?T.bg:T.surface}}>
            <TD mono><span style={{color:T.muted}}>{r.nHK||"—"}</span></TD>
            <TD mono><span style={{color:T.gold}}>{r.code}</span></TD>
            <TD>{r.description}</TD>
          </tr>
        ))}</tbody>
      </table>
    );

    // noPrice, noLog, calc0
    return (
      <table style={{width:"100%",borderCollapse:"collapse"}}>
        <THead cols={[branchN(branch),"IFB No","Descrizione","Motivo"]} sticky />
        <tbody>{panel.rows.map((r:any,i:number)=>(
          <tr key={r.id} style={{borderBottom:`1px solid ${T.border}`,background:i%2===0?T.bg:T.surface}}>
            <TD mono><span style={{color:T.muted}}>{r.nHK||"—"}</span></TD>
            <TD mono><span style={{color:T.gold}}>{r.code}</span></TD>
            <TD>{r.description}</TD>
            <TD><span style={{color:T.orange,fontSize:"11px"}}>{r.skipReason}</span></TD>
          </tr>
        ))}</tbody>
      </table>
    );
  }

  return (
    <div>
      <PageHeader title={`Dashboard · ${branch} · ${month}`} sub="Solo articoli INALCA FOOD & BEVERAGE · SEA"/>
      <div style={{display:"flex",gap:"12px",marginBottom:"20px",flexWrap:"wrap"}}>
        {STATS.map(({id,n,label,color,rows})=>{
          const isActive = activePanel===id;
          return(
            <button key={id} onClick={()=>setActivePanel(v=>v===id?null:id)}
              style={{padding:"12px 20px",
                background:isActive?`${color}20`:T.card,
                border:`2px solid ${isActive?color:color+"44"}`,
                borderRadius:"8px",minWidth:"130px",cursor:"pointer",
                fontFamily:"inherit",textAlign:"left",transition:"all 0.15s"}}>
              <div style={{fontSize:"22px",fontWeight:"bold",color}}>{n}</div>
              <div style={{fontSize:"11px",color:T.dim,marginTop:"2px"}}>{label}</div>
              <div style={{fontSize:"9px",color:`${color}88`,marginTop:"4px"}}>
                {isActive?"▲ chiudi":"▼ mostra articoli"}
              </div>
            </button>
          );
        })}
      </div>

      {panel&&(
        <Section title={`${panel.label} · ${panel.rows.length} articoli`} accent={panel.color}>
          <div style={{overflowX:"auto"}}>
            {renderPanel()}
          </div>
        </Section>
      )}
    </div>
  );
}

// ─── LOGISTICS ────────────────────────────────────────────────────────────────

function Logistics({ logistics, setLogistics, products, branch, showToast, bumpImportTs, initFilter, importLogs, setImportLogs, xrefs = [] }) {
  const[search,setSearch]=useState("");
  const[showOnlyMissing,setShowOnlyMissing]=useState(initFilter==="missing");
  const[mapStep,setMapStep]=useState("idle");
  const[logHeaders,setLogHeaders]=useState([]);
  const[logRawRows,setLogRawRows]=useState([]);
  const[colIdx,setColIdx]=useState({});

  const allIFBProducts = products.filter(p=>isIFBVendor(p.vendorName));

  function getLog(productId) { 
    return logistics.find(l=>l.productId===productId && l.branch===branch) || null; 
  }
  
  function getOrDefault(productId) {
    return getLog(productId) || {productId, branch, area:"NORD", ubicazione:"MTO", pltPerContainer:20, hasCert:false, hasAlcTax:false, alcTax:0, convFactor:1, carriage:0};
  }
  
  function update(productId, field, rawVal) {
    const existing = getLog(productId);
    if(existing) {
      showToast(`❌ ${field} non modificabile: dato importato da Work_tab`, T.red);
      return;
    }
    const val = ["ubicazione","area"].includes(field) ? rawVal : 
                ["hasCert","hasAlcTax"].includes(field) ? (rawVal === "true") : 
                (parseFloat(rawVal) || 0);
    const next = [...logistics, {...getOrDefault(productId), [field]: val}];
    setLogistics(next);
    LS.set("ifb_logistics", next);
  }

  function parseLogFile(e) {
    const file = e.target.files?.[0]; 
    if(!file) return;
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const wb = XLSX.read(ev.target.result, {type:"binary"});
        const ws = wb.Sheets[wb.SheetNames[0]];
        const raw = XLSX.utils.sheet_to_json(ws, {header:1, defval:""});
        
        let headerRowIdx = raw.findIndex(r => r.some(c => String(c||"").toLowerCase().includes("ubicazione")));
        if(headerRowIdx < 0) headerRowIdx = 0;
        
        const hdrs = raw[headerRowIdx].map(h => String(h||"").trim());
        const dataRows = raw.slice(headerRowIdx+1).filter(r => r.some(c => c !== ""));
        
        const fi = aliases => hdrs.findIndex(h => aliases.some(a => h.toLowerCase().replace(/[\s_°]/g,"").includes(a.replace(/[\s_°]/g,""))));
        
        const idx = {
          iNHK: fi(["nhk","n hk","gc"]),
          iIFB: fi(["no_(ifb)","noifb","ifb","no_"]),
          iUb: fi(["ubicazione","location","wh"]),
          iArea: fi(["area"]),
          iPlt: fi(["npltxcontainer","pltxcontainer","plt x container","nplt","pltpercontainer","n plt"]),
          iCert: fi(["healthcertificate","health certificate","cert"]),
          iTemp: fi(["rettificata","temperature","temp","trettificata"]),
          iCarriage: fi(["pltcostmedio","plt cost medio","pltcost","carriage"]),
          iAirSea: fi(["air/sea","airsea","air","sea"]),
          iAlcTax: fi(["tassa alcolica","tassaalcolica","alcolica","alctax","alc tax"]),
        };
        setColIdx(idx);
        setLogHeaders(hdrs);
        setLogRawRows(dataRows);
        setMapStep("ready");
        showToast(`File caricato: ${dataRows.length} righe`, T.gold);
      } catch(err) { 
        showToast("Errore: "+err.message, T.red); 
      }
    };
    reader.readAsBinaryString(file);
    e.target.value = "";
  }

  function applyLogFile() {
    const { iNHK, iIFB, iUb, iArea, iPlt, iCert, iTemp, iCarriage, iAirSea, iAlcTax } = colIdx;
    let next = [...logistics];
    let countLog = 0, countAir = 0;
    const currentBranch = branch;
  
    logRawRows.forEach(row => {
      // Ottieni i codici
      const nhkRaw = iNHK >= 0 ? String(row[iNHK] || "").trim() : "";
      const ifbRaw = iIFB >= 0 ? String(row[iIFB] || "").trim() : "";
      if (!nhkRaw && !ifbRaw) return;
  
      // Trova il prodotto
      const prod = findProduct(nhkRaw, products, xrefs) || findProduct(ifbRaw, products, xrefs);
      if (!prod) return;
  
      // Controlla se è AIR (salta)
      // Controlla se è AIR (salta)
      const airSeaRaw = iAirSea >= 0 ? String(row[iAirSea] || "").trim().toUpperCase() : "";
      if (airSeaRaw === "AIR") {
        countAir++;
        return;
      }
  
      // Ubicazione
      const ubRaw = iUb >= 0 ? String(row[iUb] || "").trim().toUpperCase() : "";
      let ubicazione = "MTO";
      if (ubRaw.includes("MTS")) ubicazione = "MTS";
      else if (ubRaw.includes("FOR")) ubicazione = "FOR";
  
      // Area
      const areaRaw = iArea >= 0 ? String(row[iArea] || "").trim().toUpperCase() : "";
      let area = "NORD";
      if (areaRaw.includes("SUD")) area = "SUD";
      else if (areaRaw.includes("CENTRO") || areaRaw.includes("CENTER")) area = "CENTRO";
  
      // Pallet per container
      let plt = 0;
      if (iPlt >= 0) {
        const pltVal = parseFloat(String(row[iPlt] || "0"));
        plt = isNaN(pltVal) ? 0 : pltVal;
      }
  
      // Health Certificate
      const certRaw = iCert >= 0 ? String(row[iCert] || "").trim().toUpperCase() : "";
      const hasCert = ["SI", "YES", "1", "TRUE", "SÌ", "S"].includes(certRaw);
  
      // Temperatura rettificata
      let temperatureOverride = null;
      if (iTemp >= 0) {
        const tempRaw = String(row[iTemp] || "").trim().toUpperCase();
        if (tempRaw === "DRY" || tempRaw === "SECCO") temperatureOverride = "DRY";
        else if (tempRaw === "FRESH" || tempRaw === "FRESCO") temperatureOverride = "FRESH";
        else if (tempRaw === "FROZEN" || tempRaw === "SURGELATO") temperatureOverride = "FROZEN";
      }
  
      // Carriage
      let carriage = 0;
      if (iCarriage >= 0) {
        const carrVal = parseFloat(String(row[iCarriage] || "0"));
        carriage = isNaN(carrVal) ? 0 : carrVal;
      }
  
      // Tassa alcolica
      let alcTax = 0;
      let hasAlcTax = false;
      if (iAlcTax >= 0) {
        const alcVal = parseFloat(String(row[iAlcTax] || "0"));
        alcTax = isNaN(alcVal) ? 0 : alcVal;
        hasAlcTax = alcTax > 0;
      }
  
      const entry = {
        productId: prod.id,
        branch: currentBranch,
        area,
        ubicazione,
        pltPerContainer: plt,
        hasCert,
        hasAlcTax,
        alcTax,
        convFactor: 1,
        carriage,
        temperatureOverride
      };
  
      const existIdx = next.findIndex(l => l.productId === prod.id && l.branch === currentBranch);
      if (existIdx >= 0) {
        next[existIdx] = { ...next[existIdx], ...entry };
      } else {
        next.push(entry);
      }
      countLog++;
    });
  
    setLogistics(next);
    LS.set("ifb_logistics", next);
  
    if (countAir > 0) {
      showToast(`⚠ ${countAir} articoli AIR rilevati — gestiscili da ✈ AIR Transport`, T.orange);
    }
  
    bumpImportTs();
    showToast(`Logistica aggiornata: ${countLog} prodotti per ${currentBranch} ✓`, T.gold);
  
    // Salva snapshot per storico
    const now = Date.now();
const branchEntries = next.filter((l:any) => l.branch === currentBranch);
IDB.set(`ifb_log_data_${now}`, branchEntries);
const log = { id: now, type: "logistics", date: new Date(now).toISOString(), branch: currentBranch, count: countLog };
    const newLogs = [log, ...importLogs];
    setImportLogs(newLogs);
    LS.set("ifb_importlogs", newLogs);
  
    setMapStep("idle");
    setLogHeaders([]);
    setLogRawRows([]);
  }

  const _sq = search.toLowerCase();
  const allProds = allIFBProducts.filter(p => {
    if(!search) return true;
    return p.description?.toLowerCase().includes(_sq) || p.code?.toLowerCase().includes(_sq) || p.nHK?.toLowerCase().includes(_sq);
  });
  const displayed = showOnlyMissing
  ? allProds.filter(p => !getLog(p.id))
  : allProds.filter(p => !!getLog(p.id));
  const missingCount = allIFBProducts.filter(p => !getLog(p.id)).length;
  const withCount = allIFBProducts.filter(p => getLog(p.id) !== null).length;

  return (
    <div>
      <PageHeader title={`Logistica · ${branch}`} sub={`${withCount} con logistica (read-only) · ${missingCount} senza logistica (modificabili) — totale ${allIFBProducts.length} IFB`}/>
      
      <div style={{fontSize:"11px", color:T.muted, marginBottom:"10px", padding:"6px 10px", background:`${T.gold}08`, borderRadius:"6px", border:`1px solid ${T.gold}22`}}>
        🔒 Righe <strong style={{color:T.gold}}>dorate</strong> = importate da Work_tab (sola lettura) &nbsp;·&nbsp;
        🟠 Righe arancioni = senza logistica (modificabili)
      </div>

      {mapStep === "idle" ? (
  <div style={{marginBottom:"16px", display:"flex", gap:"10px", alignItems:"center", flexWrap:"wrap"}}>
    <label style={{display:"inline-block", padding:"8px 16px", background:T.surface, border:`1px solid ${T.border}`, borderRadius:"6px", cursor:"pointer", fontSize:"12px", color:T.text}}>
      📂 Carica Work_tab (08_Work_Tab.xlsx)
      <input type="file" accept=".xlsx,.xls,.csv" onChange={parseLogFile} style={{display:"none"}}/>
    </label>
    
    {/* ✅ AGGIUNGI QUESTO DROPDOWN - Carica da storico */}
    {/* ✅ DROPDOWN CARICA DA STORICO - FUNZIONANTE */}
    {importLogs.filter((l:any) => l.type === "logistics" && l.branch === branch).length > 0 && (
      <select 
      onChange={async e => {
        if (e.target.value) {
          const snap = importLogs.find((l:any) => String(l.id) === e.target.value);
          if (!snap) return;
          if (window.confirm(`Ripristinare logistica del ${new Date(snap.id).toLocaleDateString("it-IT")} (${snap.count} righe)?`)) {
            const branchEntries = await IDB.get(`ifb_log_data_${snap.id}`, null);
            if (!branchEntries?.length) {
              showToast("Snapshot non disponibile — reimporta il file Work_tab.", T.orange);
              return;
            }
            const other = logistics.filter((l:any) => l.branch !== branch);
            const newLog = [...other, ...branchEntries];
            setLogistics(newLog);
            LS.set("ifb_logistics", newLog);
            bumpImportTs();
            showToast(`Logistica ripristinata: ${branchEntries.length} righe ✓`, T.gold);
          }
        }
        e.target.value = "";
      }}
        style={{ ...inputStyle(), width: "auto", fontSize: "12px" }}
        defaultValue=""
      >
        <option value="">📜 Carica da storico ({importLogs.filter((l:any) => l.type === "logistics" && l.branch === branch).length})</option>
        {importLogs.filter((l:any) => l.type === "logistics" && l.branch === branch).map((s: any) => (
          <option key={s.id} value={String(s.id)}>
            {new Date(s.id).toLocaleDateString("it-IT")} · {s.count} righe
          </option>
        ))}
      </select>
    )}
    
    {/* Bottone Svuota dati esistente */}
    <button
      onClick={() => {
        if(window.confirm(`⚠️ ATTENZIONE: Eliminare TUTTI i dati logistici per ${branch}?`)) {
          const newLogistics = logistics.filter((l:any) => l.branch !== branch);
          setLogistics(newLogistics);
          LS.set("ifb_logistics", newLogistics);
          bumpImportTs();
          showToast(`Dati logistici per ${branch} cancellati ✓`, T.red);
        }
      }}
      style={{
        padding:"8px 16px",
        background:"none",
        border:`1px solid ${T.red}44`,
        borderRadius:"6px",
        color:T.red,
        cursor:"pointer",
        fontSize:"12px",
        display:"flex",
        alignItems:"center",
        gap:"6px"
      }}
    >
      🗑 Svuota tutti i dati ({logistics.filter((l:any)=>l.branch===branch).length} righe)
    </button>
    
    <span style={{fontSize:"11px", color:T.muted}}>Colonne: N HK / No_(IFB) / Ubicazione / Area / Cert / Carriage / TASSA ALCOLICA / AIR/SEA</span>
  </div>
) : mapStep === "ready" ? (
  <div style={{background:T.card, border:`1px solid ${T.green}`, borderRadius:"8px", padding:"16px", marginBottom:"16px"}}>
    <div style={{color:T.green, fontWeight:"bold", fontSize:"13px", marginBottom:"8px"}}>✓ File rilevato · {logRawRows.length} righe</div>
    <div style={{fontSize:"12px", color:T.muted, marginBottom:"12px", lineHeight:"1.8"}}>
      Verranno importati per <strong style={{color:T.gold}}>{branch}</strong>: Ubicazione, Area, Plt/Container, Health Certificate, Carriage, Tassa Alcolica
    </div>
    <div style={{display:"flex", gap:"10px"}}>
      <ActionBtn label="← Annulla" onClick={() => setMapStep("idle")}/>
      <ActionBtn label={`✓ Importa logistica per ${branch} (${logRawRows.length} righe)`} onClick={applyLogFile} primary/>
    </div>
  </div>
) : null}

      <div style={{display:"flex", gap:"10px", marginBottom:"12px", alignItems:"center", flexWrap:"wrap"}}>
        <SearchBar value={search} onChange={setSearch} placeholder="🔍 Cerca prodotto IFB…"/>
        <button onClick={() => setShowOnlyMissing(v => !v)}
          style={{padding:"6px 14px", background:showOnlyMissing ? T.orange : T.surface, color:showOnlyMissing ? "#000" : T.orange, border:`1px solid ${T.orange}`, borderRadius:"6px", cursor:"pointer", fontSize:"12px", whiteSpace:"nowrap", fontWeight:showOnlyMissing ? "bold" : "normal"}}>
                    {showOnlyMissing ? `✓ Con logistica (${withCount})` : `⚠ Solo senza logistica (${missingCount})`}
        </button>
        <span style={{fontSize:"11px", color:T.muted}}>
          {showOnlyMissing ? `Mostrando ${displayed.length} senza logistica` : `${displayed.length} con logistica · ${missingCount} mancanti`}
        </span>
      </div>

      {missingCount > 0 && !showOnlyMissing && (
        <div style={{background:`${T.orange}15`, border:`1px solid ${T.orange}44`, borderRadius:"6px", padding:"10px 14px", marginBottom:"14px", fontSize:"12px", color:T.orange}}>
          ⚠ {missingCount} prodotti IFB senza parametri logistici per {branch} → Standard Cost non calcolabile.
        </div>
      )}

      {displayed.length === 0 && showOnlyMissing && (
        <div style={{padding:"32px", textAlign:"center", background:`${T.green}11`, borderRadius:"8px", color:T.green, fontSize:"13px"}}>
          ✅ PERFETTO! Tutti i {allIFBProducts.length} prodotti IFB hanno parametri logistici per {branch}!
        </div>
      )}

      {displayed.length > 0 && (
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%", borderCollapse:"collapse", fontSize:"12px"}}>
            <thead>
              <tr>
              {["IFB No","N HK","Descrizione","Ubicaz.","Area","Plt/Cont","Cert.","Alcol >30°","Carriage","Conv."].map(c=>(
                <th key={c} style={{padding:"7px 12px",background:T.card,color:T.muted,textAlign:"left",borderBottom:`1px solid ${T.border}`,fontSize:"11px",fontWeight:"normal",position:"sticky",top:0,zIndex:10}}>{c}</th>
              ))}
              </tr>
            </thead>
            <tbody>
              {displayed.map((prod, i) => {
                const l = getOrDefault(prod.id);
                const hasEntry = !!getLog(prod.id);
                return (
                  <tr key={prod.id} style={{borderBottom:`1px solid ${T.border}`, background:!hasEntry ? `${T.orange}08` : (i%2===0 ? T.bg : T.surface)}}>
                    <td style={{padding:"7px 12px", fontSize:"12px", fontFamily:"monospace"}}><span style={{color:T.gold}}>{prod.code}</span></td>
                    <td style={{padding:"7px 12px", fontSize:"12px", fontFamily:"monospace"}}><span style={{color:T.muted}}>{prod.nHK||"—"}</span></td>
                    <td style={{padding:"7px 12px", fontSize:"12px"}}>
                      {prod.description}
                      {!hasEntry && <span style={{marginLeft:"6px", fontSize:"9px", color:T.orange, fontWeight:"bold"}}>⚠ MANCANTE</span>}
                    </td>
                    {hasEntry ? (
                      <>
                        <td style={{padding:"7px 12px"}}><Chip label={l.ubicazione||"—"} color={l.ubicazione==="FOR"?T.purple:l.ubicazione==="MTS"?T.blue:T.green}/></td>
                        <td style={{padding:"7px 12px", fontSize:"12px", color:T.muted}}>{l.area||"—"}</td>
                        <td style={{padding:"7px 12px", fontSize:"12px", fontFamily:"monospace", color:T.gold}}>{l.pltPerContainer||"—"}</td>
                        <td style={{padding:"7px 12px", fontSize:"12px", color:T.muted}}>{l.hasCert?"Sì":"No"}</td>
                        <td style={{padding:"7px 12px", fontSize:"12px", color:T.muted}}>{l.hasAlcTax?"Sì":"No"}</td>
                        <td style={{padding:"7px 12px", fontSize:"12px", fontFamily:"monospace", color:T.muted}}>{l.carriage||0}</td>
                        <td style={{padding:"7px 12px", fontSize:"12px", fontFamily:"monospace", color:T.dim}}>{l.convFactor||1}</td>
                      </>
                    ) : (
                      <>
                        <td style={{padding:"7px 12px"}}>
                          <select value={l.ubicazione||"MTO"} onChange={e=>update(prod.id,"ubicazione",e.target.value)}
                            style={{background:T.card,color:T.gold,border:`1px solid ${T.border}`,borderRadius:"4px",padding:"3px 6px",fontSize:"11px",width:"70px"}}>
                            {["MTO","MTS","FOR"].map(v=><option key={v} value={v}>{v}</option>)}
                          </select>
                        </td>
                        <td style={{padding:"7px 12px"}}>
                          <select value={l.area||"NORD"} onChange={e=>update(prod.id,"area",e.target.value)}
                            style={{background:T.card,color:T.text,border:`1px solid ${T.border}`,borderRadius:"4px",padding:"3px 6px",fontSize:"11px",width:"80px"}}>
                            {["NORD","CENTRO","SUD"].map(v=><option key={v} value={v}>{v}</option>)}
                          </select>
                        </td>
                        <td style={{padding:"7px 12px"}}>
                          <input type="number" defaultValue={l.pltPerContainer||20}
                            onBlur={e=>update(prod.id,"pltPerContainer",e.target.value)}
                            style={{width:"55px",background:"transparent",color:T.gold,border:"none",textAlign:"right",fontSize:"12px",borderBottom:`1px solid ${T.border}`}}/>
                        </td>
                        <td style={{padding:"7px 12px"}}>
                          <select value={String(l.hasCert||false)} onChange={e=>update(prod.id,"hasCert",e.target.value)}
                            style={{background:T.card,color:T.text,border:`1px solid ${T.border}`,borderRadius:"4px",padding:"3px 6px",fontSize:"11px",width:"60px"}}>
                            <option value="false">No</option><option value="true">Sì</option>
                          </select>
                        </td>
                        <td style={{padding:"7px 12px"}}>
                          <select value={String(l.hasAlcTax||false)} onChange={e=>update(prod.id,"hasAlcTax",e.target.value)}
                            style={{background:T.card,color:T.text,border:`1px solid ${T.border}`,borderRadius:"4px",padding:"3px 6px",fontSize:"11px",width:"60px"}}>
                            <option value="false">No</option><option value="true">Sì</option>
                          </select>
                        </td>
                        <td style={{padding:"7px 12px"}}>
                          <input type="number" defaultValue={l.carriage||0}
                            onBlur={e=>update(prod.id,"carriage",e.target.value)}
                            style={{width:"55px",background:"transparent",color:T.gold,border:"none",textAlign:"right",fontSize:"12px",borderBottom:`1px solid ${T.border}`}}/>
                        </td>
                        <td style={{padding:"7px 12px"}}>
                          <input type="number" defaultValue={l.convFactor||1} step="0.01"
                            onBlur={e=>update(prod.id,"convFactor",e.target.value)}
                            style={{width:"50px",background:"transparent",color:T.muted,border:"none",textAlign:"right",fontSize:"11px",borderBottom:`1px solid ${T.border}`}}/>
                        </td>
                      </>
                    )}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

// ─── PRICES (con import integrato e storico) ─────────────────────────────────
function Prices({ prices, setPrices, products, branch, month, setPrices: setPricesParent, salesRows = [], xrefs = [], 
  importLogs, setImportLogs, snapshots, setSnapshots, showToast, bumpImportTs }) {
const [search, setSearch] = useState("");
const [invoiceOnly, setInvoiceOnly] = useState(false);
const [importStep, setImportStep] = useState<"idle"|"map"|"preview"|"done">("idle");
const [headers, setHeaders] = useState<string[]>([]);
const [rawRows, setRawRows] = useState<any[]>([]);
const [mapping, setMapping] = useState<any>({});
const [preview, setPreview] = useState<any[]>([]);
const [fileName, setFileName] = useState("");
const [importMonth, setImportMonth] = useState(month);
const [doneInfo, setDoneInfo] = useState<any>(null);

// Storico import listini
const priceSnaps = snapshots.filter((s: any) => s.type === "prices" && s.branch === branch);

// Funzione per verificare se un codice è valido (NON Power BI)
function isValidCode(code: string) {
if (!code) return false;
const str = String(code).trim();
if (/^P_/i.test(str)) return false;
if (/^\d{7,}$/.test(str.replace(/[^0-9]/g, ""))) return false;
if (str.includes("P_BC_")) return false;
return true;
}

function parseFile(file: File) {
setFileName(file.name);
const reader = new FileReader();
reader.onload = e => {
try {
const wb = XLSX.read(e.target.result, { type: "binary" });
const ws = wb.Sheets[wb.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
if (data.length < 2) { showToast("File vuoto", T.red); return; }

const hdrs = data[0].map((h: any) => String(h || "").trim());
const rows = data.slice(1).filter((r: any[]) => r.some(c => c !== ""));
setHeaders(hdrs);
setRawRows(rows);

// Auto-mapping dei campi
const am: any = {};
const codeAliases = ["no_", "no.", "no", "item no.", "codice", "code", "n hk", "ifb item", "ifb no", "ifb n"];
for (const h of hdrs) {
const hl = h.toLowerCase().trim();
if (codeAliases.some(a => hl === a || hl.includes(a))) {
am["code"] = h;
break;
}
}

const descAliases = ["description", "descrizione", "desc", "item description"];
for (const h of hdrs) {
const hl = h.toLowerCase().trim();
if (descAliases.some(a => hl === a || hl.includes(a))) {
am["description"] = h;
break;
}
}

const priceFields = ["mtsPrice", "fcaPrice", "fcaDiscount", "fcaDiscounted", "dapPrice", "dapDiscount", "dapDiscounted", "dapFinalDirect"];
priceFields.forEach(field => {
const aliases = PRICE_FIELD_ALIASES[field] || [];
for (const h of hdrs) {
const hl = h.toLowerCase().trim();
if (aliases.some(a => hl === a || (a.length > 3 && hl.includes(a)))) {
am[field] = h;
break;
}
}
});

setMapping(am);
setImportStep("map");
} catch (err: any) {
showToast("Errore: " + err.message, T.red);
}
};
reader.readAsBinaryString(file);
}

function buildPreview() {
const get = (row: any, field: string) => {
const col = mapping[field];
if (!col) return null;
const i = headers.indexOf(col);
return i >= 0 ? row[i] : null;
};

const mapped = rawRows.map((row, idx) => {
const rawCode = String(get(row, "code") || "").trim();
const rawDescription = String(get(row, "description") || get(row, "code") || "").trim();

if (!rawCode) return null;
if (!isValidCode(rawCode)) return null;

const prod = findProduct(rawCode, products, xrefs);

const mtsPrice = parseFloat(get(row, "mtsPrice")) || 0;
const fcaPrice = parseFloat(get(row, "fcaPrice")) || 0;
const fcaDiscount = parseFloat(get(row, "fcaDiscount")) || 0;
const fcaDiscounted = parseFloat(get(row, "fcaDiscounted")) || (fcaPrice - (fcaDiscount * fcaPrice / 100)) || 0;
const dapPrice = parseFloat(get(row, "dapPrice")) || 0;
const dapDiscount = parseFloat(get(row, "dapDiscount")) || 0;
const dapDiscounted = parseFloat(get(row, "dapDiscounted")) || (dapPrice - (dapDiscount * dapPrice / 100)) || 0;
const dapFinalDirect = parseFloat(get(row, "dapFinalDirect")) || 0;

let dapFinal = 0;
let dapNote = "";
if (dapFinalDirect !== 0) {
dapFinal = dapFinalDirect;
dapNote = "da file";
} else if (prod) {
dapFinal = dapDiscounted || 0;
dapNote = dapDiscounted ? "da DAP Disc." : "";
}

const existing = prod ? prices.find(p => p.productId === prod.id && p.branch === branch && p.month === importMonth) : null;

return {
_idx: idx,
rawCode,
ifbNo_from_file: rawCode,
description_from_file: rawDescription,
productId: prod?.id || null,
nHK_from_anag: prod?.nHK || "",
ifbNo_from_anag: prod?.code || "",
description_from_anag: prod?.description || "",
dapFinal: roundN(dapFinal),
mtsPrice: roundN(mtsPrice),
fcaDiscounted: roundN(fcaDiscounted),
dapPrice: roundN(dapPrice),
fcaPrice: roundN(fcaPrice),
dapNote,
_hasProduct: !!prod,
_existing: !!existing
};
}).filter(Boolean);

setPreview(mapped);
setImportStep("preview");
}

function executeImport() {
const snId = Date.now();
const updated = [...prices];
const diffs = [];
let count = 0, newCount = 0, changed = 0;

preview.forEach(r => {
if (!r._hasProduct) return;

const idx = updated.findIndex(p => p.productId === r.productId && p.branch === branch && p.month === importMonth);
const entry = {
productId: r.productId,
branch,
month: importMonth,
dapFinal: r.dapFinal,
mtsPrice: r.mtsPrice,
fcaDiscounted: r.fcaDiscounted,
dapPrice: r.dapPrice,
fcaPrice: r.fcaPrice
};
const prev = idx >= 0 ? updated[idx] : null;
const diffFields = [];

["dapFinal", "mtsPrice", "fcaDiscounted", "dapPrice", "fcaPrice"].forEach(f => {
const oldR = roundN(prev?.[f] || 0);
const newR = roundN(entry[f] || 0);
if (Math.abs(oldR - newR) >= 0.005) {
diffFields.push({ field: f, old: oldR, new: newR, delta: oldR > 0 ? ((newR - oldR) / oldR * 100) : null });
}
});

if (!prev) newCount++;
else if (diffFields.length > 0) changed++;

if (diffFields.length > 0 || !prev) {
diffs.push({
productId: r.productId,
nHK: r.nHK_from_anag,
ifbNo: r.ifbNo_from_anag,
description: r.description_from_anag,
isNew: !prev,
fields: diffFields
});
}

if (idx >= 0) updated[idx] = entry;
else updated.push(entry);
count++;
});

setPrices(updated);
LS.set("ifb_prices", updated);

const log = {
id: snId,
type: "prices",
fileName,
branch,
month: importMonth,
date: new Date(snId).toISOString(),
count,
newCount,
updateCount: changed,
diffs
};

const newLogs = [log, ...importLogs];
setImportLogs(newLogs);
LS.set("ifb_importlogs", newLogs);

const newSnaps = [log, ...snapshots].slice(0, 50);
setSnapshots(newSnaps);
LS.set("ifb_snapshots", newSnaps);

setDoneInfo({ count, newCount, changed, unchanged: count - newCount - changed });
bumpImportTs();
setImportStep("done");
}

function loadFromSnapshot(snap: any) {
  // Leggi direttamente dallo snapshot (senza andare in localStorage separato)
  const snapshotProducts = snap.products || [];
  
  if (snapshotProducts.length === 0) {
    showToast(`Nessun prodotto trovato nello snapshot`, T.orange);
    return;
  }

  if (window.confirm(`Caricare l'anagrafica del ${new Date(snap.id).toLocaleDateString("it-IT")} (${snapshotProducts.length} articoli)? Sostituirà i dati attuali.`)) {
    setProducts(snapshotProducts);
    LS.set(`ifb_products_${branch}`, snapshotProducts);
    showToast(`Anagrafica ripristinata da snapshot del ${new Date(snap.id).toLocaleDateString("it-IT")}`, T.gold);
    setSearch("");
    setOnlyIFB(true);
  }
}

const resetImport = () => {
setImportStep("idle");
setRawRows([]);
setHeaders([]);
setFileName("");
setMapping({});
setPreview([]);
setDoneInfo(null);
};

const invoiceProductIds = useMemo(() => {
const s = new Set();
(salesRows || []).forEach(r => {
const prod = findProduct(r.itemCode, products, xrefs);
if (prod) s.add(prod.id);
});
return s;
}, [salesRows, products, xrefs]);

const filtered = prices.filter(p => {
if (p.branch !== branch || p.month !== month) return false;
if (/^P_BC_/i.test(p.productId)) return false;
if (invoiceOnly && !invoiceProductIds.has(p.productId)) return false;
return true;
});

const displayed = filtered.filter(p => {
const prod = products.find(pr => pr.id === p.productId);
if (!search) return true;
const q = search.toLowerCase();
return prod?.description?.toLowerCase().includes(q) ||
prod?.code?.toLowerCase().includes(q) ||
prod?.nHK?.toLowerCase().includes(q) ||
String(p.productId).toLowerCase().includes(q);
});

const COLS = ["fcaPrice", "fcaDiscounted", "dapPrice", "mtsPrice", "dapFinal"];
const LABELS = ["FCA Price", "FCA Disc.", "DAP Price", "MTS Price", "DAP Final"];

// Schermata import completato
if (importStep === "done" && doneInfo) {
return (
<div>
<PageHeader title="✓ Import Listini completato" sub={fileName} />
<div style={{ padding: "20px", background: `${T.green}11`, border: `1px solid ${T.green}33`, borderRadius: "8px", marginBottom: "16px", fontSize: "13px", color: T.muted, lineHeight: "2" }}>
Mese: <strong style={{ color: T.gold }}>{importMonth}</strong> · Filiale: <strong style={{ color: T.text }}>{branch}</strong><br />
Prezzi totali: <strong style={{ color: T.text }}>{doneInfo.count}</strong> &nbsp;·&nbsp;
<span style={{ color: T.green }}>🆕 {doneInfo.newCount} nuovi</span> &nbsp;·&nbsp;
<span style={{ color: T.orange }}>✏️ {doneInfo.changed} modificati</span> &nbsp;·&nbsp;
<span style={{ color: T.dim }}>{doneInfo.unchanged} invariati</span>
</div>
<ActionBtn label="← Torna ai listini" onClick={resetImport} />
</div>
);
}

// Schermata vuota (nessun prezzo)
if (filtered.length === 0 && !invoiceOnly && importStep === "idle") {
return (
<div>
<PageHeader title={`Listini · ${branch} · ${month}`} sub="Nessun prezzo caricato" />
<div style={{ padding: "32px", textAlign: "center", color: T.muted, fontSize: "13px" }}>
Nessun prezzo per {branch} · {month}.
</div>
<div style={{ marginTop: "16px" }}>
<label style={{ display: "inline-block", padding: "10px 20px", background: T.gold, color: "#000", borderRadius: "6px", cursor: "pointer", fontWeight: "bold" }}>
📂 Carica listini (PBI / CURRENT PRICELIST)
<input type="file" accept=".xlsx,.xls,.csv" onChange={e => { const f = e.target.files?.[0]; if (f) parseFile(f); e.target.value = ""; }} style={{ display: "none" }} />
</label>
</div>
</div>
);
}

return (
<div>
<PageHeader title={`Listini · ${branch} · ${month}`} sub={`${filtered.length} prezzi caricati`} />

{/* Toolbar import */}
<div style={{ display: "flex", gap: "10px", marginBottom: "14px", alignItems: "center", flexWrap: "wrap" }}>
<label style={{ display: "inline-block", padding: "6px 14px", background: T.gold, color: "#000", borderRadius: "6px", cursor: "pointer", fontWeight: "bold", fontSize: "12px" }}>
📂 Carica listini
<input type="file" accept=".xlsx,.xls,.csv" onChange={e => { const f = e.target.files?.[0]; if (f) parseFile(f); e.target.value = ""; }} style={{ display: "none" }} />
</label>

{priceSnaps.length > 0 && (
<select onChange={e => { if (e.target.value) loadFromSnapshot(JSON.parse(e.target.value)); e.target.value = ""; }} style={{ ...inputStyle(), width: "auto", fontSize: "12px" }} defaultValue="">
<option value="">📜 Carica da storico ({priceSnaps.length})</option>
{priceSnaps.map((s: any) => (
<option key={s.id} value={JSON.stringify(s)}>
{new Date(s.id).toLocaleDateString("it-IT")} · {s.month} · {s.count} prezzi
</option>
))}
</select>
)}

<div style={{ flex: 1 }} />

<SearchBar value={search} onChange={setSearch} placeholder="🔍 Cerca prodotto…" style={{ marginBottom: 0, maxWidth: "250px" }} />

<button onClick={() => setInvoiceOnly(v => !v)} style={{ padding: "5px 12px", background: invoiceOnly ? `${T.gold}20` : T.surface, color: invoiceOnly ? T.gold : T.muted, border: `1px solid ${invoiceOnly ? T.gold : T.border}`, borderRadius: "6px", cursor: "pointer", fontSize: "11px", whiteSpace: "nowrap" }}>
{invoiceOnly ? `✓ Solo fatturati (${displayed.length})` : `📋 Solo Sales Invoice (${invoiceProductIds.size} prod.)`}
</button>

{setPricesParent && (
<button onClick={() => { if (window.confirm(`Eliminare tutti i prezzi ${branch}/${month}?`)) setPricesParent(prices.filter(p => !(p.branch === branch && p.month === month))); }} style={{ padding: "5px 12px", background: "none", border: `1px solid ${T.red}44`, borderRadius: "6px", color: T.red, cursor: "pointer", fontSize: "11px" }}>
✕ Svuota {branch}/{month}
</button>
)}
</div>

{/* Step di import - Mappa */}
{importStep === "map" && (
<div style={{ background: T.card, border: `1px solid ${T.gold}`, borderRadius: "8px", padding: "16px", marginBottom: "16px" }}>
<div style={{ color: T.gold, fontWeight: "bold", marginBottom: "12px" }}>Mappatura colonne · {fileName}</div>
<div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: "12px", marginBottom: "16px" }}>
<div>
<label style={{ fontSize: "11px", color: T.gold }}>📌 Codice *</label>
<select value={mapping["code"] || ""} onChange={e => setMapping((m: any) => ({ ...m, code: e.target.value }))} style={{ ...inputStyle(), fontSize: "12px" }}>
<option value="">— seleziona —</option>
{headers.map(h => <option key={h} value={h}>{h}</option>)}
</select>
</div>
<div>
<label style={{ fontSize: "11px", color: T.muted }}>📝 Descrizione</label>
<select value={mapping["description"] || ""} onChange={e => setMapping((m: any) => ({ ...m, description: e.target.value }))} style={{ ...inputStyle(), fontSize: "12px" }}>
<option value="">— non mappato —</option>
{headers.map(h => <option key={h} value={h}>{h}</option>)}
</select>
</div>
<div>
<label style={{ fontSize: "11px", color: T.muted }}>📅 Mese listino</label>
<input type="month" value={importMonth} onChange={e => setImportMonth(e.target.value)} style={{ ...inputStyle(), fontSize: "12px" }} />
</div>
</div>
<div style={{ display: "flex", gap: "10px" }}>
<ActionBtn label="Annulla" onClick={resetImport} />
<ActionBtn label="Preview →" onClick={buildPreview} primary disabled={!mapping["code"]} />
</div>
</div>
)}

{/* Step di import - Preview */}
{importStep === "preview" && (
<div style={{ background: T.card, border: `1px solid ${T.green}`, borderRadius: "8px", padding: "16px", marginBottom: "16px" }}>
<div style={{ color: T.green, fontWeight: "bold", marginBottom: "12px" }}>Preview import · {preview.length} righe valide</div>
<div style={{ display: "flex", gap: "12px", marginBottom: "16px", flexWrap: "wrap" }}>
{[
[preview.filter(r => r._hasProduct).length, "✅ Trovati in anagrafica", T.green],
[preview.filter(r => !r._hasProduct).length, "❌ NON trovati", T.red],
[preview.filter(r => r._existing).length, "✏️ Aggiornamenti", T.orange],
[preview.filter(r => !r._existing && r._hasProduct).length, "🆕 Nuovi", T.gold]
].map(([n, l, c]) => (
<div key={l as string} style={{ padding: "8px 12px", background: T.surface, border: `1px solid ${c}44`, borderRadius: "6px" }}>
<div style={{ fontSize: "18px", fontWeight: "bold", color: c as string }}>{n as number}</div>
<div style={{ fontSize: "10px", color: T.dim }}>{l as string}</div>
</div>
))}
</div>
<div style={{ maxHeight: "200px", overflow: "auto", marginBottom: "12px", fontSize: "11px" }}>
<table style={{ width: "100%", borderCollapse: "collapse" }}>
<thead><tr><th>Codice</th><th>Descrizione</th><th>Match</th><th>DAP Final</th><th>Stato</th></tr></thead>
<tbody>{preview.slice(0, 20).map(r => (
<tr key={r._idx} style={{ borderBottom: `1px solid ${T.border}` }}>
  <td style={{ fontFamily: "monospace", color: T.gold }}>{r.ifbNo_from_file}</td>
  <td>{r.description_from_file}</td>
  <td>{r._hasProduct ? <span style={{ color: T.green }}>✓ {r.ifbNo_from_anag}</span> : <span style={{ color: T.red }}>✗</span>}</td>
  <td style={{ fontFamily: "monospace" }}>{r.dapFinal > 0 ? `€ ${r.dapFinal.toFixed(2)}` : "—"}</td>
  <td>{r._hasProduct ? (r._existing ? <span style={{ color: T.orange }}>aggiornamento</span> : <span style={{ color: T.green }}>nuovo</span>) : <span style={{ color: T.red }}>ignorato</span>}</td>
</tr>
))}</tbody>
</table>
</div>
<div style={{ display: "flex", gap: "10px" }}>
<ActionBtn label="← Indietro" onClick={() => setImportStep("map")} />
<ActionBtn label={`✓ Importa ${preview.filter(r => r._hasProduct).length} prezzi per ${importMonth}`} onClick={executeImport} primary />
</div>
</div>
)}

{/* Tabella listini */}
<Section title={`${displayed.length} prezzi${invoiceOnly ? " (solo Sales Invoice)" : ""}`}>
<div style={{ overflowX: "auto" }}>
<table style={{ width: "100%", borderCollapse: "collapse" }}>
<THead cols={[branchN(branch),"IFB No","Descrizione","FCA Price","FCA Disc.","DAP Price","MTS Price","DAP Final"]} sticky />
<tbody>
{displayed.slice(0, 300).map((p, i) => {
const prod = products.find(pr => pr.id === p.productId);
const inInvoice = invoiceProductIds.has(p.productId);
return (
  <tr key={p.productId} style={{ borderBottom: `1px solid ${T.border}`, background: i % 2 === 0 ? T.bg : T.surface }}>
    <TD mono><span style={{ color: T.muted }}>{prod?.nHK || "—"}</span></TD>
    <TD mono>
      <span style={{ color: T.gold }}>{prod?.code || p.productId}</span>
      {inInvoice && <span style={{ marginLeft: "5px", fontSize: "9px", color: T.blue }}>📋</span>}
    </TD>
    <TD>{prod?.description || <span style={{ color: T.orange, fontSize: "11px" }}>⚠ {p.productId}</span>}</TD>
    {COLS.map(f => (
      <TD key={f} mono>
        <span style={{ color: (p[f] || 0) > 0 ? T.text : T.dim }}>
          {(p[f] || 0) > 0 ? `€ ${roundN(p[f]).toFixed(2)}` : "—"}
        </span>
      </TD>
    ))}
  </tr>
);
})}
</tbody>
</table>
</div>
{displayed.length > 300 && <div style={{ padding: "12px", textAlign: "center", color: T.muted, fontSize: "11px" }}>Mostrati 300/{displayed.length}</div>}
</Section>
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
            <THead cols={["Mese","Tasso"]} sticky />
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

function CostTable({costRows,branch,month,logistics,lastImportTs,lastCalcTs,setLastCalcTs,
  setCostHistory,initFilter,salesRows=[],products=[],xrefs=[]}: any) {

    const[search,setSearch]     = useState("");
    const[showDetail,setShowDetail] = useState<string|null>(null);
    const[invoiceOnly,setInvoiceOnly] = useState(false);
      
    // AGGIUNGI QUESTI STATI PER I FILTRI MULTIPLI
    const [filterFlags, setFilterFlags] = useState<Record<string,false|"include"|"exclude">>({
      flagged: false,
      air: false,
      noPrice: false,
      noLog: false,
      calcZero: false,
      keepOld: false,
      costCalculated: false
    });
    const cycleFilter = (key:string) => setFilterFlags(f=>({...f,[key]: f[key]===false?"include":f[key]==="include"?"exclude":false}));
  const needsRecalc = lastImportTs > lastCalcTs;
  
  // AGGIUNGI QUESTI REF PER LO SCROLL
  const topScrollRef = useRef<HTMLDivElement>(null);
  const tableScrollRef = useRef<HTMLDivElement>(null);

  // AGGIUNGI QUESTO useEffect PER SINCRONIZZARE LO SCROLL
  useEffect(() => {
    const topScroll = topScrollRef.current;
    const tableScroll = tableScrollRef.current;
    if (!topScroll || !tableScroll) return;
    
    const handleTopScroll = () => { if (tableScroll) tableScroll.scrollLeft = topScroll.scrollLeft; };
    const handleTableScroll = () => { if (topScroll) topScroll.scrollLeft = tableScroll.scrollLeft; };
    
    topScroll.addEventListener('scroll', handleTopScroll);
    tableScroll.addEventListener('scroll', handleTableScroll);
    
    return () => {
      topScroll.removeEventListener('scroll', handleTopScroll);
      tableScroll.removeEventListener('scroll', handleTableScroll);
    };
  }, []);

  const sixMonthsAgo = new Date();
  sixMonthsAgo.setMonth(sixMonthsAgo.getMonth()-6);

  // ── data ultima fattura per prodotto ──
  const lastOrderDate = useMemo(()=>{
    const map: Record<string,Date> = {};
    (salesRows||[]).forEach((row:any)=>{
      const prod = findProduct(row.itemCode, products, xrefs);
      if(!prod) return;
      const d = row.date ? new Date(row.date) : null;
      if(!d||isNaN(d.getTime())) return;
      if(!map[prod.id]||d>map[prod.id]) map[prod.id]=d;
    });
    return map;
  },[salesRows,products,xrefs]);

  // set di productId presenti in Sales Invoice
  const invoiceIds = useMemo(()=>{
    const s = new Set<string>();
    (salesRows||[]).forEach((row:any)=>{
      const prod = findProduct(row.itemCode, products, xrefs);
      if(prod) s.add(prod.id);
    });
    return s;
  },[salesRows,products,xrefs]);

  function saveSnapshot(){
    const ts=Date.now();
    const snap={ts,date:new Date(ts).toISOString(),branch,month,
      rows:costRows.map((r:any)=>({id:r.id,code:r.code,nHK:r.nHK,description:r.description,
        cost:r.cost?.step2Hkd??null,costEur:r.cost?.step2Eur??null,skipReason:r.skipReason||null}))};
    setCostHistory((prev:any)=>{const n=[snap,...(prev||[])].slice(0,60);LS.set("ifb_costhistory",n);return n;});
    setLastCalcTs(ts);LS.set("ifb_last_calc_ts",ts);
  }

  let filtered: any[] = costRows.filter((r:any)=>
  !search||r.description?.toLowerCase().includes(search.toLowerCase())||
  r.code?.includes(search)||r.nHK?.includes(search));

// APPLICA FILTRI MULTIPLI (include = mostra solo questi; exclude = nascondi questi)
  const applyFlag = (flag:false|"include"|"exclude", test:(r:any)=>boolean) => {
    if(flag==="include") filtered=filtered.filter(test);
    else if(flag==="exclude") filtered=filtered.filter(r=>!test(r));
  };
  applyFlag(filterFlags.costCalculated, r=> r.cost?.step2Hkd!=null);
  applyFlag(filterFlags.flagged,        r=> r.flagged===true);
  applyFlag(filterFlags.air,            r=> r.isAir===true);
  applyFlag(filterFlags.noPrice,        r=> !r.cost&&!r.isAir&&!!r.skipReason?.includes("NO PREZZO"));
  applyFlag(filterFlags.noLog,          r=> !r.cost&&!r.isAir&&r.skipReason==="NO LOGISTICA");
  applyFlag(filterFlags.calcZero,       r=> !r.cost&&!r.isAir&&!!r.skipReason?.includes("CALC=0"));
  applyFlag(filterFlags.keepOld,        r=> { const d=lastOrderDate[r.id]; return !!(d&&d<sixMonthsAgo); });
if (invoiceOnly) {
  filtered = filtered.filter((r:any) => invoiceIds.has(r.id));
}

filtered = filtered.filter((r: any) => r.priceInput !== 0 && r.priceInput != null);

if(initFilter==="flagged") filtered=filtered.filter((r:any)=>r.flagged===true);
else if(initFilter==="errors") filtered=filtered.filter((r:any)=>!r.cost&&!r.isAir&&r.skipReason?.includes("CALC=0"));

  const calc    = filtered.filter((r:any)=>r.cost?.step2Hkd!=null);
  const noPrice = filtered.filter((r:any)=>!r.cost&&!r.isAir&&r.skipReason?.includes("NO PREZZO"));
  const noLog   = filtered.filter((r:any)=>!r.cost&&!r.isAir&&r.skipReason==="NO LOGISTICA");

  // ── stile celle ──────────────────────────────────────────────────────────────
  const stickyTop0: React.CSSProperties = {position:"sticky",top:0,zIndex:12};
  const stickyTop22: React.CSSProperties = {position:"sticky",top:22,zIndex:12};

  const TH = ({children,accent,w,align="right",sticky=false}:any) => (
    <th style={{
      padding:"4px 6px",background:accent?`color-mix(in srgb,${accent} 15%,${T.card})`:`${T.card}`,
      color:accent||T.muted,textAlign:align as any,
      borderBottom:`1px solid ${accent?accent+"55":T.border}`,
      fontSize:"10px",fontWeight:"normal",whiteSpace:"nowrap",
      minWidth:w||undefined,
      ...(sticky?{position:"sticky",left:0,zIndex:13}:{}),
    }}>{children}</th>
  );

  const GH = ({children,span,accent}:any) => (
    <th colSpan={span} style={{
      padding:"3px 6px",background:accent?`color-mix(in srgb,${accent} 10%,${T.bg})`:`${T.bg}`,
      color:accent||T.dim,textAlign:"center",
      borderBottom:`1px solid ${accent?accent+"44":T.border}`,
      fontSize:"9px",letterSpacing:"0.08em",textTransform:"uppercase",fontWeight:"bold",
      borderRight:`1px solid ${T.border}22`,whiteSpace:"nowrap",
    }}>{children}</th>
  );

  const cell = (color?:string,bold?:boolean,minW?:number): React.CSSProperties => ({
    padding:"4px 6px",borderBottom:`1px solid ${T.border}`,fontSize:"10px",
    textAlign:"right",fontFamily:"monospace",verticalAlign:"middle",
    color:color||T.muted,fontWeight:bold?"bold":"normal",
    minWidth:minW?`${minW}px`:undefined,whiteSpace:"nowrap",
  });
  const cellL = (sticky?:boolean): React.CSSProperties => ({
    padding:"4px 6px",borderBottom:`1px solid ${T.border}`,fontSize:"10px",
    textAlign:"left",verticalAlign:"middle",whiteSpace:"nowrap",
    ...(sticky?{position:"sticky",left:0,zIndex:5,background:T.surface}:{}),
  });

  const f4=(v:number|undefined)=>v!=null&&v!==0?v.toFixed(4):"—";
  const f2=(v:number|undefined)=>v!=null&&v!==0?v.toFixed(2):"—";

  return(
    <div style={{paddingRight:"20px"}}>
      {/* ── toolbar ── */}
      <div style={{display:"flex",alignItems:"center",gap:"12px",marginBottom:"12px",flexWrap:"wrap"}}>
        <PageHeader title={`Standard Cost · ${branch} · ${month}`}
          sub={`${calc.length} calcolati · INALCA F&B · SEA`}/>
                <button onClick={saveSnapshot} disabled={!needsRecalc}
          style={{padding:"7px 16px",background:needsRecalc?T.gold:"#333",
            color:needsRecalc?"#000":T.muted,border:"none",borderRadius:"6px",
            fontWeight:"bold",cursor:needsRecalc?"pointer":"not-allowed",fontSize:"12px",marginTop:"-8px"}}>
          {needsRecalc?"⟳ Ricalcola & Salva":"✓ Aggiornato"}
        </button>
        <button onClick={()=>exportXLSX(
          filtered.filter((r:any)=>r.cost).map((r:any)=>({
            [branchN(branch)]:r.nHK||"","IFB No":r.code||"","Descrizione":r.description||"",
            "UOM":r.uom||"","Ubicazione":r.ubicazione||"",
            "Temp.":r.temperature||"","Temp. Rettif.":r.temperatureOverride||"",
            "Prezzo EUR":roundN(r.cost?.priceEur),"FOB":roundN(r.cost?.fob),
            "LIC":roundN(r.cost?.lic),"VGM":roundN(r.cost?.vgm),
            "HC":roundN(r.cost?.hc),"Pallet":roundN(r.cost?.plt),
            "Alc.Tax":roundN(r.cost?.alc),"Step1 EUR":roundN(r.cost?.step1Eur),
            "Step1 HKD":roundN(r.cost?.step1Hkd),"WH EUR":roundN(r.cost?.wh),
            "Step2 EUR":roundN(r.cost?.step2Eur,4),"Step2 HKD":roundN(r.cost?.step2Hkd),
            "Δ%":r.delta!=null?roundN(r.delta,1):"",
          })),
          "Standard Cost",`SC_${branch}_${month}.xlsx`
        )}
          style={{padding:"7px 14px",background:`${T.green}20`,border:`1px solid ${T.green}44`,
            borderRadius:"6px",color:T.green,cursor:"pointer",fontSize:"12px",marginTop:"-8px"}}>
          ⬇ Export Excel
        </button>
      </div>

      {(noLog.length>0||noPrice.length>0)&&(
        <div style={{background:`${T.orange}15`,border:`1px solid ${T.orange}44`,borderRadius:"6px",
          padding:"8px 12px",marginBottom:"10px",fontSize:"11px",color:T.orange}}>
          {noLog.length>0&&<span>⚠ {noLog.length} senza logistica &nbsp;·&nbsp;</span>}
          {noPrice.length>0&&<span>⚠ {noPrice.length} senza prezzo</span>}
        </div>
      )}

      <div style={{display:"flex",gap:"8px",marginBottom:"10px",alignItems:"center",flexWrap:"wrap"}}>
        <SearchBar value={search} onChange={setSearch} placeholder="🔍 Cerca articolo…"/>

        {/* filtro Sales Invoice */}
        <button onClick={()=>setInvoiceOnly((v:boolean)=>!v)}
          style={{padding:"5px 12px",
            background:invoiceOnly?`${T.gold}20`:T.surface,
            color:invoiceOnly?T.gold:T.muted,
            border:`1px solid ${invoiceOnly?T.gold:T.border}`,
            borderRadius:"6px",cursor:"pointer",fontSize:"11px",whiteSpace:"nowrap"}}>
          {invoiceOnly
            ? `📋 Fatturati (${filtered.length})`
            : `📋 Solo Sales Invoice (${invoiceIds.size} prod.)`}
        </button>

        {invoiceOnly&&(
  <span style={{fontSize:"10px",color:T.muted}}>
    ⚠ KEEP OLD = ultimo ordine &gt;6 mesi fa
  </span>
)}
      </div>


{/* FILTRI MULTIPLI */}
<div style={{display:"flex",gap:"8px",marginBottom:"10px",flexWrap:"wrap",alignItems:"flex-start",borderTop:`1px solid ${T.border}`,paddingTop:"10px"}}>
  <span style={{fontSize:"11px",color:T.muted,paddingTop:"6px"}}>🔍 Filtri:</span>
  {([
    {key:"costCalculated", label:"✅ Costi calcolati", col:T.gold},
    {key:"flagged",        label:"Variazioni ≥3%",    col:T.orange},
    ...(branch!=="CAN" ? [{key:"air", label:"✈ AIR", col:T.blue}] : []),
    {key:"noPrice",        label:"❌ Senza prezzo",   col:T.red},
    {key:"noLog",          label:"⚠ No logistica",    col:T.orange},
    {key:"calcZero",       label:"⚡ Calc=0",          col:T.purple},
    {key:"keepOld",        label:"⏰ KEEP OLD",        col:T.green},
  ] as Array<{key:string,label:string,col:string}>).map(({key,label,col})=>{
    const v = filterFlags[key];
    const isInclude = v==="include";
    const isExclude = v==="exclude";
    const isActive  = isInclude||isExclude;
    return(
      <div key={key} style={{display:"flex",flexDirection:"column",alignItems:"center",gap:"2px"}}>
        <button
          onClick={()=>cycleFilter(key)}
          style={{
            padding:"4px 10px",
            background: isInclude?col : isExclude?T.red : T.surface,
            color:       isInclude?"#000" : isExclude?"#fff" : col,
            border:`1px solid ${isActive?col:col+"66"}`,
            borderRadius:"6px",
            cursor:"pointer",
            fontSize:"11px",
            fontWeight: isActive ? "bold" : "normal" as React.CSSProperties["fontWeight"],
            whiteSpace:"nowrap"
          }}
        >
          {isInclude?"▲ includi":isExclude?"▼ escludi":"⬚"} {label}
        </button>
        {isActive&&(
          <span style={{fontSize:"9px",color:isInclude?col:T.red,letterSpacing:"0.03em"}}>
            {isInclude?"mostra solo questi":"nasconde questi"}
          </span>
        )}
      </div>
    );
  })}
  <button
    onClick={()=>setFilterFlags({flagged:false,air:false,noPrice:false,noLog:false,calcZero:false,keepOld:false,costCalculated:false})}
    style={{padding:"4px 10px",background:T.surface,color:T.muted,border:`1px solid ${T.border}`,borderRadius:"6px",cursor:"pointer",fontSize:"10px",alignSelf:"flex-start",marginTop:"2px"}}
  >
    ✕ Reset
  </button>
</div>

      {/* ── tabella ── */}
      {/* Barra di scorrimento orizzontale in alto */}
<div 
  ref={topScrollRef}
  style={{
    overflowX: "auto",
    overflowY: "hidden",
    marginBottom: "8px",
    paddingBottom: "4px",
    opacity: 0.5,
    transition: "opacity 0.2s"
  }}
  onMouseEnter={e => e.currentTarget.style.opacity = "1"}
  onMouseLeave={e => e.currentTarget.style.opacity = "0.5"}
>
  <div style={{ width: "max-content", height: "8px" }} />
</div>

<div ref={tableScrollRef} style={{overflowX:"auto",width:"100%"}}>
        <table style={{borderCollapse:"collapse",width:"max-content",minWidth:"100%"}}>
          <thead>
            {/* riga gruppi */}
            <tr style={stickyTop0}>
              <GH span={3}/>
              <GH span={4}/>
              <GH span={7} accent={T.blue}>Costi trasporto e dazi (€/unit)</GH>
              <GH span={2} accent={T.gold}>Step 1</GH>
              <GH span={1} accent={T.purple}>Magazzino</GH>
              <GH span={2} accent={T.green}>Step 2 finale</GH>
              <GH span={2}/>
            </tr>
            {/* riga colonne */}
            <tr style={stickyTop22}>
              <TH align="left" sticky w={70}>{branch==="CAN"?"N COMIT":"N HK"}</TH>
              <TH align="left" w={70}>IFB No</TH>
              <TH align="left" w={180}>Descrizione</TH>
              <TH w={60}>UOM</TH>
              <TH w={55}>Ubicaz.</TH>
              <TH w={55} align="center">Temp.</TH>
              <TH w={55} align="center">Rettif.</TH>
              <TH accent={T.blue} w={70}>Prezzo €</TH>
              <TH accent={T.blue} w={65}>FOB</TH>
              <TH accent={T.blue} w={65}>LIC</TH>
              <TH accent={T.blue} w={55}>VGM</TH>
              <TH accent={T.blue} w={55}>Cert.</TH>
              <TH accent={T.blue} w={60}>Pallet</TH>
              <TH accent={T.blue} w={60}>Alc.Tax</TH>
              <TH accent={T.gold} w={72}>Step1 €</TH>
              <TH accent={T.gold} w={80}>Step1 HKD</TH>
              <TH accent={T.purple} w={65}>WH €</TH>
              <TH accent={T.green} w={72}>Step2 €</TH>
              <TH accent={T.green} w={85}>Step2 HKD ✓</TH>
              <TH w={60}>Δ%</TH>
              <TH w={90}>Ultimo ordine</TH>
            </tr>
          </thead>
          <tbody>
            {filtered.map((r:any,i:number)=>{
              const c = r.cost;
              const prevHkd = r.prevCost?.step2Hkd??null;
              const hkd    = c?.step2Hkd??null;
              const pct    = hkd!=null&&prevHkd!=null&&prevHkd>0?(hkd-prevHkd)/prevHkd*100:null;
              const lastD  = lastOrderDate[r.id];
              const isOld  = lastD&&lastD<sixMonthsAgo;
              const rowBg  = i%2===0?T.bg:T.surface;
              const isSelected = showDetail===r.id;

              return(<>
                                  <tr key={r.id}
                  style={{background:isSelected?`${T.gold}08`:rowBg,opacity:r.isAir?0.45:1,cursor:"pointer"}}
                  onClick={()=>setShowDetail((v:any)=>v===r.id?null:r.id)}>

                  {/* identità */}
                  <td style={{...cellL(true),background:isSelected?`${T.gold}08`:rowBg}}>
                    <span style={{color:T.muted,fontFamily:"monospace",fontSize:"10px"}}>{r.nHK||"—"}</span>
                  </td>
                  <td style={cellL()}>
                    <span style={{color:T.gold,fontFamily:"monospace",fontSize:"10px"}}>{r.code}</span>
                  </td>
                  <td style={{...cellL(),maxWidth:"200px",overflow:"hidden",textOverflow:"ellipsis"}}>
                    {r.description}
                    {r.isAir&&<span style={{marginLeft:"4px",color:T.orange,fontSize:"9px"}}>✈</span>}
                  </td>
                  <td style={cell()}>{r.uom||"—"}</td>
                  <td style={cell()}>
                    {r.ubicazione
                      ? <Chip label={r.ubicazione} color={r.ubicazione==="FOR"?T.purple:r.ubicazione==="MTS"?T.blue:T.green}/>
                      : <span style={{color:T.dim}}>—</span>}
                  </td>

                  {/* temperatura anagrafica */}
                  <td style={{...cell(),textAlign:"center"}}>
                    {r.temperature
                      ? <Chip label={r.temperature}
                          color={r.temperature==="FROZEN"?T.blue:r.temperature==="FRESH"?T.green:T.muted}/>
                      : <span style={{color:T.dim}}>—</span>}
                  </td>

                  {/* temperatura rettificata (Work_tab) */}
                  <td style={{...cell(),textAlign:"center"}}>
                    {r.temperatureOverride && r.temperatureOverride!==r.temperature
                      ? <Chip label={r.temperatureOverride}
                          color={r.temperatureOverride==="FROZEN"?T.blue:r.temperatureOverride==="FRESH"?T.green:T.muted}/>
                      : <span style={{color:T.dim,fontSize:"9px"}}>—</span>}
                  </td>

                  {/* costi breakdown */}
                  <td style={cell(T.text)}>{c?`€${f4(c.priceEur)}`:<span style={{color:T.dim,fontSize:"9px"}}>{r.skipReason||"—"}</span>}</td>
                  <td style={cell()}>{c?f4(c.fob):"—"}</td>
                  <td style={cell()}>{c?f4(c.lic):"—"}</td>
                  <td style={cell()}>{c?f4(c.vgm):"—"}</td>
                  <td style={cell(c?.hc>0?T.orange:undefined)}>{c?(c.hc>0?f4(c.hc):"—"):"—"}</td>
                  <td style={cell()}>{c?f4(c.plt):"—"}</td>
                  <td style={cell(c?.alc>0?T.orange:undefined)}>{c?(c.alc>0?f4(c.alc):"—"):"—"}</td>

                  {/* step 1 */}
                  <td style={cell(T.gold,true)}>{c?`€${c.step1Eur.toFixed(4)}`:"—"}</td>
                  <td style={cell(T.gold,true)}>{c?`${c.step1Hkd.toFixed(2)}`:"—"}</td>

                  {/* magazzino */}
                  <td style={cell(T.purple)}>{c?(c.wh>0?f4(c.wh):"—"):"—"}</td>

                  {/* step 2 */}
                  <td style={cell(T.green,true)}>{c?`€${c.step2Eur.toFixed(4)}`:"—"}</td>
                  <td style={cell(T.green,true)}>
                    <span style={{fontSize:"11px",fontWeight:"bold"}}>
                      {hkd!=null?`${hkd.toFixed(2)}`:<span style={{color:T.dim,fontWeight:"normal",fontSize:"9px"}}>{r.skipReason||"—"}</span>}
                    </span>
                  </td>

                  {/* delta */}
                  <td style={cell(pct==null?T.dim:Math.abs(pct)>=3?(pct>0?T.red:T.green):T.muted,Math.abs(pct||0)>=3)}>
                    {pct!=null?(pct>0?"+":"")+pct.toFixed(1)+"%":"—"}
                    {Math.abs(pct||0)>=3&&" ⚡"}
                  </td>

                  {/* ultimo ordine */}
                  <td style={{...cell(),textAlign:"center"}}>
                    {!lastD
                      ? <span style={{color:T.dim}}>—</span>
                      : isOld
                        ? <div style={{lineHeight:1.2}}>
                            <div style={{color:T.orange,fontWeight:"bold",fontSize:"9px"}}>⚠ KEEP OLD</div>
                            <div style={{color:T.dim,fontSize:"9px"}}>{lastD.toLocaleDateString("it-IT")}</div>
                          </div>
                        : <span style={{color:T.muted}}>{lastD.toLocaleDateString("it-IT")}</span>
                    }
                  </td>
                </tr>

                {/* ── riga dettaglio espansa ── */}
                {isSelected&&c&&(
                  <tr key={r.id+"_detail"}>
                    <td colSpan={21} style={{padding:"8px 16px",background:`${T.gold}06`,
                      borderBottom:`1px solid ${T.gold}33`}}>
                      <div style={{display:"flex",flexWrap:"wrap",gap:"6px",fontSize:"10px"}}>
                        {([
                          ["Prezzo acquisto",`€ ${c.priceEur.toFixed(4)}`,T.text],
                          ["FOB/unit",`€ ${c.fob.toFixed(4)}`,T.blue],
                          ["LIC/unit",`€ ${c.lic.toFixed(4)}`,T.blue],
                          ["VGM/unit",`€ ${c.vgm.toFixed(4)}`,T.blue],
                          ["HC/unit",c.hc>0?`€ ${c.hc.toFixed(4)}`:"—",c.hc>0?T.orange:T.dim],
                          ["Pallet/unit",`€ ${c.plt.toFixed(4)}`,T.blue],
                          ["AlcTax/unit",c.alc>0?`€ ${c.alc.toFixed(4)}`:"—",c.alc>0?T.orange:T.dim],
                          ["Step1 €",`€ ${c.step1Eur.toFixed(4)}`,T.gold],
                          ["Step1 HKD",`${c.step1Hkd.toFixed(2)}`,T.gold],
                          ["WH/unit",c.wh>0?`€ ${c.wh.toFixed(4)}`:"—",T.purple],
                          ["Step2 €",`€ ${c.step2Eur.toFixed(4)}`,T.green],
                          ["Step2 HKD ✓",`${c.step2Hkd.toFixed(2)}`,T.green],
                          ["Rate",`${c.rate}`,T.muted],
                          ["Units/plt",`${c.unitsPerPlt||"—"}`,T.muted],
                        ] as [string,string,string][]).map(([k,v,col])=>(
                          <div key={k} style={{padding:"3px 8px",background:T.card,
                            borderRadius:"4px",border:`1px solid ${T.border}`}}>
                            <span style={{color:T.dim}}>{k}: </span>
                            <span style={{color:col,fontWeight:"bold"}}>{v}</span>
                          </div>
                        ))}
                      </div>
                    </td>
                  </tr>
                )}
              </>);
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}



function InvoiceAndCosts({rows,setRows,branch,airList,products,xrefs,costRows,snapshots,setSnapshots,importLogs,setImportLogs,showToast,bumpImportTs}) {
  const [step,setStep]         = useState(()=>rows?.length?"view":"upload");
  const [preview,setPreview]   = useState<any[]>([]);
  const [headers,setHeaders]   = useState<string[]>([]);
  const [mapping,setMapping]   = useState<any>({});
  const [rawRows,setRawRows]   = useState<any[]>([]);
  const [fileName,setFileName] = useState("");
  const [excludeAir,setExcludeAir]     = useState(false);
  const [newHkdFilter,setNewHkdFilter] = useState<"all"|"ok"|"mancante"|"air">("all");
  const [filterTransport,setFilterTransport] = useState("all");
  const [filterNHK,setFilterNHK]   = useState("");
  const [filterIFBNo,setFilterIFBNo] = useState("");
  const [search,setSearch]     = useState("");
  const [sortDir,setSortDir]   = useState<"desc"|"asc">("desc");

  useEffect(()=>{ if(rows?.length&&step==="upload") setStep("view"); },[rows]);

  function saveRows(data:any[]) {
    setRows(data);
    IDB.set(`ifb_sales_invoice_${branch}`, data);
  }

  function parseFile(e:any) {
    const file=e.target.files?.[0]; if(!file) return;
    setFileName(file.name);
    const reader=new FileReader();
    reader.onload=ev=>{
      try{
        const wb=XLSX.read((ev.target as any).result,{type:"binary"});
        const ws=wb.Sheets[wb.SheetNames[0]];
        const data:any[][]=XLSX.utils.sheet_to_json(ws,{header:1,defval:""});
        if(data.length<2){showToast("File vuoto",T.red);return;}
        const hdrs=data[0].map((h:any)=>String(h||"").trim());
        const raws=data.slice(1).filter((r:any[])=>r.some(c=>c!==""));
        setHeaders(hdrs);setRawRows(raws);
        const norm=(s:string)=>s.toLowerCase().replace(/[\s_()\-\.]/g,"");
        const am:any={};
        hdrs.forEach(h=>{
          const n=norm(h);
          if(!am.itemCode&&(n==="no"||n==="no_"||["itemno","codice","code"].some(a=>n.includes(a))))am.itemCode=h;
          else if(!am.description&&["description","descrizione"].some(a=>n.includes(a)))am.description=h;
          else if(!am.date&&["postingdate","invoicedate","shipdate","lastpostingdate","date","data"].some(a=>n.includes(a)))am.date=h;
          else if(!am.qty&&(n==="qty"||n==="quantity"||n.includes("quantit")))am.qty=h;
          else if(!am.unitPrice&&["unitprice","salesprice","listprice","prezzounit","unitamount","price"].some(a=>n.includes(a.replace(/\s/g,""))))am.unitPrice=h;
          else if(!am.location&&["location","locationcode","ubicazione","warehouse","magazzino"].some(a=>n.includes(a)))am.location=h;
        });
        setMapping(am);setStep("map");
      }catch(err:any){showToast("Errore: "+err.message,T.red);}
    };
    reader.readAsBinaryString(file);
    e.target.value="";
  }

  function buildPreview(){
    const get=(row:any,field:string)=>{const col=mapping[field];if(!col)return"";const i=headers.indexOf(col);return i>=0?row[i]:"";};
    const parsed=rawRows.map(r=>{
      const code=String(get(r,"itemCode")||"").trim();
      const description=String(get(r,"description")||code||"").trim();
      if(!code)return null;
      if(isExcludedDesc(description)||isAccountingCode(code))return null;
      const dateRaw=get(r,"date");
      let dateStr="";
      if(dateRaw){const d=dateRaw instanceof Date?dateRaw:new Date(dateRaw);if(!isNaN(d.getTime()))dateStr=d.toISOString().slice(0,10);else dateStr=String(dateRaw).slice(0,10);}
      const qty=parseFloat(get(r,"qty"))||0;
      const unitPrice=parseFloat(get(r,"unitPrice"))||0;
      const location=String(get(r,"location")||"").trim();
      const prod=findProduct(code,products,xrefs);
      const nHK=prod?.nHK||(xrefs.find((x:any)=>x.ifbNo===code)?.nHK)||"";
      const isAirProd=prod&&airList.some((a:any)=>a.productId===prod.id||(a.code&&a.code===prod.code)||(a.nHK&&prod.nHK&&a.nHK===prod.nHK));
      return{itemCode:code,description,date:dateStr,qty,unitPrice,isSample:qty>0&&unitPrice===0,location,nHK,transport:isAirProd?"AIR":"SEA",_prodFound:!!prod};
    }).filter(Boolean);
    setPreview(parsed);setStep("preview");
  }

  function executeImport(){
    const now=Date.now();
    const data=preview.map((r:any)=>({...r,branch}));
    saveRows(data);
    IDB.set(`ifb_sales_data_${now}`,data);
    const log={id:now,type:"sales",date:new Date(now).toISOString(),count:preview.length,diffs:[],branch};
    const newLogs=[log,...importLogs];setImportLogs(newLogs);LS.set("ifb_importlogs",newLogs);
    const newSnaps=[log,...snapshots].slice(0,50);setSnapshots(newSnaps);LS.set("ifb_snapshots",newSnaps);
    bumpImportTs();showToast(`${preview.length} righe importate ✓`,T.gold);
    setStep("view");
  }

  const activeRows=useMemo(()=>(rows||[]).filter((r:any)=>!r.branch||r.branch===branch),[rows,branch]);

  const enriched=useMemo(()=>{
    return [...activeRows]
      .sort((a:any,b:any)=>{
        if(!a.date&&!b.date)return 0;if(!a.date)return 1;if(!b.date)return -1;
        return sortDir==="desc"?b.date.localeCompare(a.date):a.date.localeCompare(b.date);
      })
      .map((r:any)=>{
        const prod=findProduct(r.itemCode,products,xrefs);
        const cr=prod?costRows.find((c:any)=>c.id===prod.id):null;
        const isAir=r.transport==="AIR"||cr?.isAir===true||cr?.skipReason==="AIR";
        const locationIsNCJ=String(r.location||"").toUpperCase().includes("NCJ");
        const mismatch=(isAir&&!locationIsNCJ)||(!isAir&&locationIsNCJ);
        const newHkd=cr?.cost?.step2Hkd??null;
        const oldHkd=cr?.prevCost?.step2Hkd??null;
        const pct=newHkd!=null&&oldHkd!=null&&oldHkd>0?(newHkd-oldHkd)/oldHkd*100:null;
        const skipReason=isAir?"AIR":cr?.skipReason||(!prod?"NON IN ANAGRAFICA":"");
        return{...r,nHK:prod?.nHK||r.nHK||"",ifbNo:prod?.code||r.itemCode||"",
          description:r.description||prod?.description||"",ubicazione:cr?.ubicazione||"",
          isAir,locationIsNCJ,mismatch,newHkd,oldHkd,pct,skipReason};
      });
  },[activeRows,costRows,products,xrefs,sortDir]);

  const mismatches  = enriched.filter((r:any)=>r.mismatch);
  const airCount    = enriched.filter((r:any)=>r.isAir).length;
  const uniqueNHK   = [...new Set(enriched.map((r:any)=>r.nHK).filter(Boolean))].sort() as string[];
  const uniqueIFBNo = [...new Set(enriched.map((r:any)=>r.ifbNo).filter(Boolean))].sort() as string[];

  let displayed=enriched as any[];
  if(excludeAir)               displayed=displayed.filter(r=>!r.isAir);
  if(filterTransport==="air")      displayed=displayed.filter(r=>r.isAir);
  else if(filterTransport==="sea") displayed=displayed.filter(r=>!r.isAir);
  else if(filterTransport==="mismatch") displayed=displayed.filter(r=>r.mismatch);
  if(newHkdFilter==="ok")          displayed=displayed.filter(r=>r.newHkd!==null&&!r.isAir);
  else if(newHkdFilter==="mancante")displayed=displayed.filter(r=>r.newHkd===null&&!r.isAir);
  else if(newHkdFilter==="air")    displayed=displayed.filter(r=>r.isAir);
  if(filterNHK)   displayed=displayed.filter(r=>r.nHK===filterNHK);
  if(filterIFBNo) displayed=displayed.filter(r=>r.ifbNo===filterIFBNo);
  if(search){const q=search.toLowerCase();displayed=displayed.filter(r=>r.description?.toLowerCase().includes(q)||r.itemCode?.toLowerCase().includes(q)||r.nHK?.toLowerCase().includes(q)||r.location?.toLowerCase().includes(q));}

  // ── STEPS IMPORT ──────────────────────────────────────────────────────────
  if(step==="map") return(
    <div>
      <PageHeader title="📋 Fatture & Costi · Mappatura" sub={`${fileName} · ${rawRows.length} righe`}/>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:"12px",marginBottom:"20px",maxWidth:"800px"}}>
        {([["itemCode","Codice articolo *",true],["description","Descrizione",false],["date","Data fattura",false],["qty","Quantità",false],["unitPrice","Prezzo unitario",false],["location","Location",false]] as [string,string,boolean][]).map(([field,label,req])=>(
          <div key={field}>
            <label style={{display:"block",fontSize:"11px",color:req?T.gold:T.muted,marginBottom:"5px"}}>{label}{req?" *":""}</label>
            <select value={mapping[field]||""} onChange={e=>setMapping((m:any)=>({...m,[field]:e.target.value||null}))}
              style={{...inputStyle(),cursor:"pointer",borderColor:req&&!mapping[field]?T.red+"88":T.border}}>
              <option value="">— non mappato —</option>
              {headers.map(h=><option key={h} value={h}>{h}</option>)}
            </select>
          </div>
        ))}
      </div>
      <div style={{display:"flex",gap:"10px"}}>
        <ActionBtn label="← Ricarica" onClick={()=>setStep("upload")}/>
        <ActionBtn label="Preview →" onClick={buildPreview} primary disabled={!mapping["itemCode"]}/>
      </div>
    </div>
  );

  if(step==="preview") return(
    <div>
      <PageHeader title="📋 Fatture & Costi · Preview" sub={`${fileName} · ${preview.length} righe valide`}/>
      {rows?.length>0&&<div style={{background:`${T.orange}15`,border:`1px solid ${T.orange}44`,borderRadius:"6px",padding:"10px 14px",marginBottom:"14px",fontSize:"12px",color:T.orange}}>⚠ Questo import sostituirà i dati attuali ({rows.length} righe).</div>}
      <div style={{display:"flex",gap:"12px",marginBottom:"16px",flexWrap:"wrap"}}>
        {[[preview.length,"Totale",T.text],[preview.filter((r:any)=>r.transport==="AIR").length,"✈ AIR",T.orange],[preview.filter((r:any)=>r.transport==="SEA").length,"⛴ SEA",T.blue],[preview.filter((r:any)=>!r._prodFound).length,"⚠ Non in anagrafica",T.red]].map(([n,l,c])=>(
          <div key={l as string} style={{padding:"10px 16px",background:T.card,border:`1px solid ${T.border}`,borderRadius:"8px"}}>
            <div style={{fontSize:"20px",fontWeight:"bold",color:c as string}}>{n as number}</div>
            <div style={{fontSize:"10px",color:T.dim}}>{l as string}</div>
          </div>
        ))}
      </div>
      <div style={{display:"flex",gap:"10px",marginBottom:"16px"}}>
        <ActionBtn label="← Mappa" onClick={()=>setStep("map")}/>
        <ActionBtn label={`✓ Importa ${preview.length} righe`} onClick={executeImport} primary/>
      </div>
    </div>
  );

  if(step==="upload") return(
    <div>
      <PageHeader title="📋 Fatture & Costi" sub="Carica il file fattura"/>
      {importLogs.filter((l:any)=>l.type==="sales"&&l.branch===branch).length>0&&(
        <select onChange={async e=>{
          if(!e.target.value)return;
          const snap=importLogs.find((l:any)=>String(l.id)===e.target.value);
          if(!snap)return;
          if(window.confirm(`Ripristinare fattura del ${new Date(snap.id).toLocaleDateString("it-IT")}?`)){
            const r=await IDB.get(`ifb_sales_data_${snap.id}`,null);
            if(!r?.length){showToast("Snapshot non disponibile",T.orange);return;}
            saveRows(r);setStep("view");showToast(`Ripristinata: ${snap.count} righe ✓`,T.gold);
          }
          e.target.value="";
        }} style={{...inputStyle(),width:"auto",fontSize:"12px",marginBottom:"16px"}} defaultValue="">
          <option value="">📜 Carica da storico ({importLogs.filter((l:any)=>l.type==="sales"&&l.branch===branch).length})</option>
          {importLogs.filter((l:any)=>l.type==="sales"&&l.branch===branch).map((s:any)=>(
            <option key={s.id} value={String(s.id)}>{new Date(s.id).toLocaleDateString("it-IT")} · {s.count} righe</option>
          ))}
        </select>
      )}
      <Section title="Carica file fattura">
        <DropZone onFile={(f:File)=>{const e={target:{files:[f],value:""}} as any;parseFile(e);}} label="Trascina o clicca — Excel / CSV fattura"/>
      </Section>
    </div>
  );

  // ── VIEW ─────────────────────────────────────────────────────────────────
  return(
    <div>
      <PageHeader title={`Fatture & Costi · ${branch}`} sub={`${enriched.length} righe · ${fileName||"dati caricati"}`}/>

      {/* Mismatch banner */}
      {mismatches.length>0&&(
        <div style={{background:`${T.orange}15`,border:`1px solid ${T.orange}`,borderRadius:"6px",padding:"10px 16px",marginBottom:"14px",display:"flex",alignItems:"center",gap:"12px",flexWrap:"wrap"}}>
          <span style={{color:T.orange,fontWeight:"bold"}}>
            ⚠ {mismatches.filter((r:any)=>r.isAir&&!r.locationIsNCJ).length} AIR senza NCJ &nbsp;·&nbsp;
            {mismatches.filter((r:any)=>!r.isAir&&r.locationIsNCJ).length} NCJ ma SEA
          </span>
          <button onClick={()=>setFilterTransport(v=>v==="mismatch"?"all":"mismatch")}
            style={{padding:"4px 12px",background:filterTransport==="mismatch"?T.purple:T.surface,color:filterTransport==="mismatch"?"#fff":T.purple,border:`1px solid ${T.purple}`,borderRadius:"4px",cursor:"pointer",fontSize:"12px",fontWeight:"bold"}}>
            {filterTransport==="mismatch"?"Mostra tutte":"Mostra mismatch"}
          </button>
        </div>
      )}

      {/* Toolbar */}
      <div style={{display:"flex",gap:"8px",marginBottom:"12px",alignItems:"center",flexWrap:"wrap"}}>
        <label style={{display:"inline-block",padding:"6px 14px",background:T.surface,border:`1px solid ${T.border}`,borderRadius:"6px",cursor:"pointer",fontSize:"12px",color:T.text}}>
          📂 Ricarica
          <input type="file" accept=".xlsx,.xls,.csv" onChange={parseFile} style={{display:"none"}}/>
        </label>
        {importLogs.filter((l:any)=>l.type==="sales"&&l.branch===branch).length>0&&(
          <select onChange={async e=>{
            if(!e.target.value)return;
            const snap=importLogs.find((l:any)=>String(l.id)===e.target.value);
            if(!snap)return;
            if(window.confirm(`Ripristinare fattura del ${new Date(snap.id).toLocaleDateString("it-IT")}?`)){
              const r=await IDB.get(`ifb_sales_data_${snap.id}`,null);
              if(!r?.length){showToast("Snapshot non disponibile",T.orange);return;}
              saveRows(r);showToast(`${snap.count} righe ripristinate ✓`,T.gold);
            }
            e.target.value="";
          }} style={{...inputStyle(),width:"auto",fontSize:"12px"}} defaultValue="">
            <option value="">📜 Storico ({importLogs.filter((l:any)=>l.type==="sales"&&l.branch===branch).length})</option>
            {importLogs.filter((l:any)=>l.type==="sales"&&l.branch===branch).map((s:any)=>(
              <option key={s.id} value={String(s.id)}>{new Date(s.id).toLocaleDateString("it-IT")} · {s.count} righe</option>
            ))}
          </select>
        )}
        <button onClick={()=>exportXLSX(
          displayed.map((r:any)=>({
            "Data":r.date||"",[branchN(branch)]:r.nHK||"","IFB No":r.ifbNo||"","Descrizione":r.description||"",
            "Qty":r.qty||"","Prezzo Unit.":r.unitPrice||"","Location":r.location||"",
            "Mismatch":r.mismatch?"⚠ "+( r.isAir&&!r.locationIsNCJ?"AIR senza NCJ":"NCJ ma SEA"):"",
            "Mag./Trasp.":r.isAir?"AIR":r.ubicazione||"",
            "Old HKD":r.oldHkd!=null?roundN(r.oldHkd):"","New HKD":r.isAir?"AIR":r.newHkd!=null?roundN(r.newHkd):"MANCANTE",
            "Δ%":r.pct!=null?roundN(r.pct,1):"","Motivo":r.skipReason||"",
          })),
          "Fatture & Costi",`Fatture_${branch}.xlsx`
        )} style={{padding:"6px 14px",background:`${T.green}20`,border:`1px solid ${T.green}44`,borderRadius:"6px",color:T.green,cursor:"pointer",fontSize:"11px"}}>
          ⬇ Export Excel
        </button>
        <button onClick={()=>setExcludeAir(v=>!v)}
          style={{padding:"6px 14px",background:excludeAir?`${T.orange}20`:T.surface,color:excludeAir?T.orange:T.muted,border:`1px solid ${excludeAir?T.orange:T.border}`,borderRadius:"6px",cursor:"pointer",fontSize:"11px",fontWeight:excludeAir?"bold":"normal"}}>
          {excludeAir?`✓ AIR esclusi (${airCount})`:`✈ Escludi AIR (${airCount})`}
        </button>
        <button onClick={()=>setSortDir(d=>d==="desc"?"asc":"desc")}
          style={{padding:"6px 14px",background:T.surface,color:T.muted,border:`1px solid ${T.border}`,borderRadius:"6px",cursor:"pointer",fontSize:"11px"}}>
          Data {sortDir==="desc"?"↓":"↑"}
        </button>
        <button onClick={()=>{if(window.confirm(`Eliminare i dati (${activeRows.length} righe)?`)){saveRows([]);setStep("upload");}}}
          style={{padding:"6px 12px",background:"none",border:`1px solid ${T.red}44`,borderRadius:"6px",color:T.red,cursor:"pointer",fontSize:"11px"}}>
          ✕ Svuota
        </button>
      </div>

      {/* Filtri transport */}
      <div style={{display:"flex",gap:"6px",marginBottom:"10px",flexWrap:"wrap"}}>
        {([["all",`Tutte (${enriched.length})`,T.text],["air",`✈ AIR (${airCount})`,T.orange],["sea",`⛴ SEA (${enriched.length-airCount})`,T.blue],["mismatch",`⚠ Mismatch (${mismatches.length})`,T.purple]] as [string,string,string][]).map(([v,l,c])=>(
          <button key={v} onClick={()=>setFilterTransport(v)}
            style={{padding:"5px 12px",background:filterTransport===v?`${c}20`:T.surface,color:filterTransport===v?c:T.muted,border:`1px solid ${filterTransport===v?c:T.border}`,borderRadius:"6px",cursor:"pointer",fontSize:"11px",fontWeight:filterTransport===v?"bold":"normal"}}>
            {l}
          </button>
        ))}
      </div>

      <SearchBar value={search} onChange={setSearch} placeholder="🔍 Cerca codice, N HK, descrizione, location…"/>

      <Section title={`${displayed.length} righe`}>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse"}}>
            <thead><tr>
              {["Data",branchN(branch)+" ▾","IFB No ▾","Descrizione","Qty","Prezzo","Location","Mag./Trasp.","Old HKD","New HKD ▾","Δ%","Motivo"].map((c,ci)=>{
                if(c===branchN(branch)+" ▾") return(
                  <th key={c} style={{padding:"4px 8px",background:T.card,borderBottom:`1px solid ${T.border}`,position:"sticky",top:0,zIndex:10}}>
                    <select value={filterNHK} onChange={e=>setFilterNHK(e.target.value)}
                      style={{background:filterNHK?`${T.gold}22`:T.card,color:filterNHK?T.gold:T.muted,border:`1px solid ${filterNHK?T.gold:T.border}`,borderRadius:"4px",padding:"3px 6px",fontSize:"10px",cursor:"pointer",fontFamily:"inherit",outline:"none",maxWidth:"110px"}}>
                      <option value="">{branchN(branch)} ▾</option>
                      {uniqueNHK.map(v=><option key={v} value={v}>{v}</option>)}
                    </select>
                  </th>
                );
                if(c==="IFB No ▾") return(
                  <th key={c} style={{padding:"4px 8px",background:T.card,borderBottom:`1px solid ${T.border}`,position:"sticky",top:0,zIndex:10}}>
                    <select value={filterIFBNo} onChange={e=>setFilterIFBNo(e.target.value)}
                      style={{background:filterIFBNo?`${T.gold}22`:T.card,color:filterIFBNo?T.gold:T.muted,border:`1px solid ${filterIFBNo?T.gold:T.border}`,borderRadius:"4px",padding:"3px 6px",fontSize:"10px",cursor:"pointer",fontFamily:"inherit",outline:"none",maxWidth:"110px"}}>
                      <option value="">IFB No ▾</option>
                      {uniqueIFBNo.map(v=><option key={v} value={v}>{v}</option>)}
                    </select>
                  </th>
                );
                if(c==="New HKD ▾") return(
                  <th key={c} style={{padding:"4px 8px",background:T.card,borderBottom:`1px solid ${T.border}`,position:"sticky",top:0,zIndex:10}}>
                    <select value={newHkdFilter} onChange={e=>setNewHkdFilter(e.target.value as any)}
                      style={{background:newHkdFilter!=="all"?`${T.gold}22`:T.card,color:newHkdFilter!=="all"?T.gold:T.muted,border:`1px solid ${newHkdFilter!=="all"?T.gold:T.border}`,borderRadius:"4px",padding:"3px 6px",fontSize:"10px",cursor:"pointer",fontFamily:"inherit",outline:"none"}}>
                      <option value="all">New HKD ▾</option>
                      <option value="ok">✅ Con costo</option>
                      <option value="mancante">❌ MANCANTE</option>
                      <option value="air">✈ AIR</option>
                    </select>
                  </th>
                );
                return <th key={c} style={{padding:"7px 10px",background:T.card,color:T.muted,textAlign:"left",borderBottom:`1px solid ${T.border}`,fontSize:"11px",fontWeight:"normal",whiteSpace:"nowrap",position:"sticky",top:0,zIndex:10}}>{c}</th>;
              })}
            </tr></thead>
            <tbody>
              {displayed.slice(0,1000).map((r:any,i:number)=>{
                const mismatchType=r.mismatch?(r.isAir&&!r.locationIsNCJ?"AIR senza NCJ":"NCJ ma SEA"):"";
                return(
                  <tr key={i} style={{borderBottom:`1px solid ${T.border}`,background:r.mismatch?`${T.purple}10`:i%2===0?T.bg:T.surface}}>
                    <td style={{padding:"6px 10px",fontSize:"11px",fontFamily:"monospace",whiteSpace:"nowrap"}}><span style={{color:T.gold,fontWeight:"bold"}}>{r.date||"—"}</span></td>
                    <td style={{padding:"6px 10px",fontSize:"11px",fontFamily:"monospace"}}><span style={{color:T.muted}}>{r.nHK||"—"}</span></td>
                    <td style={{padding:"6px 10px",fontSize:"11px",fontFamily:"monospace"}}><span style={{color:T.gold}}>{r.ifbNo||r.itemCode||"—"}</span></td>
                    <td style={{padding:"6px 10px",fontSize:"12px",maxWidth:"220px",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                      <span style={{color:r._prodFound===false?T.orange:T.text}}>{r.description}</span>
                    </td>
                    <td style={{padding:"6px 10px",fontSize:"11px",fontFamily:"monospace",textAlign:"right"}}><span style={{color:T.muted}}>{r.qty||"—"}</span></td>
                    <td style={{padding:"6px 10px",fontSize:"11px",fontFamily:"monospace",textAlign:"right"}}>
                      {r.isSample?<Chip label="SAMPLE" color={T.purple}/>:<span style={{color:T.muted}}>{r.unitPrice>0?r.unitPrice.toFixed(2):"—"}</span>}
                    </td>
                    <td style={{padding:"6px 10px",fontSize:"11px",fontFamily:"monospace"}}><span style={{color:r.mismatch?T.purple:T.muted}}>{r.location||"—"}</span></td>
                    <td style={{padding:"6px 10px",fontSize:"11px",whiteSpace:"nowrap"}}>
                      {r.isAir
                        ? <><Chip label="✈ AIR" color={r.locationIsNCJ?T.green:T.orange}/>{r.mismatch&&<span style={{marginLeft:"5px",fontSize:"9px",color:T.purple}}>⚠ {mismatchType}</span>}</>
                        : <><Chip label={r.ubicazione||"—"} color={r.ubicazione==="FOR"?T.purple:r.ubicazione==="MTS"?T.blue:T.green}/>{r.mismatch&&<span style={{marginLeft:"5px",fontSize:"9px",color:T.purple}}>⚠ {mismatchType}</span>}</>
                      }
                    </td>
                    <td style={{padding:"6px 10px",fontSize:"11px",fontFamily:"monospace",textAlign:"right"}}><span style={{color:T.muted}}>{r.oldHkd!=null?r.oldHkd.toFixed(2):"—"}</span></td>
                    <td style={{padding:"6px 10px",fontSize:"11px",fontFamily:"monospace",textAlign:"right"}}>
                      {r.isAir
                        ? <span style={{color:T.orange,fontWeight:"bold"}}>AIR</span>
                        : r.newHkd!=null
                          ? <span style={{color:T.gold,fontWeight:"bold"}}>{r.newHkd.toFixed(2)}</span>
                          : <span style={{color:T.red,fontWeight:"bold"}}>MANCANTE</span>
                      }
                    </td>
                    <td style={{padding:"6px 10px",fontSize:"11px",textAlign:"right"}}>
                      {r.pct!=null?<span style={{color:r.pct>3?T.red:r.pct<-3?T.green:T.text,fontWeight:Math.abs(r.pct)>3?"bold":"normal"}}>{r.pct>0?"+":""}{r.pct.toFixed(1)}%</span>:<span style={{color:T.dim}}>—</span>}
                    </td>
                    <td style={{padding:"6px 10px",fontSize:"10px",maxWidth:"160px"}}>
                      {r.skipReason?<span style={{color:r.skipReason==="AIR"?T.orange:T.orange,fontStyle:"italic"}}>{r.skipReason}</span>:<span style={{color:T.dim}}>—</span>}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
          {displayed.length>1000&&<div style={{padding:"12px",textAlign:"center",color:T.muted,fontSize:"11px"}}>Mostrate 1000/{displayed.length} righe</div>}
        </div>
      </Section>
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



// ─── STORICO ──────────────────────────────────────────────────────────────────
function Storico({snapshots,setSnapshots,costHistory,setCostHistory,branch,showToast}) {
  const[sel,setSel]=useState<any>(null);
  const[sortDir,setSortDir]=useState("asc");
  const[deltaFilter,setDeltaFilter]=useState("all");
  const[showModified,setShowModified]=useState(false);
  const[showNew,setShowNew]=useState(false);
  const[selCostSnap,setSelCostSnap]=useState<any>(null);
  const costSnaps = (costHistory || []).filter((s:any) => !s.branch || s.branch === branch);

  const branchSnaps = snapshots.filter((s:any)=>
    !s.branch || s.branch==="ALL" || s.branch===branch
  );
  const snapDate=(s:any)=>new Date(s.date||s.id).toLocaleDateString("it-IT",{day:"2-digit",month:"2-digit",year:"numeric"});

  function deleteSnap(id:any){
    const next=snapshots.filter((s:any)=>s.id!==id);
    setSnapshots(next);LS.set("ifb_snapshots",next);
    if(sel?.id===id) setSel(null);
  }

  const ICON:any={prices:"💶",anagrafica:"◈",xref:"⇄",air:"✈",sales:"📋"};
  const LABEL:any={prices:"Import Listini",anagrafica:"Import Anagrafica",xref:"Import XRef",air:"Import AIR",sales:"Sales Invoice"};

  return(
    <div>
      <PageHeader title="Storico & Diff" sub="Snapshot import e Standard Cost"/>

      {/* ── COST HISTORY ── */}
      {costSnaps.length>0&&(
        <Section title={`📊 Storico Standard Cost · ${branch}`} accent={T.gold}>
          <div style={{fontSize:"12px",color:T.muted,marginBottom:"10px"}}>
            Clicca una data per vedere i costi salvati in quel momento
          </div>
          <div style={{display:"flex",flexDirection:"column",gap:"8px",marginBottom:"12px"}}>
            {costSnaps.map((s:any, idx:number) => {
              // Conta quante volte appare questa data
              const dateStr = new Date(s.ts).toLocaleDateString("it-IT");
              const sameDates = costSnaps.filter((x:any) => 
                new Date(x.ts).toLocaleDateString("it-IT") === dateStr
              ).length;
              const isDuplicate = sameDates > 1;
              
              return (
                <div key={s.ts} style={{display:"flex",alignItems:"center",gap:"8px"}}>
                  <button 
                    onClick={() => setSelCostSnap(selCostSnap?.ts === s.ts ? null : s)}
                    style={{
                      flex:1,
                      padding:"8px 12px",
                      background:selCostSnap?.ts === s.ts ? T.gold : T.card,
                      color:selCostSnap?.ts === s.ts ? "#000" : T.text,
                      border:`1px solid ${selCostSnap?.ts === s.ts ? T.gold : T.border}`,
                      borderRadius:"6px",
                      cursor:"pointer",
                      fontSize:"12px",
                      textAlign:"left",
                      display:"flex",
                      justifyContent:"space-between",
                      alignItems:"center"
                    }}
                  >
                    <span>
                      {s.month || "?"} · {dateStr}
                      {isDuplicate && (
                        <span style={{
                          marginLeft:"8px",
                          fontSize:"10px",
                          color:T.orange,
                          background:`${T.orange}20`,
                          padding:"2px 6px",
                          borderRadius:"4px"
                        }}>
                          ⚠ duplicato #{idx + 1}
                        </span>
                      )}
                    </span>
                    <span style={{
                      fontSize:"10px",
                      color:T.muted,
                      fontFamily:"monospace"
                    }}>
                      {new Date(s.ts).toLocaleTimeString("it-IT", {hour:"2-digit", minute:"2-digit", second:"2-digit"})}
                    </span>
                  </button>
                  
                  {/* Bottone elimina */}
                  <button
                    onClick={() => {
                      if(window.confirm(`Eliminare lo snapshot del ${dateStr} alle ${new Date(s.ts).toLocaleTimeString("it-IT")}?`)) {
                        const newCostHistory = costHistory.filter((c:any) => c.ts !== s.ts);
                        setCostHistory(newCostHistory);
                        LS.set("ifb_costhistory", newCostHistory);
                        if(selCostSnap?.ts === s.ts) setSelCostSnap(null);
                        showToast(`Snapshot del ${dateStr} eliminato`, T.red);
                      }
                    }}
                    style={{
                      padding:"6px 12px",
                      background:"none",
                      border:`1px solid ${T.red}44`,
                      borderRadius:"6px",
                      color:T.red,
                      cursor:"pointer",
                      fontSize:"11px",
                      whiteSpace:"nowrap"
                    }}
                    title="Elimina questo snapshot"
                  >
                    🗑 Elimina
                  </button>
                </div>
              );
            })}
          </div>
          
          {/* Bottone per eliminare TUTTI gli snapshot */}
          {costSnaps.length > 1 && (
            <div style={{marginBottom:"16px"}}>
              <button
                onClick={() => {
                  if(window.confirm(`⚠️ ATTENZIONE: Eliminare TUTTI i ${costSnaps.length} snapshot di Standard Cost per ${branch}? Questa operazione è irreversibile.`)) {
                    const newCostHistory = costHistory.filter((c:any) => c.branch !== branch);
                    setCostHistory(newCostHistory);
                    LS.set("ifb_costhistory", newCostHistory);
                    setSelCostSnap(null);
                    showToast(`Eliminati ${costSnaps.length} snapshot per ${branch}`, T.red);
                  }
                }}
                style={{
                  padding:"6px 14px",
                  background:`${T.red}15`,
                  border:`1px solid ${T.red}`,
                  borderRadius:"6px",
                  color:T.red,
                  cursor:"pointer",
                  fontSize:"11px",
                  fontWeight:"bold"
                }}
              >
                🗑 Elimina TUTTI gli snapshot ({costSnaps.length})
              </button>
            </div>
          )}
    
    {/* Visualizzazione dettaglio snapshot selezionato */}
    {selCostSnap && (
      <div style={{overflowX:"auto", marginTop:"16px"}}>
        <div style={{marginBottom:"8px", fontSize:"11px", color:T.muted}}>
          Snapshot del {new Date(selCostSnap.ts).toLocaleString("it-IT")}
        </div>
        <table style={{width:"100%",borderCollapse:"collapse"}}>
          <THead cols={[branchN(branch),"IFB No","Descrizione","Costo HKD","Note"]} sticky />
          <tbody>{(selCostSnap.rows||[]).map((r:any,i:number)=>(
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
  </Section>
)}

      {/* ── IMPORT SNAPSHOTS LIST ── */}
      <Section title="📥 Storico Import">
        {snapshots.length===0
          ? <div style={{padding:"24px",textAlign:"center",color:T.dim,fontSize:"13px"}}>Nessuno snapshot ancora.</div>
          : <div style={{display:"flex",flexDirection:"column",gap:"6px",maxHeight:"320px",overflowY:"auto"}}>
              {snapshots.map((s:any)=>(
                <div key={s.id}
                  style={{display:"flex",alignItems:"center",gap:"8px",padding:"8px 12px",
                    background:sel?.id===s.id?`${T.gold}15`:T.card,
                    border:`1px solid ${sel?.id===s.id?T.gold:T.border}`,
                    borderRadius:"6px",cursor:"pointer"}}
                  onClick={()=>{setSel(sel?.id===s.id?null:s);setShowModified(false);setShowNew(false);}}>
                  <span style={{fontSize:"16px"}}>{ICON[s.type]||"📥"}</span>
                  <div style={{flex:1,minWidth:0}}>
                    <div style={{fontSize:"12px",color:sel?.id===s.id?T.gold:T.text,fontWeight:"bold"}}>
                      {LABEL[s.type]||"Import"}
                      {s.branch&&s.branch!=="ALL"&&<span style={{marginLeft:"6px",color:T.muted,fontWeight:"normal"}}>· {s.branch}</span>}
                      {s.month&&<span style={{marginLeft:"6px",color:T.gold,fontWeight:"normal",fontSize:"11px"}}>· {s.month}</span>}
                    </div>
                    <div style={{fontSize:"11px",color:T.muted,marginTop:"2px"}}>
                      {snapDate(s)} alle {new Date(s.date||s.id).toLocaleTimeString("it-IT",{hour:"2-digit",minute:"2-digit"})}
                      {" · "}{s.count} voci
                      {s.diffs?.length>0&&(
                        <span style={{color:T.orange}}>
                          {" · "}{s.diffs.filter((d:any)=>d.isNew).length} nuovi
                          {" · "}{s.diffs.filter((d:any)=>!d.isNew&&d.fields?.length>0).length} modif.
                        </span>
                      )}
                    </div>
                  </div>
                  <button onClick={e=>{e.stopPropagation();deleteSnap(s.id);}}
                    style={{background:"none",border:"none",color:T.muted,cursor:"pointer",fontSize:"13px",padding:"2px 6px"}}>✕</button>
                </div>
              ))}
            </div>
        }
      </Section>

      {/* ── DETAIL PANEL ── */}
      {sel&&(()=>{
        const diffs:any[]=sel.diffs||[];
        const newItems=diffs.filter((d:any)=>d.isNew);
        const priceFields=["fcaPrice","fcaDiscounted","dapPrice","dapDiscounted","mtsPrice","dapFinal"];

        const realModified=diffs
          .filter((d:any)=>!d.isNew&&d.fields?.length>0)
          .map((d:any)=>({...d,fields:d.fields.filter((f:any)=>{
            if(!priceFields.includes(f.field)) return true;
            return Math.abs(roundN(f.new||0)-roundN(f.old||0))>=0.005;
          })}))
          .filter((d:any)=>d.fields.length>0);

        const getPct=(d:any)=>{
          const pf=d.fields.find((f:any)=>f.field==="dapFinal")||d.fields[0];
          if(!pf||!pf.old||pf.old===0) return 0;
          return(roundN(pf.new)-roundN(pf.old))/Math.abs(pf.old)*100;
        };

        let shownDiffs=realModified;
        if(deltaFilter==="minus") shownDiffs=realModified.filter((d:any)=>getPct(d)<-3);
        else if(deltaFilter==="plus") shownDiffs=realModified.filter((d:any)=>getPct(d)>3);
        shownDiffs=[...shownDiffs].sort((a,b)=>sortDir==="asc"?getPct(a)-getPct(b):getPct(b)-getPct(a));

        const thisDate=snapDate(sel);
        const prevSnap=snapshots.find((s:any)=>s.id!==sel.id&&s.type===sel.type&&s.branch===sel.branch&&s.month===sel.month);
        const prevDate=prevSnap?snapDate(prevSnap):"—";

        return(
          <div style={{background:T.card,borderRadius:"8px",padding:"16px",border:`1px solid ${T.border}`,marginTop:"16px"}}>
            <h3 style={{color:T.gold,marginTop:0,marginBottom:"12px",fontSize:"13px"}}>
              {LABEL[sel.type]||sel.type} · {thisDate} · {sel.branch||"ALL"} · {sel.count} voci
            </h3>

            <div style={{display:"flex",gap:"8px",flexWrap:"wrap",alignItems:"center",marginBottom:"14px"}}>
              <button onClick={()=>{setShowNew(v=>!v);setShowModified(false);}}
                style={{padding:"6px 14px",background:showNew?T.green:T.surface,
                  color:showNew?"#000":T.green,border:`1px solid ${T.green}`,
                  borderRadius:"6px",cursor:"pointer",fontSize:"12px",fontWeight:"bold"}}>
                🆕 {newItems.length} nuovi
              </button>
              <button onClick={()=>{setShowModified(v=>!v);setShowNew(false);}}
                style={{padding:"6px 14px",background:showModified?T.orange:T.surface,
                  color:showModified?"#000":T.orange,border:`1px solid ${T.orange}`,
                  borderRadius:"6px",cursor:"pointer",fontSize:"12px",fontWeight:"bold"}}>
                ✏️ {realModified.length} modificati
              </button>
              {showModified&&<>
                <button onClick={()=>setDeltaFilter(f=>f==="minus"?"all":"minus")}
                  style={{padding:"4px 10px",background:deltaFilter==="minus"?T.red:T.surface,
                    color:deltaFilter==="minus"?"#fff":T.red,border:`1px solid ${T.red}`,
                    borderRadius:"4px",cursor:"pointer",fontSize:"11px"}}>{"< −3%"}</button>
                <button onClick={()=>setDeltaFilter(f=>f==="plus"?"all":"plus")}
                  style={{padding:"4px 10px",background:deltaFilter==="plus"?T.green:T.surface,
                    color:deltaFilter==="plus"?"#fff":T.green,border:`1px solid ${T.green}`,
                    borderRadius:"4px",cursor:"pointer",fontSize:"11px"}}>{">"} +3%</button>
                <button onClick={()=>setSortDir(d=>d==="asc"?"desc":"asc")}
                  style={{padding:"4px 10px",background:T.surface,color:T.muted,
                    border:`1px solid ${T.border}`,borderRadius:"4px",cursor:"pointer",fontSize:"11px"}}>
                  Δ {sortDir==="asc"?"↑":"↓"}
                </button>
                {deltaFilter!=="all"&&<span style={{fontSize:"11px",color:T.muted}}>({shownDiffs.length}/{realModified.length})</span>}
              </>}
            </div>

            {showNew&&(
              <div style={{marginBottom:"14px"}}>
                <div style={{color:T.green,fontSize:"12px",fontWeight:"bold",marginBottom:"8px"}}>Nuovi ({newItems.length})</div>
                {newItems.length===0
                  ? <div style={{color:T.dim,fontSize:"12px"}}>Nessuno.</div>
                  : <table style={{width:"100%",borderCollapse:"collapse"}}>
                      <THead cols={["IFB No","Descrizione"]} sticky />
                      <tbody>{newItems.map((d:any,i:number)=>(
                        <tr key={i} style={{borderBottom:`1px solid ${T.border}`}}>
                          <TD mono><span style={{color:T.gold}}>{d.id||d.productId}</span></TD>
                          <TD>{d.description}</TD>
                        </tr>
                      ))}</tbody>
                    </table>
                }
              </div>
            )}

            {showModified&&(
              <div>
                <div style={{color:T.orange,fontSize:"12px",fontWeight:"bold",marginBottom:"8px"}}>Modifiche ({shownDiffs.length})</div>
                {shownDiffs.length===0
                  ? <div style={{color:T.dim,fontSize:"12px"}}>Nessuna variazione reale.</div>
                  : <div style={{overflowX:"auto"}}>
                      <table style={{width:"100%",borderCollapse:"collapse"}}>
                        <THead cols={[`IFB No / ${branchN(branch)}`,"Descrizione","Campo",`Vecchio (${prevDate})`,`Nuovo (${thisDate})`,"Δ%"]} sticky />
                        <tbody>{shownDiffs.map((d:any,i:number)=>
                          d.fields.map((f:any,j:number)=>{
                            const oldR=roundN(f.old||0),newR=roundN(f.new||0);
                            const pct=oldR!==0?(newR-oldR)/Math.abs(oldR)*100:null;
                            return(
                              <tr key={`${i}-${j}`} style={{
                                borderBottom:j===d.fields.length-1?`1px solid ${T.border}`:`1px solid ${T.border}44`,
                                background:i%2===0?T.bg:T.surface}}>
                                {j===0&&<>
                                  <td rowSpan={d.fields.length} style={{padding:"6px 12px",borderBottom:`1px solid ${T.border}`,verticalAlign:"top",fontFamily:"monospace",fontSize:"12px",color:T.gold}}>
                                    {d.ifbNo||d.id}<br/>
                                    <span style={{color:T.muted,fontSize:"10px"}}>{d.nHK||""}</span>
                                  </td>
                                  <td rowSpan={d.fields.length} style={{padding:"6px 12px",borderBottom:`1px solid ${T.border}`,verticalAlign:"top",fontSize:"12px",color:T.text}}>
                                    {d.description}
                                  </td>
                                </>}
                                <TD><span style={{color:T.muted,fontSize:"11px"}}>{f.field}</span></TD>
                                <TD mono>{oldR.toFixed(2)}</TD>
                                <TD mono>{newR.toFixed(2)}</TD>
                                <TD><span style={{color:pct==null?T.dim:pct>0?T.red:T.green,fontWeight:"bold"}}>
                                  {pct!=null?(pct>0?"+":"")+pct.toFixed(1)+"%":"—"}
                                </span></TD>
                              </tr>
                            );
                          })
                        )}</tbody>
                      </table>
                    </div>
                }
              </div>
            )}
          </div>
        );
      })()}
    </div>
  );
}

// ─── PRODUCTS (con import integrato e storico) ─────────────────────────────
function Products({ products, setProducts, branch, importLogs, setImportLogs, snapshots, setSnapshots, showToast, bumpImportTs }) {
  const [search, setSearch] = useState("");
  const [onlyIFB, setOnlyIFB] = useState(true);
  const [importStep, setImportStep] = useState<"idle" | "map" | "preview">("idle");
  const [headers, setHeaders] = useState<string[]>([]);
  const [rawRows, setRawRows] = useState<any[]>([]);
  const [map, setMap] = useState<any>({});
  const [preview, setPreview] = useState<any[]>([]);
  const [fileName, setFileName] = useState("");

  // Storico import anagrafica
  const anagSnaps = snapshots.filter((s: any) => s.type === "anagrafica" && (!s.branch || s.branch === "ALL" || s.branch === branch));

  const FIELDS = ["nHK", "code", "description", "category", "uom", "qtyPerBox", "boxPerPallet", "kgPerBox", "kgxplt", "temperature", "active", "vendorName", "vendorName2"];
  const FLABELS = {
    nHK: "N HK (No_)",
    code: "IFB Item *",
    description: "Descrizione *",
    category: "Section",
    uom: "UOM",
    qtyPerBox: "Qty/Cartone",
    boxPerPallet: "Cartoni/Pallet",
    kgPerBox: "Kg per Cartone",
    temperature: "Product Type",
    active: "Bloccato",
    vendorName: "Vendor Name",
    vendorName2: "Vendor Name 2"
  };

  const LOCAL_ALIASES = {
    nHK: ["no", "no_"],
    code: ["ifbitem", "ifb item", "ifb no", "ifb n"],
    description: ["description"],
    category: ["sectiondescription", "section description", "section"],
    uom: ["salesunitofmeasure", "sales unit of measure"],
    qtyPerBox: ["quantityxpackaging", "quantity x packaging"],
    boxPerPallet: ["packagingxpallet", "packaging x pallet"],
    kgPerBox: ["netweight", "net weight"],
    kgxplt: ["kgxplt", "kg x pallet", "kg per pallet", "kgperpallet"],
    temperature: ["producttype", "product type", "product type rettificato"],
    active: ["blocked"],
    vendorName: ["vendorname", "vendor name"],
    vendorName2: ["vendorname2", "vendor name 2"],
  };

  // ✅ MAPBCVAL DEFINITA DENTRO IL COMPONENTE
  const mapBCVal = (field: string, raw: string) => {
    const maps: any = {
      category: {
        "food": "FOOD",
        "alimenti": "FOOD",
        "beverage": "WINE",
        "wine": "WINE",
        "spirits": "SPIRITS",
        "vino": "WINE",
        "meat": "MEAT",
        "carni": "MEAT",
        "salumi": "MEAT",
        "milk and dairy products": "FOOD",
        "cow cheese": "FOOD",
        "sheep cheese": "FOOD",
        "stretched-curd cheese": "FOOD",
        "pork meat": "MEAT",
        "ham": "MEAT",
        "other cured meats": "MEAT",
        "poultry and rabbit meat": "MEAT",
        "egg products": "FOOD",
        "eggs": "FOOD",
        "flour and groats": "FOOD",
        "preserved fish": "MEAT",
        "fish processing": "MEAT",
        "shellfish": "MEAT",
        "molluscs and mussels": "MEAT",
        "soft drinks": "FOOD",
        "oil and fats": "FOOD",
        "pasta and rice": "FOOD",
        "condiments": "FOOD",
        "meat processing": "MEAT"
      },
      uom: {
        "pcs": "PCS",
        "pz": "PCS",
        "piece": "PCS",
        "pezzi": "PCS",
        "box": "BOX",
        "ctn": "BOX",
        "cartone": "BOX",
        "collo": "BOX",
        "kg": "KG",
        "kgs": "KG",
        "kilogram": "KG"
      },
      temperature: {
        "dry": "DRY",
        "secco": "DRY",
        "ambient": "DRY",
        "amb": "DRY",
        "fresh": "FRESH",
        "fresco": "FRESH",
        "chilled": "FRESH",
        "refrigerated": "FRESH",
        "frozen": "FROZEN",
        "surgelato": "FROZEN",
        "congelato": "FROZEN"
      },
    };
    if (!maps[field]) return raw;
    return maps[field][String(raw || "").toLowerCase().trim()] || raw;
  };

  function autoMap(hdrs: string[]) {
    const m: any = {};
    for (const field of FIELDS) {
      const aliases = LOCAL_ALIASES[field] || [];
      const h = hdrs.find(h => aliases.some(a => h.toLowerCase().replace(/[\s_]/g, "") === a.replace(/[\s_]/g, "")));
      if (h) { m[field] = h; continue; }
      const h2 = hdrs.find(h => aliases.some(a => h.toLowerCase().replace(/[\s_]/g, "").includes(a.replace(/[\s_]/g, "")) && a.length > 3));
      if (h2) m[field] = h2;
    }
    return m;
  }

  function parseFile(file: File) {
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const wb = XLSX.read((ev.target as any).result, { type: "binary" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
        if (!data.length) { showToast("File vuoto", T.red); return; }

        let hi = 0;
        for (let i = 0; i < Math.min(5, data.length); i++) {
          const rNorm = data[i].map((c: any) => String(c || "").toLowerCase().replace(/[\s_]/g, ""));
          if (rNorm.some((c: string) => ["no", "ifbitem", "vendorname", "description"].includes(c))) { hi = i; break; }
        }
        const hdrs = data[hi].map((h: any) => String(h || "").trim());
        setHeaders(hdrs);
        setRawRows(data.slice(hi + 1).filter((r: any[]) => r.some((c: any) => c !== "")));
        setMap(autoMap(hdrs));
        setImportStep("map");
      } catch (err: any) { 
        showToast("Errore lettura file: " + err.message, T.red); 
      }
    };
    reader.readAsBinaryString(file);
  }

  function buildPreview() {
    const mapped = rawRows.map(r => {
      const obj: any = {};
      for (const field of FIELDS) {
        const col = map[field];
        if (col) {
          const idx = headers.indexOf(col);
          obj[field] = idx >= 0 ? String(r[idx] || "").trim() : "";
        } else obj[field] = "";
      }
      return obj;
    }).filter(r => r.code || r.nHK);
    setPreview(mapped);
    setImportStep("preview");
  }

  function executeImport() {
    const now = Date.now();
    const newProds = preview.map((r: any) => ({
      id: r.code || r.nHK,
      code: r.code,
      nHK: r.nHK,
      description: r.description,
      category: mapBCVal("category", r.category),
      uom: mapBCVal("uom", r.uom),
      qtyPerBox: parseFloat(r.qtyPerBox) || 0,
      boxPerPallet: parseFloat(r.boxPerPallet) || 0,
      kgPerBox: parseFloat(r.kgPerBox) || 0,
      temperature: mapBCVal("temperature", r.temperature),
      kgxplt: parseFloat(r.kgxplt) > 0 ? parseFloat(r.kgxplt) : roundN((parseFloat(r.kgPerBox) || 0) * (parseFloat(r.qtyPerBox) || 1) * (parseFloat(r.boxPerPallet) || 0)),
      active: !["true", "1", "yes"].includes(String(r.active || "").toLowerCase()),
      vendorName: r.vendorName || "",
      vendorName2: r.vendorName2 || "",
    }));
  
    setProducts(newProds);
    IDB.set(`ifb_products_${branch}`, newProds);
    IDB.set(`ifb_anag_data_${now}`, newProds);
    const log = { id:now, type:"anagrafica", date:new Date(now).toISOString(), count:newProds.length, branch:"ALL" };
    const newLogs = [log,...importLogs];
    setImportLogs(newLogs); LS.set("ifb_importlogs", newLogs);
    const newSnaps = [log,...snapshots].slice(0,50);
    setSnapshots(newSnaps); LS.set("ifb_snapshots", newSnaps);
  
    bumpImportTs();
    showToast(`Importati ${newProds.length} articoli`, T.gold);
    setImportStep("idle");
    setPreview([]);
    setRawRows([]);
  }

  async function loadFromSnapshot(snap: any) {
    const snapshotProducts = await IDB.get(`ifb_anag_data_${snap.id}`, []);
    if (!snapshotProducts?.length) {
      showToast(`Snapshot non disponibile — reimporta il file`, T.orange);
      return;
    }
    if (window.confirm(`Caricare l'anagrafica del ${new Date(snap.id).toLocaleDateString("it-IT")} (${snapshotProducts.length} articoli)?`)) {
      setProducts(snapshotProducts);
      await IDB.set(`ifb_products_${branch}`, snapshotProducts);
      showToast(`Anagrafica ripristinata: ${snapshotProducts.length} articoli ✓`, T.gold);
      setSearch(""); setOnlyIFB(true);
    }
  }

  const base = onlyIFB ? products.filter((p: any) => isIFBVendor(p.vendorName)) : products;
  const q = search.trim().toLowerCase();
  const filtered = !q ? base : base.filter((p: any) =>
    String(p.description||"").toLowerCase().includes(q) ||
    String(p.code||"").toLowerCase().includes(q) ||
    String(p.nHK||"").toLowerCase().includes(q) ||
    String(p.id||"").toLowerCase().includes(q)
  );

  return (
    <div>
      <PageHeader title="Anagrafica Articoli" sub={`${products.length} articoli · ${products.filter((p: any) => isIFBVendor(p.vendorName)).length} INALCA F&B`} />

      {/* Toolbar import */}
      <div style={{ display: "flex", gap: "10px", marginBottom: "16px", alignItems: "center", flexWrap: "wrap" }}>
        <label style={{ display: "inline-block", padding: "8px 16px", background: T.gold, color: "#000", borderRadius: "6px", cursor: "pointer", fontWeight: "bold", fontSize: "12px" }}>
          📂 Carica anagrafica (BC export)
          <input type="file" accept=".xlsx,.xls,.csv" onChange={e => { const f = e.target.files?.[0]; if (f) parseFile(f); e.target.value = ""; }} style={{ display: "none" }} />
        </label>

        {anagSnaps.length > 0 && (
          <select
            onChange={e => {
              if (e.target.value) {
                const snap = anagSnaps.find((s: any) => String(s.id) === e.target.value);
                if (snap) loadFromSnapshot(snap);
              }
              e.target.value = "";
            }}
            style={{ ...inputStyle(), width: "auto", fontSize: "12px" }}
            defaultValue=""
          >
            <option value="">📜 Carica da storico ({anagSnaps.length})</option>
            {anagSnaps.map((s: any) => (
              <option key={s.id} value={String(s.id)}>
                {new Date(s.id).toLocaleDateString("it-IT")} · {s.count} articoli
              </option>
            ))}
          </select>
        )}

        <div style={{ flex: 1 }} />
        
        <button
          onClick={() => {
            if (window.confirm(`Eliminare tutti i ${products.length} articoli dall'anagrafica?`)) {
              setProducts([]);
              LS.set(`ifb_products_${branch}`, []);
              bumpImportTs();
              showToast("Anagrafica svuotata", T.red);
            }
          }}
          style={{
            padding: "5px 12px",
            background: "none",
            border: `1px solid ${T.red}44`,
            borderRadius: "6px",
            color: T.red,
            cursor: "pointer",
            fontSize: "11px"
          }}
        >
          🗑 Svuota anagrafica ({products.length})
        </button>

        <button
          onClick={() => setOnlyIFB((v: boolean) => !v)}
          style={{
            padding: "5px 12px",
            background: onlyIFB ? `${T.gold}20` : T.surface,
            color: onlyIFB ? T.gold : T.muted,
            border: `1px solid ${onlyIFB ? T.gold : T.border}`,
            borderRadius: "6px",
            cursor: "pointer",
            fontSize: "11px"
          }}
        >
          {onlyIFB ? `✓ Solo IF&B (${base.length})` : `Mostra tutti (${products.length})`}
        </button>
      </div>

      {/* Step di import - Mappa */}
      {importStep === "map" && (
        <div style={{ background: T.card, border: `1px solid ${T.gold}`, borderRadius: "8px", padding: "16px", marginBottom: "16px" }}>
          <div style={{ color: T.gold, fontWeight: "bold", marginBottom: "12px" }}>Mappatura colonne · {fileName}</div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: "10px", marginBottom: "16px" }}>
            {FIELDS.slice(0, 9).map(f => (
              <div key={f}>
                <label style={{ fontSize: "10px", color: T.muted }}>{FLABELS[f]}</label>
                <select
                  value={map[f] || ""}
                  onChange={e => setMap((m: any) => ({ ...m, [f]: e.target.value }))}
                  style={{ ...inputStyle(), fontSize: "11px", padding: "4px 6px" }}
                >
                  <option value="">—</option>
                  {headers.map(h => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>
            ))}
          </div>
          <div style={{ display: "flex", gap: "10px" }}>
            <ActionBtn label="Annulla" onClick={() => setImportStep("idle")} />
            <ActionBtn label="Preview →" onClick={buildPreview} primary />
          </div>
        </div>
      )}

      {/* Step di import - Preview */}
      {importStep === "preview" && (
        <div style={{ background: T.card, border: `1px solid ${T.green}`, borderRadius: "8px", padding: "16px", marginBottom: "16px" }}>
          <div style={{ color: T.green, fontWeight: "bold", marginBottom: "12px" }}>Preview · {preview.length} articoli</div>
          <div style={{ maxHeight: "200px", overflow: "auto", marginBottom: "12px", fontSize: "11px" }}>
            <table style={{ width: "100%", borderCollapse: "collapse" }}>
              <thead>
                <tr>
                  <th style={{ textAlign: "left", padding: "4px" }}>N HK</th>
                  <th style={{ textAlign: "left", padding: "4px" }}>IFB No</th>
                  <th style={{ textAlign: "left", padding: "4px" }}>Descrizione</th>
                </tr>
              </thead>
              <tbody>
                {preview.slice(0, 20).map((r, i) => (
                  <tr key={i}>
                    <td style={{ padding: "2px 4px" }}>{r.nHK}</td>
                    <td style={{ padding: "2px 4px", color: T.gold }}>{r.code}</td>
                    <td style={{ padding: "2px 4px" }}>{r.description}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div style={{ display: "flex", gap: "10px" }}>
            <ActionBtn label="← Indietro" onClick={() => setImportStep("map")} />
            <ActionBtn label={`✓ Importa ${preview.length} articoli`} onClick={executeImport} primary />
          </div>
        </div>
      )}

      {/* Barra di ricerca */}
      <SearchBar value={search} onChange={setSearch} placeholder="🔍 Cerca articolo…" />

      {/* Tabella */}
      <Section title={`${filtered.length} articoli`}>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr>
                <th style={{ padding: "7px 12px", background: T.card, color: T.muted, textAlign: "left", borderBottom: `1px solid ${T.border}`, fontSize: "11px" }}>{branchN(branch)}</th>
                <th style={{ padding: "7px 12px", background: T.card, color: T.muted, textAlign: "left", borderBottom: `1px solid ${T.border}`, fontSize: "11px" }}>IFB No</th>
                <th style={{ padding: "7px 12px", background: T.card, color: T.muted, textAlign: "left", borderBottom: `1px solid ${T.border}`, fontSize: "11px" }}>Descrizione</th>
                <th style={{ padding: "7px 12px", background: T.card, color: T.muted, textAlign: "left", borderBottom: `1px solid ${T.border}`, fontSize: "11px" }}>Vendor</th>
                <th style={{ padding: "7px 12px", background: T.card, color: T.muted, textAlign: "left", borderBottom: `1px solid ${T.border}`, fontSize: "11px" }}>Categoria</th>
                <th style={{ padding: "7px 12px", background: T.card, color: T.muted, textAlign: "left", borderBottom: `1px solid ${T.border}`, fontSize: "11px" }}>UOM</th>
                <th style={{ padding: "7px 12px", background: T.card, color: T.muted, textAlign: "left", borderBottom: `1px solid ${T.border}`, fontSize: "11px" }}>Qty/Box</th>
                <th style={{ padding: "7px 12px", background: T.card, color: T.muted, textAlign: "left", borderBottom: `1px solid ${T.border}`, fontSize: "11px" }}>Box/Plt</th>
                <th style={{ padding: "7px 12px", background: T.card, color: T.muted, textAlign: "left", borderBottom: `1px solid ${T.border}`, fontSize: "11px" }}>Kg/Box</th>
                <th style={{ padding: "7px 12px", background: T.card, color: T.muted, textAlign: "left", borderBottom: `1px solid ${T.border}`, fontSize: "11px" }}>Kg/Plt</th>
                <th style={{ padding: "7px 12px", background: T.card, color: T.muted, textAlign: "left", borderBottom: `1px solid ${T.border}`, fontSize: "11px" }}>Temp</th>
                <th style={{ padding: "7px 12px", background: T.card, color: T.muted, textAlign: "left", borderBottom: `1px solid ${T.border}`, fontSize: "11px" }}>Attivo</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((p: any, i: number) => {
                const kgxplt = p.kgxplt || roundN((parseFloat(p.kgPerBox) || 0) * (parseFloat(p.boxPerPallet) || 0));
                return (
                  <tr key={p.id} style={{ borderBottom: `1px solid ${T.border}`, background: i % 2 === 0 ? T.bg : T.surface }}>
                    <td style={{ padding: "7px 12px", fontSize: "12px", fontFamily: "monospace" }}>
                      <span style={{ color: T.muted }}>{p.nHK || "—"}</span>
                    </td>
                    <td style={{ padding: "7px 12px", fontSize: "12px", fontFamily: "monospace" }}>
                      <span style={{ color: T.gold }}>{p.code}</span>
                    </td>
                    <td style={{ padding: "7px 12px", fontSize: "12px" }}>{p.description}</td>
                    <td style={{ padding: "7px 12px", fontSize: "12px" }}>
                      <span style={{ fontSize: "11px", color: isIFBVendor(p.vendorName) ? T.gold : T.muted }}>
                        {p.vendorName || "—"}
                      </span>
                    </td>
                    <td style={{ padding: "7px 12px", fontSize: "12px" }}>
                      <Chip label={p.category || "—"} color={p.category === "WINE" ? T.purple : p.category === "MEAT" ? T.red : T.blue} />
                    </td>
                    <td style={{ padding: "7px 12px", fontSize: "12px" }}>
                      <Chip label={p.uom || "—"} color={T.muted} />
                    </td>
                    <td style={{ padding: "7px 12px", fontSize: "12px", fontFamily: "monospace" }}>{p.qtyPerBox || "—"}</td>
                    <td style={{ padding: "7px 12px", fontSize: "12px", fontFamily: "monospace" }}>{p.boxPerPallet || "—"}</td>
                    <td style={{ padding: "7px 12px", fontSize: "12px", fontFamily: "monospace" }}>{p.kgPerBox || "—"}</td>
                    <td style={{ padding: "7px 12px", fontSize: "12px", fontFamily: "monospace" }}>
                      <span style={{ color: kgxplt > 0 ? T.text : T.dim }}>{kgxplt > 0 ? kgxplt : "—"}</span>
                    </td>
                    <td style={{ padding: "7px 12px", fontSize: "12px" }}>
                      <Chip label={p.temperature || "—"} color={p.temperature === "FROZEN" ? T.blue : p.temperature === "FRESH" ? T.green : T.muted} />
                    </td>
                    <td style={{ padding: "7px 12px", fontSize: "12px" }}>
                      <Chip label={p.active ? "Sì" : "No"} color={p.active ? T.green : T.red} />
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </Section>
    </div>
  );
}

// ─── MEAT PRICE LIST ──────────────────────────────────────────────────────────
function MeatPriceListPage({meatPrices,setMeatPrices,products,xrefs,importLogs,setImportLogs,snapshots,setSnapshots,showToast,bumpImportTs}) {
  const [step, setStep] = useState<"main"|"map"|"preview">("main");
  const [headers, setHeaders] = useState<string[]>([]);
  const [rawRows, setRawRows] = useState<any[]>([]);
  const [mapping, setMapping] = useState<any>({});
  const [preview, setPreview] = useState<any[]>([]);
  const [fileName, setFileName] = useState("");
  const [search, setSearch] = useState("");

  const meatSnaps = (importLogs||[]).filter((l:any) => l.type === "meatlist");

  function parsePrice(raw:any): number {
    if(raw == null) return 0;
    if(typeof raw === "number") return raw;
    const s = String(raw).replace(/[€\s]/g,"").replace(",",".");
    return parseFloat(s) || 0;
  }

  function parseFile(file:File) {
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const wb = XLSX.read((ev.target as any).result, {type:"binary"});
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data:any[][] = XLSX.utils.sheet_to_json(ws, {header:1, defval:""});
        if(data.length < 2) { showToast("File vuoto", T.red); return; }
        const hdrs = data[0].map((h:any) => String(h||"").trim());
        const rows = data.slice(1).filter((r:any[]) => r.some(c => c !== ""));
        setHeaders(hdrs);
        setRawRows(rows);
        // Auto-map
        const am:any = {};
        hdrs.forEach(h => {
          const hl = h.toLowerCase().replace(/[\s_()\/]/g,"");
          if(!am.code        && ["codice","code","no","no_","ifbn"].some(a=>hl===a||hl.includes(a))) am.code = h;
          if(!am.description && ["descrizione","description","desc"].some(a=>hl.includes(a))) am.description = h;
          if(!am.pricePerKg  && ["prezzo","price","€","eur","kg"].some(a=>hl.includes(a))) am.pricePerKg = h;
          if(!am.fonte       && hl.includes("fonte")) am.fonte = h;
          if(!am.foglio      && hl.includes("foglio")) am.foglio = h;
        });
        setMapping(am);
        setStep("map");
      } catch(err:any) { showToast("Errore: "+err.message, T.red); }
    };
    reader.readAsBinaryString(file);
  }

  function buildPreview() {
    const get = (row:any, field:string) => {
      const col = mapping[field]; if(!col) return "";
      const i = headers.indexOf(col); return i >= 0 ? row[i] : "";
    };
    const mapped = rawRows.map((row,idx) => {
      const code = String(get(row,"code")||"").trim();
      if(!code) return null;
      const pricePerKg = parsePrice(get(row,"pricePerKg"));
      const prod = findProduct(code, products, xrefs);
      return {
        _idx: idx,
        code,
        description: String(get(row,"description")||"").trim() || prod?.description || code,
        pricePerKg,
        fonte: String(get(row,"fonte")||"").trim(),
        foglio: String(get(row,"foglio")||"").trim(),
        _prodFound: !!prod,
        _prodCode: prod?.code || "",
        _prodNHK: prod?.nHK || "",
      };
    }).filter(Boolean);
    setPreview(mapped);
    setStep("preview");
  }

  function executeImport() {
    const now = Date.now();
    const entries = preview.map((r:any) => ({
      code: r.code,
      description: r.description,
      pricePerKg: r.pricePerKg,
      fonte: r.fonte,
      foglio: r.foglio,
    }));
    setMeatPrices(entries);
    LS.set("ifb_meatprices", entries);
    IDB.set(`ifb_meatprices_data_${now}`, entries);
    const log = {id:now, type:"meatlist", date:new Date(now).toISOString(), count:entries.length, diffs:[], branch:"ALL"};
    const newLogs = [log,...importLogs]; setImportLogs(newLogs); LS.set("ifb_importlogs",newLogs);
    const newSnaps = [log,...snapshots].slice(0,50); setSnapshots(newSnaps); LS.set("ifb_snapshots",newSnaps);
    bumpImportTs();
    showToast(`Listino carne: ${entries.length} prezzi importati ✓`, T.gold);
    setStep("main"); setPreview([]); setRawRows([]);
  }

  const _sq = search.toLowerCase();
  const displayed = meatPrices.filter((m:any) =>
    !search ||
    m.code?.toLowerCase().includes(_sq) ||
    m.description?.toLowerCase().includes(_sq) ||
    m.foglio?.toLowerCase().includes(_sq)
  );

  return (
    <div>
      <PageHeader title="🥩 Listino Carne" sub={`${meatPrices.length} prezzi · usato come fallback se l'articolo non è nei listini principali`}/>

      {/* Toolbar */}
      <div style={{display:"flex",gap:"10px",marginBottom:"16px",alignItems:"center",flexWrap:"wrap"}}>
        <label style={{display:"inline-block",padding:"8px 16px",background:T.gold,color:"#000",borderRadius:"6px",cursor:"pointer",fontWeight:"bold",fontSize:"12px"}}>
          📂 Carica listino carne
          <input type="file" accept=".xlsx,.xls,.csv" onChange={e=>{const f=e.target.files?.[0];if(f)parseFile(f);e.target.value="";}} style={{display:"none"}}/>
        </label>

        {meatSnaps.length > 0 && (
          <select onChange={async e=>{
            if(!e.target.value) return;
            const snap = importLogs.find((l:any)=>String(l.id)===e.target.value);
            if(!snap) return;
            if(window.confirm(`Ripristinare listino del ${new Date(snap.id).toLocaleDateString("it-IT")} (${snap.count} prezzi)?`)) {
              const data = await IDB.get(`ifb_meatprices_data_${snap.id}`, null);
              if(!data?.length){ showToast("Snapshot non disponibile — reimporta il file", T.orange); return; }
              setMeatPrices(data); LS.set("ifb_meatprices", data);
              showToast(`Listino carne: ${data.length} prezzi ripristinati ✓`, T.gold);
            }
            e.target.value="";
          }} style={{...inputStyle(),width:"auto",fontSize:"12px"}} defaultValue="">
            <option value="">📜 Storico ({meatSnaps.length})</option>
            {meatSnaps.map((s:any)=>(
              <option key={s.id} value={String(s.id)}>{new Date(s.id).toLocaleDateString("it-IT")} · {s.count} prezzi</option>
            ))}
          </select>
        )}

        {meatPrices.length > 0 && (
          <button onClick={()=>{if(window.confirm(`Eliminare tutti i ${meatPrices.length} prezzi del listino carne?`)){setMeatPrices([]);LS.set("ifb_meatprices",[]);}}}
            style={{padding:"8px 16px",background:"none",border:`1px solid ${T.red}44`,borderRadius:"6px",color:T.red,cursor:"pointer",fontSize:"12px"}}>
            🗑 Svuota ({meatPrices.length})
          </button>
        )}

        <span style={{fontSize:"11px",color:T.muted}}>Colonne attese: Codice · Descrizione · Prezzo (€/kg) · opzionali: Fonte, Foglio</span>
      </div>

      {/* Step: map */}
      {step === "map" && (
        <div style={{background:T.card,border:`1px solid ${T.gold}`,borderRadius:"8px",padding:"16px",marginBottom:"16px"}}>
          <div style={{color:T.gold,fontWeight:"bold",marginBottom:"12px"}}>Mappatura · {fileName} · {rawRows.length} righe</div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:"12px",marginBottom:"16px"}}>
            {([
              ["code","Codice IFB *",true],
              ["description","Descrizione",false],
              ["pricePerKg","Prezzo €/kg *",true],
              ["fonte","Fonte",false],
              ["foglio","Foglio / Categoria",false],
            ] as [string,string,boolean][]).map(([f,l,req])=>(
              <div key={f}>
                <label style={{display:"block",fontSize:"11px",color:req?T.gold:T.muted,marginBottom:"4px"}}>{l}</label>
                <select value={mapping[f]||""} onChange={e=>setMapping((m:any)=>({...m,[f]:e.target.value}))}
                  style={{...inputStyle(),cursor:"pointer",borderColor:req&&!mapping[f]?T.red+"88":T.border}}>
                  <option value="">— non mappato —</option>
                  {headers.map(h=><option key={h} value={h}>{h}</option>)}
                </select>
              </div>
            ))}
          </div>
          <div style={{display:"flex",gap:"10px"}}>
            <ActionBtn label="← Annulla" onClick={()=>setStep("main")}/>
            <ActionBtn label="Preview →" onClick={buildPreview} primary disabled={!mapping["code"]||!mapping["pricePerKg"]}/>
          </div>
        </div>
      )}

      {/* Step: preview */}
      {step === "preview" && (
        <div style={{background:T.card,border:`1px solid ${T.green}`,borderRadius:"8px",padding:"16px",marginBottom:"16px"}}>
          <div style={{color:T.green,fontWeight:"bold",marginBottom:"12px"}}>Preview · {preview.length} righe</div>
          <div style={{display:"flex",gap:"12px",marginBottom:"14px",flexWrap:"wrap"}}>
            {[
              [preview.filter((r:any)=>r._prodFound).length, "✅ Trovati in anagrafica", T.green],
              [preview.filter((r:any)=>!r._prodFound).length, "⚠ Non in anagrafica", T.orange],
              [preview.filter((r:any)=>r.pricePerKg>0).length, "💶 Con prezzo", T.gold],
            ].map(([n,l,c])=>(
              <div key={l as string} style={{padding:"8px 14px",background:T.surface,border:`1px solid ${c}44`,borderRadius:"6px"}}>
                <div style={{fontSize:"18px",fontWeight:"bold",color:c as string}}>{n as number}</div>
                <div style={{fontSize:"10px",color:T.dim}}>{l as string}</div>
              </div>
            ))}
          </div>
          <div style={{maxHeight:"200px",overflow:"auto",marginBottom:"14px"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:"11px"}}>
              <THead cols={["Codice","Descrizione","€/kg","Match anagrafica","Foglio"]} sticky/>
              <tbody>{preview.slice(0,30).map((r:any,i:number)=>(
                <tr key={i} style={{borderBottom:`1px solid ${T.border}`,background:r._prodFound?T.bg:`${T.orange}08`}}>
                  <td style={{padding:"4px 8px",fontFamily:"monospace",color:T.gold}}>{r.code}</td>
                  <td style={{padding:"4px 8px"}}>{r.description}</td>
                  <td style={{padding:"4px 8px",fontFamily:"monospace",color:r.pricePerKg>0?T.green:T.red}}>
                    {r.pricePerKg>0?`€ ${r.pricePerKg.toFixed(2)}`:"—"}
                  </td>
                  <td style={{padding:"4px 8px"}}>
                    {r._prodFound
                      ? <span style={{color:T.green,fontSize:"10px"}}>✓ {r._prodCode}</span>
                      : <span style={{color:T.orange,fontSize:"10px"}}>⚠ non trovato</span>}
                  </td>
                  <td style={{padding:"4px 8px",color:T.muted,fontSize:"10px"}}>{r.foglio||"—"}</td>
                </tr>
              ))}</tbody>
            </table>
          </div>
          <div style={{display:"flex",gap:"10px"}}>
            <ActionBtn label="← Torna" onClick={()=>setStep("map")}/>
            <ActionBtn label={`✓ Importa ${preview.length} prezzi`} onClick={executeImport} primary/>
          </div>
        </div>
      )}

      {/* Tabella listino */}
      {step === "main" && (
        <>
          <SearchBar value={search} onChange={setSearch} placeholder="🔍 Cerca codice, descrizione, foglio…"/>
          {meatPrices.length === 0 ? (
            <div style={{padding:"32px",textAlign:"center",color:T.dim,fontSize:"13px"}}>
              Nessun listino carne caricato. Clicca "Carica listino carne" per iniziare.
            </div>
          ) : (
            <Section title={`${displayed.length} prezzi`}>
              <div style={{overflowX:"auto"}}>
                <table style={{width:"100%",borderCollapse:"collapse"}}>
                  <THead cols={["Codice IFB","Descrizione","€/kg","Foglio","Fonte"]} sticky/>
                  <tbody>
                    {displayed.map((m:any,i:number)=>{
                      const prod = findProduct(m.code, products, xrefs);
                      return(
                        <tr key={i} style={{borderBottom:`1px solid ${T.border}`,background:i%2===0?T.bg:T.surface}}>
                          <td style={{padding:"7px 12px",fontSize:"12px",fontFamily:"monospace"}}>
                            <span style={{color:T.gold}}>{m.code}</span>
                            {prod && <span style={{marginLeft:"6px",fontSize:"9px",color:T.green}}>✓ {prod.code}</span>}
                          </td>
                          <td style={{padding:"7px 12px",fontSize:"12px"}}>{m.description}</td>
                          <td style={{padding:"7px 12px",fontSize:"12px",fontFamily:"monospace"}}>
                            <span style={{color:T.green,fontWeight:"bold"}}>€ {m.pricePerKg?.toFixed(2)||"—"}</span>
                          </td>
                          <td style={{padding:"7px 12px",fontSize:"12px"}}>
                            {m.foglio && <Chip label={m.foglio} color={T.blue}/>}
                          </td>
                          <td style={{padding:"7px 12px",fontSize:"11px",color:T.muted}}>{m.fonte||"—"}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </Section>
          )}
        </>
      )}
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
        onChange={async e=>{const f=e.target.files?.[0];if(f)onFile(f);e.target.value="";}} style={{display:"none"}}/>
    </div>
  );
}
function SearchBar({value,onChange,placeholder}){
  return<input value={value} onChange={e=>onChange(e.target.value)} placeholder={placeholder||"Cerca..."}
    style={{...inputStyle(),maxWidth:"320px",marginBottom:"14px"}}/>;
}
function THead({cols, sticky=false}: any) {
  return (
    <thead>
      <tr>
        {cols.map(c => (
          <th key={c} style={{
            padding: "7px 12px",
            background: T.card,
            color: T.muted,
            textAlign: "left",
            borderBottom: `1px solid ${T.border}`,
            fontSize: "11px",
            fontWeight: "normal",
            letterSpacing: "0.05em",
            whiteSpace: "nowrap",
            ...(sticky ? { position: "sticky" as const, top: 0, zIndex: 10 } : {})
          }}>
            {c}
          </th>
        ))}
      </tr>
    </thead>
  );
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