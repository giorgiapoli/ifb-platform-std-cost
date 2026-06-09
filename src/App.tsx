import { useState, useEffect, useMemo } from "react";
import * as XLSX from "xlsx";

const T = {
  bg:"#0B0F14", surface:"#111720", card:"#161E28",
  border:"rgba(255,255,255,0.07)", borderHi:"rgba(255,255,255,0.14)",
  text:"#E2D9CC", muted:"rgba(226,217,204,0.45)", dim:"rgba(226,217,204,0.22)",
  gold:"#C9A84C", goldDim:"rgba(201,168,76,0.18)",
  blue:"#4A8FB5", green:"#4BA87A", red:"#B5534A", orange:"#C47A3B", purple:"#7B5AC4",
};
const BRANCH_CFG = {
  HK:  { label:"Hong Kong", flag:"🇭🇰", color:T.gold,   currency:"HKD", defaultRate:9.1437 },
  CAN: { label:"Canarie",   flag:"🇮🇨", color:T.blue,   currency:"EUR", defaultRate:1      },
  AUS: { label:"Australia", flag:"🇦🇺", color:T.orange, currency:"AUD", defaultRate:1.6420 },
};
const NOW = () => new Date().toISOString().slice(0,7);
const FMT = (n,d=2) => typeof n==="number" ? n.toFixed(d) : "—";
const PCT = (a,b) => b ? ((a-b)/b*100) : null;

// ─── findProduct: cerca per IFB No, N HK, o via tabella XRef ────────────────
function findProduct(code, products, xrefs = []) {
  if (!code) return null;
  const c = String(code).trim();
  let p = products.find(pr => pr.code === c);
  if (p) return p;
  p = products.find(pr => pr.nHK && pr.nHK === c);
  if (p) return p;
  const xr = xrefs.find(x => x.nHK === c);
  if (xr) { p = products.find(pr => pr.code === xr.ifbNo); if (p) return p; }
  const xr2 = xrefs.find(x => x.ifbNo === c);
  if (xr2) { p = products.find(pr => pr.nHK === xr2.nHK); if (p) return p; }
  return null;
}

// ─────────────────────────────────────────────────────────────────────────────
const SEED_PRODUCTS = [
  {id:"P001",code:"BAL-001",description:"Balsamic Vinegar 500ml CASA EMILIA",category:"FOOD",uom:"PCS",qtyPerBox:12,boxPerPallet:70, kgPerBox:null,temperature:"DRY",  active:true},
  {id:"P002",code:"BAL-002",description:"Balsamic Cream 250ml",              category:"FOOD",uom:"PCS",qtyPerBox:12,boxPerPallet:130,kgPerBox:null,temperature:"DRY",  active:true},
  {id:"P003",code:"CHE-001",description:"Bavarian Emmental Cepparo 3kg",     category:"FOOD",uom:"KG", qtyPerBox:10,boxPerPallet:60, kgPerBox:10,  temperature:"FRESH",active:true},
  {id:"P004",code:"CHE-002",description:"Parmigiano Reggiano 24M 2kg",       category:"FOOD",uom:"KG", qtyPerBox:8, boxPerPallet:99, kgPerBox:8,   temperature:"FRESH",active:true},
  {id:"P005",code:"PAS-001",description:"Tortelli Ricotta & Spinach 3kg",    category:"FOOD",uom:"BOX",qtyPerBox:1, boxPerPallet:160,kgPerBox:null,temperature:"FROZEN",active:true},
  {id:"P006",code:"WIN-001",description:"Lambrusco Dell'Emilia IGT 75cl",    category:"WINE",uom:"PCS",qtyPerBox:6, boxPerPallet:80, kgPerBox:null,temperature:"FRESH",active:true},
  {id:"P007",code:"WIN-002",description:"Prosecco DOC Extra Dry 75cl",       category:"WINE",uom:"PCS",qtyPerBox:6, boxPerPallet:96, kgPerBox:null,temperature:"FRESH",active:true},
  {id:"P008",code:"MEA-001",description:"Mortadella BLU 3kg",                category:"MEAT",uom:"KG", qtyPerBox:6, boxPerPallet:72, kgPerBox:6,   temperature:"FRESH",active:true},
  {id:"P009",code:"MEA-002",description:"Spianata Calabra 2kg",              category:"MEAT",uom:"KG", qtyPerBox:4, boxPerPallet:150,kgPerBox:4,   temperature:"FRESH",active:true},
  {id:"P010",code:"FRZ-001",description:"Arancini Mozzarella & Tomato 3kg",  category:"FOOD",uom:"BOX",qtyPerBox:1, boxPerPallet:162,kgPerBox:null,temperature:"FROZEN",active:true},
];
const SEED_LOGISTIC = [
  {productId:"P001",branch:"HK",area:"NORD",ubicazione:"MTO",pltPerContainer:25,hasCert:false,hasAlcTax:false,alcTax:0,convFactor:1},
  {productId:"P002",branch:"HK",area:"NORD",ubicazione:"MTO",pltPerContainer:25,hasCert:false,hasAlcTax:false,alcTax:0,convFactor:1},
  {productId:"P003",branch:"HK",area:"NORD",ubicazione:"MTO",pltPerContainer:23,hasCert:true, hasAlcTax:false,alcTax:0,convFactor:1},
  {productId:"P004",branch:"HK",area:"NORD",ubicazione:"FOR",pltPerContainer:23,hasCert:true, hasAlcTax:false,alcTax:0,convFactor:1},
  {productId:"P005",branch:"HK",area:"NORD",ubicazione:"MTO",pltPerContainer:23,hasCert:false,hasAlcTax:false,alcTax:0,convFactor:1},
  {productId:"P006",branch:"HK",area:"NORD",ubicazione:"MTO",pltPerContainer:23,hasCert:false,hasAlcTax:true, alcTax:0.45,convFactor:1,carriage:60},
  {productId:"P007",branch:"HK",area:"NORD",ubicazione:"MTO",pltPerContainer:23,hasCert:false,hasAlcTax:true, alcTax:0.38,convFactor:1,carriage:60},
  {productId:"P008",branch:"HK",area:"NORD",ubicazione:"MTO",pltPerContainer:23,hasCert:true, hasAlcTax:false,alcTax:0,convFactor:1},
  {productId:"P009",branch:"HK",area:"NORD",ubicazione:"MTO",pltPerContainer:23,hasCert:true, hasAlcTax:false,alcTax:0,convFactor:1},
  {productId:"P010",branch:"HK",area:"NORD",ubicazione:"MTS",pltPerContainer:23,hasCert:false,hasAlcTax:false,alcTax:0,convFactor:1},
];
const SEED_PRICES = [
  {productId:"P001",branch:"HK",month:"2026-05",dapFinal:1.4941,mtsPrice:0,fcaDiscounted:0},
  {productId:"P002",branch:"HK",month:"2026-05",dapFinal:2.1200,mtsPrice:0,fcaDiscounted:0},
  {productId:"P003",branch:"HK",month:"2026-05",dapFinal:5.5556,mtsPrice:0,fcaDiscounted:0},
  {productId:"P004",branch:"HK",month:"2026-05",dapFinal:19.181,mtsPrice:0,fcaDiscounted:0},
  {productId:"P005",branch:"HK",month:"2026-05",dapFinal:19.000,mtsPrice:0,fcaDiscounted:0},
  {productId:"P006",branch:"HK",month:"2026-05",dapFinal:3.2000,mtsPrice:0,fcaDiscounted:0},
  {productId:"P007",branch:"HK",month:"2026-05",dapFinal:2.9500,mtsPrice:0,fcaDiscounted:0},
  {productId:"P008",branch:"HK",month:"2026-05",dapFinal:5.2828,mtsPrice:0,fcaDiscounted:0},
  {productId:"P009",branch:"HK",month:"2026-05",dapFinal:8.5172,mtsPrice:0,fcaDiscounted:0},
  {productId:"P010",branch:"HK",month:"2026-05",dapFinal:18.809,mtsPrice:0,fcaDiscounted:0},
  {productId:"P001",branch:"HK",month:"2026-06",dapFinal:1.5510,mtsPrice:0,fcaDiscounted:0},
  {productId:"P002",branch:"HK",month:"2026-06",dapFinal:2.1200,mtsPrice:0,fcaDiscounted:0},
  {productId:"P003",branch:"HK",month:"2026-06",dapFinal:5.8000,mtsPrice:0,fcaDiscounted:0},
  {productId:"P004",branch:"HK",month:"2026-06",dapFinal:19.181,mtsPrice:0,fcaDiscounted:0},
  {productId:"P005",branch:"HK",month:"2026-06",dapFinal:19.000,mtsPrice:0,fcaDiscounted:0},
  {productId:"P006",branch:"HK",month:"2026-06",dapFinal:3.0500,mtsPrice:0,fcaDiscounted:0},
  {productId:"P007",branch:"HK",month:"2026-06",dapFinal:2.9500,mtsPrice:0,fcaDiscounted:0},
  {productId:"P008",branch:"HK",month:"2026-06",dapFinal:5.2828,mtsPrice:0,fcaDiscounted:0},
  {productId:"P009",branch:"HK",month:"2026-06",dapFinal:8.9000,mtsPrice:0,fcaDiscounted:0},
  {productId:"P010",branch:"HK",month:"2026-06",dapFinal:18.809,mtsPrice:0,fcaDiscounted:0},
];
const SEED_FX = [
  {branch:"HK", month:"2026-05",rate:9.1200},
  {branch:"HK", month:"2026-06",rate:9.1437},
  {branch:"CAN",month:"2026-05",rate:1.0},
  {branch:"CAN",month:"2026-06",rate:1.0},
  {branch:"AUS",month:"2026-05",rate:1.6200},
  {branch:"AUS",month:"2026-06",rate:1.6420},
];

// ─────────────────────────────────────────────────────────────────────────────
// COST ENGINE
// ─────────────────────────────────────────────────────────────────────────────
const COSTS = {
  FOB: {
    DRY:   {NORD:2000,   CENTRO:0,    SUD:1108.55},
    FRESH: {NORD:3500,   CENTRO:3500, SUD:0},
    FROZEN:{NORD:4000,   CENTRO:0,    SUD:0},
  },
  LIC_HKD: 4100 + 3800,
  VGM: 100,
  HC:  80,
  PLT: 30,
  CARRIAGE_WINE: 60,
  VENDOR_CARRIAGE: {
    "WINE":60,"SPIRITS":60,
    "ALICO SRL":70,"ANTICO PASTIFICIO MORELLI SRL":80,"AZ. AGRICOLA MANCINI SRL AGRICOLA":70,
    "BONOMI SPA":30,"CAPURSO AZIENDA CASEARIA S.R.L.":90,"CECCHINI DARIO SRL":20,
    "CONSERVAS ANGELACHU S.L.":200,"DELIZIA 2000 SRL":75,"GRA-COM S.R.L.":90,
    "GREENS FOOD SPA":30,"INALCA S.P.A. A SOCIO UNICO":40,"ITALPIZZA S.R.L.":30,
    "OILALA' SRL":100,"QUANTOBASTA S.R.L.":140,"VALLE FINE FOODS ITALIA S.R.L.S.":150,
  },
  MTO:   {DRY:8.16,    FRESH:10.2,    FROZEN:12.24},
  MTS_D: {DRY:14.4228, FRESH:16.4832, FROZEN:24.7248},
  MTS_I: {DRY:2.5755,  FRESH:3.6057,  FROZEN:3.6057},
  MTS_P: {DRY:0.303,   FRESH:0.3434,  FROZEN:0.3535},
};

function calcHK({ priceInput, ubicazione, product, logistic, eurToHkd }) {
  const { uom, qtyPerBox, boxPerPallet, kgPerBox, temperature } = product;
  const { area, pltPerContainer, hasCert, hasAlcTax, alcTax, convFactor=1, carriage, vendorName="", category="" } = logistic;
  const carriagePlt = carriage > 0 ? carriage :
    (category==="WINE"||category==="SPIRITS") ? COSTS.VENDOR_CARRIAGE["WINE"] :
    COSTS.VENDOR_CARRIAGE[vendorName] || 0;
  const priceEur = (priceInput || 0) * convFactor;
  let unitsPerPlt;
  if      (uom==="BOX") unitsPerPlt = boxPerPallet;
  else if (uom==="KG")  unitsPerPlt = (kgPerBox||qtyPerBox) * boxPerPallet;
  else                  unitsPerPlt = qtyPerBox * boxPerPallet;
  const divisoreCollo = uom==="BOX" ? 1 : uom==="KG" ? (kgPerBox||qtyPerBox) : qtyPerBox;
  const totalUnits = unitsPerPlt * pltPerContainer;
  if (!totalUnits) return null;
  const fob = (COSTS.FOB[temperature]?.[area] ?? 0) / totalUnits;
  const lic = (COSTS.LIC_HKD / eurToHkd) / totalUnits;
  const vgm = COSTS.VGM / totalUnits;
  const hc  = hasCert ? COSTS.HC / totalUnits : 0;
  const plt = COSTS.PLT / unitsPerPlt;
  const alc = hasAlcTax ? (alcTax || 0) : 0;
  const carriageUnit = carriagePlt > 0 ? carriagePlt / unitsPerPlt : 0;
  const step1Eur = priceEur + fob + lic + vgm + hc + plt + alc + carriageUnit;
  let wh = 0, whDetail = {};
  if (ubicazione==="MTO") {
    wh = COSTS.MTO[temperature] / unitsPerPlt;
    whDetail = {type:"MTO", mto:wh};
  } else if (ubicazione==="MTS") {
    const dep     = COSTS.MTS_D[temperature] / unitsPerPlt;
    const inbound = COSTS.MTS_I[temperature] / unitsPerPlt;
    const picking = COSTS.MTS_P[temperature] / divisoreCollo;
    wh = dep + inbound + picking;
    whDetail = {type:"MTS", dep, inbound, picking};
  } else {
    whDetail = {type:"FOR"};
  }
  const step2Eur = step1Eur + wh;
  return {
    priceEur, fob, lic, vgm, hc, plt, alc, carriageUnit,
    step1Eur, step1Hkd: step1Eur * eurToHkd,
    wh, whDetail,
    step2Eur, step2Hkd: Math.round(step2Eur * eurToHkd * 100) / 100,
    rate: eurToHkd,
  };
}

// Replica formula Excel: FOR→col11(fcaDiscounted), MTO→col18(dapFinal), MTS→col8 se≠0 altrimenti col18
function selectPrice(priceRow, ubicazione) {
  if (!priceRow) return 0;
  if (ubicazione==="FOR")  return priceRow.fcaDiscounted || priceRow.dapFinal || 0;
  if (ubicazione==="MTO")  return priceRow.dapFinal || 0;
  if (ubicazione==="MTS") {
    const mts = priceRow.mtsPrice || 0;
    return mts !== 0 ? mts : (priceRow.dapFinal || 0);
  }
  return priceRow.dapFinal || 0;
}

// ─────────────────────────────────────────────────────────────────────────────
// BC COLUMN MAPPING
// ─────────────────────────────────────────────────────────────────────────────
const BC_FIELD_ALIASES = {
  nHK:          ["n hk","nhk","hk code","hk no","n_hk","codice hk","gc code","gc no","hong kong no"],
  code:         ["no_","no.","no","item no.","item no","ifb no","ifb n","codice","code"],
  description:  ["description","descrizione","desc","item description"],
  category:     ["section description","section desc","section","item category code","item category","category","categoria"],
  uom:          ["base unit of measure","uom","unit of measure","base uom","unit"],
  qtyPerBox:    ["quantity x packaging","units per parcel","qty per box","qty/box","pz per cartone"],
  boxPerPallet: ["packaging x pallet","parcels per pallet","box per pallet","cartoni per pallet"],
  kgPerBox:     ["quantity x packaging","kg per box","net weight","peso netto","kg per cartone"],
  temperature:  ["product type","product type rettificato","product type - anagrafica","item tracking code","temperatura","temperature","storage"],
  active:       ["blocked","bloccato","active","attivo"],
};
const BC_VALUE_MAP = {
  category: {
    "food":"FOOD","alimenti":"FOOD","f&b":"FOOD","beverage":"WINE","wine":"WINE","spirits":"WINE",
    "vino":"WINE","bevande":"WINE","meat":"MEAT","carni":"MEAT","fish":"MEAT","pesce":"MEAT",
    "salumi":"MEAT","beef":"MEAT","milk and dairy products":"FOOD","cow cheese":"FOOD",
    "sheep cheese":"FOOD","stretched-curd cheese":"FOOD","pork meat":"MEAT","ham":"MEAT",
    "other cured meats":"MEAT","poultry and rabbit meat":"MEAT","egg products":"FOOD",
    "eggs":"FOOD","flour and groats":"FOOD","preserved fish":"MEAT","fish processing":"MEAT",
    "shellfish":"MEAT","molluscs and mussels":"MEAT","pate":"MEAT",
    "soft drinks":"FOOD","oil and fats":"FOOD","pasta and rice":"FOOD","condiments":"FOOD",
    "meat processing":"MEAT",
  },
  uom: {
    "pcs":"PCS","pz":"PCS","pezzo":"PCS","pezzi":"PCS","piece":"PCS",
    "box":"BOX","ctn":"BOX","cartone":"BOX","collo":"BOX",
    "kg":"KG","kgs":"KG","kilogram":"KG",
  },
  temperature: {
    "dry":"DRY","secco":"DRY","ambient":"DRY","amb":"DRY",
    "fresh":"FRESH","fresco":"FRESH","chilled":"FRESH","refrigerated":"FRESH",
    "frozen":"FROZEN","surgelato":"FROZEN","congelato":"FROZEN",
  },
  active: {"true":false,"yes":false,"si":false,"1":false,"false":true,"no":true,"0":true,"":true},
};
function mapBCValue(field,raw) {
  if (!BC_VALUE_MAP[field]) return raw;
  const key = String(raw||"").toLowerCase().trim();
  return BC_VALUE_MAP[field][key] !== undefined ? BC_VALUE_MAP[field][key] : raw;
}
function detectBCColumn(headers,field) {
  const aliases = BC_FIELD_ALIASES[field]||[];
  for (const h of headers) {
    const hLow = h.toLowerCase().trim();
    if (aliases.some(a => hLow===a||hLow.includes(a))) return h;
  }
  return null;
}

const PRICE_FIELD_ALIASES = {
  code:           ["no_","no.","no","item no.","codice","code","n hk","ifb item","ifb no","ifb n"],
  vendorName:     ["vendor name 3","vendor name 3","vendor name","vendor","fornitore","vendor name 2"],
  section:        ["section description","section desc","section","sezione","categoria","category","sectiondescription"],
  mtsPrice:       ["mts price","mts","mts price (eur)"," mts price ","mts "],
  fcaPrice:       ["fca price","fca"," fca price "],
  fcaDiscount:    ["fca discount","fca disc","fca discount %"],
  fcaDiscounted:  ["fca discounted","fca disc.","fca final"," fca discounted "],
  dapPrice:       ["dap price","dap"," dap price "],
  dapDiscount:    ["dap discount","dap disc"],
  dapDiscounted:  ["dap discounted","dap final discounted"," dap discounted "],
  dapFinalDirect: ["dap final","dap final price","final price","prezzo acquisto"],
};

const FOR_VENDORS = new Set([
  "ALICO SRL","ANTICO PASTIFICIO MORELLI SRL","AZ. AGRICOLA MANCINI SRL AGRICOLA",
  "BONOMI SPA","CAPURSO AZIENDA CASEARIA S.R.L.","CECCHINI DARIO SRL",
  "CONSERVAS ANGELACHU S.L.","DELIZIA 2000 SRL","GRA-COM S.R.L.",
  "GREENS FOOD SPA","INALCA S.P.A. A SOCIO UNICO","ITALPIZZA S.R.L.",
  "OILALA' SRL","QUANTOBASTA S.R.L.","VALLE FINE FOODS ITALIA S.R.L.S",
]);

function calcDAPFinal({ dapDiscounted, fcaPrice, fcaDiscounted, mtsPrice, vendorName, section, products, code }) {
  const prod = products.find(p => p.code === code);
  let unitsPerPlt = 1;
  if (prod) {
    const { uom, qtyPerBox, boxPerPallet, kgPerBox } = prod;
    if (uom === "BOX") unitsPerPlt = boxPerPallet;
    else if (uom === "KG") unitsPerPlt = (kgPerBox || qtyPerBox) * boxPerPallet;
    else unitsPerPlt = qtyPerBox * boxPerPallet;
  }
  const sec = (section || "").toUpperCase();
  const isWineSpirits = sec === "WINE" || sec === "SPIRITS";
  const isX = isWineSpirits || FOR_VENDORS.has(vendorName || "");
  const pltCost = isWineSpirits ? 60 : (COSTS.VENDOR_CARRIAGE[vendorName] || 0);
  const carriageUnit = unitsPerPlt > 0 ? pltCost / unitsPerPlt : 0;
  const dapDisc = dapDiscounted || 0;
  const fcaP   = fcaPrice      || 0;
  const fcaD   = fcaDiscounted || 0;
  if (dapDisc !== 0) return { dapFinal: dapDisc, carriageUnit, note: "DAP Disc." };
  if (!isX)          return { dapFinal: 0,        carriageUnit: 0, note: "non-X" };
  if (isWineSpirits) return { dapFinal: fcaP !== 0 ? fcaP + carriageUnit : 0, carriageUnit, note: "Wine FCA+C" };
  return               { dapFinal: fcaD !== 0 ? fcaD + carriageUnit : 0, carriageUnit, note: "FCA Disc+C" };
}

// ─────────────────────────────────────────────────────────────────────────────
const LS = {
  get:(k,def)=>{try{const v=localStorage.getItem(k);return v?JSON.parse(v):def;}catch{return def;}},
  set:(k,v)=>{try{localStorage.setItem(k,JSON.stringify(v));}catch{}},
};

// ─────────────────────────────────────────────────────────────────────────────
// ROOT APP
// ─────────────────────────────────────────────────────────────────────────────
export default function App() {
  const [products,   setProducts]   = useState(()=>LS.get("ifb_products",   SEED_PRODUCTS));
  const [logistics,  setLogistics]  = useState(()=>LS.get("ifb_logistics",  SEED_LOGISTIC));
  const [prices,     setPrices]     = useState(()=>LS.get("ifb_prices",     SEED_PRICES));
  const [fx,         setFx]         = useState(()=>LS.get("ifb_fx",         SEED_FX));
  const [xrefs,      setXrefs]      = useState(()=>LS.get("ifb_xrefs",      []));
  const [salesRows,  setSalesRows]  = useState(()=>LS.get("ifb_sales_invoice", []));
  const [sentMails,  setSentMails]  = useState(()=>LS.get("ifb_mails",      []));
  const [importLogs, setImportLogs] = useState(()=>LS.get("ifb_importlogs", []));
  const [snapshots,  setSnapshots]  = useState(()=>LS.get("ifb_snapshots",  []));
  const [page,   setPage]   = useState("dashboard");
  const [branch, setBranch] = useState("HK");
  const [month,  setMonth]  = useState(NOW());
  const [toast,  setToast]  = useState(null);

  useEffect(()=>{LS.set("ifb_products",  products);  },[products]);
  useEffect(()=>{LS.set("ifb_logistics", logistics); },[logistics]);
  useEffect(()=>{LS.set("ifb_prices",    prices);    },[prices]);
  useEffect(()=>{LS.set("ifb_fx",        fx);        },[fx]);
  useEffect(()=>{LS.set("ifb_xrefs",     xrefs);     },[xrefs]);
  useEffect(()=>{LS.set("ifb_sales_invoice", salesRows); },[salesRows]);
  useEffect(()=>{LS.set("ifb_mails",     sentMails); },[sentMails]);
  useEffect(()=>{LS.set("ifb_importlogs",importLogs);},[importLogs]);
  useEffect(()=>{LS.set("ifb_snapshots", snapshots); },[snapshots]);

  const showToast=(msg,color=T.green)=>{setToast({msg,color});setTimeout(()=>setToast(null),3500);};

  const costRows = useMemo(()=>{
    const fxRate = fx.find(f=>f.branch===branch&&f.month===month)?.rate||BRANCH_CFG[branch].defaultRate;
    const prevMonth = month.slice(0,4)+"-"+String(parseInt(month.slice(5))-1).padStart(2,"0");
    return products.filter(p=>p.active).map(prod=>{
      const log    = logistics.find(l=>l.productId===prod.id&&l.branch===branch);
      const pr     = prices.find(p=>p.productId===prod.id&&p.branch===branch&&p.month===month);
      const prPrev = prices.find(p=>p.productId===prod.id&&p.branch===branch&&p.month===prevMonth);
      if (!log||!pr) return {...prod,cost:null,prevCost:null,log};
      const ubicazione = log.ubicazione;
      const priceInput = selectPrice(pr, ubicazione);
      const prevInput  = prPrev ? selectPrice(prPrev, ubicazione) : null;
      const cost     = calcHK({priceInput,      ubicazione,product:prod,logistic:{...log,category:prod.category},eurToHkd:fxRate});
      const prevCost = prevInput!=null ? calcHK({priceInput:prevInput,ubicazione,product:prod,logistic:{...log,category:prod.category},eurToHkd:fxRate}) : null;
      const delta    = prevCost ? PCT(cost.step2Eur,prevCost.step2Eur) : null;
      return {...prod,cost,prevCost,delta,priceInput,isNew:!prPrev,flagged:delta!==null&&Math.abs(delta)>=3,ubicazione};
    });
  },[products,logistics,prices,fx,branch,month]);

  const branchCfg = BRANCH_CFG[branch];
  const NAV = [
    {id:"dashboard",  icon:"⬡", label:"Dashboard"},
    {id:"products",   icon:"◈", label:"Anagrafica"},
    {id:"importAnag", icon:"⇪", label:"Import Anagrafica", badge:"BC"},
    {id:"xref",       icon:"⇄", label:"XRef N HK / IFB"},
    {id:"logistics",  icon:"◎", label:"Logistica"},
    {id:"prices",     icon:"◉", label:"Listini"},
    {id:"importPrice",icon:"💶", label:"Import Listini",    badge:"BC"},
    {id:"fx",         icon:"◌", label:"Cambi"},
    {id:"costs",      icon:"◆", label:"Standard Cost"},
    {id:"sales",      icon:"📋", label:"Sales Invoice"},
    {id:"storico",    icon:"⧖", label:"Storico & Diff"},
    {id:"mail",       icon:"◻", label:"Mail Mensile"},
    {id:"notes",      icon:"📝", label:"Note & Ambiguità"},
  ];
  const pages = {
    dashboard:   <Dashboard    costRows={costRows} branch={branch} month={month} branchCfg={branchCfg} setPage={setPage} />,
    products:    <Products     products={products} setProducts={setProducts} showToast={showToast} />,
    importAnag:  <ImportBC     products={products} setProducts={setProducts} importLogs={importLogs} setImportLogs={setImportLogs} snapshots={snapshots} setSnapshots={setSnapshots} showToast={showToast} branch={branch} />,
    xref:        <XRefPage     xrefs={xrefs} setXrefs={setXrefs} snapshots={snapshots} setSnapshots={setSnapshots} importLogs={importLogs} setImportLogs={setImportLogs} showToast={showToast} />,
    logistics:   <Logistics    logistics={logistics} setLogistics={setLogistics} products={products} branch={branch} showToast={showToast} />,
    prices:      <Prices       prices={prices} setPrices={setPrices} products={products} branch={branch} month={month} showToast={showToast} />,
    importPrice: <ImportPrices prices={prices} setPrices={setPrices} products={products} xrefs={xrefs} branch={branch} month={month} importLogs={importLogs} setImportLogs={setImportLogs} snapshots={snapshots} setSnapshots={setSnapshots} showToast={showToast} />,
    fx:          <FxRates      fx={fx} setFx={setFx} branch={branch} month={month} showToast={showToast} />,
    costs:       <CostTable    costRows={costRows} branch={branch} month={month} branchCfg={branchCfg} />,
    sales:       <SalesInvoice products={products} prices={prices} branch={branch} rows={salesRows} setRows={setSalesRows} snapshots={snapshots} setSnapshots={setSnapshots} importLogs={importLogs} setImportLogs={setImportLogs} showToast={showToast} />,
    storico:     <Storico      snapshots={snapshots} setSnapshots={setSnapshots} />,
    mail:        <MailGen      costRows={costRows} branch={branch} month={month} branchCfg={branchCfg} sentMails={sentMails} setSentMails={setSentMails} showToast={showToast} />,
    notes:       <NotesPage />,
  };

  return (
    <div style={{display:"flex",height:"100vh",width:"100vw",background:T.bg,fontFamily:"'Palatino Linotype','Book Antiqua',Palatino,serif",color:T.text,overflow:"hidden"}}>
      <div style={{width:"200px",flexShrink:0,background:T.surface,borderRight:`1px solid ${T.border}`,display:"flex",flexDirection:"column",height:"100vh",overflow:"hidden"}}>
        <div style={{padding:"18px 16px 14px",borderBottom:`1px solid ${T.border}`}}>
          <div style={{fontSize:"9px",letterSpacing:"3px",color:T.gold,textTransform:"uppercase",marginBottom:"3px"}}>IFB Platform</div>
          <div style={{fontSize:"14px",fontWeight:"bold",lineHeight:1.2}}>Cost Intelligence</div>
        </div>
        <div style={{padding:"10px 12px",borderBottom:`1px solid ${T.border}`}}>
          <div style={{fontSize:"9px",letterSpacing:"2px",color:T.dim,textTransform:"uppercase",marginBottom:"6px"}}>Filiale</div>
          <div style={{display:"flex",flexDirection:"column",gap:"3px"}}>
            {Object.entries(BRANCH_CFG).map(([key,cfg])=>(
              <button key={key} onClick={()=>setBranch(key)} style={{padding:"5px 8px",background:branch===key?`${cfg.color}20`:"transparent",border:`1px solid ${branch===key?cfg.color:"transparent"}`,borderRadius:"6px",color:branch===key?cfg.color:T.muted,cursor:"pointer",fontFamily:"inherit",fontSize:"11px",textAlign:"left",display:"flex",alignItems:"center",gap:"6px",transition:"all 0.2s"}}>
                <span>{cfg.flag}</span>{cfg.label}
                {key==="AUS"&&<span style={{fontSize:"7px",color:T.orange,marginLeft:"auto",background:`${T.orange}22`,padding:"1px 4px",borderRadius:"3px"}}>SOON</span>}
              </button>
            ))}
          </div>
        </div>
        <div style={{padding:"10px 12px",borderBottom:`1px solid ${T.border}`}}>
          <div style={{fontSize:"9px",letterSpacing:"2px",color:T.dim,textTransform:"uppercase",marginBottom:"5px"}}>Mese</div>
          <input type="month" value={month} onChange={e=>setMonth(e.target.value)} style={{width:"100%",padding:"5px 7px",background:"rgba(255,255,255,0.05)",border:`1px solid ${T.border}`,borderRadius:"6px",color:T.text,fontFamily:"inherit",fontSize:"11px",outline:"none",boxSizing:"border-box"}} />
        </div>
        <nav style={{flex:1,padding:"8px",display:"flex",flexDirection:"column",gap:"1px",overflowY:"auto"}}>
          {NAV.map(n=>(
            <button key={n.id} onClick={()=>setPage(n.id)} style={{padding:"7px 10px",background:page===n.id?T.goldDim:"transparent",border:`1px solid ${page===n.id?T.gold+"44":"transparent"}`,borderRadius:"6px",color:page===n.id?T.gold:T.muted,cursor:"pointer",fontFamily:"inherit",fontSize:"11px",textAlign:"left",display:"flex",alignItems:"center",gap:"7px",transition:"all 0.18s"}}>
              <span style={{fontSize:"10px",opacity:0.8}}>{n.icon}</span>{n.label}
              {n.badge&&<span style={{marginLeft:"auto",fontSize:"7px",background:`${T.blue}33`,color:T.blue,padding:"1px 4px",borderRadius:"3px"}}>{n.badge}</span>}
            </button>
          ))}
        </nav>
        <div style={{padding:"10px 12px",borderTop:`1px solid ${T.border}`,fontSize:"9px",color:T.dim,lineHeight:1.5}}>
          Inalca Food & Beverage<br/>© 2026 · v3.4
        </div>
      </div>
      <div style={{flex:1,display:"flex",flexDirection:"column",minWidth:0,overflow:"hidden"}}>
        <div style={{padding:"10px 24px",borderBottom:`1px solid ${T.border}`,background:T.surface,display:"flex",alignItems:"center",gap:"10px",flexShrink:0,zIndex:10}}>
          <span style={{fontSize:"16px"}}>{branchCfg.flag}</span>
          <span style={{fontSize:"13px",fontWeight:"bold"}}>{branchCfg.label}</span>
          <span style={{color:T.dim}}>·</span>
          <span style={{fontSize:"12px",color:T.muted}}>{NAV.find(n=>n.id===page)?.label}</span>
          <span style={{color:T.dim}}>·</span>
          <span style={{fontSize:"11px",color:T.gold}}>{month}</span>
          <div style={{marginLeft:"auto",display:"flex",gap:"6px"}}>
            <button onClick={()=>setPage("importAnag")} style={{padding:"5px 12px",background:`${T.blue}15`,border:`1px solid ${T.blue}44`,borderRadius:"5px",color:T.blue,cursor:"pointer",fontFamily:"inherit",fontSize:"10px"}}>⇪ Anagrafica</button>
            <button onClick={()=>setPage("importPrice")} style={{padding:"5px 12px",background:`${T.purple}15`,border:`1px solid ${T.purple}44`,borderRadius:"5px",color:T.purple,cursor:"pointer",fontFamily:"inherit",fontSize:"10px"}}>💶 Listini</button>
            <button onClick={()=>setPage("costs")} style={{padding:"5px 12px",background:"rgba(255,255,255,0.05)",border:`1px solid ${T.border}`,borderRadius:"5px",color:T.muted,cursor:"pointer",fontFamily:"inherit",fontSize:"10px"}}>◆ Costi</button>
            <button onClick={()=>setPage("mail")} style={{padding:"5px 12px",background:T.gold,border:`1px solid ${T.gold}`,borderRadius:"5px",color:T.bg,cursor:"pointer",fontFamily:"inherit",fontSize:"10px",fontWeight:"bold"}}>✉ Mail</button>
          </div>
        </div>
        <div style={{flex:1,padding:"20px 28px",overflow:"auto"}}>{pages[page]}</div>
      </div>
      {toast&&<div style={{position:"fixed",bottom:"24px",right:"24px",padding:"10px 18px",background:toast.color,borderRadius:"8px",color:"#fff",fontSize:"12px",fontWeight:"bold",boxShadow:"0 8px 24px rgba(0,0,0,0.4)",zIndex:1000}}>{toast.msg}</div>}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// XREF PAGE — mappa N HK ↔ IFB N
// ─────────────────────────────────────────────────────────────────────────────
function XRefPage({ xrefs, setXrefs, snapshots, setSnapshots, importLogs, setImportLogs, showToast }) {
  const [step, setStep] = useState("main");
  const [rawRows, setRawRows] = useState([]);
  const [headers, setHeaders] = useState([]);
  const [fileName, setFileName] = useState("");
  const [colNHK, setColNHK] = useState("");
  const [colIFB, setColIFB] = useState("");
  const [preview, setPreview] = useState([]);
  const [search, setSearch] = useState("");

  const parseFile = (file) => {
    if (!file) return;
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, {type:"binary"});
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(ws, {header:1, defval:""});
        if (data.length < 2) { showToast("File vuoto", T.red); return; }
        const hdrs = data[0].map(h => String(h).trim()).filter(h => h);
        const rows = data.slice(1).filter(r => r.some(c => c !== ""));
        setHeaders(hdrs);
        setRawRows(rows);
        // Auto-detect
        const nhkCandidates = ["n hk","nhk","hk","n_hk","gc code","gc no","hk code","hk no","hong kong"];
        const ifbCandidates = ["ifb n","ifb no","ifb no.","no_","no.","no","codice","code","item no"];
        const autoNHK = hdrs.find(h => nhkCandidates.some(a => h.toLowerCase().includes(a))) || "";
        const autoIFB = hdrs.find(h => ifbCandidates.some(a => h.toLowerCase() === a || h.toLowerCase().includes(a))) || "";
        setColNHK(autoNHK);
        setColIFB(autoIFB);
        setStep("map");
      } catch(err) { showToast("Errore: " + err.message, T.red); }
    };
    reader.readAsBinaryString(file);
  };

  const buildPreview = () => {
    const iNHK = headers.indexOf(colNHK);
    const iIFB = headers.indexOf(colIFB);
    const mapped = rawRows.map((row, idx) => {
      const nHK = String(row[iNHK] || "").trim();
      const ifbNo = String(row[iIFB] || "").trim();
      if (!nHK && !ifbNo) return null;
      const existing = xrefs.find(x => x.nHK === nHK);
      return {
        _idx: idx, nHK, ifbNo,
        _isNew: !existing,
        _changed: existing && existing.ifbNo !== ifbNo,
        _oldIFB: existing?.ifbNo,
      };
    }).filter(Boolean);
    setPreview(mapped);
    setStep("preview");
  };

  const executeImport = () => {
    const snapshotId = Date.now();
    const diffs = [];
    const incoming = preview.filter(r => r.nHK && r.ifbNo);
    incoming.forEach(r => {
      if (r._isNew) diffs.push({ nHK: r.nHK, ifbNo: r.ifbNo, isNew: true });
      else if (r._changed) diffs.push({ nHK: r.nHK, ifbNo: r.ifbNo, oldIFB: r._oldIFB, isNew: false, changed: true });
    });
    // Merge: replace same nHK
    const kept = xrefs.filter(x => !incoming.find(i => i.nHK === x.nHK));
    setXrefs([...incoming.map(r => ({nHK: r.nHK, ifbNo: r.ifbNo})), ...kept]);
    const logEntry = {
      id: snapshotId, type:"xref", fileName, importedAt: new Date().toISOString(),
      count: incoming.length, newCount: diffs.filter(d=>d.isNew).length,
      updateCount: diffs.filter(d=>d.changed).length, diffs,
    };
    setImportLogs(logs => [logEntry, ...logs]);
    setSnapshots(snaps => [logEntry, ...snaps].slice(0, 50));
    showToast(`XRef aggiornata: ${incoming.length} voci · ${diffs.filter(d=>d.isNew).length} nuove ✓`, T.gold);
    setStep("main");
    setPreview([]); setRawRows([]); setHeaders([]);
  };

  const displayed = xrefs.filter(x =>
    !search || x.nHK?.toLowerCase().includes(search.toLowerCase()) || x.ifbNo?.toLowerCase().includes(search.toLowerCase())
  );

  return (
    <div>
      <PageHeader title="⇄ Cross Reference N HK ↔ IFB N" sub="Tabella di corrispondenza tra codici Hong Kong e codici IFB · usata per il matching automatico dei listini" />

      {step === "map" && (
        <Section title={`Mappatura colonne — ${fileName} · ${rawRows.length} righe`} mb="20px">
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"16px",marginBottom:"16px"}}>
            <div>
              <label style={{display:"block",fontSize:"11px",color:T.gold,marginBottom:"5px"}}>Colonna N HK *</label>
              <select value={colNHK} onChange={e=>setColNHK(e.target.value)} style={{...inputStyle(),cursor:"pointer",borderColor:!colNHK?T.red+"88":T.border}}>
                <option value="">— seleziona —</option>
                {headers.map(h=><option key={h} value={h}>{h}</option>)}
              </select>
            </div>
            <div>
              <label style={{display:"block",fontSize:"11px",color:T.gold,marginBottom:"5px"}}>Colonna IFB N *</label>
              <select value={colIFB} onChange={e=>setColIFB(e.target.value)} style={{...inputStyle(),cursor:"pointer",borderColor:!colIFB?T.red+"88":T.border}}>
                <option value="">— seleziona —</option>
                {headers.map(h=><option key={h} value={h}>{h}</option>)}
              </select>
            </div>
          </div>
          <div style={{display:"flex",gap:"10px"}}>
            <ActionBtn label="← Ricarica" onClick={()=>setStep("main")} />
            <ActionBtn label="Preview →" onClick={buildPreview} primary disabled={!colNHK||!colIFB} />
          </div>
        </Section>
      )}

      {step === "preview" && (
        <div>
          <div style={{display:"flex",gap:"12px",marginBottom:"20px"}}>
            {[[preview.filter(r=>r._isNew).length,"Nuove",T.gold],[preview.filter(r=>r._changed).length,"Modificate",T.orange],[preview.filter(r=>!r._isNew&&!r._changed).length,"Invariate",T.dim],[preview.length,"Totale",T.text]].map(([n,l,c])=>(
              <div key={l} style={{padding:"10px 16px",background:T.card,border:`1px solid ${T.border}`,borderRadius:"8px"}}>
                <div style={{fontSize:"18px",fontWeight:"bold",color:c}}>{n}</div>
                <div style={{fontSize:"10px",color:T.dim,marginTop:"2px"}}>{l}</div>
              </div>
            ))}
          </div>
          <Section title="Preview (prime 50)">
            <table style={{width:"100%",borderCollapse:"collapse"}}>
              <THead cols={["N HK","IFB N","Stato"]} />
              <tbody>
                {preview.slice(0,50).map(r=>(
                  <tr key={r._idx} style={{borderBottom:`1px solid ${T.border}`,background:r._isNew?`${T.gold}07`:r._changed?`${T.orange}07`:""}}>
                    <TD mono><span style={{color:T.gold}}>{r.nHK}</span></TD>
                    <TD mono>{r.ifbNo}</TD>
                    <TD>
                      {r._isNew && <Chip label="NUOVO" color={T.gold} />}
                      {r._changed && <><Chip label="MODIF." color={T.orange} /><span style={{fontSize:"10px",color:T.dim,marginLeft:"6px"}}>{r._oldIFB} → {r.ifbNo}</span></>}
                      {!r._isNew && !r._changed && <span style={{color:T.dim,fontSize:"11px"}}>=</span>}
                    </TD>
                  </tr>
                ))}
              </tbody>
            </table>
          </Section>
          <div style={{display:"flex",gap:"10px",marginTop:"16px"}}>
            <ActionBtn label="← Torna" onClick={()=>setStep("map")} />
            <ActionBtn label={`✓ Aggiorna XRef (${preview.filter(r=>r.nHK&&r.ifbNo).length} voci)`} onClick={executeImport} primary />
          </div>
        </div>
      )}

      {step === "main" && (
        <>
          <div style={{display:"flex",gap:"10px",marginBottom:"16px",alignItems:"center"}}>
            <div style={{border:`2px dashed ${T.borderHi}`,borderRadius:"10px",padding:"20px 28px",textAlign:"center",cursor:"pointer"}} onClick={()=>document.getElementById("_xref_input").click()}>
              <div style={{fontSize:"24px",marginBottom:"6px"}}>⇄</div>
              <div style={{fontSize:"13px",color:T.text,marginBottom:"4px"}}>Carica file XRef (Excel/CSV)</div>
              <div style={{fontSize:"11px",color:T.muted}}>Due colonne: N HK · IFB N</div>
              <input id="_xref_input" type="file" accept=".xlsx,.xls,.csv" onChange={e=>{const f=e.target.files[0];if(f)parseFile(f);e.target.value="";}} style={{display:"none"}} />
            </div>
            <div style={{fontSize:"12px",color:T.muted,maxWidth:"300px",lineHeight:"1.6"}}>
              Carica il file PBI con le corrispondenze N HK ↔ IFB N.<br/>
              Questa tabella viene usata automaticamente nell'import listini e nell'anagrafica per risolvere i codici.
            </div>
          </div>
          <SearchBar value={search} onChange={setSearch} placeholder="🔍  Cerca per N HK o IFB N…" />
          {xrefs.length === 0 ? (
            <Section title="Nessuna XRef caricata">
              <div style={{padding:"32px",textAlign:"center",color:T.dim,fontSize:"13px"}}>Carica un file con N HK e IFB N per abilitare il matching automatico.</div>
            </Section>
          ) : (
            <Section title={`${displayed.length} / ${xrefs.length} corrispondenze`}>
              <table style={{width:"100%",borderCollapse:"collapse"}}>
                <THead cols={["N HK","IFB N","Azioni"]} />
                <tbody>
                  {displayed.map((x,i)=>(
                    <tr key={x.nHK+i} style={{borderBottom:`1px solid ${T.border}`}}>
                      <TD mono><span style={{color:T.gold}}>{x.nHK}</span></TD>
                      <TD mono>{x.ifbNo}</TD>
                      <TD><MiniBtn label="✕" onClick={()=>setXrefs(xs=>xs.filter((_,j)=>xs.indexOf(x)!==xs.indexOf(x)||j!==xrefs.indexOf(x)))} color={T.red} /></TD>
                    </tr>
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

// ─────────────────────────────────────────────────────────────────────────────
// NOTES PAGE
// ─────────────────────────────────────────────────────────────────────────────
function NotesPage() {
  const sections = [
    {
      title: "🔴 Problemi noti / Da risolvere",
      color: T.red,
      items: [
        "Import Listini: il codice nel file PBI deve essere N HK o IFB N — usa la pagina XRef per mappare i codici se il matching non funziona automaticamente.",
        "Colonna 'Prezzo €' in Standard Cost: mostra '—' se il prodotto non ha un prezzo caricato per il mese/filiale corrente. Verifica Import Listini.",
        "Canarie (CAN): i parametri di costo (FOB, LIC, magazzino) non sono ancora definiti. La filiale CAN non calcola Standard Cost.",
        "Australia (AUS): marcata come 'SOON', i parametri di costo non sono implementati.",
        "convFactor: nei seed data è sempre 1. Se un fornitore usa un'unità di misura diversa (es. cassa da 6 vs pezzo), aggiornare manualmente in Logistica.",
      ]
    },
    {
      title: "🟡 Ambiguità / Chiarimenti necessari",
      color: T.orange,
      items: [
        "Colonna 'kgPerBox': usata solo per prodotti con UOM=KG. È il peso netto per cartone, coincide con 'Quantity x Packaging' in BC? → Confermare mapping.",
        "Formula DAP Final: replica =SE(N≠0,N, SE(O='X', SE(WINE, I+P, K+P), 0)). La colonna 'O' (isX) dipende dall'elenco FOR_VENDORS hardcodato nel codice — aggiornare se cambiano i fornitori.",
        "VENDOR_CARRIAGE: tabella hardcodata. Se cambiano i costi plt per fornitore, aggiornare la costante COSTS.VENDOR_CARRIAGE nel codice.",
        "LIC (Local Import Charges): calcolato come (4100+3800) HKD / tasso cambio. Confermare che questi valori siano fissi o se cambiano per container.",
        "Sezione vs Category: nel file PBI, 'Section Description' viene mappata a category (FOOD/WINE/MEAT). Alcune sezioni come 'Milk and Dairy Products' → FOOD, 'Pork Meat' → MEAT. Verificare se tutti i mapping sono corretti.",
        "Health Certificate (HC): configurato manualmente per prodotto in Logistica (€80/container). Confermare importo fisso o variabile.",
        "FOR vs MTO vs MTS: l'ubicazione determina quale colonna prezzo usare (fcaDiscounted / dapFinal / mtsPrice). Verificare che i prodotti siano configurati con l'ubicazione corretta in Logistica.",
      ]
    },
    {
      title: "🟢 Funzionalità mancanti (future)",
      color: T.green,
      items: [
        "Export Standard Cost a Excel (.xlsx) con la stessa struttura del modello originale.",
        "Export mail a PDF o testo copiabile per Outlook.",
        "Filtro per categoria (FOOD / WINE / MEAT) nelle pagine Listini e Standard Cost.",
        "Grafico andamento prezzi per prodotto nel tempo.",
        "Gestione multi-container (diversi pltPerContainer per lo stesso prodotto in periodi diversi).",
        "Validazione: avviso se il prezzo importato è >20% diverso dal mese precedente.",
        "Import Sales Invoice: confronto automatico tra file diversi nel tempo (diff fatture).",
        "Backup/restore completo dello stato (tutti i localStorage) in un singolo file JSON.",
      ]
    },
    {
      title: "ℹ️ Note tecniche",
      color: T.blue,
      items: [
        "I dati sono salvati in localStorage del browser. Se si cancella la cache, si perdono i dati. Considerare export periodico.",
        "Limite localStorage: ~5-10 MB. Con molti prodotti e mesi di storico potrebbe riempirsi. Controllare via DevTools → Application → Local Storage.",
        "Il calcolo Standard Cost replica il modello Excel HK. Per CAN/AUS la struttura è simile ma i parametri FOB/LIC/WH sono diversi e non ancora implementati.",
        "La formula 'Standard Cost 1' include: prezzo + FOB + LIC + VGM + HC + pallet + alcol + carriage. La 'Standard Cost 2' aggiunge i costi magazzino (MTO/MTS/FOR).",
        "XRef table: se lo stesso N HK appare più volte nel file, viene tenuto l'ultimo. Stesso per IFB N.",
      ]
    }
  ];
  return (
    <div>
      <PageHeader title="📝 Note & Ambiguità" sub="Stato del progetto · problemi noti · chiarimenti necessari" />
      {sections.map(s=>(
        <Section key={s.title} title={s.title} accent={s.color} mb="20px">
          <ul style={{margin:0,padding:"0 0 0 18px",listStyle:"disc"}}>
            {s.items.map((item,i)=>(
              <li key={i} style={{fontSize:"12px",color:T.muted,lineHeight:"1.8",marginBottom:"4px"}}>{item}</li>
            ))}
          </ul>
        </Section>
      ))}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// IMPORT LISTINI
// ─────────────────────────────────────────────────────────────────────────────
function ImportPrices({ prices, setPrices, products, xrefs, branch, month, importLogs, setImportLogs, snapshots, setSnapshots, showToast }) {
  const [step, setStep] = useState("upload");
  const [rawRows, setRawRows] = useState([]);
  const [headers, setHeaders] = useState([]);
  const [fileName, setFileName] = useState("");
  const [mapping, setMapping] = useState({});
  const [preview, setPreview] = useState([]);
  const [importMonth, setImportMonth] = useState(month);

  const parseFile = (file) => {
    if (!file) return;
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, {type:"binary"});
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(ws, {header:1,defval:""});
        if (data.length<2) {showToast("File vuoto",T.red);return;}
        const hdrs = data[0].map(h=>String(h).trim()).filter(h=>h);
        const rows = data.slice(1).filter(r=>r.some(c=>c!==""));
        setHeaders(hdrs); setRawRows(rows);
        const autoMap={};
        Object.keys(PRICE_FIELD_ALIASES).forEach(field=>{
          const aliases = PRICE_FIELD_ALIASES[field];
          for (const h of hdrs) { const hLow=h.toLowerCase().trim(); if (aliases.some(a=>hLow===a)) {autoMap[field]=h;break;} }
          if (!autoMap[field]) { for (const h of hdrs) { const hLow=h.toLowerCase().trim(); if (aliases.some(a=>hLow.includes(a)&&a.length>3)) {autoMap[field]=h;break;} } }
        });
        setMapping(autoMap); setStep("map");
      } catch(err){showToast("Errore: "+err.message,T.red);}
    };
    reader.readAsBinaryString(file);
  };

  const buildPreview = () => {
    const get=(row,field)=>{const col=mapping[field];if(!col)return null;const i=headers.indexOf(col);return i>=0?row[i]:null;};
    const mapped = rawRows.map((row,idx)=>{
      const rawCode = String(get(row,"code")||"").trim();
      if (!rawCode) return null;
      // Usa findProduct per supportare N HK, IFB N e XRef
      const prod = findProduct(rawCode, products, xrefs);
      const mtsPrice      = parseFloat(get(row,"mtsPrice"))      || 0;
      const fcaPrice      = parseFloat(get(row,"fcaPrice"))      || 0;
      const fcaDiscount   = parseFloat(get(row,"fcaDiscount"))   || 0;
      const fcaDiscounted = parseFloat(get(row,"fcaDiscounted")) || (fcaPrice - fcaDiscount*fcaPrice/100) || 0;
      const dapPrice      = parseFloat(get(row,"dapPrice"))      || 0;
      const dapDiscount   = parseFloat(get(row,"dapDiscount"))   || 0;
      const dapDiscounted = parseFloat(get(row,"dapDiscounted")) || (dapPrice - dapDiscount*dapPrice/100) || 0;
      const vendorName    = String(get(row,"vendorName")||"").trim();
      const section       = String(get(row,"section")||"").trim();
      const dapFinalDirect = parseFloat(get(row,"dapFinalDirect")) || 0;
      let dapFinal = 0, carriageUnit = 0, dapNote = "";
      if (dapFinalDirect !== 0) { dapFinal = dapFinalDirect; dapNote = "da file"; }
      else {
        const calc = calcDAPFinal({ dapDiscounted, fcaPrice, fcaDiscounted, mtsPrice, vendorName, section, products, code: prod?.code||rawCode });
        dapFinal = calc.dapFinal; carriageUnit = calc.carriageUnit; dapNote = calc.note;
      }
      const existing = prices.find(p=>p.productId===(prod?.id||rawCode)&&p.branch===branch&&p.month===importMonth);
      return {
        _idx:idx, rawCode, productId:prod?.id||null,
        nHK: prod?.nHK || rawCode,
        ifbNo: prod?.code || rawCode,
        description:prod?.description||rawCode,
        vendorName, section,
        dapFinal: Math.round(dapFinal*100)/100,
        mtsPrice: Math.round(mtsPrice*100)/100,
        fcaDiscounted: Math.round(fcaDiscounted*100)/100,
        dapPrice: Math.round(dapPrice*100)/100,
        fcaPrice: Math.round(fcaPrice*100)/100,
        carriageUnit: Math.round(carriageUnit*100)/100,
        dapNote,
        _hasProduct:!!prod, _existing:!!existing,
      };
    }).filter(Boolean);
    setPreview(mapped); setStep("preview");
  };

  const executeImport = () => {
    let count=0;
    const snapshotId = Date.now();
    const diffs = [];
    const updated=[...prices];
    preview.filter(r=>r._hasProduct).forEach(r=>{
      const idx=updated.findIndex(p=>p.productId===r.productId&&p.branch===branch&&p.month===importMonth);
      const entry={productId:r.productId,branch,month:importMonth,
        dapFinal:r.dapFinal,mtsPrice:r.mtsPrice,fcaDiscounted:r.fcaDiscounted,
        dapPrice:r.dapPrice,fcaPrice:r.fcaPrice};
      const prev = idx>=0 ? updated[idx] : null;
      const diffFields = [];
      ["dapFinal","mtsPrice","fcaDiscounted","dapPrice","fcaPrice"].forEach(f=>{
        const oldVal = prev?.[f]||0, newVal = entry[f]||0;
        if (Math.abs(oldVal-newVal)>0.0001) {
          diffFields.push({field:f, old:oldVal, new:newVal, delta: oldVal>0?((newVal-oldVal)/oldVal*100):null});
        }
      });
      if (diffFields.length>0 || !prev) {
        diffs.push({
          productId:r.productId, nHK:r.nHK, ifbNo:r.ifbNo,
          description:r.description, isNew:!prev, fields:diffFields,
        });
      }
      if (idx>=0) updated[idx]=entry; else updated.push(entry);
      count++;
    });
    setPrices(updated);
    const logEntry={id:snapshotId,type:"prices",fileName,branch,month:importMonth,
      importedAt:new Date().toISOString(),count,newCount:diffs.filter(d=>d.isNew).length,
      updateCount:diffs.filter(d=>!d.isNew).length,diffs};
    setImportLogs(logs=>[logEntry,...logs]);
    setSnapshots(snaps=>[logEntry,...snaps].slice(0,50));
    showToast(`Import listini: ${count} prezzi per ${importMonth} ✓`,T.gold);
    setStep("done");
  };

  const reset=()=>{setStep("upload");setRawRows([]);setHeaders([]);setFileName("");setMapping({});setPreview([]);};

  return (
    <div>
      <PageHeader title="💶 Import Listini da Business Central / Power BI" sub={`Carica prezzi di acquisto · ${branch} · mese selezionabile`} />
      <StepBar steps={["upload","map","preview","done"]} current={step} />

      {step==="upload"&&(
        <div>
          <Section title="Mese di riferimento" mb="16px">
            <div style={{display:"flex",alignItems:"center",gap:"12px"}}>
              <div style={{fontSize:"12px",color:T.muted}}>Mese a cui assegnare i prezzi:</div>
              <input type="month" value={importMonth} onChange={e=>setImportMonth(e.target.value)} style={{...inputStyle(),width:"160px"}} />
            </div>
          </Section>
          <Section title="Carica file export (CURRENT PRICELIST o report PBI)" mb="20px">
            <DropZone onFile={parseFile} />
          </Section>
          <Section title="Logica selezione prezzo (replica formula Excel)">
            <div style={{padding:"10px 14px",background:`${T.blue}11`,border:`1px solid ${T.blue}33`,borderRadius:"8px",fontSize:"12px",color:T.muted}}>
              <strong style={{color:T.text}}>FOR</strong> → FCA Discounted (col 11) &nbsp;·&nbsp;
              <strong style={{color:T.text}}>MTO</strong> → DAP Final (col 18) &nbsp;·&nbsp;
              <strong style={{color:T.text}}>MTS</strong> → MTS Price (col 8) se ≠ 0, altrimenti DAP Final<br/>
              <span style={{fontSize:"11px",color:T.dim,marginTop:"4px",display:"block"}}>Il codice nel file può essere N HK o IFB N — viene usata la tabella XRef per risolvere automaticamente.</span>
            </div>
          </Section>
        </div>
      )}

      {step==="map"&&(
        <Section title={`Mappatura — ${fileName} · ${rawRows.length} righe`} mb="20px">
          <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:"12px",marginBottom:"18px"}}>
            {Object.keys(PRICE_FIELD_ALIASES).map(field=>{
              const labels={code:"Codice * (N HK o IFB N)",vendorName:"Vendor Name",section:"Section",
                mtsPrice:"MTS Price",fcaPrice:"FCA Price",fcaDiscount:"FCA Discount %",
                fcaDiscounted:"FCA Discounted",dapPrice:"DAP Price",dapDiscount:"DAP Discount %",
                dapDiscounted:"DAP Discounted",dapFinalDirect:"DAP Final (già calcolato)"};
              return (
                <div key={field}>
                  <label style={{display:"block",fontSize:"11px",color:field==="code"?T.gold:T.muted,marginBottom:"5px"}}>{labels[field]}</label>
                  <select value={mapping[field]||""} onChange={e=>setMapping(m=>({...m,[field]:e.target.value||null}))}
                    style={{...inputStyle(),cursor:"pointer",borderColor:!mapping[field]&&field==="code"?T.red+"88":T.border}}>
                    <option value="">— non mappato —</option>
                    {headers.map(h=><option key={h} value={h}>{h}</option>)}
                  </select>
                </div>
              );
            })}
          </div>
          {rawRows.length>0&&(
            <div style={{overflowX:"auto",marginBottom:"16px"}}>
              <div style={{fontSize:"10px",color:T.dim,marginBottom:"6px"}}>Prime 3 righe</div>
              <table style={{borderCollapse:"collapse",fontSize:"11px",whiteSpace:"nowrap"}}>
                <thead><tr>{headers.map(h=><th key={h} style={{padding:"4px 10px",color:T.muted,borderBottom:`1px solid ${T.border}`,textAlign:"left",fontWeight:"normal"}}>{h}</th>)}</tr></thead>
                <tbody>{rawRows.slice(0,3).map((row,i)=><tr key={i} style={{borderBottom:`1px solid ${T.border}`}}>{headers.map((h,j)=><td key={j} style={{padding:"4px 10px",color:T.text,fontFamily:"monospace"}}>{String(row[j]||"")}</td>)}</tr>)}</tbody>
              </table>
            </div>
          )}
          <div style={{display:"flex",gap:"10px"}}>
            <ActionBtn label="← Ricarica" onClick={reset} />
            <ActionBtn label="Preview →" onClick={buildPreview} primary disabled={!mapping["code"]} />
          </div>
        </Section>
      )}

      {step==="preview"&&(
        <div>
          <div style={{display:"flex",gap:"12px",marginBottom:"20px"}}>
            {[[preview.filter(r=>r._hasProduct).length,"Prodotti trovati",T.green],
              [preview.filter(r=>!r._hasProduct).length,"Non trovati",T.red],
              [preview.filter(r=>r._existing).length,"Sovrascrittura",T.orange],
              [preview.length,"Totale",T.text]].map(([n,l,c])=>(
              <div key={l} style={{padding:"10px 16px",background:T.card,border:`1px solid ${T.border}`,borderRadius:"8px"}}>
                <div style={{fontSize:"20px",fontWeight:"bold",color:c}}>{n}</div>
                <div style={{fontSize:"10px",color:T.dim,marginTop:"2px"}}>{l}</div>
              </div>
            ))}
          </div>
          <Section title={`Preview · ${importMonth} · ${branch}`}>
            <div style={{overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:"12px"}}>
                <THead cols={["N HK","IFB N","Descrizione","DAP Final","MTS Price","FCA Disc.","Stato"]} />
                <tbody>
                  {preview.map(r=>(
                    <tr key={r._idx} style={{borderBottom:`1px solid ${T.border}`,opacity:r._hasProduct?1:0.4,background:r._existing?`${T.orange}08`:""}}>
                      <TD mono><span style={{color:T.gold}}>{r.nHK||"—"}</span></TD>
                      <TD mono>{r.ifbNo}</TD>
                      <TD>{r.description}</TD>
                      <TD mono><span style={{color:T.gold}}>{r.dapFinal>0?`€ ${r.dapFinal.toFixed(2)}`:"—"}</span>{r.dapNote&&<span style={{marginLeft:"4px",fontSize:"9px",color:T.dim}}>({r.dapNote})</span>}</TD>
                      <TD mono><span style={{color:T.blue}}>{r.mtsPrice>0?`€ ${r.mtsPrice.toFixed(2)}`:"—"}</span></TD>
                      <TD mono><span style={{color:T.muted}}>{r.fcaDiscounted>0?`€ ${r.fcaDiscounted.toFixed(2)}`:"—"}</span></TD>
                      <TD>
                        {!r._hasProduct&&<Chip label="NOT FOUND" color={T.red} />}
                        {r._hasProduct&&r._existing&&<Chip label="AGGIORNA" color={T.orange} />}
                        {r._hasProduct&&!r._existing&&<Chip label="NUOVO" color={T.green} />}
                      </TD>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </Section>
          <div style={{display:"flex",gap:"10px",marginTop:"18px"}}>
            <ActionBtn label="← Torna" onClick={()=>setStep("map")} />
            <ActionBtn label={`✓ Importa ${preview.filter(r=>r._hasProduct).length} prezzi`} onClick={executeImport} primary />
          </div>
        </div>
      )}

      {step==="done"&&(
        <Section title="✓ Import completato" accent={T.green}>
          <div style={{padding:"20px",background:`${T.green}11`,border:`1px solid ${T.green}33`,borderRadius:"8px",marginBottom:"16px",fontSize:"13px",color:T.muted,lineHeight:"1.8"}}>
            File: <strong style={{color:T.text}}>{fileName}</strong><br/>
            Mese: <strong style={{color:T.gold}}>{importMonth}</strong> · Filiale: <strong style={{color:T.text}}>{branch}</strong><br/>
            Prezzi salvati: <strong style={{color:T.green}}>{importLogs[0]?.count}</strong>
          </div>
          <div style={{display:"flex",gap:"10px"}}>
            <ActionBtn label="💶 Nuovo import" onClick={reset} primary />
          </div>
        </Section>
      )}

      {importLogs.filter(l=>l.type==="prices").length>0&&step!=="done"&&(
        <Section title="Storico import listini" mt="24px">
          <table style={{width:"100%",borderCollapse:"collapse"}}>
            <THead cols={["File","Filiale","Mese","Data","Importati","Nuovi","Aggiornati"]} />
            <tbody>
              {importLogs.filter(l=>l.type==="prices").map(log=>(
                <tr key={log.id} style={{borderBottom:`1px solid ${T.border}`}}>
                  <TD mono>{log.fileName}</TD>
                  <TD>{BRANCH_CFG[log.branch]?.flag} {log.branch}</TD>
                  <TD mono>{log.month}</TD>
                  <TD mono>{new Date(log.importedAt).toLocaleDateString("it-IT")}</TD>
                  <TD><Chip label={String(log.count)} color={T.gold} /></TD>
                  <TD><Chip label={String(log.newCount||0)} color={T.green} /></TD>
                  <TD><Chip label={String(log.updateCount||0)} color={T.orange} /></TD>
                </tr>
              ))}
            </tbody>
          </table>
        </Section>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// IMPORT ANAGRAFICA BC
// ─────────────────────────────────────────────────────────────────────────────
function ImportBC({ products, setProducts, importLogs, setImportLogs, snapshots, setSnapshots, showToast, branch }) {
  const [step,setStep]=useState("upload");
  const [rawRows,setRawRows]=useState([]);
  const [headers,setHeaders]=useState([]);
  const [fileName,setFileName]=useState("");
  const [mapping,setMapping]=useState({});
  const [preview,setPreview]=useState([]);
  const [actions,setActions]=useState({});

  const parseFile=(file)=>{
    if(!file)return;
    setFileName(file.name);
    const reader=new FileReader();
    reader.onload=(e)=>{
      try{
        const wb=XLSX.read(e.target.result,{type:"binary"});
        const ws=wb.Sheets[wb.SheetNames[0]];
        const data=XLSX.utils.sheet_to_json(ws,{header:1,defval:""});
        if(data.length<2){showToast("File vuoto",T.red);return;}
        const hdrs=data[0].map(h=>String(h).trim()).filter(h=>h);
        const rows=data.slice(1).filter(r=>r.some(c=>c!==""));
        setHeaders(hdrs);setRawRows(rows);
        const autoMap={};
        Object.keys(BC_FIELD_ALIASES).forEach(field=>{const d=detectBCColumn(hdrs,field);if(d)autoMap[field]=d;});
        setMapping(autoMap);setStep("map");
      }catch(err){showToast("Errore: "+err.message,T.red);}
    };
    reader.readAsBinaryString(file);
  };

  const buildPreview=()=>{
    const get=(row,field)=>{const col=mapping[field];if(!col)return null;const i=headers.indexOf(col);return i>=0?row[i]:null;};
    const mapped=rawRows.map((row,idx)=>{
      const code=String(get(row,"code")||"").trim();
      if(!code)return null;
      const existing=products.find(p=>p.code===code);
      return {
        _idx:idx,
        nHK: String(get(row,"nHK")||existing?.nHK||"").trim(),
        code,
        description:String(get(row,"description")||"").trim()||(existing?.description||""),
        category:mapBCValue("category",get(row,"category"))||existing?.category||"FOOD",
        uom:mapBCValue("uom",get(row,"uom"))||existing?.uom||"PCS",
        temperature:mapBCValue("temperature",get(row,"temperature"))||existing?.temperature||"DRY",
        active:get(row,"active")!==null?mapBCValue("active",get(row,"active")):(existing?.active??true),
        qtyPerBox:parseFloat(get(row,"qtyPerBox"))||existing?.qtyPerBox||1,
        boxPerPallet:parseFloat(get(row,"boxPerPallet"))||existing?.boxPerPallet||1,
        kgPerBox:parseFloat(get(row,"kgPerBox"))||existing?.kgPerBox||null,
        _isNew:!existing,_existing:existing,
      };
    }).filter(Boolean);
    setPreview(mapped);
    const defaultActions={};
    mapped.forEach(r=>{defaultActions[r._idx]=r._isNew?"NEW":"UPDATE";});
    setActions(defaultActions);setStep("preview");
  };

  const executeImport=()=>{
    const toProcess=preview.filter(r=>actions[r._idx]!=="SKIP");
    let newCount=0,updateCount=0;
    const snapshotId=Date.now();
    const diffs=[];
    const updated=[...products];
    toProcess.forEach(r=>{
      if(actions[r._idx]==="NEW"){
        updated.push({id:"P_BC_"+Date.now()+"_"+r._idx,nHK:r.nHK||"",code:r.code,description:r.description,category:r.category,uom:r.uom,qtyPerBox:r.qtyPerBox,boxPerPallet:r.boxPerPallet,kgPerBox:r.kgPerBox,temperature:r.temperature,active:r.active});
        diffs.push({nHK:r.nHK,ifbNo:r.code,description:r.description,isNew:true,fields:[]});
        newCount++;
      }else if(actions[r._idx]==="UPDATE"){
        const i=updated.findIndex(p=>p.code===r.code);
        if(i>=0){
          const prev=updated[i];
          const diffFields=[];
          ["nHK","description","category","uom","qtyPerBox","boxPerPallet","kgPerBox","temperature","active"].forEach(f=>{
            if(String(prev[f]??null)!==String(r[f]??null))diffFields.push({field:f,old:prev[f],new:r[f]});
          });
          updated[i]={...prev,nHK:r.nHK||prev.nHK||"",description:r.description,category:r.category,uom:r.uom,qtyPerBox:r.qtyPerBox,boxPerPallet:r.boxPerPallet,kgPerBox:r.kgPerBox,temperature:r.temperature,active:r.active};
          if(diffFields.length>0)diffs.push({nHK:r.nHK,ifbNo:r.code,description:r.description,isNew:false,fields:diffFields});
          updateCount++;
        }
      }
    });
    setProducts(updated);
    const logEntry={id:snapshotId,type:"anagrafica",fileName,branch,importedAt:new Date().toISOString(),
      totalRows:rawRows.length,newCount,updateCount,
      skippedCount:preview.filter(r=>actions[r._idx]==="SKIP").length,diffs};
    setImportLogs(logs=>[logEntry,...logs]);
    setSnapshots(snaps=>[logEntry,...snaps].slice(0,50));
    showToast(`Import: ${newCount} nuovi, ${updateCount} aggiornati ✓`,T.gold);
    setStep("done");
  };

  const reset=()=>{setStep("upload");setRawRows([]);setHeaders([]);setFileName("");setMapping({});setPreview([]);setActions({});};

  return (
    <div>
      <PageHeader title="⇪ Import Anagrafica da Business Central" sub="Carica export Excel/CSV da Power BI o BC" />
      <StepBar steps={["upload","map","preview","done"]} current={step} />
      {step==="upload"&&(<div><Section title="Carica file" mb="20px"><DropZone onFile={parseFile} /></Section></div>)}
      {step==="map"&&(
        <Section title={`Mappatura — ${fileName} · ${rawRows.length} righe`} mb="20px">
          <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:"12px",marginBottom:"16px"}}>
            {Object.keys(BC_FIELD_ALIASES).map(field=>{
              const isReq=["code","description"].includes(field);
              const labels={nHK:"N HK",code:"IFB No *",description:"Descrizione *",category:"Categoria",uom:"UOM",qtyPerBox:"Qty/cartone",boxPerPallet:"Cartoni/pallet",kgPerBox:"Kg per cartone",temperature:"Temperatura",active:"Bloccato/Attivo"};
              return(<div key={field}><label style={{display:"block",fontSize:"11px",color:isReq?T.gold:T.muted,marginBottom:"5px"}}>{labels[field]}</label><select value={mapping[field]||""} onChange={e=>setMapping(m=>({...m,[field]:e.target.value||null}))} style={{...inputStyle(),cursor:"pointer",borderColor:!mapping[field]&&isReq?T.red+"88":T.border}}><option value="">— non mappato —</option>{headers.map(h=><option key={h} value={h}>{h}</option>)}</select></div>);
            })}
          </div>
          <div style={{display:"flex",gap:"10px"}}><ActionBtn label="← Ricarica" onClick={reset} /><ActionBtn label="Preview →" onClick={buildPreview} primary disabled={["code","description"].some(f=>!mapping[f])} /></div>
        </Section>
      )}
      {step==="preview"&&(
        <div>
          <div style={{display:"flex",gap:"12px",marginBottom:"20px"}}>
            {[[preview.filter(r=>r._isNew).length,"Nuovi",T.gold],[preview.filter(r=>!r._isNew).length,"Aggiornare",T.blue],[preview.filter(r=>actions[r._idx]==="SKIP").length,"Skip",T.muted],[preview.length,"Totale",T.text]].map(([n,l,c])=>(
              <div key={l} style={{padding:"10px 16px",background:T.card,border:`1px solid ${T.border}`,borderRadius:"8px"}}><div style={{fontSize:"20px",fontWeight:"bold",color:c}}>{n}</div><div style={{fontSize:"10px",color:T.dim,marginTop:"2px"}}>{l}</div></div>
            ))}
          </div>
          <Section title={`Preview · ${preview.length} prodotti`}>
            <div style={{overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:"12px"}}>
                <THead cols={["N HK","IFB N","Descrizione","Cat.","UOM","Temp.","Qty/Box","Box/Plt","Stato","Azione"]} />
                <tbody>
                  {preview.map(r=>{
                    const act=actions[r._idx]||"SKIP";
                    return(
                      <tr key={r._idx} style={{borderBottom:`1px solid ${T.border}`,opacity:act==="SKIP"?0.4:1,background:r._isNew?`${T.gold}07`:act==="UPDATE"?`${T.blue}07`:""}}>
                        <TD mono><span style={{color:T.gold}}>{r.nHK||<span style={{color:T.dim}}>—</span>}</span></TD>
                        <TD mono>{r.code}</TD>
                        <TD>{r.description}</TD>
                        <TD><Chip label={r.category} /></TD>
                        <TD><Chip label={r.uom} color={T.blue} /></TD>
                        <TD><TempChip t={r.temperature} /></TD>
                        <TD mono>{r.qtyPerBox}</TD>
                        <TD mono>{r.boxPerPallet}</TD>
                        <TD><Chip label={r.active?"Attivo":"Sospeso"} color={r.active?T.green:T.red} /></TD>
                        <TD>
                          <div style={{display:"flex",gap:"4px"}}>
                            {(r._isNew?["NEW","SKIP"]:["UPDATE","SKIP","NEW"]).map(a=>(
                              <button key={a} onClick={()=>setActions(ac=>({...ac,[r._idx]:a}))} style={{padding:"2px 8px",fontSize:"10px",fontWeight:act===a?"bold":"normal",cursor:"pointer",fontFamily:"inherit",background:act===a?`${a==="NEW"?T.gold:a==="UPDATE"?T.blue:T.muted}22`:"transparent",border:`1px solid ${act===a?(a==="NEW"?T.gold:a==="UPDATE"?T.blue:T.muted):T.border}`,borderRadius:"4px",color:act===a?(a==="NEW"?T.gold:a==="UPDATE"?T.blue:T.muted):T.dim}}>{a}</button>
                            ))}
                          </div>
                        </TD>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
            <div style={{display:"flex",gap:"8px",marginTop:"14px",paddingTop:"14px",borderTop:`1px solid ${T.border}`}}>
              <MiniBtn label="Tutto UPDATE" onClick={()=>setActions(a=>{const n={...a};preview.filter(r=>!r._isNew).forEach(r=>{n[r._idx]="UPDATE"});return n;})} />
              <MiniBtn label="Skip esistenti" onClick={()=>setActions(a=>{const n={...a};preview.filter(r=>!r._isNew).forEach(r=>{n[r._idx]="SKIP"});return n;})} />
            </div>
          </Section>
          <div style={{display:"flex",gap:"10px",marginTop:"18px"}}>
            <ActionBtn label="← Mappatura" onClick={()=>setStep("map")} />
            <ActionBtn label={`✓ Esegui (${preview.filter(r=>actions[r._idx]!=="SKIP").length})`} onClick={executeImport} primary />
          </div>
        </div>
      )}
      {step==="done"&&importLogs.length>0&&(
        <Section title="✓ Import completato" accent={T.green}>
          <div style={{padding:"20px",background:`${T.green}11`,border:`1px solid ${T.green}33`,borderRadius:"8px",fontSize:"13px",color:T.muted,lineHeight:"1.8",marginBottom:"16px"}}>
            File: <strong style={{color:T.text}}>{importLogs[0]?.fileName}</strong><br/>
            Nuovi: <strong style={{color:T.gold}}>{importLogs[0]?.newCount}</strong> · Aggiornati: <strong style={{color:T.blue}}>{importLogs[0]?.updateCount}</strong>
          </div>
          <ActionBtn label="⇪ Nuovo import" onClick={reset} primary />
        </Section>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// DASHBOARD
// ─────────────────────────────────────────────────────────────────────────────
function Dashboard({costRows,branch,month,branchCfg,setPage}) {
  const flagged=costRows.filter(r=>r.flagged);
  const newItems=costRows.filter(r=>r.isNew&&r.cost);
  const missing=costRows.filter(r=>!r.cost);
  return (
    <div>
      <PageHeader title={`Dashboard — ${branchCfg.flag} ${branchCfg.label}`} sub={`Riepilogo Standard Cost · ${month}`} />
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:"14px",marginBottom:"28px"}}>
        <KPI label="Prodotti attivi"   value={costRows.length} color={T.blue}   icon="◈" />
        <KPI label="Variazioni ≥ ±3%" value={flagged.length}  color={flagged.length?T.red:T.green} icon="◉" />
        <KPI label="Nuovi prodotti"    value={newItems.length} color={T.gold}   icon="+" />
        <KPI label="Senza dati"        value={missing.length}  color={missing.length?T.orange:T.green} icon="◌" />
      </div>
      {flagged.length>0&&(
        <Section title="⚠️ Variazioni significative (≥ ±3%)" accent={T.red} mb="20px">
          <table style={{width:"100%",borderCollapse:"collapse"}}>
            <THead cols={["N HK","IFB N","Descrizione","Ubic.","SC2 prec. €","SC2 nuovo €","Δ %","Trend"]} />
            <tbody>
              {flagged.map(r=>(
                <tr key={r.id} style={{borderBottom:`1px solid ${T.border}`}}>
                  <TD mono><span style={{color:T.gold}}>{r.nHK||"—"}</span></TD>
                  <TD mono>{r.code}</TD>
                  <TD>{r.description}</TD>
                  <TD><UbicChip u={r.ubicazione} /></TD>
                  <TD mono>{FMT(r.prevCost?.step2Eur)}</TD>
                  <TD mono bold>{FMT(r.cost?.step2Eur)}</TD>
                  <TD><DeltaBadge delta={r.delta} /></TD>
                  <TD>{r.delta>0?"📈":"📉"}</TD>
                </tr>
              ))}
            </tbody>
          </table>
        </Section>
      )}
      {newItems.length>0&&(
        <Section title="✦ Nuovi prodotti questo mese" accent={T.gold} mb="20px">
          <table style={{width:"100%",borderCollapse:"collapse"}}>
            <THead cols={["N HK","IFB N","Descrizione","Cat.","Temp.","Ubic.","SC2 €",`SC2 ${branchCfg.currency}`]} />
            <tbody>
              {newItems.map(r=>(
                <tr key={r.id} style={{borderBottom:`1px solid ${T.border}`}}>
                  <TD mono><span style={{color:T.gold}}>{r.nHK||"—"}</span></TD>
                  <TD mono>{r.code}</TD>
                  <TD>{r.description}</TD>
                  <TD><Chip label={r.category} /></TD>
                  <TD><TempChip t={r.temperature} /></TD>
                  <TD><UbicChip u={r.ubicazione} /></TD>
                  <TD mono bold>{FMT(r.cost?.step2Eur)}</TD>
                  <TD mono>{FMT(r.cost?.step2Hkd,2)}</TD>
                </tr>
              ))}
            </tbody>
          </table>
        </Section>
      )}
      <div style={{display:"flex",gap:"12px",marginTop:"8px",flexWrap:"wrap"}}>
        <ActionBtn label="📊 Vedi tutti i costi" onClick={()=>setPage("costs")} primary />
        <ActionBtn label="✉ Genera mail mensile" onClick={()=>setPage("mail")} />
        <ActionBtn label="⇪ Import Anagrafica"   onClick={()=>setPage("importAnag")} />
        <ActionBtn label="💶 Import Listini"      onClick={()=>setPage("importPrice")} />
      </div>
    </div>
  );
}
// ─────────────────────────────────────────────────────────────────────────────
// PRODUCTS
// ─────────────────────────────────────────────────────────────────────────────
function Products({products,setProducts,showToast}) {
  const EMPTY={id:"",nHK:"",code:"",description:"",category:"FOOD",uom:"PCS",qtyPerBox:"",boxPerPallet:"",kgPerBox:"",temperature:"DRY",active:true};
  const [form,setForm]=useState(EMPTY);
  const [editId,setEditId]=useState(null);
  const [search,setSearch]=useState("");
  const setF=(k,v)=>setForm(f=>({...f,[k]:v}));
  const handleSave=()=>{
    if(!form.code||!form.description){showToast("IFB No e descrizione obbligatori",T.red);return;}
    if(editId){setProducts(ps=>ps.map(p=>p.id===editId?{...form,id:editId}:p));showToast("Aggiornato ✓");}
    else{setProducts(ps=>[...ps,{...form,id:"P"+Date.now()}]);showToast("Aggiunto ✓");}
    setForm(EMPTY);setEditId(null);
  };
  const filtered=products.filter(p=>
    !search||p.nHK?.toLowerCase().includes(search.toLowerCase())||
    p.code?.toLowerCase().includes(search.toLowerCase())||
    p.description?.toLowerCase().includes(search.toLowerCase())
  );
  return (
    <div>
      <PageHeader title="◈ Anagrafica Prodotti" sub="Master data condiviso tra tutte le filiali" />
      <Section title={editId?"Modifica prodotto":"Nuovo prodotto"} mb="16px">
        <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:"10px"}}>
          <FormField label="N HK (codice Hong Kong)" value={form.nHK||""} onChange={v=>setF("nHK",v)} placeholder="es. GC0059" />
          <FormField label="IFB No (codice sistema) *" value={form.code} onChange={v=>setF("code",v)} placeholder="es. MMA01" />
          <FormField label="Descrizione *" value={form.description} onChange={v=>setF("description",v)} placeholder="Nome prodotto" span={2} />
          <SelectField label="Categoria" value={form.category} onChange={v=>setF("category",v)} opts={[["FOOD","Food"],["WINE","Beverage/Wine"],["MEAT","Meat/Fish"]]} />
          <SelectField label="UOM" value={form.uom} onChange={v=>setF("uom",v)} opts={[["PCS","PCS"],["BOX","BOX"],["KG","KG"]]} />
          <SelectField label="Temperatura" value={form.temperature} onChange={v=>setF("temperature",v)} opts={[["DRY","DRY"],["FRESH","FRESH"],["FROZEN","FROZEN"]]} />
          <FormField label={form.uom==="KG"?"Kg per cartone":"Pz per cartone"} value={form.qtyPerBox} onChange={v=>setF("qtyPerBox",v)} type="number" />
          <FormField label="Cartoni per pallet" value={form.boxPerPallet} onChange={v=>setF("boxPerPallet",v)} type="number" />
          {form.uom==="KG"&&<FormField label="Kg netti per cartone" value={form.kgPerBox} onChange={v=>setF("kgPerBox",v)} type="number" />}
        </div>
        <div style={{display:"flex",gap:"10px",marginTop:"12px"}}>
          <ActionBtn label={editId?"Salva modifiche":"Aggiungi"} onClick={handleSave} primary />
          {editId&&<ActionBtn label="Annulla" onClick={()=>{setForm(EMPTY);setEditId(null);}} />}
        </div>
      </Section>
      <SearchBar value={search} onChange={setSearch} placeholder="🔍  Filtra per N HK, IFB No o descrizione…" />
      <Section title={`${filtered.length} prodotti`}>
        <table style={{width:"100%",borderCollapse:"collapse"}}>
          <THead cols={["N HK","IFB N","Descrizione","Cat.","UOM","Temp.","Qty/Box","Box/Plt","Stato","Azioni"]} />
          <tbody>
            {filtered.map(p=>(
              <tr key={p.id} style={{borderBottom:`1px solid ${T.border}`,opacity:p.active?1:0.45}}>
                <TD mono><span style={{color:T.gold}}>{p.nHK||<span style={{color:T.dim}}>—</span>}</span></TD>
                <TD mono>{p.code}</TD>
                <TD>{p.description}</TD>
                <TD><Chip label={p.category} /></TD>
                <TD><Chip label={p.uom} color={T.blue} /></TD>
                <TD><TempChip t={p.temperature} /></TD>
                <TD mono>{p.qtyPerBox}</TD>
                <TD mono>{p.boxPerPallet}</TD>
                <TD><Chip label={p.active?"Attivo":"Sospeso"} color={p.active?T.green:T.red} /></TD>
                <TD>
                  <div style={{display:"flex",gap:"6px"}}>
                    <MiniBtn label="✎" onClick={()=>{setForm(p);setEditId(p.id);}} />
                    <MiniBtn label={p.active?"⊘":"✓"} onClick={()=>setProducts(ps=>ps.map(pp=>pp.id===p.id?{...pp,active:!pp.active}:pp))} color={p.active?T.red:T.green} />
                  </div>
                </TD>
              </tr>
            ))}
          </tbody>
        </table>
      </Section>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// LOGISTICS
// ─────────────────────────────────────────────────────────────────────────────
function Logistics({logistics,setLogistics,products,branch,showToast}) {
  const [editId,setEditId]=useState(null);
  const [form,setForm]=useState({});
  const [search,setSearch]=useState("");
  const setF=(k,v)=>setForm(f=>({...f,[k]:v}));
  const branchLog=logistics.filter(l=>l.branch===branch);
  const coveredIds=branchLog.map(l=>l.productId);
  const uncovered=products.filter(p=>p.active&&!coveredIds.includes(p.id));
  const startEdit=(log)=>{setEditId(log.productId);setForm({...log});};
  const startNew=(prod)=>{setEditId("NEW_"+prod.id);setForm({productId:prod.id,branch,area:"NORD",ubicazione:"MTO",pltPerContainer:23,hasCert:false,hasAlcTax:false,alcTax:0,convFactor:1,carriage:0,vendorName:""});};
  const handleSave=()=>{
    if(editId?.startsWith("NEW_"))setLogistics(ls=>[...ls,{...form}]);
    else setLogistics(ls=>ls.map(l=>l.productId===editId&&l.branch===branch?form:l));
    showToast("Salvato ✓");setEditId(null);setForm({});
  };
  return (
    <div>
      <PageHeader title="◎ Parametri Logistici" sub={`FOR / MTO / MTS — ${branch}`} />
      {editId&&(
        <Section title="Modifica parametri" mb="22px">
          <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:"12px"}}>
            <SelectField label="Ubicazione" value={form.ubicazione} onChange={v=>setF("ubicazione",v)} opts={[["MTO","MTO — Cross-docking"],["FOR","FOR — Franco fornitore"],["MTS","MTS — Magazzino Unifreddo"]]} />
            <SelectField label="Area fornitore" value={form.area} onChange={v=>setF("area",v)} opts={[["NORD","Nord Italia"],["CENTRO","Centro Italia"],["SUD","Sud Italia"]]} />
            <FormField label="Pallet per container" value={form.pltPerContainer} onChange={v=>setF("pltPerContainer",parseFloat(v)||23)} type="number" />
            <FormField label="Fattore conversione UM" value={form.convFactor||1} onChange={v=>setF("convFactor",parseFloat(v)||1)} type="number" />
            <FormField label="Fornitore (per carriage auto)" value={form.vendorName||""} onChange={v=>setF("vendorName",v)} placeholder="es. BONOMI SPA" span={2} />
            <FormField label="Carriage €/plt (0=auto)" value={form.carriage||0} onChange={v=>setF("carriage",parseFloat(v)||0)} type="number" />
            <div style={{display:"flex",alignItems:"center",gap:"10px",paddingTop:"18px"}}><CheckBox label="Health Certificate" checked={form.hasCert||false} onChange={v=>setF("hasCert",v)} /></div>
            <div style={{display:"flex",alignItems:"center",gap:"10px",paddingTop:"18px"}}><CheckBox label="Tassa alcolici (>30°)" checked={form.hasAlcTax||false} onChange={v=>setF("hasAlcTax",v)} /></div>
            {form.hasAlcTax&&<FormField label="Importo tassa alcolici (€/u)" value={form.alcTax||0} onChange={v=>setF("alcTax",parseFloat(v)||0)} type="number" />}
          </div>
          <div style={{display:"flex",gap:"10px",marginTop:"14px"}}>
            <ActionBtn label="Salva" onClick={handleSave} primary />
            <ActionBtn label="Annulla" onClick={()=>{setEditId(null);setForm({});}} />
          </div>
        </Section>
      )}
      {uncovered.length>0&&(
        <Section title={`${uncovered.length} prodotti senza configurazione logistica`} accent={T.orange} mb="20px">
          <table style={{width:"100%",borderCollapse:"collapse"}}>
            <THead cols={["N HK","IFB N","Descrizione","Temp.","Azione"]} />
            <tbody>
              {uncovered.map(p=>(
                <tr key={p.id} style={{borderBottom:`1px solid ${T.border}`}}>
                  <TD mono><span style={{color:T.gold}}>{p.nHK||"—"}</span></TD>
                  <TD mono>{p.code}</TD>
                  <TD>{p.description}</TD>
                  <TD><TempChip t={p.temperature} /></TD>
                  <TD><MiniBtn label="+ Configura" onClick={()=>startNew(p)} color={T.gold} /></TD>
                </tr>
              ))}
            </tbody>
          </table>
        </Section>
      )}
      <SearchBar value={search} onChange={setSearch} />
      <Section title="Configurazioni attive">
        <table style={{width:"100%",borderCollapse:"collapse"}}>
          <THead cols={["N HK","IFB N","Descrizione","Ubic.","Area","Plt/Cont","Conv.","Carriage","HC","Alc.","✎"]} />
          <tbody>
            {branchLog.filter(l=>{
              if(!search)return true;
              const p=products.find(pp=>pp.id===l.productId);
              const s=search.toLowerCase();
              return p?.code?.toLowerCase().includes(s)||p?.nHK?.toLowerCase().includes(s)||p?.description?.toLowerCase().includes(s);
            }).map(l=>{
              const p=products.find(pp=>pp.id===l.productId);
              if(!p)return null;
              return(
                <tr key={l.productId} style={{borderBottom:`1px solid ${T.border}`}}>
                  <TD mono><span style={{color:T.gold}}>{p.nHK||"—"}</span></TD>
                  <TD mono>{p.code}</TD>
                  <TD>{p.description}</TD>
                  <TD><UbicChip u={l.ubicazione} /></TD>
                  <TD><Chip label={l.area} color={T.muted} /></TD>
                  <TD mono>{l.pltPerContainer}</TD>
                  <TD mono>{l.convFactor||1}</TD>
                  <TD mono>{l.carriage>0?`€${l.carriage}(M)`:`€${COSTS.VENDOR_CARRIAGE[l.vendorName]||"—"}(A)`}</TD>
                  <TD>{l.hasCert?"✓":"—"}</TD>
                  <TD>{l.hasAlcTax?`€${l.alcTax}`:"—"}</TD>
                  <TD><MiniBtn label="✎" onClick={()=>startEdit(l)} /></TD>
                </tr>
              );
            })}
          </tbody>
        </table>
      </Section>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// PRICES
// ─────────────────────────────────────────────────────────────────────────────
function Prices({prices,setPrices,products,branch,month,showToast}) {
  const [editingId,setEditingId]=useState(null);
  const [tempVals,setTempVals]=useState({dapFinal:"",mtsPrice:"",fcaDiscounted:""});
  const [search,setSearch]=useState("");
  const monthPrices=prices.filter(p=>p.branch===branch&&p.month===month);
  const prevMonth=month.slice(0,4)+"-"+String(parseInt(month.slice(5))-1).padStart(2,"0");
  const prevPrices=prices.filter(p=>p.branch===branch&&p.month===prevMonth);
  const handleSave=(productId)=>{
    const dap=parseFloat(tempVals.dapFinal);
    if(!dap||dap<=0){showToast("DAP Final obbligatorio",T.red);return;}
    const entry={productId,branch,month,dapFinal:Math.round(dap*100)/100,mtsPrice:Math.round((parseFloat(tempVals.mtsPrice)||0)*100)/100,fcaDiscounted:Math.round((parseFloat(tempVals.fcaDiscounted)||0)*100)/100,dapPrice:0,fcaPrice:0};
    const exists=prices.find(p=>p.productId===productId&&p.branch===branch&&p.month===month);
    if(exists)setPrices(ps=>ps.map(p=>p.productId===productId&&p.branch===branch&&p.month===month?entry:p));
    else setPrices(ps=>[...ps,entry]);
    showToast("Prezzo salvato ✓");setEditingId(null);
  };
  return(
    <div>
      <PageHeader title="◉ Listini — Prezzi di Acquisto" sub={`${branch} · ${month}`} />
      <Section title="Prezzi mensili">
        <SearchBar value={search} onChange={setSearch} />
        <table style={{width:"100%",borderCollapse:"collapse"}}>
          <THead cols={["N HK","IFB N","Descrizione","UOM","DAP Final (MTO)","MTS Price","FCA Disc. (FOR)","Δ %"]} />
          <tbody>
            {products.filter(p=>p.active&&(!search||p.code?.toLowerCase().includes(search.toLowerCase())||p.nHK?.toLowerCase().includes(search.toLowerCase())||p.description?.toLowerCase().includes(search.toLowerCase()))).map(prod=>{
              const cur=monthPrices.find(p=>p.productId===prod.id);
              const prev=prevPrices.find(p=>p.productId===prod.id);
              const delta=cur&&prev&&cur.dapFinal>0&&prev.dapFinal>0?PCT(cur.dapFinal,prev.dapFinal):null;
              const isEditing=editingId===prod.id;
              return(
                <tr key={prod.id} style={{borderBottom:`1px solid ${T.border}`,background:!cur?"rgba(196,122,59,0.07)":""}}>
                  <TD mono><span style={{color:T.gold}}>{prod.nHK||"—"}</span></TD>
                  <TD mono>{prod.code}</TD>
                  <TD>{prod.description}</TD>
                  <TD><Chip label={prod.uom} color={T.blue} /></TD>
                  {isEditing?(
                    <>
                      <td style={{padding:"6px 8px"}}><input autoFocus type="number" placeholder="DAP Final *" value={tempVals.dapFinal} onChange={e=>setTempVals(v=>({...v,dapFinal:e.target.value}))} style={{...inputStyle(),width:"110px",padding:"4px 8px",fontSize:"12px"}} /></td>
                      <td style={{padding:"6px 8px"}}><input type="number" placeholder="MTS" value={tempVals.mtsPrice} onChange={e=>setTempVals(v=>({...v,mtsPrice:e.target.value}))} style={{...inputStyle(),width:"90px",padding:"4px 8px",fontSize:"12px"}} /></td>
                      <td style={{padding:"6px 8px"}}>
                        <div style={{display:"flex",gap:"4px",alignItems:"center"}}>
                          <input type="number" placeholder="FCA" value={tempVals.fcaDiscounted} onChange={e=>setTempVals(v=>({...v,fcaDiscounted:e.target.value}))} style={{...inputStyle(),width:"90px",padding:"4px 8px",fontSize:"12px"}} />
                          <MiniBtn label="✓" onClick={()=>handleSave(prod.id)} color={T.green} />
                          <MiniBtn label="✕" onClick={()=>setEditingId(null)} color={T.red} />
                        </div>
                      </td>
                    </>
                  ):(
                    <>
                      <td style={{padding:"9px 10px"}}><span onClick={()=>{setEditingId(prod.id);setTempVals({dapFinal:cur?.dapFinal||"",mtsPrice:cur?.mtsPrice||"",fcaDiscounted:cur?.fcaDiscounted||""});}} style={{cursor:"pointer",color:cur?T.gold:T.orange,textDecoration:"underline dotted",fontFamily:"monospace",fontSize:"13px"}}>{cur&&cur.dapFinal>0?`€ ${cur.dapFinal.toFixed(2)}`:"— inserisci"}</span></td>
                      <td style={{padding:"9px 10px",fontFamily:"monospace",fontSize:"13px",color:T.blue}}>{cur?.mtsPrice>0?`€ ${cur.mtsPrice.toFixed(2)}`:<span style={{color:T.dim}}>—</span>}</td>
                      <td style={{padding:"9px 10px",fontFamily:"monospace",fontSize:"13px",color:T.muted}}>{cur?.fcaDiscounted>0?`€ ${cur.fcaDiscounted.toFixed(2)}`:<span style={{color:T.dim}}>—</span>}</td>
                    </>
                  )}
                  <TD>{delta!==null?<DeltaBadge delta={delta} small />:<span style={{color:T.dim}}>—</span>}</TD>
                </tr>
              );
            })}
          </tbody>
        </table>
      </Section>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// FX RATES
// ─────────────────────────────────────────────────────────────────────────────
function FxRates({fx,setFx,branch,month,showToast}) {
  const [val,setVal]=useState("");
  const cur=fx.find(f=>f.branch===branch&&f.month===month);
  const cfg=BRANCH_CFG[branch];
  const handleSave=()=>{
    const v=parseFloat(val);
    if(!v||v<=0){showToast("Tasso non valido",T.red);return;}
    if(cur)setFx(fs=>fs.map(f=>f.branch===branch&&f.month===month?{...f,rate:v}:f));
    else setFx(fs=>[...fs,{branch,month,rate:v}]);
    showToast("Tasso aggiornato ✓");setVal("");
  };
  const history=fx.filter(f=>f.branch===branch).sort((a,b)=>b.month.localeCompare(a.month));
  return(
    <div>
      <PageHeader title="◌ Tassi di Cambio" sub={`EUR → ${cfg.currency} · ${branch}`} />
      <Section title={`Tasso corrente — ${month}`} mb="22px">
        <div style={{display:"flex",alignItems:"flex-end",gap:"16px"}}>
          <div>
            <div style={{fontSize:"11px",color:T.muted,marginBottom:"6px"}}>EUR/{cfg.currency}</div>
            <div style={{fontSize:"32px",fontWeight:"bold",color:cfg.color}}>{cur?cur.rate.toFixed(4):<span style={{color:T.dim}}>—</span>}</div>
            <div style={{fontSize:"11px",color:T.dim,marginTop:"3px"}}>Default: {cfg.defaultRate.toFixed(4)}</div>
          </div>
          <div style={{display:"flex",gap:"8px",alignItems:"center",paddingBottom:"4px"}}>
            <input type="number" placeholder={`es. ${cfg.defaultRate}`} value={val} onChange={e=>setVal(e.target.value)} style={{...inputStyle(),width:"140px"}} />
            <ActionBtn label="Aggiorna" onClick={handleSave} primary />
          </div>
        </div>
      </Section>
      <Section title={`Storico EUR/${cfg.currency}`}>
        <table style={{width:"100%",borderCollapse:"collapse"}}>
          <THead cols={["Mese","Tasso","Δ %"]} />
          <tbody>
            {history.map((f,i)=>{const prev=history[i+1];const delta=prev?PCT(f.rate,prev.rate):null;return(
              <tr key={f.month} style={{borderBottom:`1px solid ${T.border}`}}>
                <TD mono>{f.month}</TD><TD mono bold>{f.rate.toFixed(4)}</TD>
                <TD>{delta!==null?<DeltaBadge delta={delta} small />:<span style={{color:T.dim}}>—</span>}</TD>
              </tr>
            );})}
          </tbody>
        </table>
      </Section>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// COST TABLE
// ─────────────────────────────────────────────────────────────────────────────
function CostTable({costRows,branch,month,branchCfg}) {
  const [filter,setFilter]=useState("ALL");
  const [expand,setExpand]=useState(null);
  const [search,setSearch]=useState("");
  const filtered=costRows.filter(r=>{
    const s=search.toLowerCase();
    const matchSearch=!search||r.code?.toLowerCase().includes(s)||r.nHK?.toLowerCase().includes(s)||r.description?.toLowerCase().includes(s);
    if(!matchSearch)return false;
    if(filter==="FLAGGED")return r.flagged;
    if(filter==="NEW")return r.isNew&&r.cost;
    if(filter==="MISSING")return !r.cost;
    return true;
  });
  return(
    <div>
      <PageHeader title="◆ Standard Cost" sub={`Calcolo completo · ${branch} · ${month}`} />
      <div style={{display:"flex",gap:"8px",marginBottom:"12px",flexWrap:"wrap"}}>
        {[["ALL","Tutti"],["FLAGGED","Variazioni ≥ ±3%"],["NEW","Nuovi"],["MISSING","Mancanti"]].map(([k,l])=>(
          <button key={k} onClick={()=>setFilter(k)} style={{padding:"6px 14px",background:filter===k?T.goldDim:"rgba(255,255,255,0.04)",border:`1px solid ${filter===k?T.gold:T.border}`,borderRadius:"6px",color:filter===k?T.gold:T.muted,cursor:"pointer",fontFamily:"inherit",fontSize:"12px"}}>{l}</button>
        ))}
      </div>
      <SearchBar value={search} onChange={setSearch} />
      <Section title={`${filtered.length} prodotti`}>
        <table style={{width:"100%",borderCollapse:"collapse"}}>
          <THead cols={["N HK","IFB N","Descrizione","Temp.","Ubic.","Prezzo €","SC1 €","SC2 €",`SC2 ${branchCfg.currency}`,"Δ %","Det."]} />
          <tbody>
            {filtered.map(r=>(
              <>
                <tr key={r.id} style={{borderBottom:`1px solid ${T.border}`,background:r.flagged?"rgba(181,83,74,0.05)":r.isNew?"rgba(201,168,76,0.05)":""}}>
                  <TD mono><span style={{color:T.gold}}>{r.nHK||"—"}</span></TD>
                  <TD mono>{r.code}</TD>
                  <TD>{r.description}</TD>
                  <TD><TempChip t={r.temperature} /></TD>
                  <TD><UbicChip u={r.ubicazione||r.log?.ubicazione} /></TD>
                  <TD mono>{r.cost&&r.cost.priceEur>0?<span style={{color:T.text}}>€ {r.cost.priceEur.toFixed(2)}</span>:<span style={{color:T.orange,fontSize:"11px"}}>— mancante</span>}</TD>
                  <TD mono>{r.cost?FMT(r.cost.step1Eur):<span style={{color:T.dim}}>—</span>}</TD>
                  <TD mono bold>{r.cost?FMT(r.cost.step2Eur):<span style={{color:T.dim}}>—</span>}</TD>
                  <TD mono>{r.cost?FMT(r.cost.step2Hkd,2):<span style={{color:T.dim}}>—</span>}</TD>
                  <TD>{r.delta!==null?<DeltaBadge delta={r.delta} />:r.isNew?<Chip label="NUOVO" color={T.gold} />:<span style={{color:T.dim}}>—</span>}</TD>
                  <TD>{r.cost&&<MiniBtn label={expand===r.id?"▲":"▼"} onClick={()=>setExpand(expand===r.id?null:r.id)} />}</TD>
                </tr>
                {expand===r.id&&r.cost&&(
                  <tr key={r.id+"_exp"} style={{background:"rgba(255,255,255,0.02)"}}>
                    <td colSpan={11} style={{padding:"14px 20px"}}><CostBreakdown cost={r.cost} /></td>
                  </tr>
                )}
              </>
            ))}
          </tbody>
        </table>
      </Section>
    </div>
  );
}

function CostBreakdown({cost}) {
  const {whDetail}=cost;
  const rows=[
    ["Prezzo acquisto (× conv.factor)", cost.priceEur],
    ["FOB + Nolo",                      cost.fob],
    ["Local Import Charges (LIC)",      cost.lic],
    ["VGM Charges",                     cost.vgm],
    cost.hc>0&&["Health Certificate",   cost.hc],
    ["Costo Pallet",                    cost.plt],
    cost.alc>0&&["Tassa alcolici",      cost.alc],
    cost.carriageUnit>0&&["Carriage",   cost.carriageUnit],
    ["→ Standard Cost 1 (pre-WH)",      cost.step1Eur, true],
    whDetail.type==="MTO"&&["MTO — Cross docking", cost.wh],
    whDetail.type==="MTS"&&["MTS — Deposito",       whDetail.dep],
    whDetail.type==="MTS"&&["MTS — Inbound",        whDetail.inbound],
    whDetail.type==="MTS"&&["MTS — Picking",        whDetail.picking],
    whDetail.type==="FOR"&&["FOR — nessun costo WH",0],
    ["→ Standard Cost 2 (finale) €",   cost.step2Eur, true],
  ].filter(Boolean);
  return(
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"5px 24px"}}>
      {rows.map(([label,eur,bold],i)=>(
        <div key={i} style={{display:"flex",justifyContent:"space-between",fontSize:"12px",fontWeight:bold?"bold":"normal",color:bold?T.gold:T.muted,borderTop:bold?`1px solid ${T.border}`:"none",paddingTop:bold?"4px":0}}>
          <span>{label}</span>
          <span style={{fontFamily:"monospace"}}>€ {typeof eur==="number"?eur.toFixed(2):"—"}</span>
        </div>
      ))}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// MAIL GENERATOR
// ─────────────────────────────────────────────────────────────────────────────
function MailGen({costRows,branch,month,branchCfg,sentMails,setSentMails,showToast}) {
  const [sent,setSent]=useState(false);
  const [viewing,setViewing]=useState(null);
  const flagged=costRows.filter(r=>r.flagged&&r.cost);
  const ups=flagged.filter(r=>r.delta>0);
  const downs=flagged.filter(r=>r.delta<0);
  const newItems=costRows.filter(r=>r.isNew&&r.cost);
  const monthLabel=new Date(month+"-01").toLocaleDateString("it-IT",{month:"long",year:"numeric"});
  const handleSend=()=>{
    setSentMails(ms=>[{id:Date.now(),branch,month,sentAt:new Date().toISOString(),newCount:newItems.length,changedCount:flagged.length},...ms]);
    setSent(true);showToast("Mail registrata ✓",T.gold);
  };
  if(viewing!==null){
    const m=sentMails[viewing];
    return(<div><PageHeader title="✉ Mail storico" sub={`${m.branch} · ${m.month}`} /><ActionBtn label="← Torna" onClick={()=>setViewing(null)} /><div style={{marginTop:"16px",padding:"16px",background:T.card,borderRadius:"8px",fontSize:"12px",color:T.muted}}>Inviata {new Date(m.sentAt).toLocaleString("it-IT")} · {m.newCount} nuovi · {m.changedCount} variazioni</div></div>);
  }
  return(
    <div>
      <PageHeader title="✉ Mail Mensile" sub={`Standard Cost Update · ${branchCfg.flag} ${branchCfg.label} · ${monthLabel}`} />
      <Section title="Anteprima comunicazione" mb="24px">
        <div style={{background:"#fff",borderRadius:"10px",padding:"40px 48px",color:"#1A1A1A",fontFamily:"'Helvetica Neue',Arial,sans-serif",maxWidth:"700px",boxShadow:"0 4px 32px rgba(0,0,0,0.4)"}}>
          <div style={{borderBottom:"3px solid #1A2B3C",paddingBottom:"20px",marginBottom:"28px",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
            <div>
              <div style={{fontSize:"11px",letterSpacing:"3px",color:"#C9A84C",textTransform:"uppercase",marginBottom:"4px"}}>Inalca Food & Beverage</div>
              <div style={{fontSize:"22px",fontWeight:"bold",color:"#1A2B3C"}}>Standard Cost Update</div>
              <div style={{fontSize:"13px",color:"#666",marginTop:"3px"}}>{branchCfg.flag} {branchCfg.label} · {monthLabel}</div>
            </div>
            <div style={{fontSize:"36px"}}>{branchCfg.flag}</div>
          </div>
          <p style={{fontSize:"13px",color:"#333",lineHeight:"1.7",marginBottom:"24px"}}>Si comunica l'aggiornamento mensile degli <strong>Standard Cost</strong> per la filiale di <strong>{branchCfg.label}</strong> con decorrenza <strong>{monthLabel}</strong>.</p>
          {newItems.length>0&&(<div style={{marginBottom:"28px"}}><div style={{fontSize:"12px",fontWeight:"bold",color:"#C9A84C",textTransform:"uppercase",letterSpacing:"2px",marginBottom:"12px"}}>✦ Nuovi Prodotti ({newItems.length})</div><MailTable rows={newItems} currency={branchCfg.currency} showPrev={false} /></div>)}
          {ups.length>0&&(<div style={{marginBottom:"28px"}}><div style={{fontSize:"12px",fontWeight:"bold",color:"#B5534A",textTransform:"uppercase",letterSpacing:"2px",marginBottom:"12px"}}>↑ Aumenti ≥ +3% ({ups.length})</div><MailTable rows={ups} currency={branchCfg.currency} showPrev={true} /></div>)}
          {downs.length>0&&(<div style={{marginBottom:"28px"}}><div style={{fontSize:"12px",fontWeight:"bold",color:"#2D7A50",textTransform:"uppercase",letterSpacing:"2px",marginBottom:"12px"}}>↓ Riduzioni ≥ -3% ({downs.length})</div><MailTable rows={downs} currency={branchCfg.currency} showPrev={true} /></div>)}
          {newItems.length===0&&flagged.length===0&&(<div style={{padding:"20px",background:"#F5F5F0",borderRadius:"8px",textAlign:"center",color:"#666",fontSize:"13px"}}>Nessuna variazione significativa questo mese.</div>)}
          <div style={{borderTop:"1px solid #E8E8E0",paddingTop:"18px",marginTop:"28px",fontSize:"11px",color:"#999"}}>Generato da <strong style={{color:"#1A2B3C"}}>IFB Cost Intelligence Platform</strong> · {new Date().toLocaleDateString("it-IT")}</div>
        </div>
      </Section>
      <div style={{display:"flex",gap:"12px"}}>
        <ActionBtn label={sent?"✓ Registrata":"✉ Registra invio"} onClick={handleSend} primary disabled={sent} />
      </div>
      {sentMails.length>0&&(
        <Section title="Storico mail" mt="28px">
          <table style={{width:"100%",borderCollapse:"collapse"}}>
            <THead cols={["Filiale","Mese","Data","Nuovi","Var.","Azione"]} />
            <tbody>
              {sentMails.map((m,i)=>(
                <tr key={m.id} style={{borderBottom:`1px solid ${T.border}`}}>
                  <TD>{BRANCH_CFG[m.branch]?.flag} {m.branch}</TD><TD mono>{m.month}</TD>
                  <TD mono>{new Date(m.sentAt).toLocaleDateString("it-IT")}</TD>
                  <TD mono>{m.newCount}</TD><TD mono>{m.changedCount}</TD>
                  <TD><MiniBtn label="Vedi" onClick={()=>setViewing(i)} /></TD>
                </tr>
              ))}
            </tbody>
          </table>
        </Section>
      )}
    </div>
  );
}

function MailTable({rows,currency,showPrev}) {
  const cols=showPrev?["N HK","IFB N","Descrizione","Ubic.","SC2 prec. €","SC2 nuovo €","Δ %"]:["N HK","IFB N","Descrizione","Temp.","Ubic.","SC2 (€)",`SC2 (${currency})`];
  return(
    <table style={{width:"100%",borderCollapse:"collapse",fontSize:"12px"}}>
      <thead><tr style={{background:"#F5F5F0"}}>{cols.map(h=><th key={h} style={{padding:"8px 10px",textAlign:"left",color:"#444",fontWeight:"bold",fontSize:"11px"}}>{h}</th>)}</tr></thead>
      <tbody>
        {rows.map((r,i)=>(
          <tr key={r.id} style={{background:i%2===0?"#fff":"#FAFAF8",borderBottom:"1px solid #E8E8E0"}}>
            <td style={{padding:"8px 10px",fontFamily:"monospace",fontSize:"11px",color:"#C9A84C"}}>{r.nHK||"—"}</td>
            <td style={{padding:"8px 10px",fontFamily:"monospace",fontSize:"11px"}}>{r.code}</td>
            <td style={{padding:"8px 10px"}}>{r.description}</td>
            {showPrev?(
              <>
                <td style={{padding:"8px 10px",fontSize:"11px"}}>{r.ubicazione||"—"}</td>
                <td style={{padding:"8px 10px",fontFamily:"monospace"}}>€ {FMT(r.prevCost?.step2Eur)}</td>
                <td style={{padding:"8px 10px",fontFamily:"monospace",fontWeight:"bold"}}>€ {FMT(r.cost?.step2Eur)}</td>
                <td style={{padding:"8px 10px"}}><span style={{padding:"2px 8px",borderRadius:"4px",fontSize:"11px",fontWeight:"bold",background:r.delta>0?"#FDECEA":"#E8F5EE",color:r.delta>0?"#B5534A":"#2D7A50"}}>{r.delta>0?"+":""}{r.delta?.toFixed(1)}%</span></td>
              </>
            ):(
              <>
                <td style={{padding:"8px 10px"}}>{r.temperature}</td>
                <td style={{padding:"8px 10px",fontSize:"11px"}}>{r.ubicazione||"—"}</td>
                <td style={{padding:"8px 10px",fontFamily:"monospace"}}>€ {FMT(r.cost?.step2Eur)}</td>
                <td style={{padding:"8px 10px",fontFamily:"monospace"}}>{currency} {FMT(r.cost?.step2Hkd,2)}</td>
              </>
            )}
          </tr>
        ))}
      </tbody>
    </table>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// SALES INVOICE — stato sollevato a App level
// ─────────────────────────────────────────────────────────────────────────────
function SalesInvoice({ products, prices, branch, rows, setRows, snapshots, setSnapshots, importLogs, setImportLogs, showToast }) {
  const [search,  setSearch]  = useState("");
  const [secFilter, setSecFilter] = useState("ALL");
  const [step,    setStep]    = useState("main");
  const [headers, setHeaders] = useState([]);
  const [rawRows, setRawRows] = useState([]);
  const [mapping, setMapping] = useState({});
  const [preview, setPreview] = useState([]);
  const [fileName, setFileName] = useState("");

  const ALIASES = {
    ifbNo:       ["no_","no.","no","item no","item no.","ifb no","ifb n","codice"],
    description: ["description","descrizione","desc","item description"],
    vendor:      ["vendor name","vendor","fornitore","vendor name 3","vendor name 2","italia alimentari"],
    lastDate:    ["last posting date","last date","data","posting date","last invoice date","data fattura"],
    quantity:    ["quantity","qty","quantità","qt"],
    price:       ["price","prezzo","unit price","prezzo unitario"],
    location:    ["location code","location","magazzino","loc"],
    section:     ["section description","section","sezione","categoria"],
  };

  const autoDetect = (hdrs) => {
    const m = {};
    Object.keys(ALIASES).forEach(field => {
      for (const h of hdrs) { const hl=h.toLowerCase().trim(); if(ALIASES[field].some(a=>hl===a)){m[field]=h;break;} }
      if(!m[field]){for(const h of hdrs){const hl=h.toLowerCase().trim();if(ALIASES[field].some(a=>a.length>3&&hl.includes(a))){m[field]=h;break;}}}
    });
    return m;
  };

  const handleFile = (e) => {
    const file = e.target.files[0]; if (!file) return;
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = ev => {
      const wb = XLSX.read(ev.target.result, { type:"binary" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(ws, { header:1, defval:"" });
      if (data.length < 2) return;
      const hdrs = data[0].map(String);
      const raws = data.slice(1).filter(r => r.some(c => c !== ""));
      setHeaders(hdrs); setRawRows(raws);
      setMapping(autoDetect(hdrs)); setStep("map");
    };
    reader.readAsBinaryString(file);
  };

  const get = (row, field) => { const col=mapping[field];if(!col)return"";const i=headers.indexOf(col);return i>=0?row[i]:""; };

  const buildPreview = () => {
    const p = rawRows.map((row, idx) => {
      const ifbNo = String(get(row,"ifbNo")||"").trim();
      if (!ifbNo) return null;
      const prod = products.find(p => p.code === ifbNo);
      const rawDate = String(get(row,"lastDate")||"").trim();
      const dateObj = rawDate ? new Date(rawDate) : null;
      const dateStr = dateObj && !isNaN(dateObj) ? dateObj.toISOString().slice(0,10) : rawDate;
      return {
        _idx:idx, ifbNo,
        nHKDisplay: prod ? (prod.nHK || prod.code) : null,
        description: String(get(row,"description")||prod?.description||ifbNo),
        vendor:    String(get(row,"vendor")||"").trim(),
        lastDate:  dateStr,
        quantity:  parseFloat(String(get(row,"quantity")||"").replace(",","."))||0,
        price:     Math.round((parseFloat(String(get(row,"price")||"").replace(",","."))||0)*100)/100,
        location:  String(get(row,"location")||"").trim(),
        section:   String(get(row,"section")||prod?.category||"").trim(),
        _found:    !!prod,
      };
    }).filter(Boolean);
    setPreview(p); setStep("preview");
  };

  const executeImport = () => {
    // Calcola diff rispetto ai rows esistenti
    const diffs = [];
    preview.forEach(r => {
      const existing = rows.find(e => e.ifbNo === r.ifbNo);
      if (!existing) {
        diffs.push({ ifbNo: r.ifbNo, nHK: r.nHKDisplay||"—", description: r.description, isNew: true, fields: [] });
      } else {
        const diffFields = [];
        [["price","Prezzo"],["lastDate","Ultima data"],["quantity","Qty"],["vendor","Fornitore"]].forEach(([f,label])=>{
          if (String(existing[f]||"") !== String(r[f]||"")) {
            diffFields.push({field:label, old:existing[f], new:r[f], delta: f==="price"&&existing[f]>0?PCT(r[f],existing[f]):null});
          }
        });
        if (diffFields.length > 0) diffs.push({ ifbNo: r.ifbNo, nHK: r.nHKDisplay||"—", description: r.description, isNew: false, fields: diffFields });
      }
    });

    const incoming = preview.map(r => ({
      ifbNo:r.ifbNo, nHK:r.nHKDisplay, description:r.description,
      vendor:r.vendor, lastDate:r.lastDate, quantity:r.quantity,
      price:r.price, location:r.location, section:r.section, _found:r._found,
    }));
    const kept = rows.filter(r => !incoming.find(i => i.ifbNo === r.ifbNo));
    setRows([...incoming, ...kept]);

    const logEntry = {
      id: Date.now(), type:"sales", fileName, branch, importedAt: new Date().toISOString(),
      count: incoming.length, newCount: diffs.filter(d=>d.isNew).length,
      updateCount: diffs.filter(d=>!d.isNew&&d.fields.length>0).length, diffs,
    };
    setImportLogs(logs => [logEntry, ...logs]);
    setSnapshots(snaps => [logEntry, ...snaps].slice(0, 50));
    if (showToast) showToast(`Sales Invoice: ${incoming.length} righe importate ✓`, T.gold);
    setStep("main");
    setPreview([]); setRawRows([]); setHeaders([]); setMapping({}); setFileName("");
  };

  const sections = ["ALL", ...Array.from(new Set(rows.map(r=>r.section||"—"))).sort()];
  const displayed = rows
    .filter(r=>{
      const s=search.toLowerCase();
      const matchS=!search||r.ifbNo?.toLowerCase().includes(s)||r.description?.toLowerCase().includes(s)||r.vendor?.toLowerCase().includes(s);
      const matchSec=secFilter==="ALL"||r.section===secFilter;
      return matchS&&matchSec;
    })
    .sort((a,b)=>(b.lastDate||"").localeCompare(a.lastDate||""));

  if (step==="upload") return (
    <div>
      <PageHeader title="📋 Sales on Invoice — IFB" sub="Carica report fatture" />
      <Section title="Carica file"><DropZone onFile={f=>{setFileName(f.name);const r=new FileReader();r.onload=ev=>{const wb=XLSX.read(ev.target.result,{type:"binary"});const ws=wb.Sheets[wb.SheetNames[0]];const data=XLSX.utils.sheet_to_json(ws,{header:1,defval:""});if(data.length<2)return;const hdrs=data[0].map(String);const raws=data.slice(1).filter(r=>r.some(c=>c!==""));setHeaders(hdrs);setRawRows(raws);setMapping(autoDetect(hdrs));setStep("map");};r.readAsBinaryString(f);}} /></Section>
      <div style={{marginTop:"12px"}}><ActionBtn label="← Annulla" onClick={()=>setStep("main")} /></div>
    </div>
  );

  if (step==="map") return (
    <div>
      <PageHeader title="📋 Sales on Invoice — Mappatura" sub={fileName} />
      <Section title="Mappatura colonne" mb="20px">
        <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:"12px"}}>
          {Object.keys(ALIASES).map(field=>(
            <div key={field}>
              <div style={{fontSize:"10px",color:T.gold,marginBottom:"4px",textTransform:"uppercase",letterSpacing:"1px"}}>{field}</div>
              <select value={mapping[field]||""} onChange={e=>setMapping(m=>({...m,[field]:e.target.value}))} style={{width:"100%",background:T.surface,border:`1px solid ${mapping[field]?T.gold:T.border}`,borderRadius:"6px",padding:"7px 10px",color:mapping[field]?T.text:T.dim,fontSize:"12px",fontFamily:"inherit"}}>
                <option value="">— non mappare —</option>
                {headers.map(h=><option key={h} value={h}>{h}</option>)}
              </select>
            </div>
          ))}
        </div>
      </Section>
      <div style={{display:"flex",gap:"8px"}}>
        <ActionBtn label="← Indietro" onClick={()=>setStep("upload")} />
        <ActionBtn label="Preview →" onClick={buildPreview} primary disabled={!mapping["ifbNo"]} />
      </div>
    </div>
  );

  if (step==="preview") return (
    <div>
      <PageHeader title="📋 Sales on Invoice — Preview" sub={`${preview.length} righe · ${preview.filter(r=>r._found).length} trovati in anagrafica`} />
      <div style={{display:"flex",gap:"12px",marginBottom:"20px"}}>
        {[[preview.filter(r=>r._found).length,"In anagrafica",T.green],[preview.filter(r=>!r._found).length,"Non trovati",T.orange],[preview.length,"Totale",T.text]].map(([n,l,c])=>(
          <div key={l} style={{padding:"10px 16px",background:T.card,border:`1px solid ${T.border}`,borderRadius:"8px"}}><div style={{fontSize:"18px",fontWeight:"bold",color:c}}>{n}</div><div style={{fontSize:"10px",color:T.dim,marginTop:"2px"}}>{l}</div></div>
        ))}
      </div>
      <Section title="Preview (max 50)">
        <table style={{width:"100%",borderCollapse:"collapse"}}>
          <THead cols={["IFB N","N HK","Descrizione","Fornitore","Data","Qty","Prezzo","Sezione","✓"]} />
          <tbody>
            {preview.slice(0,50).map(r=>(
              <tr key={r._idx} style={{borderBottom:`1px solid ${T.border}`,background:r._found?"":` ${T.orange}08`}}>
                <TD mono>{r.ifbNo}</TD>
                <TD mono><span style={{color:T.gold}}>{r.nHKDisplay||"—"}</span></TD>
                <TD>{r.description}</TD>
                <TD style={{fontSize:"11px"}}>{r.vendor}</TD>
                <TD mono>{r.lastDate}</TD>
                <TD mono>{r.quantity}</TD>
                <TD mono>{r.price.toFixed(2)}</TD>
                <TD><Chip label={r.section||"—"} color={T.blue} /></TD>
                <TD>{r._found?<span style={{color:T.green}}>✓</span>:<span style={{color:T.orange}}>?</span>}</TD>
              </tr>
            ))}
          </tbody>
        </table>
      </Section>
      <div style={{display:"flex",gap:"8px",marginTop:"14px"}}>
        <ActionBtn label="← Mappatura" onClick={()=>setStep("map")} />
        <ActionBtn label={`✓ Importa ${preview.length} righe`} onClick={executeImport} primary />
      </div>
    </div>
  );

  return (
    <div>
      <PageHeader title="📋 Sales on Invoice — IFB" sub="Fatture ordinate per data decrescente" />
      <div style={{display:"flex",gap:"8px",marginBottom:"16px",alignItems:"center"}}>
        <ActionBtn label="⬆ Carica / Aggiorna" onClick={()=>setStep("upload")} primary />
        {rows.length>0&&<span style={{fontSize:"11px",color:T.dim}}>{rows.length} righe</span>}
        {rows.length>0&&<MiniBtn label="✕ Svuota" onClick={()=>{if(window.confirm("Eliminare tutti i dati?"))setRows([]);}} color={T.red} />}
      </div>
      {rows.length===0&&(
        <Section title="Nessun dato"><div style={{padding:"32px",textAlign:"center",color:T.dim,fontSize:"13px"}}>Carica il report Sales on Invoice da Business Central.</div></Section>
      )}
      {rows.length>0&&(
        <>
          <div style={{display:"flex",gap:"6px",marginBottom:"12px",flexWrap:"wrap"}}>
            {sections.map(s=>(
              <button key={s} onClick={()=>setSecFilter(s)} style={{padding:"4px 12px",background:secFilter===s?T.goldDim:"rgba(255,255,255,0.04)",border:`1px solid ${secFilter===s?T.gold:T.border}`,borderRadius:"5px",color:secFilter===s?T.gold:T.muted,cursor:"pointer",fontFamily:"inherit",fontSize:"11px"}}>{s}</button>
            ))}
          </div>
          <SearchBar value={search} onChange={setSearch} placeholder="🔍  Filtra per IFB N, descrizione o fornitore…" />
          {(()=>{
            const withXref=displayed.filter(r=>products.find(p=>p.code===r.ifbNo));
            const withStdCost=withXref.filter(r=>{const p=products.find(pp=>pp.code===r.ifbNo);if(!p)return false;const pp=prices.filter(pr=>pr.productId===p.id&&pr.branch===branch);return pp.length>0&&(pp[0].dapFinal>0||pp[0].mtsPrice>0||pp[0].fcaDiscounted>0);});
            return(<div style={{display:"flex",gap:"10px",marginBottom:"14px"}}>
              {[[displayed.length,"Righe",T.text],[withXref.length,"In anagrafica",T.green],[withStdCost.length,"Con Std Cost",T.blue],[withXref.length-withStdCost.length,"Std Cost mancante",withXref.length-withStdCost.length>0?T.orange:T.dim]].map(([n,l,c])=>(
                <div key={l} style={{padding:"8px 14px",background:T.card,border:`1px solid ${c===T.text||c===T.dim?T.border:c+"44"}`,borderRadius:"7px",minWidth:"100px"}}>
                  <div style={{fontSize:"18px",fontWeight:"bold",color:c}}>{n}</div>
                  <div style={{fontSize:"9px",color:T.dim,marginTop:"2px",textTransform:"uppercase",letterSpacing:"0.5px"}}>{l}</div>
                </div>
              ))}
            </div>);
          })()}
          <Section title={`${displayed.length} fatture`}>
            <table style={{width:"100%",borderCollapse:"collapse"}}>
              <THead cols={["IFB N","N HK","Descrizione","Fornitore","Ultima Fattura","Qty","Prezzo €","Location","Sezione","Std Cost €","Ultimo SC","Xref"]} />
              <tbody>
                {displayed.map((r,i)=>{
                  const prod=products.find(p=>p.code===r.ifbNo);
                  const prodPrices=prod?prices.filter(p=>p.productId===prod.id&&p.branch===branch).sort((a,b)=>b.month.localeCompare(a.month)):[];
                  const latestPrice=prodPrices[0]||null;
                  const hasStdCost=latestPrice&&(latestPrice.dapFinal>0||latestPrice.mtsPrice>0||latestPrice.fcaDiscounted>0);
                  const stdCostVal=latestPrice?(latestPrice.dapFinal||latestPrice.mtsPrice||latestPrice.fcaDiscounted||0):0;
                  return(
                    <tr key={r.ifbNo+"_"+i} style={{borderBottom:`1px solid ${T.border}`,background:prod?"":` ${T.orange}07`}}>
                      <TD mono>{r.ifbNo}</TD>
                      <TD mono><span style={{color:T.gold}}>{prod?.nHK||"—"}</span></TD>
                      <td style={{padding:"7px 9px",fontSize:"12px",color:T.text,maxWidth:"200px"}}>{r.description}</td>
                      <td style={{padding:"7px 9px",fontSize:"11px",color:T.muted,maxWidth:"130px",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.vendor}</td>
                      <TD mono>{r.lastDate}</TD>
                      <TD mono>{r.quantity}</TD>
                      <TD mono>{r.price.toFixed(2)}</TD>
                      <TD><Chip label={r.location||"—"} color={r.location==="NCJ"?T.purple:T.blue} /></TD>
                      <TD><Chip label={r.section||"—"} color={T.blue} /></TD>
                      <TD mono bold>{hasStdCost?<span style={{color:T.green}}>€ {stdCostVal.toFixed(2)}</span>:<span style={{color:T.orange,fontSize:"11px"}}>— mancante</span>}</TD>
                      <TD mono><span style={{color:T.muted,fontSize:"11px"}}>{latestPrice?latestPrice.month:"—"}</span></TD>
                      <TD>{prod?<span style={{color:T.green,fontSize:"11px"}}>✓</span>:<span style={{color:T.orange,fontSize:"11px"}}>⚠</span>}</TD>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </Section>
        </>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// STORICO & DIFF
// ─────────────────────────────────────────────────────────────────────────────
function Storico({ snapshots, setSnapshots }) {
  const [selected, setSelected] = useState(null);
  const [typeFilter, setTypeFilter] = useState("ALL");
  const [search, setSearch] = useState("");

  const filtered = snapshots.filter(s => typeFilter==="ALL" || s.type===typeFilter);
  const snap = selected !== null ? snapshots.find(s=>s.id===selected) : null;

  const FIELD_LABELS = {
    dapFinal:"DAP Final",mtsPrice:"MTS Price",fcaDiscounted:"FCA Discounted",
    dapPrice:"DAP Price",fcaPrice:"FCA Price",description:"Descrizione",
    category:"Categoria",uom:"UOM",qtyPerBox:"Qty/cartone",boxPerPallet:"Cartoni/plt",
    kgPerBox:"Kg/cartone",temperature:"Temperatura",active:"Attivo",nHK:"N HK",
    "Prezzo":"Prezzo","Ultima data":"Ultima data","Qty":"Qty","Fornitore":"Fornitore",
  };

  const filteredDiffs = snap?.diffs?.filter(d=>
    !search||d.nHK?.toLowerCase().includes(search.toLowerCase())||
    d.ifbNo?.toLowerCase().includes(search.toLowerCase())||
    d.description?.toLowerCase().includes(search.toLowerCase())
  )||[];
  const changedOnly = filteredDiffs.filter(d=>d.fields?.length>0||d.isNew);

  const typeLabel = (t) => t==="prices"?"Listini":t==="anagrafica"?"Anagrafica":t==="sales"?"Sales Inv.":t==="xref"?"XRef":"?";
  const typeColor = (t) => t==="prices"?T.purple:t==="anagrafica"?T.blue:t==="sales"?T.green:t==="xref"?T.gold:T.muted;

  // Branch label per snapshot
  const snapBranch = (s) => {
    if (s.branch) return `${BRANCH_CFG[s.branch]?.flag||""} ${s.branch}`;
    return "—";
  };

  return (
    <div>
      <PageHeader title="⧖ Storico & Diff" sub="Confronto tra import successivi — campi modificati evidenziati" />
      {snapshots.length===0&&(
        <Section title="Nessun import registrato">
          <div style={{padding:"32px",textAlign:"center",color:T.dim,fontSize:"13px"}}>Gli snapshot si creano automaticamente dopo ogni import (Anagrafica, Listini, Sales Invoice, XRef).</div>
        </Section>
      )}
      {snapshots.length>0&&!snap&&(
        <div>
          <div style={{display:"flex",gap:"8px",marginBottom:"16px"}}>
            {[["ALL","Tutti"],["anagrafica","Anagrafica"],["prices","Listini"],["sales","Sales Inv."],["xref","XRef"]].map(([k,l])=>(
              <button key={k} onClick={()=>setTypeFilter(k)} style={{padding:"6px 14px",background:typeFilter===k?T.goldDim:"rgba(255,255,255,0.04)",border:`1px solid ${typeFilter===k?T.gold:T.border}`,borderRadius:"6px",color:typeFilter===k?T.gold:T.muted,cursor:"pointer",fontFamily:"inherit",fontSize:"12px"}}>{l}</button>
            ))}
          </div>
          <Section title={`${filtered.length} import registrati`}>
            <table style={{width:"100%",borderCollapse:"collapse"}}>
              <THead cols={["Tipo","File","Filiale/Mese","Data","Nuovi","Aggiornati","Con diff","Azione"]} />
              <tbody>
                {filtered.map(s=>{
                  const changed=s.diffs?.filter(d=>d.fields?.length>0||d.isNew).length||0;
                  return(
                    <tr key={s.id} style={{borderBottom:`1px solid ${T.border}`}}>
                      <TD><Chip label={typeLabel(s.type)} color={typeColor(s.type)} /></TD>
                      <td style={{padding:"9px 10px",fontSize:"11px",fontFamily:"monospace",color:T.text,maxWidth:"180px",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{s.fileName}</td>
                      <TD mono>
                        {s.type==="prices"?`${BRANCH_CFG[s.branch]?.flag||""} ${s.branch} · ${s.month}`:
                         s.type==="sales"?snapBranch(s):
                         s.type==="anagrafica"?snapBranch(s):"—"}
                      </TD>
                      <TD mono>{new Date(s.importedAt).toLocaleString("it-IT",{day:"2-digit",month:"2-digit",year:"2-digit",hour:"2-digit",minute:"2-digit"})}</TD>
                      <TD><Chip label={String(s.newCount||0)} color={T.gold} /></TD>
                      <TD><Chip label={String(s.updateCount||s.count||0)} color={T.blue} /></TD>
                      <TD>{changed>0?<Chip label={String(changed)} color={T.red} />:<span style={{color:T.dim,fontSize:"12px"}}>—</span>}</TD>
                      <TD>
                        <div style={{display:"flex",gap:"6px"}}>
                          <MiniBtn label="🔍 Diff" onClick={()=>setSelected(s.id)} color={T.gold} />
                          <MiniBtn label="✕" onClick={()=>setSnapshots(ss=>ss.filter(x=>x.id!==s.id))} color={T.red} />
                        </div>
                      </TD>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </Section>
        </div>
      )}
      {snap&&(
        <div>
          <div style={{display:"flex",alignItems:"center",gap:"12px",marginBottom:"20px"}}>
            <ActionBtn label="← Torna" onClick={()=>{setSelected(null);setSearch("");}} />
            <div style={{fontSize:"13px",color:T.muted}}>
              <Chip label={typeLabel(snap.type)} color={typeColor(snap.type)} />
              &nbsp;<strong style={{color:T.text}}>{snap.fileName}</strong>
              &nbsp;·&nbsp;{new Date(snap.importedAt).toLocaleString("it-IT")}
              {snap.branch&&<>&nbsp;·&nbsp;{snapBranch(snap)}</>}
              {snap.month&&<>&nbsp;·&nbsp;{snap.month}</>}
            </div>
          </div>
          <div style={{display:"flex",gap:"12px",marginBottom:"20px"}}>
            {[[snap.diffs?.filter(d=>d.isNew).length||0,"Nuovi",T.gold],[snap.diffs?.filter(d=>!d.isNew&&(d.fields?.length>0||d.changed)).length||0,"Modificati",T.orange],[snap.diffs?.length||0,"Totale processati",T.text]].map(([n,l,c])=>(
              <div key={l} style={{padding:"10px 16px",background:T.card,border:`1px solid ${T.border}`,borderRadius:"8px"}}><div style={{fontSize:"20px",fontWeight:"bold",color:c}}>{n}</div><div style={{fontSize:"10px",color:T.dim,marginTop:"2px"}}>{l}</div></div>
            ))}
          </div>
          <div style={{marginBottom:"12px"}}><input placeholder="🔍  Cerca per N HK, IFB N o descrizione…" value={search} onChange={e=>setSearch(e.target.value)} style={inputStyle()} /></div>
          {changedOnly.length===0&&(<Section title="Nessuna modifica"><div style={{padding:"20px",textAlign:"center",color:T.dim,fontSize:"13px"}}>Tutti i valori erano identici.</div></Section>)}
          {changedOnly.length>0&&(
            <Section title={`${changedOnly.length} voci con modifiche`}>
              <table style={{width:"100%",borderCollapse:"collapse"}}>
                <THead cols={["N HK","IFB N","Descrizione","Stato","Campo","Valore prec.","Valore nuovo","Δ %"]} />
                <tbody>
                  {changedOnly.map(d=>(
                    d.isNew?(
                      <tr key={(d.nHK||d.ifbNo||d.code)+"_new"} style={{borderBottom:`1px solid ${T.border}`,background:`${T.gold}07`}}>
                        <TD mono><span style={{color:T.gold}}>{d.nHK||"—"}</span></TD>
                        <TD mono>{d.ifbNo||d.code||"—"}</TD>
                        <TD>{d.description}</TD>
                        <td colSpan={5} style={{padding:"8px 10px"}}><Chip label="NUOVO" color={T.gold} /></td>
                      </tr>
                    ):(d.fields||[]).map((f,fi)=>(
                      <tr key={(d.nHK||d.ifbNo||d.code)+"_"+fi} style={{borderBottom:fi===(d.fields.length-1)?`1px solid ${T.border}`:`1px solid ${T.border}22`,background:`${T.orange}05`}}>
                        {fi===0&&<TD mono style={{verticalAlign:"top"}} rowSpan={d.fields.length}><span style={{color:T.gold}}>{d.nHK||"—"}</span></TD>}
                        {fi===0&&<TD mono style={{verticalAlign:"top"}} rowSpan={d.fields.length}>{d.ifbNo||d.code||"—"}</TD>}
                        {fi===0&&<TD style={{verticalAlign:"top"}} rowSpan={d.fields.length}>{d.description}</TD>}
                        {fi===0&&<td style={{padding:"8px 10px",verticalAlign:"top"}} rowSpan={d.fields.length}><Chip label="MODIF." color={T.orange} /></td>}
                        <TD style={{fontSize:"11px",color:T.muted}}>{FIELD_LABELS[f.field]||f.field}</TD>
                        <TD mono style={{textDecoration:"line-through",color:T.dim}}>{typeof f.old==="number"?f.old.toFixed(2):String(f.old??"—")}</TD>
                        <TD mono bold>{typeof f.new==="number"?f.new.toFixed(2):String(f.new??"—")}</TD>
                        <TD>{f.delta!=null?<DeltaBadge delta={f.delta} small />:<span style={{color:T.dim}}>—</span>}</TD>
                      </tr>
                    ))
                  ))}
                </tbody>
              </table>
            </Section>
          )}
        </div>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// SHARED COMPONENTS
// ─────────────────────────────────────────────────────────────────────────────
function StepBar({steps,current}) {
  const idx=steps.indexOf(current);
  return(
    <div style={{display:"flex",alignItems:"center",gap:0,marginBottom:"28px"}}>
      {steps.map((s,i)=>(
        <div key={s} style={{display:"flex",alignItems:"center"}}>
          <div style={{display:"flex",flexDirection:"column",alignItems:"center",gap:"4px"}}>
            <div style={{width:"28px",height:"28px",borderRadius:"50%",background:current===s?T.gold:i<idx?`${T.gold}44`:"rgba(255,255,255,0.06)",border:`2px solid ${current===s?T.gold:i<idx?`${T.gold}66`:T.border}`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:"11px",fontWeight:"bold",color:current===s?T.bg:i<idx?T.gold:T.dim}}>{i+1}</div>
            <span style={{fontSize:"10px",color:current===s?T.gold:T.dim,whiteSpace:"nowrap"}}>{s==="upload"?"1. Carica":s==="map"?"2. Mappa":s==="preview"?"3. Preview":"4. Fine"}</span>
          </div>
          {i<steps.length-1&&<div style={{width:"60px",height:"1px",background:T.border,margin:"0 4px",marginBottom:"18px"}} />}
        </div>
      ))}
    </div>
  );
}
function DropZone({onFile}) {
  return(
    <div onDragOver={e=>e.preventDefault()} onDrop={e=>{e.preventDefault();const f=e.dataTransfer.files[0];if(f)onFile(f);}} style={{border:`2px dashed ${T.borderHi}`,borderRadius:"12px",padding:"48px 32px",textAlign:"center",background:"rgba(255,255,255,0.02)",cursor:"pointer"}} onClick={()=>document.getElementById("_dropzone_input").click()}>
      <div style={{fontSize:"40px",marginBottom:"12px"}}>⇪</div>
      <div style={{fontSize:"15px",color:T.text,marginBottom:"6px"}}>Trascina qui il file Excel o CSV</div>
      <div style={{fontSize:"12px",color:T.muted,marginBottom:"18px"}}>Supportati: .xlsx, .xls, .csv — Prima riga = intestazioni</div>
      <ActionBtn label="Sfoglia file…" onClick={e=>{e.stopPropagation();document.getElementById("_dropzone_input").click();}} primary />
      <input id="_dropzone_input" type="file" accept=".xlsx,.xls,.csv" onChange={e=>{const f=e.target.files[0];if(f)onFile(f);e.target.value="";}} style={{display:"none"}} />
    </div>
  );
}
function SearchBar({value,onChange,placeholder="🔍  Filtra…"}) {
  return(
    <div style={{marginBottom:"14px"}}>
      <input value={value} onChange={e=>onChange(e.target.value)} placeholder={placeholder} style={{...inputStyle(),width:"280px",padding:"7px 12px",fontSize:"12px"}} />
      {value&&<button onClick={()=>onChange("")} style={{marginLeft:"6px",background:"none",border:"none",color:T.dim,cursor:"pointer",fontSize:"14px",lineHeight:1}}>✕</button>}
    </div>
  );
}
function PageHeader({title,sub}){return(<div style={{marginBottom:"16px"}}><h1 style={{fontSize:"18px",fontWeight:"bold",margin:0,marginBottom:"3px",color:T.text}}>{title}</h1><div style={{fontSize:"11px",color:T.muted}}>{sub}</div></div>);}
function Section({title,accent,children,mb,mt}){return(<div style={{marginBottom:mb||"0",marginTop:mt||"0"}}>{title&&<div style={{fontSize:"9px",letterSpacing:"2px",textTransform:"uppercase",color:accent||T.gold,marginBottom:"8px",display:"flex",alignItems:"center",gap:"6px"}}><div style={{height:"1px",width:"14px",background:accent||T.gold,opacity:0.5}} />{title}</div>}<div style={{background:T.card,border:`1px solid ${T.border}`,borderRadius:"8px",padding:"14px"}}>{children}</div></div>);}
function KPI({label,value,color,icon}){return(<div style={{background:T.card,border:`1px solid ${T.border}`,borderRadius:"10px",padding:"20px"}}><div style={{fontSize:"24px",marginBottom:"4px"}}>{icon}</div><div style={{fontSize:"28px",fontWeight:"bold",color}}>{value}</div><div style={{fontSize:"11px",color:T.muted,marginTop:"2px"}}>{label}</div></div>);}
function THead({cols}){return(<thead><tr>{cols.map(c=><th key={c} style={{padding:"6px 9px",textAlign:"left",fontSize:"9px",letterSpacing:"1px",textTransform:"uppercase",color:T.dim,borderBottom:`1px solid ${T.border}`,fontWeight:"normal",whiteSpace:"nowrap"}}>{c}</th>)}</tr></thead>);}
function TD({children,mono,bold,style:s,rowSpan}){return <td rowSpan={rowSpan} style={{padding:"7px 9px",fontSize:"12px",fontFamily:mono?"monospace":"inherit",fontWeight:bold?"bold":"normal",color:T.text,...s}}>{children}</td>;}
function Chip({label,color}){const c=color||(label==="FOOD"?T.gold:label==="WINE"?T.purple:label==="MEAT"?T.red:T.muted);return <span style={{padding:"2px 7px",borderRadius:"4px",fontSize:"10px",background:`${c}22`,color:c,border:`1px solid ${c}33`,letterSpacing:"0.5px"}}>{label}</span>;}
function TempChip({t}){const c=t==="DRY"?T.gold:t==="FRESH"?T.blue:T.purple;const e=t==="DRY"?"🌾":t==="FRESH"?"❄️":"🧊";return <Chip label={`${e} ${t}`} color={c} />;}
function UbicChip({u}){if(!u)return <span style={{color:T.dim}}>—</span>;const c=u==="MTO"?T.blue:u==="FOR"?T.green:T.orange;return <Chip label={u} color={c} />;}
function DeltaBadge({delta,small}){if(delta===null||delta===undefined)return null;const up=delta>0;return <span style={{padding:small?"1px 6px":"3px 9px",borderRadius:"5px",fontSize:small?"10px":"12px",fontWeight:"bold",background:up?"rgba(181,83,74,0.18)":"rgba(75,168,122,0.18)",color:up?T.red:T.green}}>{up?"+":""}{delta.toFixed(1)}%</span>;}
function ActionBtn({label,onClick,primary,disabled}){return <button onClick={onClick} disabled={disabled} style={{padding:"9px 18px",background:disabled?"rgba(255,255,255,0.04)":primary?T.gold:"rgba(255,255,255,0.06)",border:`1px solid ${disabled?T.border:primary?T.gold:T.borderHi}`,borderRadius:"7px",color:disabled?T.dim:primary?T.bg:T.text,cursor:disabled?"not-allowed":"pointer",fontFamily:"inherit",fontSize:"13px",fontWeight:primary?"bold":"normal",transition:"all 0.18s",opacity:disabled?0.6:1}}>{label}</button>;}
function MiniBtn({label,onClick,color}){return <button onClick={onClick} style={{padding:"3px 9px",background:"rgba(255,255,255,0.04)",border:`1px solid ${color||T.border}`,borderRadius:"5px",color:color||T.muted,cursor:"pointer",fontFamily:"inherit",fontSize:"11px",whiteSpace:"nowrap"}}>{label}</button>;}
function FormField({label,value,onChange,placeholder,type,span}){return(<div style={{gridColumn:span?`span ${span}`:undefined}}><label style={{display:"block",fontSize:"11px",color:T.muted,marginBottom:"5px"}}>{label}</label><input type={type||"text"} placeholder={placeholder} value={value||""} onChange={e=>onChange(e.target.value)} style={inputStyle()} /></div>);}
function SelectField({label,value,onChange,opts}){return(<div><label style={{display:"block",fontSize:"11px",color:T.muted,marginBottom:"5px"}}>{label}</label><select value={value} onChange={e=>onChange(e.target.value)} style={{...inputStyle(),cursor:"pointer"}}>{opts.map(([v,l])=><option key={v} value={v}>{l}</option>)}</select></div>);}
function CheckBox({label,checked,onChange}){return(<div onClick={()=>onChange(!checked)} style={{display:"flex",alignItems:"center",gap:"8px",cursor:"pointer",userSelect:"none"}}><div style={{width:"16px",height:"16px",borderRadius:"4px",background:checked?T.gold:"rgba(255,255,255,0.08)",border:`1px solid ${checked?T.gold:T.border}`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:"10px",color:T.bg,flexShrink:0}}>{checked?"✓":""}</div><span style={{fontSize:"12px",color:checked?T.text:T.muted}}>{label}</span></div>);}
function inputStyle(){return {width:"100%",padding:"6px 10px",background:"rgba(255,255,255,0.05)",border:`1px solid ${T.border}`,borderRadius:"6px",color:T.text,fontFamily:"inherit",fontSize:"12px",outline:"none",boxSizing:"border-box"};}