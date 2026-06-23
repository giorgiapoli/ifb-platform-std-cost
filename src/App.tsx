import React, { useState, useMemo, useEffect, useRef } from "react";
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
  MAC: { label:"Macao",     flag:"🇲🇴", color:T.green,  currency:"MOP", defaultRate:9.08,   active:true  },
  CAN: { label:"Canarie",   flag:"🇮🇨", color:T.blue,   currency:"EUR", defaultRate:1,      active:true },
  AUS: { label:"Australia", flag:"🇦🇺", color:T.orange, currency:"AUD", defaultRate:1.6420, active:false },
};
const IFB_VENDOR = "INALCA FOOD & BEVERAGE";

const NOW = () => new Date().toISOString().slice(0,7);
const roundN = (n, d=2) => Math.round((n||0)*Math.pow(10,d))/Math.pow(10,d);
const EXCLUDED_INVOICE_DESC = [
  "freight","health certificate","handling costs","freight cost",
  "interest on intercompany","pallet","vendita prodotti finiti",
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

// ─── CANARIE COST ENGINE ─────────────────────────────────────────────────────
// Source: 05_Modello_Standard_Cost.xlsx — COSTS (LOG) sheet
const COSTS_CAN = {
  // GOMMA: Verona → Barcellona, costo per pallet (COSTS(LOG)!D5 = 62.5 = 2000/32plt)
  VERONA_BARC_PLT: 62.5,
  // GOMMA: Barcellona → island per pallet, by temperature (COSTS(LOG) rows 9-11)
  BARC: {
    DRY:    { GC:133.692, TF:133.692, LAN:190.254, FUE:190.254 },
    FRESH:  { GC:117,     TF:117,     LAN:190.254, FUE:190.254 },
    FROZEN: { GC:133.692, TF:133.692, LAN:271.498, FUE:271.498 },
  },
  // GOMMA: Assicurazione (Seguro) = 0.5% sul valore merce (COSTS(LOG)!F15)
  ASSICURAZIONE: 0.005,
  // MARE: all-in freight per container per island, by area (COSTS(LOG) rows 20-23)
  // Formula Excel: D20 / (Y6 * AA6) = rate / (unitsPerPlt × pltPerContainer)
  MARE: {
    NORD:   { GC:2580.72, TF:2580.72, LAN:3521.44, FUE:3521.44 },
    CENTRO: { GC:0,       TF:0,       LAN:0,       FUE:0       },
    SUD:    { GC:2258.53, TF:2258.53, LAN:0,       FUE:0       },
  },
  // MARE: Inland + Custom Clearance per container per island (COSTS(LOG) rows 27-30)
  // Formula Excel: D27 / (Y6 * AA6) = rate / totalUnits
  INLAND_DRY: { GC:762, TF:762, LAN:830, FUE:830 },
  INLAND_FF:  { GC:0,   TF:0,   LAN:0,   FUE:0   },
  // Pallet (COSTS(LOG)!I1 = 15, formula: I1/Y6)
  PLT: 15,
  // MTO (CROSS DOCKING)
  MTO:   { DRY:8.16,   FRESH:10.2,  FROZEN:12.24 },
  MTS_D: { DRY:14.42,  FRESH:16.48, FROZEN:24.72 },
  MTS_I: { DRY:2.58,   FRESH:3.61,  FROZEN:3.61  },
  MTS_P: { DRY:0.30,   FRESH:0.34,  FROZEN:0.35  },
};

const CAN_ISLANDS = ["GC","TF","LAN","FUE"] as const;

function calcCAN({ priceInput, ubicazione, product, logistic, bevData }: any) {
  const { uom, qtyPerBox, boxPerPallet, kgPerBox, kgxplt, temperature, aiem: prodAiem } = product;
  const { pltPerContainer, area, hasAlcTax, alcTax, convFactor, transport } = logistic || {};

  const cf = Number(convFactor||1) || 1;

  // unitsPerPlt (Y6 in modello): formula Excel IF(J="PCS",Q*R,IF(J="BOX",R,...)) / CM6
  let unitsPerPlt: number;
  if (uom==="BOX")      unitsPerPlt = Number(boxPerPallet) / cf;
  else if (uom==="KG")  unitsPerPlt = (Number(kgxplt)>0 ? Number(kgxplt) : 300) / cf;
  else                  unitsPerPlt = (Number(qtyPerBox) * Number(boxPerPallet)) / cf; // PCS

  // divisoreCollo (AC in modello): per MTS picking
  const divisoreCollo = uom==="BOX" ? 1 : uom==="KG" ? Number(kgPerBox||qtyPerBox) : Number(qtyPerBox);

  const plt_n = Math.max(Number(pltPerContainer)||1, 1);
  const totalUnits = unitsPerPlt * plt_n;
  if (!unitsPerPlt || !totalUnits) return null;

  const priceEur = Number(priceInput||0);
  if (!priceEur) return null;

  const temp: string = temperature || "DRY";
  const areaKey: string = area || "NORD";
  const isMARE = transport === "MARE";
  const isFF = temp === "FRESH" || temp === "FROZEN";

  // Pallet (BO): COSTS(LOG)!I1 / Y6 = 15 / unitsPerPlt
  const plt = COSTS_CAN.PLT / unitsPerPlt;

  // AIEM: se bevData fornisce TOTALE BOTTIGLIA (importo fisso/unit per alcolici), usa quello.
  // Altrimenti usa la % da Anagrafica (prodAiem) o logistic (alcTax).
  const aiemFixed: number = (bevData?.totaleBottiglia ?? 0) > 0 ? Number(bevData.totaleBottiglia) : 0;
  const aiemPct = aiemFixed > 0 ? 0
    : (Number(prodAiem)||0) > 0 ? Number(prodAiem) / 100
    : (hasAlcTax ? (Number(alcTax)||0) / 100 : 0);

  // Costi MARE per isola (AM/AO/AQ/AS + AU/AW/AY/BA)
  // Formula: rate / (Y6 * AA6) = rate / totalUnits
  const inlandTbl = isFF ? COSTS_CAN.INLAND_FF : COSTS_CAN.INLAND_DRY;
  const freightPerIsland = (isl: string): number =>
    isMARE ? (COSTS_CAN.MARE[areaKey]?.[isl] ?? 0) / totalUnits : 0;
  const inlandPerIsland = (isl: string): number =>
    isMARE ? (inlandTbl[isl] ?? 0) / totalUnits : 0;

  // Costi GOMMA (BC/BE/BG/BI/BK + BM)
  const veronaBarcUnit = isMARE ? 0 : COSTS_CAN.VERONA_BARC_PLT / unitsPerPlt;
  const barcPerIsland = (isl: string): number =>
    isMARE ? 0 : (COSTS_CAN.BARC[temp]?.[isl] ?? 0) / unitsPerPlt;
  const assicUnit = isMARE ? 0 : priceEur * COSTS_CAN.ASSICURAZIONE;

  // Trasporto per isola (escluso pallet e AIEM)
  const transpPerIsland = (isl: string): number =>
    freightPerIsland(isl) + inlandPerIsland(isl) + veronaBarcUnit + barcPerIsland(isl) + assicUnit;

  // AIEM per isola: fisso (alcolici) o % su (prezzo + trasporto)
  const aiemGCTF   = aiemFixed > 0 ? aiemFixed : (priceEur + transpPerIsland("GC"))  * aiemPct;
  const aiemLANFUE = aiemFixed > 0 ? aiemFixed : (priceEur + transpPerIsland("LAN")) * aiemPct;
  const aiemForIsl = (isl: string) => (isl==="LAN"||isl==="FUE") ? aiemLANFUE : aiemGCTF;

  // Step1 per isola: BV=AL+AM+AU+BC+BE+BM+BO+BR (MARE o GOMMA, appropriato)
  const step1: Record<string,number> = {};
  const step2: Record<string,number> = {};

  // Warehouse
  let wh = 0;
  if (ubicazione==="MTO") {
    wh = (COSTS_CAN.MTO[temp] ?? 0) / unitsPerPlt;
  } else if (ubicazione==="MTS") {
    wh = (COSTS_CAN.MTS_D[temp]??0)/unitsPerPlt + (COSTS_CAN.MTS_I[temp]??0)/unitsPerPlt + (COSTS_CAN.MTS_P[temp]??0)/divisoreCollo;
  }

  for (const isl of CAN_ISLANDS) {
    step1[isl] = priceEur + transpPerIsland(isl) + plt + aiemForIsl(isl);
    step2[isl] = step1[isl] + wh;
  }

  // Per display: breakdown per GC
  const freightGC = freightPerIsland("GC");
  const inlandGC  = inlandPerIsland("GC");
  const barcUnitGC = barcPerIsland("GC");
  const aiemUnit = aiemGCTF; // GC canonical

  return {
    priceEur, plt, aiemUnit, wh, transport: transport||"GOMMA", unitsPerPlt,
    veronaBarcUnit, barcUnitGC, assicUnit, freightGC, inlandGC,
    // per-island breakdown (GC=TF share rates; LAN=FUE share rates)
    freightLAN: freightPerIsland("LAN"),
    barcUnitLAN: barcPerIsland("LAN"),
    aiemGCTF, aiemLANFUE,
    isMARE,
    step1GC: step1.GC, step1TF: step1.TF, step1LAN: step1.LAN, step1FUE: step1.FUE,
    step2GC: step2.GC, step2TF: step2.TF, step2LAN: step2.LAN, step2FUE: step2.FUE,
    // compat con codice esistente (GC come canonico)
    step1Eur: step1.GC, step1Hkd: step1.GC,
    step2Eur: step2.GC, step2Hkd: step2.GC, rate:1,
    fob:0, lic:0, vgm:0, hc:0, alc: aiemUnit,
  };
}

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

// ─── MACAO CONSTANTS & CALC ───────────────────────────────────────────────────
const HKD_TO_MOP = 1.03;          // fixed exchange rate HKD → MOP
const MAC_MARKUP = { hoff: 0.03, nonHoff: 0.10 };  // 3% HOFF, 10% non-HOFF
// ALL-IN logistics cost per KG (MOP): BV Warehouse→HK Port + Ferry + Customs + Delivery Macao
const MAC_LOG_PER_KG: any = { DRY: 3, FRESH: 5, FROZEN: 8 };

function calcMAC({ hkCost, isHoff, macToHkConv = 1, temperature = "DRY", kgPerMacUom = 0 }: any) {
  if (!hkCost?.step2Hkd) return null;
  const markup = isHoff ? MAC_MARKUP.hoff : MAC_MARKUP.nonHoff;
  const conv = Number(macToHkConv) > 0 ? Number(macToHkConv) : 1;
  const hkNewSC = hkCost.step2Hkd;
  const baseInMop = hkNewSC * conv * (1 + markup) * HKD_TO_MOP;
  // Costo logistico ALL-IN per MAC UOM (MOP/kg × kg per MAC UOM)
  const logPerKg = MAC_LOG_PER_KG[String(temperature||"DRY").toUpperCase()] ?? 3;
  const logPerUom = kgPerMacUom > 0 ? logPerKg * kgPerMacUom : logPerKg; // se UOM=KG → kgPerMacUom=1
  const macNewSC = baseInMop + logPerUom;
  return {
    hkNewSC, markup: markup * 100, isHoff, macNewSC, macToHkConv: conv,
    baseInMop, logPerKg, logPerUom, temperature,
    step2Hkd: macNewSC,
    step2Eur: 0,
    priceEur: hkCost.priceEur || 0,
    unitsPerPlt: hkCost.unitsPerPlt || 0,
    rate: HKD_TO_MOP,
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
class ErrorBoundary extends React.Component<{children:any},{err:any}> {
  constructor(p:any){super(p);this.state={err:null};}
  static getDerivedStateFromError(e:any){return{err:e};}
  render(){
    if(this.state.err) return(
      <div style={{padding:"40px",fontFamily:"monospace",color:"#ff6b6b",background:"#0B0F14",minHeight:"100vh"}}>
        <h2 style={{color:"#C9A84C"}}>⚠ Runtime Error</h2>
        <pre style={{whiteSpace:"pre-wrap",wordBreak:"break-all",fontSize:"12px"}}>{String(this.state.err?.message||this.state.err)}</pre>
        <pre style={{whiteSpace:"pre-wrap",wordBreak:"break-all",fontSize:"10px",color:"#aaa",marginTop:"12px"}}>{this.state.err?.stack}</pre>
      </div>
    );
    return this.props.children;
  }
}

export default function App() {
  const[products,setProducts]   = useState<any[]>([]);
  const[logistics,setLogistics] = useState<any[]>(SEED_LOGISTIC);
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
  const [meatPrices, setMeatPrices] = useState<any[]>([]);
  const [bevInfo, setBevInfo] = useState<any[]>([]); // CAN: dati alcolici per AIEM fisso
  const [priceExceptions, setPriceExceptions] = useState<any[]>(() => LS.get(`ifb_exceptions_${LS.get("ifb_branch","")}`, []));
  const [scAttuali, setScAttuali] = useState<any[]>([]);
  const [macHkCostRows, setMacHkCostRows] = useState<any[]>([]); // HK costs loaded for MAC derivation

  const navigate = (pageName, filter=null) => { setPageFilter(filter); setPage(pageName); };

  const branchRef = useRef(branch);
  useEffect(()=>{ branchRef.current = branch; },[branch]);
  const branchLoadedRef = useRef<string>("");
  const globalLoadedRef = useRef(false); // blocks global saves until IDB load completes

  // Load global data (logistics, meatPrices) from IDB on mount
  useEffect(()=>{
    (async()=>{
      setLogistics(await IDB.get("ifb_logistics", SEED_LOGISTIC));
      setMeatPrices(await IDB.get("ifb_meatprices", []));
      setBevInfo(await IDB.get("ifb_bevinfo", []));
      globalLoadedRef.current = true;
    })();
  },[]);

  // Reload price exceptions when branch changes
  useEffect(()=>{ if(branch) setPriceExceptions(LS.get(`ifb_exceptions_${branch}`,[])); },[branch]);
  // Save effects — only fire after load is complete
  useEffect(()=>{ if(branchRef.current) LS.set(`ifb_exceptions_${branchRef.current}`, priceExceptions); },[priceExceptions]);
  useEffect(()=>{ if(branchRef.current&&branchLoadedRef.current===branchRef.current) IDB.set(`ifb_products_${branchRef.current}`, products); },[products]);
  useEffect(()=>{ if(globalLoadedRef.current) IDB.set("ifb_logistics", logistics); }, [logistics]);
  useEffect(()=>{ if(branchRef.current&&branchLoadedRef.current===branchRef.current) IDB.set(`ifb_airlist_${branchRef.current}`, airList); },[airList]);
  useEffect(()=>{ if(branchRef.current&&branchLoadedRef.current===branchRef.current) IDB.set(`ifb_xrefs_${branchRef.current}`, xrefs); },[xrefs]);
  useEffect(()=>{ if(branchRef.current&&branchLoadedRef.current===branchRef.current) IDB.set(`ifb_sales_invoice_${branchRef.current}`, salesRows); },[salesRows]);
  useEffect(()=>{ if(branchRef.current&&branchLoadedRef.current===branchRef.current) IDB.set(`ifb_scattuali_${branchRef.current}`, scAttuali); },[scAttuali]);
  // MAC: load saved HK costRows when switching to MAC branch
  useEffect(()=>{ if(branch==="MAC") IDB.get("ifb_hk_costrows_for_mac",[]).then((d:any[])=>setMacHkCostRows(d)); },[branch]);
  useEffect(()=>{ if(prices.length) LS.set("ifb_prices", prices); }, [prices]);
  useEffect(()=>{ if(branch) LS.set("ifb_branch",branch); },[branch]);
  useEffect(()=>{ if(globalLoadedRef.current) IDB.set("ifb_meatprices", meatPrices); }, [meatPrices]);
  useEffect(()=>{ if(globalLoadedRef.current) IDB.set("ifb_bevinfo", bevInfo); }, [bevInfo]);
  // Ricarica dati branch-specifici ad ogni cambio filiale
  useEffect(()=>{
    if(!branch) return;
    branchLoadedRef.current = ""; // reset — block saves while loading
    (async()=>{
      setProducts(await IDB.get(`ifb_products_${branch}`,[]));
      setXrefs(await IDB.get(`ifb_xrefs_${branch}`,[]));
      setAirList(await IDB.get(`ifb_airlist_${branch}`,[]));
      setSalesRows(await IDB.get(`ifb_sales_invoice_${branch}`,[]));
      setScAttuali(await IDB.get(`ifb_scattuali_${branch}`,[]));
      branchLoadedRef.current = branch; // unblock saves
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

    // MAC: calcola da standardCostHkd importato nel file anagrafica MAC
    if (branch === "MAC") {
      const eligible = products.filter((p:any) => p.active && p.standardCostHkd > 0);
      return eligible.map((prod:any) => {
        const isHoff = prod.isHoff ?? false;
        const macUom = prod.uom  || "";
        const hkUom  = prod.hkUom || "";
        // Conversione automatica UOM: se MAC vende a BOX e HK a PCS, conv = qtyPerBox
        const uomDiffers = macUom && hkUom && macUom.toUpperCase() !== hkUom.toUpperCase();
        let macToHkConv = 1;
        if (uomDiffers) {
          const hkU = hkUom.toUpperCase();
          const mU  = macUom.toUpperCase();
          const qty = Number(prod.qtyPerBox) || 1;
          if (hkU==="PCS" && (mU==="BOX"||mU==="CTN")) macToHkConv = qty;
          else if ((hkU==="BOX"||hkU==="CTN") && mU==="PCS") macToHkConv = 1/qty;
        }
        // kg per MAC UOM (per costo logistica ALL-IN)
        // Se MAC vende a KG → 1; se a BOX → kgPerBox; se a PCS → kgPerBox/qtyPerBox
        const mU = macUom.toUpperCase();
        const kgPerMacUom = mU==="KG" ? 1
          : mU==="BOX"||mU==="CTN" ? (Number(prod.kgPerBox)||0)
          : mU==="PCS" ? (Number(prod.kgPerBox)||0) / Math.max(Number(prod.qtyPerBox)||1,1)
          : 1;
        const temperature = prod.temperature || "DRY";
        // Base = standardCostHkd dal file MAC (costo FOB che Macao paga a HK)
        const hkCostBase = { step2Hkd: prod.standardCostHkd, step2Eur:0, priceEur:0, unitsPerPlt:0, rate: HKD_TO_MOP };
        const macCost = calcMAC({ hkCost: hkCostBase, isHoff, macToHkConv, temperature, kgPerMacUom });
        return {
          ...prod, id: prod.id||prod.code, code: prod.code, nHK: prod.nHK,
          description: prod.description, uom: macUom,
          isHoff, macUom, hkUom, uomDiffers, macToHkConv,
          cost: macCost, prevCost: null,
        };
      });
    }

    const fxRate = fx.find(f=>f.branch===branch&&f.month===month)?.rate || BRANCH_CFG[branch]?.defaultRate || 9.1437;
    const [yr,mo] = month.split("-").map(Number);
    const prevM = mo===1 ? `${yr-1}-12` : `${yr}-${String(mo-1).padStart(2,"0")}`;
    const eligible = products.filter(p => p.active && isIFBVendor(p.vendorName));

    return eligible.map(prod => {
      // Eccezione prezzo manuale: ha priorità assoluta su listino e carne
      const exc = priceExceptions.find((e:any) =>
        e.branch === branch && (
          e.productId === prod.id ||
          (e.code && e.code === prod.code) ||
          (e.nHK && prod.nHK && e.nHK === prod.nHK)
        )
      );

      const airEntry = airList.find((a:any)=>
          a.productId === prod.id ||
          (a.code && a.code === prod.code) ||
          (a.nHK && prod.nHK && a.nHK === prod.nHK)
        );
      if(airEntry && isAirTransport(airEntry.transportation))
        return { ...prod, cost:null, prevCost:null, priceInput:null, isAir:true, skipReason:"AIR" };

      const logRaw = logistics.find(l=>l.productId===prod.id&&l.branch===branch);
      if(!logRaw) return { ...prod, cost:null, prevCost:null, priceInput:null, skipReason:"NO LOGISTICA" };

      // Apply pltPerContainer default: CAN uses fixed values (MARE=24, GOMMA=32); other branches use temp-based formula
      const pltFromFile = logRaw.pltPerContainer || 0;
      const effectiveTemp = logRaw.temperatureOverride || prod.temperature;
      const canDefaultPlt = branch === "CAN"
        ? (logRaw.transport === "MARE" ? 24 : 32)
        : pltDefault(effectiveTemp);
      const plt = pltFromFile > 0 ? pltFromFile : canDefaultPlt;
      const log = { ...logRaw, pltPerContainer: plt };

      const pr     = prices.find(p=>p.productId===prod.id&&p.branch===branch&&p.month===month);
      const prPrev = prices.find(p=>p.productId===prod.id&&p.branch===branch&&p.month===prevM);

      const ub = log.ubicazione;
      const effectiveProd = log.temperatureOverride ? { ...prod, temperature: log.temperatureOverride } : prod;

      // Branch-agnostic calc helper
      const isCAN_b = branch === "CAN";
      const bevData = isCAN_b ? bevInfo.find((b:any) => b.ifbNo === prod.code) : null;
      const calcCost = (pi: number) =>
        isCAN_b
          ? calcCAN({ priceInput:pi, ubicazione:ub, product:effectiveProd, logistic:log, bevData })
          : calcHK({ priceInput:pi, ubicazione:ub, product:effectiveProd, logistic:{...log,category:prod.category}, eurToHkd:fxRate });

      // Eccezione prezzo: bypassa listino e carne
      if(exc && exc.price > 0) {
        const costE = calcCost(exc.price);
        const deltaE = costE ? null : null; // no prev for exceptions
        return { ...prod, cost:costE, prevCost:null, delta:null, priceInput:exc.price,
          flagged:false, ubicazione:ub, pltUsed:plt, area:log.area||"NORD", pltPerContainer:plt,
          temperatureOverride:log.temperatureOverride||null, _fromException:true, skipReason: costE?undefined:"CALC=0" };
      }

      // Meat list fallback helper
      const meatFallback = () => {
        const meat = meatPrices.find((m:any) =>
          m.code===prod.code || m.code===String(prod.id) || (prod.nHK&&m.code===prod.nHK));
        if (!meat) return null;
        const kgPerUnit =
          prod.uom==="KG" ? 1 :
          prod.uom==="BOX" ? (Number(prod.kgPerBox)||0) :
          (Number(prod.kgPerBox)||0) / Math.max(Number(prod.qtyPerBox)||1,1);
        return { pi: meat.pricePerKg * kgPerUnit * 1.01 }; // +1% intercompany
      };

      if(!pr) {
        const mf = meatFallback();
        if(!mf) return { ...prod, cost:null, prevCost:null, priceInput:null, ubicazione:ub, skipReason:`NO PREZZO (${branch}/${month})` };
        const cost2 = calcCost(mf.pi);
        return { ...prod, cost:cost2, prevCost:null, delta:null, priceInput:mf.pi, isNew:true,
          flagged:false, ubicazione:ub, pltUsed:plt, area:log.area||"NORD", pltPerContainer:plt,
          temperatureOverride:log.temperatureOverride||null,
          skipReason: cost2 ? undefined : "CALC=0", _fromMeatList:true };
      }

      const pi  = selectPrice(pr, ub);
      const piP = prPrev ? selectPrice(prPrev, ub) : null;

      // Prezzo zero → fallback listino carne
      if (!pi || pi === 0) {
        const mf = meatFallback();
        if (mf) {
          const costM = calcCost(mf.pi);
          return { ...prod, cost:costM, prevCost:null, delta:null, priceInput:mf.pi, isNew:true,
            flagged:false, ubicazione:ub, pltUsed:plt, area:log.area||"NORD", pltPerContainer:plt,
            temperatureOverride:log.temperatureOverride||null,
            skipReason: costM ? undefined : "CALC=0", _fromMeatList:true };
        }
      }

      const cost = calcCost(pi);
      if(!cost) return { ...prod, cost:null, prevCost:null, priceInput:pi,
        skipReason: !pi||pi===0 ? "PREZZO ZERO" : `CALC=0 (qty=${prod.qtyPerBox} box/plt=${prod.boxPerPallet} plt=${plt} uom=${prod.uom})` };

      const prevCost = piP!=null ? calcCost(piP) : null;
      const delta    = cost&&prevCost ? (cost.step2Hkd-prevCost.step2Hkd)/prevCost.step2Hkd*100 : null;
      return { ...prod, cost, prevCost, delta, priceInput:pi, isNew:!prPrev,
        flagged: delta!==null&&Math.abs(delta)>=3, ubicazione:ub, pltUsed:plt,
        area:log.area||"NORD", pltPerContainer:plt,
        temperatureOverride: log.temperatureOverride||null };
    });
  }, [products,logistics,prices,fx,airList,meatPrices,priceExceptions,branch,month,bevInfo]);

  // MAC: save HK costRows to IDB whenever they're computed (declared after costRows useMemo to avoid TDZ)
  useEffect(()=>{ if(branch==="HK" && costRows.length>0) IDB.set("ifb_hk_costrows_for_mac", costRows); },[costRows,branch]);

  const isCAN = branch === "CAN";
  const isMAC = branch === "MAC";

  const NAV_ALL = [
    {id:"dashboard",  icon:"⬡", label:"Dashboard"},
    {id:"products",   icon:"◈", label:"Anagrafica", badge:"⇪"},
    {id:"xref",       icon:"⇄", label:isCAN?"XRef N COMIT / IFB":"XRef N / IFB"},
    ...(!isMAC ? [{id:"logistics", icon:"◎", label:isCAN?"Work Tab (Logistica)":"Logistica"}] : []),
    ...(!isMAC ? [{id:"prices",    icon:"◉", label:"Listini", badge:"💶"}] : []),
    ...(!isMAC ? [{id:"meatlist",  icon:"🥩", label:"Listino Carne"}] : []),
    ...(isCAN ? [{id:"bevinfo", icon:"🍷", label:"Beverage Info (AIEM)"}] : []),
    ...(!isCAN&&!isMAC ? [{id:"fx",  icon:"◌", label:"Cambi"}] : []),
    ...(!isCAN&&!isMAC ? [{id:"air", icon:"✈", label:"AIR Transport"}] : []),
    ...(!isMAC ? [{id:"exceptions", icon:"⚡", label:"Eccezioni Prezzi"}] : []),
    {id:"costs",      icon:"◆", label:"Standard Cost"},
    ...(!isMAC ? [{id:"invoice", icon:"📋", label:"Fatture & Costi", badge:"⇪"}] : []),
    ...(!isMAC ? [{id:"scattuali", icon:"📊", label:"SC Attuali", badge:scAttuali.length>0?String(scAttuali.length):undefined}] : []),
    {id:"storico",    icon:"⧖", label:"Storico & Diff"},
    ...(!isMAC ? [{id:"check", icon:"📅", label:"Check Mensile"}] : []),
    {id:"notes",      icon:"📝", label:"Guida & Istruzioni"},
  ];
  const NAV = NAV_ALL;

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
    bevinfo: <BeverageInfoPage bevInfo={bevInfo} setBevInfo={setBevInfo} products={products} showToast={showToast}/>,
    exceptions:  <PriceExceptions branch={branch} products={products} xrefs={xrefs} priceExceptions={priceExceptions} setPriceExceptions={setPriceExceptions}/>,
    costs:       <CostTable costRows={costRows} branch={branch} month={month} logistics={logistics} lastImportTs={lastImportTs} lastCalcTs={lastCalcTs} setLastCalcTs={setLastCalcTs} setCostHistory={setCostHistory} initFilter={pageFilter} salesRows={salesRows} products={products} xrefs={xrefs}/>,
    invoice: <InvoiceAndCosts rows={salesRows} setRows={setSalesRows} branch={branch} airList={airList} products={products} xrefs={xrefs} costRows={costRows} logistics={logistics} snapshots={snapshots} setSnapshots={setSnapshots} importLogs={importLogs} setImportLogs={setImportLogs} showToast={showToast} bumpImportTs={bumpImportTs}/>,
    scattuali: <ScAttualiPage scAttuali={scAttuali} setScAttuali={setScAttuali} branch={branch} showToast={showToast}/>,
    storico: <Storico
      snapshots={snapshots}
      setSnapshots={setSnapshots}
      costHistory={costHistory}
      setCostHistory={setCostHistory}
      branch={branch}
      showToast={showToast}
      macHkCostRows={macHkCostRows}
    />,
    check: <CheckMensile costRows={costRows} branch={branch} salesRows={salesRows} xrefs={xrefs} scAttuali={scAttuali} products={products}/>,
    mail:  <MailGen costRows={costRows} branch={branch} month={month}/>,
    notes:       <NotesPage/>,
  };

  return (
    <ErrorBoundary>
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
    </ErrorBoundary>
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
        // Alias per la colonna filiale (N COMIT per CAN, N HK per altri)
        const nhkA = branch==="CAN"
          ? ["n comit","ncomit","comit","canarie","can no","can n","n°","numero comit","codice comit","cod comit","codcan","n_comit"]
          : ["n hk","nhk","hk","n_hk","gc code","gc no","hk code","hk no","hong kong"];
        // Alias per colonna IFB (espliciti prima, generici dopo)
        const ifbA=["ifb n","ifb no","ifb no.","ifb item","bv no","bv n","no_ifb","ifb","no_","code","item no"];
        // Match esatto o inclusione, alias filiale prima degli alias IFB (evita false positiv)
        const normH = (h:string) => h.toLowerCase().replace(/[°\s_]/g,"");
        const pickCol = (aliases:string[]) =>
          hdrs.find(h=>aliases.some(a=>normH(h)===normH(a))) ||
          hdrs.find(h=>aliases.some(a=>normH(h).includes(normH(a))));
        setColNHK(pickCol(nhkA)||"");
        setColIFB(pickCol(ifbA)||"");
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
    setXrefs(next);IDB.set(`ifb_xrefs_${branch}`,next);
    const log={id,type:"xref",fileName,date:new Date(id).toISOString(),count:incoming.length,diffs,branch};
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
              <button onClick={()=>{if(window.confirm(`Eliminare tutte le ${xrefs.length} XRef di ${branch}?`)){setXrefs([]);IDB.set(`ifb_xrefs_${branch}`,[]);}}}
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
                    <TD><MiniBtn label="✕" onClick={()=>{const n=xrefs.filter((_,j)=>j!==xrefs.indexOf(x));setXrefs(n);IDB.set(`ifb_xrefs_${branch}`,n);}} color={T.red}/></TD>
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
  const [open, setOpen] = useState<number|null>(null);
  const toggle = (i:number) => setOpen(o=>o===i?null:i);

  // ── FLUSSO SETUP (una tantum per filiale)
  const setup = [
    { icon:"◈", label:"Anagrafica", color:T.muted,
      desc:"Carica il file export da BC/Navision con tutti gli articoli. Contiene: codice, UOM, dimensioni pallet, temperatura, fornitore, AIEM (CAN).",
      steps:["Pagina Anagrafica → Carica anagrafica","Mappa le colonne (rilevamento automatico)","Preview → Importa","✓ Da rifare solo se cambiano i prodotti"] },
    { icon:"⇄", label:"XRef", color:T.muted,
      desc:"Collega il codice filiale (N HK / N COMIT) al codice IFB. Necessario per incrociare fatture e standard cost.",
      steps:["Pagina XRef → Carica file con 2 colonne: N filiale + IFB N","Seleziona le colonne → Preview → Importa","✓ Da rifare solo se cambiano i codici"] },
    { icon:"◎", label:"Work Tab (Logistica)", color:T.muted,
      desc:"Parametri logistici per articolo: ubicazione MTS/MTO/FOR, MARE/GOMMA (CAN), numero pallet per container, area geografica.",
      steps:["Pagina Work Tab (Logistica) → Carica Work Tab","Il sistema rileva automaticamente le colonne","Importa","✓ Da rifare se cambiano i parametri logistici"] },
    { icon:"⚡", label:"Eccezioni Prezzi", color:T.muted,
      desc:"Override manuale del prezzo per un articolo specifico, con priorità su listino e listino carne.",
      steps:["Pagina Eccezioni Prezzi → cerca articolo per codice o descrizione","Inserisci il prezzo manuale e una nota","Salva → il prezzo verrà usato nel calcolo SC"] },
  ];

  // ── FLUSSO MENSILE
  const monthly = [
    { icon:"💰", label:"1. Listino prezzi", color:T.green,
      desc:"Prezzi DAP/FCA del mese dal sistema (Power BI / CURRENT PRICELIST). Necessario per calcolare il costo di acquisto.",
      steps:["Pagina Listini → Carica file listino","Seleziona il mese di riferimento","Preview → Importa"] },
    { icon:"🧾", label:"2. Fatture del mese", color:T.green,
      desc:"Sales Invoice export da BC/Navision. L'app legge le posting date per filtrare per mese. I dati sono cumulativi — basta caricare il file più aggiornato.",
      steps:["Pagina Fatture & Costi → Carica file","Mappa le colonne (rilevamento automatico)","Importa","✓ Le fatture vecchie vengono conservate, le nuove aggiunte"] },
    { icon:"📊", label:"3. SC Attuali", color:T.green,
      desc:"Report degli Standard Cost correnti dal sistema (BC per HK, Navision per CAN). Serve per il confronto mensile (soglia > +3% o < -3%).",
      steps:["Pagina SC Attuali → Carica report","Il formato HK o CAN viene rilevato automaticamente","Importa","✓ Da aggiornare ogni mese"] },
    { icon:"◆", label:"4. Calcola Standard Cost", color:T.blue,
      desc:"Il calcolo usa listino + logistica + parametri fissi (FOB, LIC, VGM, PLT…) per produrre il New Standard Cost per articolo, con dettaglio di ogni voce.",
      steps:["Pagina Standard Cost → clicca ⟳ Ricalcola","Attendi il calcolo (pochi secondi)","Verifica la tabella: ogni riga mostra il breakdown completo","Clicca una riga per il dettaglio"] },
    { icon:"📅", label:"5. Check Mensile → Export", color:T.gold,
      desc:"Confronta lo SC calcolato con gli SC Attuali. Identifica articoli NUOVI (nessun SC in sistema) e DA AGGIORNARE (variazione > +3% o < -3%). Esporta il file Excel pronto.",
      steps:["Pagina Check Mensile","Seleziona il mese dalla lista (derivato dalle fatture caricate)","Verifica la lista: NUOVO ARTICOLO / DA AGGIORNARE / OK","Clicca 📥 Esporta Excel → file STDC_Analisi_BRANCH_MESE.xlsx"] },
  ];

  const Card = ({icon,label,color,desc,steps,idx,isOpen}:{icon:string,label:string,color:string,desc:string,steps:string[],idx:number,isOpen:boolean}) => (
    <div style={{borderRadius:"8px",border:`1px solid ${isOpen?color:T.border}`,overflow:"hidden",transition:"border-color 0.15s"}}>
      <button onClick={()=>toggle(idx)}
        style={{width:"100%",display:"flex",alignItems:"center",gap:"10px",padding:"12px 16px",
          background:isOpen?`${color}12`:"transparent",border:"none",cursor:"pointer",textAlign:"left",fontFamily:"inherit"}}>
        <span style={{fontSize:"15px",width:"22px",textAlign:"center"}}>{icon}</span>
        <span style={{flex:1,fontSize:"13px",fontWeight:"bold",color:isOpen?color:T.text}}>{label}</span>
        <span style={{fontSize:"10px",color:T.dim}}>{isOpen?"▲":"▼"}</span>
      </button>
      {isOpen&&(
        <div style={{padding:"0 16px 14px 16px",borderTop:`1px solid ${color}22`}}>
          <p style={{fontSize:"12px",color:T.muted,margin:"10px 0 10px",lineHeight:"1.6"}}>{desc}</p>
          <ol style={{margin:0,paddingLeft:"18px",fontSize:"12px",color:T.text,lineHeight:"2"}}>
            {steps.map((s,i)=><li key={i} style={{marginBottom:"2px"}}>{s}</li>)}
          </ol>
        </div>
      )}
    </div>
  );

  return(
    <div style={{maxWidth:"800px"}}>
      <PageHeader title="📝 Guida rapida" sub="Come si usa la piattaforma — passo per passo"/>

      {/* FLUSSO VISIVO */}
      <Section title="">
        <div style={{display:"flex",alignItems:"center",gap:"6px",flexWrap:"wrap",marginBottom:"8px"}}>
          {[
            {label:"Setup",sub:"una tantum",c:T.muted},
            {label:"→"},
            {label:"Listino",sub:"mensile",c:T.green},
            {label:"→"},
            {label:"Fatture",sub:"mensile",c:T.green},
            {label:"→"},
            {label:"SC Attuali",sub:"mensile",c:T.green},
            {label:"→"},
            {label:"Calcola SC",sub:"mensile",c:T.blue},
            {label:"→"},
            {label:"Check Mensile",sub:"export",c:T.gold},
          ].map((item,i)=>item.label==="→"
            ? <span key={i} style={{color:T.dim,fontSize:"14px"}}>→</span>
            : <div key={i} style={{background:`${item.c}18`,border:`1px solid ${item.c}44`,borderRadius:"6px",padding:"5px 10px",textAlign:"center"}}>
                <div style={{fontSize:"11px",fontWeight:"bold",color:item.c}}>{item.label}</div>
                {item.sub&&<div style={{fontSize:"9px",color:T.dim}}>{item.sub}</div>}
              </div>
          )}
        </div>
        <p style={{fontSize:"12px",color:T.dim,margin:0}}>
          Il <strong style={{color:T.muted}}>Setup</strong> si fa una volta sola per filiale. Ogni mese si ripetono solo i 5 step a destra.
        </p>
      </Section>

      {/* SETUP */}
      <Section title="Setup — da fare una volta per filiale" accent={T.muted}>
        <div style={{display:"flex",flexDirection:"column",gap:"6px"}}>
          {setup.map((s,i)=><Card key={i} {...s} idx={i} isOpen={open===i}/>)}
        </div>
      </Section>

      {/* MENSILE */}
      <Section title="Flusso mensile — da ripetere ogni mese" accent={T.green}>
        <div style={{display:"flex",flexDirection:"column",gap:"6px"}}>
          {monthly.map((s,i)=><Card key={i+100} {...s} idx={i+100} isOpen={open===i+100}/>)}
        </div>
      </Section>

      {/* NOTE RAPIDE */}
      <Section title="Regole da sapere" accent={T.blue}>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"10px"}}>
          {[
            {t:"Solo INALCA F&B",d:"Il calcolo SC è attivo solo per articoli con fornitore INALCA FOOD & BEVERAGE."},
            {t:"Soglia > +3% o < -3%",d:"Una variazione viene segnalata come DA AGGIORNARE solo se superiore a +3% o inferiore a -3% rispetto allo SC attuale."},
            {t:"New Standard Cost",d:"New SC = costo acquisto + trasporti + dazi + pallet + AIEM (CAN) + costo magazzino (MTS/MTO). Clicca una riga per il breakdown completo."},
            {t:"MARE vs GOMMA (CAN)",d:"MARE: costo container diviso per unità totali. GOMMA: costo per pallet diviso per unità/pallet."},
            {t:"Eccezione prezzo",d:"Ha priorità assoluta su listino e listino carne. Usarla per accordi speciali o campioni."},
            {t:"Fatture cumulative",d:"Il file fatture include tutti i mesi. L'app filtra per posting date — non serve caricare ogni mese un file diverso."},
          ].map(({t,d},i)=>(
            <div key={i} style={{background:T.card,borderRadius:"8px",padding:"12px 14px",border:`1px solid ${T.border}`}}>
              <div style={{fontSize:"11px",fontWeight:"bold",color:T.text,marginBottom:"4px"}}>{t}</div>
              <div style={{fontSize:"11px",color:T.muted,lineHeight:"1.6"}}>{d}</div>
            </div>
          ))}
        </div>
      </Section>
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
              <label style={{display:"block", fontSize:"11px", color:T.gold, marginBottom:"5px"}}>📌 Codice * ({branchN(branch)} o IFB N)</label>
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

  const FIELDS=["nHK","code","description","category","uom","qtyPerBox","boxPerPallet","kgPerBox","kgxplt","temperature","aiem","isHoff","macUom","hkUom","macToHkConv","active","vendorName","vendorName2"];
  const FLABELS={nHK:`${branchN(branch)} (No_)`,code:"IFB Item / BV No *",description:"Descrizione *",category:"Section",uom:"UOM",qtyPerBox:"Qty/Cartone",boxPerPallet:"Cartoni/Pallet",kgPerBox:"Kg/Cartone (Net Weight)",kgxplt:"Kg x PLT",temperature:"Product Type",aiem:"★ AIEM % (CAN — col. W anagrafica)",isHoff:"HOFF Flag (MAC: 1=HOFF)",macUom:"MAC UOM di vendita",hkUom:"HK/BV UOM di vendita",macToHkConv:"Fattore conversione MAC÷HK (es. 6 se HK=PCS e MAC=BOX6)",active:"Bloccato",vendorName:"Vendor Name",vendorName2:"Vendor Name 2"};

  const LOCAL_ALIASES = {
    nHK:         ["no","no_","macaono","macao no","macao_no","macaomastercode","macao mastercode","macaoitemno"],
    code:        ["ifbitem","ifb item","ifb no","ifb n","bvno","bv no","bvmastercode","bv mastercode"],
    description: ["description"],
    category:    ["sectiondescription","section description","section"],
    uom:         ["salesunitofmeasure","sales unit of measure"],
    qtyPerBox:   ["quantityxpackaging","quantity x packaging"],
    boxPerPallet:["packagingxpallet","packaging x pallet"],
    kgPerBox:    ["netweight","net weight"],
    temperature: ["producttype","product type","product type rettificato","product type - anagrafica"],
    active:      ["blocked"],
    kgxplt:      ["kgxplt","kg x pallet","kg per pallet","kgperpallet","kgplt"],
    vendorName:  ["vendorname","vendor name"],
    vendorName2: ["vendorname2","vendor name 2"],
    aiem:        ["aiem","igic","alim","aiem%","aiem_perc","aiem_canarie","aiemperc"],
    isHoff:      ["ishoff","hoff","hofflag","hoff flag","hoff_flag","is hoff"],
    macUom:      ["macaosalesunitofmeasure","macao salesunitofmeasure","macaouom","macao uom","mac uom","macuom"],
    hkUom:       ["bvsalesunitofmeasure","bv salesunitofmeasure","bvuom","hk uom","hkuom"],
    macToHkConv: ["mactoHkconv","conversionfactor","conv factor","conversion","fattoreconv","macaotoHkconv"],
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
      kgxplt: parseFloat(r.kgxplt||"") > 0
        ? parseFloat(r.kgxplt)
        : roundN((parseFloat(r.kgPerBox)||0) * (parseFloat(r.qtyPerBox)||1) * (parseFloat(r.boxPerPallet)||0)),
      active:!["true","1","yes"].includes(String(r.active||"").toLowerCase()),
      vendorName: r.vendorName || "",
      vendorName2: r.vendorName2 || "",
      aiem: parseFloat(r.aiem)||0,
      isHoff: ["true","1","yes","hoff","si","sì","vero","x"].includes(String(r.isHoff||"").toLowerCase()),
      macUom: r.macUom ? String(r.macUom).trim().toUpperCase() : "",
      hkUom:  r.hkUom  ? String(r.hkUom).trim().toUpperCase()  : "",
      macToHkConv: parseFloat(r.macToHkConv)>0 ? parseFloat(r.macToHkConv) : 1,
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
    const snap={id:now,type:"anagrafica",date:new Date(now).toISOString(),count:newProds.length,diffs,products:newProds,branch};
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
      <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:"12px",maxWidth:"960px",marginBottom:"20px"}}>
        {FIELDS.map(f=>{
          const isRequired = f==="code"||f==="description";
          const isCurrent = branch==="CAN"&&f==="aiem";
          const isMacField = ["macUom","hkUom","macToHkConv"].includes(f);
          const isHoffField = f==="isHoff";
          const labelColor = isRequired?T.gold:isCurrent?T.orange:(isMacField||isHoffField)?T.purple:T.muted;
          return(
          <div key={f} style={isCurrent?{border:`1px solid ${T.orange}33`,borderRadius:"6px",padding:"4px 6px",background:`${T.orange}08`}:{}}>
            <label style={{display:"block",fontSize:"11px",color:labelColor,marginBottom:"5px"}}>{FLABELS[f]}</label>
            <select value={map[f]||""} onChange={e=>setMap(m=>({...m,[f]:e.target.value}))} style={{...inputStyle(),borderColor:map[f]?T.gold:(isCurrent&&!map[f])?T.orange:T.border}}>
              <option value="">— non mappato —</option>
              {headers.map(h=><option key={h} value={h}>{h}</option>)}
            </select>
          </div>
          );
        })}
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
    setAirList(next); IDB.set(`ifb_airlist_${branch}`, next);
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
            <label style={{display:"block",fontSize:"11px",color:T.gold,marginBottom:"5px"}}>Colonna Codice * ({branchN(branch)} o IFB N)</label>
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
              setAirList(next);IDB.set(`ifb_airlist_${branch}`,next);
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
    setAirList([]);IDB.set(`ifb_airlist_${branch}`,[]);
  }}}
     style={{padding:"8px 16px",background:"none",border:`1px solid ${T.red}44`,borderRadius:"6px",color:T.red,cursor:"pointer",fontSize:"12px"}}>
     ✕ Svuota lista ({branchAir.length})
   </button>
 )}
 <span style={{fontSize:"11px",color:T.muted}}>Colonna richiesta: {branchN(branch)} o IFB N · ogni import sostituisce la lista precedente</span>
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
                      <TD><MiniBtn label="✕ Rimuovi" onClick={()=>{const n=airList.filter((_,j)=>j!==airList.indexOf(a));setAirList(n);IDB.set(`ifb_airlist_${branch}`,n);}} color={T.red}/></TD>
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
    { id:"flagged", n:flagged.length,  label:"Variazioni ≥3% New SC",        color:T.orange, rows:flagged  },
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
        <THead cols={[branchN(branch),"IFB No","Descrizione","Ubicaz.",branch==="CAN"?"Area":"","Step2 HKD","Prec. HKD","Δ%"]}sticky/>
        <tbody>{panel.rows.map((r:any,i:number)=>{
          const pct = r.cost&&r.prevCost&&r.prevCost.step2Hkd>0
            ? (r.cost.step2Hkd-r.prevCost.step2Hkd)/r.prevCost.step2Hkd*100 : null;
          return(
            <tr key={r.id} style={{borderBottom:`1px solid ${T.border}`,background:i%2===0?T.bg:T.surface}}>
              <TD mono><span style={{color:T.muted}}>{r.nHK||"—"}</span></TD>
              <TD mono><span style={{color:T.gold}}>{r.code}</span></TD>
              <TD>{r.description}</TD>
              <TD><Chip label={r.ubicazione||"—"} color={r.ubicazione==="FOR"?T.purple:r.ubicazione==="MTS"?T.blue:T.green}/></TD>
              {branch==="CAN" ? <TD><span style={{color:T.muted,fontSize:"11px"}}>{r.area||"NORD"}</span></TD> : <TD/>}
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
    IDB.set("ifb_logistics", next);
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
          iNHK: fi(["nhk","n hk","n comit","comit","ncomit"]),
          iIFB: fi(["no_(ifb)","noifb","ifb","no_"]),
          iUb: fi(["mts/mto","mtsmto","ubicazione","location","wh"]),
          iArea: fi(["area","zona","portoimbarco","porto imbarco","porto di partenza","nord/sud","nordsud"]),
          iPlt: fi(["npltxcontainer","pltxcontainer","plt x container","nplt","pltpercontainer","n plt","palletpercontainer","numeropallet","pallet per container","npalletcontainer","palletcontainer"]),
          iCert: fi(["healthcertificate","health certificate","cert"]),
          iTemp: fi(["rettificata","temperature","temp","trettificata","camion"]),
          iCarriage: fi(["pltcostmedio","plt cost medio","pltcost","carriage"]),
          iAirSea: fi(["air/sea","airsea"]),
          iTransport: fi(["trasporto","transport","air/sea","airsea","air","sea"]),
          iAlcTax: fi(["tassa alcolica","tassaalcolica","alcolica","alctax","alc tax","aiem"]),
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
    const { iNHK, iIFB, iUb, iArea, iPlt, iCert, iTemp, iCarriage, iAirSea, iTransport, iAlcTax } = colIdx;
    let next = [...logistics];
    let countLog = 0, countAir = 0;
    const currentBranch = branch;
  
    logRawRows.forEach(row => {
      // Ottieni i codici
      // N COMIT / N HK — try dedicated col first, fallback to first column (CAN "N" col)
      const nhkRaw = iNHK >= 0 ? String(row[iNHK] || "").trim() : String(row[0] || "").trim();
      const ifbRaw = iIFB >= 0 ? String(row[iIFB] || "").trim() : "";
      if (!nhkRaw && !ifbRaw) return;

      // Trova il prodotto
      const prod = findProduct(nhkRaw, products, xrefs) || findProduct(ifbRaw, products, xrefs);
      if (!prod) return;

      // Trasporto: colonna dedicata "Trasporto" o colonna "AIR/SEA"
      const transportRaw = (iTransport >= 0 ? String(row[iTransport] || "") : "").trim().toUpperCase();
      if (transportRaw === "AIR") { countAir++; return; }
      const airSeaRaw = iAirSea >= 0 ? String(row[iAirSea] || "").trim().toUpperCase() : "";
      if (airSeaRaw === "AIR") { countAir++; return; }
      // Normalise transport value
      let transport = "";
      if (transportRaw === "MARE" || transportRaw === "SEA") transport = "MARE";
      else if (transportRaw === "GOMMA" || transportRaw === "TRUCK" || transportRaw === "ROAD") transport = "GOMMA";
  
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
  
      // Se Area è valorizzata (NORD/SUD/CENTRO) → MARE; se vuota e transport non specificato → GOMMA
      const effectiveTransport = transport || (area !== "NORD" || areaRaw ? (areaRaw ? "MARE" : "GOMMA") : "");
      const finalTransport = areaRaw ? "MARE" : transport || "";

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
        temperatureOverride,
        ...(finalTransport ? { transport: finalTransport } : {}),
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
    IDB.set("ifb_logistics", next);
  
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
  // Quando si cerca: mostra tutti i risultati. Senza ricerca: rispetta filtro missing/con logistica
  const displayed = search
    ? allProds
    : showOnlyMissing
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
            IDB.set("ifb_logistics", newLog);
            bumpImportTs();
            showToast(`Logistica ripristinata: ${branchEntries.length} righe ✓`, T.gold);
          }
        }
        e.target.value = "";
      }}
      style={{ ...inputStyle(), width: "auto", fontSize: "12px" }}
      defaultValue=""
    >
      {importLogs.filter((l:any) => l.type === "logistics" && l.branch === branch).length === 0
        ? <option value="">📜 Storico — nessun import precedente</option>
        : <>
            <option value="">📜 Carica da storico ({importLogs.filter((l:any) => l.type === "logistics" && l.branch === branch).length})</option>
            {importLogs.filter((l:any) => l.type === "logistics" && l.branch === branch).map((s: any) => (
              <option key={s.id} value={String(s.id)}>
                {new Date(s.id).toLocaleDateString("it-IT")} · {s.count} righe
              </option>
            ))}
          </>
      }
    </select>
    
    {/* Bottone Svuota dati esistente */}
    <button
      onClick={() => {
        if(window.confirm(`⚠️ ATTENZIONE: Eliminare TUTTI i dati logistici per ${branch}?`)) {
          const newLogistics = logistics.filter((l:any) => l.branch !== branch);
          setLogistics(newLogistics);
          IDB.set("ifb_logistics", newLogistics);
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
    
    <span style={{fontSize:"11px", color:T.muted}}>Colonne: {branchN(branch)} / No_(IFB) / Ubicazione / Area / Cert / Carriage / TASSA ALCOLICA / AIR/SEA</span>
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
              {["IFB No",branchN(branch),"Descrizione","Ubicaz.","Area","Plt/Cont","Cert.","Alcol >30°","Carriage","Conv."].map(c=>(
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
  const topScrollInnerRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const topScroll = topScrollRef.current;
    const tableScroll = tableScrollRef.current;
    const topInner = topScrollInnerRef.current;
    if (!topScroll || !tableScroll || !topInner) return;

    // sync scroll positions
    const handleTopScroll = () => { tableScroll.scrollLeft = topScroll.scrollLeft; };
    const handleTableScroll = () => { topScroll.scrollLeft = tableScroll.scrollLeft; };
    topScroll.addEventListener('scroll', handleTopScroll);
    tableScroll.addEventListener('scroll', handleTableScroll);

    // sync phantom width to actual table scroll width
    const syncWidth = () => { topInner.style.width = tableScroll.scrollWidth + "px"; };
    syncWidth();
    const ro = new ResizeObserver(syncWidth);
    ro.observe(tableScroll);

    return () => {
      topScroll.removeEventListener('scroll', handleTopScroll);
      tableScroll.removeEventListener('scroll', handleTableScroll);
      ro.disconnect();
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
          sub={branch==="MAC"?`${calc.length} articoli · SC HKD × markup × ${HKD_TO_MOP} HKD/MOP`:`${calc.length} calcolati · INALCA F&B · SEA`}/>
      {branch==="MAC"&&costRows.length===0&&(
        <div style={{padding:"32px",textAlign:"center",color:T.muted,fontSize:"13px"}}>
          ⚠ Nessun articolo MAC con Standard Cost. Vai in <strong>Anagrafica</strong> e importa il file Macao (con colonna <code>standardcost</code>).
        </div>
      )}
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
    {key:"flagged",        label:"Variazioni ≥3% New SC",    col:T.orange},
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
  <div ref={topScrollInnerRef} style={{ width: "1px", height: "12px" }} />
</div>

<div ref={tableScrollRef} style={{overflowX:"auto",width:"100%"}}>
        <table style={{borderCollapse:"collapse",width:"max-content",minWidth:"100%"}}>
          <thead>
            {/* riga gruppi */}
            <tr style={stickyTop0}>
              {branch==="MAC" ? <>
                <GH span={3}/>
                <GH span={2}/>
                <GH span={1} accent={T.blue}>HK Reference</GH>
                <GH span={1} accent={T.dim}>Markup</GH>
                <GH span={1} accent={T.green}>New Standard Cost</GH>
                <GH span={2}/>
              </> : <>
                <GH span={3}/>
                <GH span={branch==="CAN"?5:4}/>
                {branch==="CAN"
                  ? <GH span={4} accent={T.blue}>Costi trasporto (€/unit)</GH>
                  : <GH span={7} accent={T.blue}>Costi trasporto e dazi (€/unit)</GH>}
                <GH span={1} accent={T.purple}>Magazzino</GH>
                {branch==="CAN"
                  ? <GH span={4} accent={T.green}>New Standard Cost</GH>
                  : <GH span={2} accent={T.green}>New Standard Cost</GH>}
                <GH span={2}/>
              </>}
            </tr>
            {/* riga colonne */}
            <tr style={stickyTop22}>
              {branch==="MAC" ? <>
                <TH align="left" sticky w={70}>N HK</TH>
                <TH align="left" w={70}>IFB No</TH>
                <TH align="left" w={180}>Descrizione</TH>
                <TH w={55}>UOM</TH>
                <TH accent={T.orange} w={75} align="center">HOFF</TH>
                <TH accent={T.blue} w={80}>HK SC (HKD)</TH>
                <TH accent={T.dim} w={60}>Markup</TH>
                <TH accent={T.green} w={90}>New SC (MOP) ✓</TH>
                <TH w={60}>Δ%</TH>
                <TH w={90}>Ultimo ordine</TH>
              </> : <>
              <TH align="left" sticky w={70}>{branchN(branch)}</TH>
              <TH align="left" w={70}>IFB No</TH>
              <TH align="left" w={180}>Descrizione</TH>
              <TH w={60}>UOM</TH>
              <TH w={55}>Ubicaz.</TH>
              {branch==="CAN"&&<TH w={65} align="center">Tratta</TH>}
              <TH w={55} align="center">Temp.</TH>
              <TH w={55} align="center">Rettif.</TH>
              <TH accent={T.blue} w={70}>Prezzo €</TH>
              {branch==="CAN" ? <>
                <TH accent={T.blue} w={65}>Trasp.</TH>
                <TH accent={T.blue} w={60}>Pallet</TH>
                <TH accent={T.blue} w={60}>AIEM</TH>
              </> : <>
                <TH accent={T.blue} w={65}>FOB</TH>
                <TH accent={T.blue} w={65}>LIC</TH>
                <TH accent={T.blue} w={55}>VGM</TH>
                <TH accent={T.blue} w={55}>Cert.</TH>
                <TH accent={T.blue} w={60}>Pallet</TH>
                <TH accent={T.blue} w={60}>Alc.Tax</TH>
              </>}
              <TH accent={T.purple} w={65}>WH €</TH>
              {branch==="CAN" ? <>
                <TH accent={T.green} w={72}>New SC GC</TH>
                <TH accent={T.green} w={72}>New SC TF</TH>
                <TH accent={T.green} w={72}>New SC LAN</TH>
                <TH accent={T.green} w={85}>New SC FUE ✓</TH>
              </> : <>
                <TH accent={T.green} w={72}>New SC €</TH>
                <TH accent={T.green} w={85}>New SC HKD ✓</TH>
              </>}
              <TH w={60}>Δ%</TH>
              <TH w={90}>Ultimo ordine</TH>
              </>}
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

              // ── MAC row ──────────────────────────────────────────────────
              if (branch==="MAC") {
                const mc = r.cost;
                return (<React.Fragment key={r.id}>
                  <tr style={{background:isSelected?`${T.gold}08`:rowBg,cursor:"pointer"}}
                    onClick={()=>setShowDetail((v:any)=>v===r.id?null:r.id)}>
                    <td style={{...cellL(true),background:isSelected?`${T.gold}08`:rowBg}}>
                      <span style={{color:T.muted,fontFamily:"monospace",fontSize:"10px"}}>{r.nHK||r.code||"—"}</span>
                    </td>
                    <td style={cellL()}><span style={{color:T.gold,fontFamily:"monospace",fontSize:"10px"}}>{r.code}</span></td>
                    <td style={{...cellL(),maxWidth:"200px",overflow:"hidden",textOverflow:"ellipsis"}}>{r.description}</td>
                    <td style={cell()}>{r.uom||"—"}</td>
                    <td style={{...cell(),textAlign:"center"}}>
                      <Chip label={r.isHoff?"HOFF":"NON-HOFF"} color={r.isHoff?T.orange:T.blue}/>
                    </td>
                    <td style={cell(T.blue)}>{mc?`${mc.hkNewSC.toFixed(2)}`:"—"}</td>
                    <td style={cell(T.dim)}>{mc?`+${mc.markup.toFixed(0)}%`:"—"}</td>
                    <td style={cell(mc&&r.uomDiffers&&mc.macToHkConv===1?T.orange:T.green,true)}>
                      <span style={{fontSize:"11px",fontWeight:"bold"}}>
                        {mc?<>
                          {`${mc.macNewSC.toFixed(2)}`}
                          {mc.macToHkConv>1&&<span style={{fontSize:"8px",color:T.purple,marginLeft:"3px"}}>×{mc.macToHkConv}</span>}
                          {r.uomDiffers&&mc.macToHkConv===1&&<span style={{fontSize:"8px",color:T.orange,marginLeft:"3px"}}>⚠UOM</span>}
                        </>:<span style={{color:T.dim,fontSize:"9px"}}>{r.skipReason||"—"}</span>}
                      </span>
                    </td>
                    <td style={cell(pct==null?T.dim:Math.abs(pct)>=3?(pct>0?T.red:T.green):T.muted,Math.abs(pct||0)>=3)}>
                      {pct!=null?(pct>0?"+":"")+pct.toFixed(1)+"%":"—"}
                    </td>
                    <td style={{...cell(),textAlign:"center"}}>
                      {!lastD?<span style={{color:T.dim}}>—</span>:isOld?<div style={{lineHeight:1.2}}><div style={{color:T.orange,fontWeight:"bold",fontSize:"9px"}}>⚠ KEEP OLD</div><div style={{color:T.dim,fontSize:"9px"}}>{lastD.toLocaleDateString("it-IT")}</div></div>:<span style={{color:T.muted}}>{lastD.toLocaleDateString("it-IT")}</span>}
                    </td>
                  </tr>
                  {isSelected&&mc&&(
                    <tr key={r.id+"_detail"}>
                      <td colSpan={10} style={{padding:"10px 20px",background:`${T.gold}06`,borderBottom:`1px solid ${T.gold}33`}}>
                        <div style={{fontSize:"9px",color:T.gold,letterSpacing:"2px",textTransform:"uppercase",marginBottom:"8px"}}>
                          Breakdown New Standard Cost · Macao · {r.isHoff?"HOFF (House of Fine Foods)":"NON-HOFF"}
                        </div>
                        <table style={{borderCollapse:"collapse"}}>
                          <tbody>
                            {[
                              ["SC HK / FOB (HKD)",`HKD ${mc.hkNewSC.toFixed(4)}`,`Costo pagato da Macao a HK${mc.macToHkConv>1?` (per ${r.hkUom||"HK UOM"})`:""}`,T.blue],
                              ...(mc.macToHkConv>1?[[`Conv. UOM (×${mc.macToHkConv})`,`HKD ${(mc.hkNewSC*mc.macToHkConv).toFixed(4)}`,`${r.hkUom||"HK UOM"} → ${r.macUom||"MAC UOM"}: × ${mc.macToHkConv}`,T.purple]]:[]),
                              [`Markup ${r.isHoff?"HOFF":"non-HOFF"}`,`+ ${mc.markup.toFixed(0)}%`,r.isHoff?"HOFF (House of Fine Foods): +3%":"Non-HOFF: +10%",T.orange],
                              ["Tasso HKD → MOP",`× ${HKD_TO_MOP}`,"Tasso di cambio fisso",T.muted],
                              ["Base in MOP",`MOP ${mc.baseInMop?.toFixed(4)||"—"}`,`HKD ${mc.hkNewSC.toFixed(4)} × ${mc.macToHkConv>1?mc.macToHkConv+"×":""}${(1+mc.markup/100).toFixed(2)} × ${HKD_TO_MOP}`,T.muted],
                              [`Logistica ALL-IN (${mc.temperature||"DRY"})`,`+ MOP ${mc.logPerUom?.toFixed(4)||"—"}`,`${mc.logPerKg} MOP/kg × ${(mc.logPerUom/mc.logPerKg||1).toFixed(3)} kg/${r.macUom||"UOM"} · include: BV Whs→HK Port, Ferry, Dogana, Consegna Macao`,T.blue],
                              ["NEW SC MAC (MOP)",`MOP ${mc.macNewSC.toFixed(4)}`,`Base MOP + Logistica`,T.green],
                            ].map(([k,v,f,col]:any[])=>(
                              <tr key={String(k)}>
                                <td style={{padding:"3px 12px 3px 0",fontSize:"11px",color:T.muted,whiteSpace:"nowrap"}}>{k}</td>
                                <td style={{padding:"3px 10px",fontSize:"11px",color:col,fontWeight:"bold",fontFamily:"monospace",textAlign:"right"}}>{v}</td>
                                <td style={{padding:"3px 0 3px 14px",fontSize:"10px",color:T.dim,fontStyle:"italic"}}>{f}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                        <div style={{marginTop:"6px",fontSize:"9px",color:T.dim}}>
                          u/plt: {(mc.unitsPerPlt||0).toFixed(2)} · HK UOM: {r.hkUom||"—"} · MAC UOM: {r.macUom||"—"}
                        </div>
                      </td>
                    </tr>
                  )}
                </React.Fragment>);
              }
              // ── END MAC row ───────────────────────────────────────────────

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
                  {branch==="CAN"&&(()=>{
                    const logEntry = logistics?.find((l:any)=>l.productId===r.id&&l.branch===branch);
                    const tr = logEntry?.transport||"";
                    return (
                      <td style={{...cell(),textAlign:"center"}}>
                        {tr==="MARE"
                          ? <Chip label="🚢 MARE" color={T.blue}/>
                          : tr==="GOMMA"
                          ? <Chip label="🚛 GOMMA" color={T.orange}/>
                          : <span style={{color:T.dim,fontSize:"9px"}}>—</span>}
                      </td>
                    );
                  })()}

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
                  {branch==="CAN" ? <>
                    {/* trasporto (GC come rappresentativo) */}
                    <td style={cell()}>{c?f4(c.freightGC+c.inlandGC+c.veronaBarcUnit+c.barcUnitGC+c.assicUnit):"—"}</td>
                    <td style={cell()}>{c?f4(c.plt):"—"}</td>
                    <td style={cell(c?.aiemUnit>0?T.orange:undefined)}>{c?(c.aiemUnit>0?f4(c.aiemUnit):"—"):"—"}</td>
                  </> : <>
                    <td style={cell()}>{c?f4(c.fob):"—"}</td>
                    <td style={cell()}>{c?f4(c.lic):"—"}</td>
                    <td style={cell()}>{c?f4(c.vgm):"—"}</td>
                    <td style={cell(c?.hc>0?T.orange:undefined)}>{c?(c.hc>0?f4(c.hc):"—"):"—"}</td>
                    <td style={cell()}>{c?f4(c.plt):"—"}</td>
                    <td style={cell(c?.alc>0?T.orange:undefined)}>{c?(c.alc>0?f4(c.alc):"—"):"—"}</td>
                  </>}

                  {/* magazzino */}
                  <td style={cell(T.purple)}>{c?(c.wh>0?f4(c.wh):"—"):"—"}</td>

                  {/* step 2 */}
                  {branch==="CAN" ? <>
                    <td style={cell(T.green,true)}>{c?`€${c.step2GC.toFixed(4)}`:"—"}</td>
                    <td style={cell(T.green,true)}>{c?`€${c.step2TF.toFixed(4)}`:"—"}</td>
                    <td style={cell(T.green,true)}>{c?`€${c.step2LAN.toFixed(4)}`:"—"}</td>
                    <td style={cell(T.green,true)}>
                      <span style={{fontSize:"11px",fontWeight:"bold"}}>
                        {c?`€${c.step2FUE.toFixed(4)}`:<span style={{color:T.dim,fontWeight:"normal",fontSize:"9px"}}>{r.skipReason||"—"}</span>}
                      </span>
                    </td>
                  </> : <>
                    <td style={cell(T.green,true)}>{c?`€${c.step2Eur.toFixed(4)}`:"—"}</td>
                    <td style={cell(T.green,true)}>
                      <span style={{fontSize:"11px",fontWeight:"bold"}}>
                        {hkd!=null?`${hkd.toFixed(2)}`:<span style={{color:T.dim,fontWeight:"normal",fontSize:"9px"}}>{r.skipReason||"—"}</span>}
                      </span>
                    </td>
                  </>}

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
                    <td colSpan={branch==="CAN"?19:19} style={{padding:"10px 20px",background:`${T.gold}06`,
                      borderBottom:`1px solid ${T.gold}33`}}>
                      {branch==="CAN" ? (() => {
                        const isMARE = c.isMARE || c.transport==="MARE";
                        const f4=(v:number)=>`€ ${v.toFixed(4)}`;
                        const row=(label:string,gcTf:string,lanFue:string,formula:string,color=T.text)=>(
                          <tr key={label}>
                            <td style={{padding:"3px 10px 3px 0",fontSize:"11px",color:T.muted,whiteSpace:"nowrap"}}>{label}</td>
                            <td style={{padding:"3px 8px",fontSize:"11px",color:color,fontWeight:"bold",fontFamily:"monospace",textAlign:"right"}}>{gcTf}</td>
                            <td style={{padding:"3px 8px",fontSize:"11px",color:color,fontWeight:"bold",fontFamily:"monospace",textAlign:"right"}}>{lanFue}</td>
                            <td style={{padding:"3px 0 3px 14px",fontSize:"10px",color:T.dim,fontStyle:"italic"}}>{formula}</td>
                          </tr>
                        );
                        const sep=(label:string)=>(
                          <tr key={"sep"+label}><td colSpan={4} style={{padding:"2px 0",borderTop:`1px solid ${T.border}`,fontSize:"9px",color:T.dim,letterSpacing:"1px",textTransform:"uppercase",paddingTop:"6px"}}>{label}</td></tr>
                        );
                        return (
                          <div style={{display:"flex",gap:"32px",flexWrap:"wrap"}}>
                            <div>
                              <div style={{fontSize:"9px",color:T.gold,letterSpacing:"2px",textTransform:"uppercase",marginBottom:"8px"}}>Breakdown New Standard Cost · {c.transport||"GOMMA"} · {r.temperature||"DRY"} · {r.ubicazione||"MTS"}</div>
                              <table style={{borderCollapse:"collapse"}}>
                                <thead><tr>
                                  <th style={{padding:"2px 10px 4px 0",fontSize:"9px",color:T.dim,textAlign:"left",letterSpacing:"1px"}}>VOCE</th>
                                  <th style={{padding:"2px 8px 4px",fontSize:"9px",color:T.dim,textAlign:"right",letterSpacing:"1px"}}>GC / TF</th>
                                  <th style={{padding:"2px 8px 4px",fontSize:"9px",color:T.dim,textAlign:"right",letterSpacing:"1px"}}>LAN / FUE</th>
                                  <th style={{padding:"2px 0 4px 14px",fontSize:"9px",color:T.dim,textAlign:"left",letterSpacing:"1px"}}>FORMULA</th>
                                </tr></thead>
                                <tbody>
                                  {sep("Costo acquisto")}
                                  {row("Prezzo acquisto",f4(c.priceEur),f4(c.priceEur),r.ubicazione==="FOR"?"Da listino (FCA)":r.ubicazione==="MTS"?"Da listino (MTS)":"Da listino (DAP Verona)",T.text)}
                                  {isMARE ? sep("Trasporto MARE") : sep("Trasporto GOMMA")}
                                  {isMARE ? <>
                                    {row("Freight MARE",f4(c.freightGC),f4(c.freightLAN||0),`MARE[${r.area||"NORD"}] ÷ (u/plt × plt/cont)`,T.blue)}
                                    {row("Inland",f4(c.inlandGC),f4(c.freightLAN||0),"INLAND ÷ (u/plt × plt/cont)",T.blue)}
                                  </> : <>
                                    {row("Verona → Barcellona",f4(c.veronaBarcUnit),f4(c.veronaBarcUnit),"62,50 € ÷ u/plt (COSTS LOG!D5)",T.blue)}
                                    {row("Barc → Isola",f4(c.barcUnitGC),f4(c.barcUnitLAN||0),"BARC[temp][isola] ÷ u/plt",T.blue)}
                                    {row("Assicurazione",f4(c.assicUnit),f4(c.assicUnit),"Prezzo × 0,5%",T.blue)}
                                  </>}
                                  {sep("Pallet & AIEM")}
                                  {row("Pallet",f4(c.plt),f4(c.plt),"15 € ÷ u/plt (COSTS LOG!I1)",T.blue)}
                                  {c.aiemGCTF>0&&row("AIEM",f4(c.aiemGCTF),f4(c.aiemLANFUE||0),"(Prezzo + Trasporto isola) × AIEM%",T.orange)}
                                  {sep("Magazzino")}
                                  {row("WH / unit",c.wh>0?f4(c.wh):"—",c.wh>0?f4(c.wh):"—",
                                    r.ubicazione==="MTO"?"MTO[temp] ÷ u/plt":r.ubicazione==="MTS"?"MTS-D + MTS-I ÷ u/plt + MTS-P ÷ collo":"—",T.purple)}
                                  {sep("New Standard Cost")}
                                  <tr>
                                    <td style={{padding:"4px 10px 4px 0",fontSize:"12px",color:T.green,fontWeight:"bold"}}>NEW SC GC</td>
                                    <td colSpan={2} style={{padding:"4px 8px",fontSize:"13px",color:T.green,fontWeight:"bold",fontFamily:"monospace",textAlign:"right"}}>{f4(c.step2GC)}</td>
                                    <td style={{padding:"4px 0 4px 14px",fontSize:"10px",color:T.dim,fontStyle:"italic"}}>Prezzo + Trasp.GC + Pallet + AIEM + WH</td>
                                  </tr>
                                  <tr>
                                    <td style={{padding:"2px 10px 2px 0",fontSize:"11px",color:T.green}}>NEW SC TF</td>
                                    <td colSpan={2} style={{padding:"2px 8px",fontSize:"11px",color:T.green,fontFamily:"monospace",textAlign:"right"}}>{f4(c.step2TF)}</td>
                                    <td style={{padding:"2px 0 2px 14px",fontSize:"10px",color:T.dim,fontStyle:"italic"}}>= GC (stesse tariffe)</td>
                                  </tr>
                                  <tr>
                                    <td style={{padding:"2px 10px 2px 0",fontSize:"11px",color:T.green}}>NEW SC LAN</td>
                                    <td colSpan={2} style={{padding:"2px 8px",fontSize:"11px",color:T.green,fontFamily:"monospace",textAlign:"right"}}>{f4(c.step2LAN)}</td>
                                    <td style={{padding:"2px 0 2px 14px",fontSize:"10px",color:T.dim,fontStyle:"italic"}}>Prezzo + Trasp.LAN + Pallet + AIEM + WH</td>
                                  </tr>
                                  <tr>
                                    <td style={{padding:"2px 10px 2px 0",fontSize:"11px",color:T.green}}>NEW SC FUE</td>
                                    <td colSpan={2} style={{padding:"2px 8px",fontSize:"11px",color:T.green,fontFamily:"monospace",textAlign:"right"}}>{f4(c.step2FUE)}</td>
                                    <td style={{padding:"2px 0 2px 14px",fontSize:"10px",color:T.dim,fontStyle:"italic"}}>= LAN (stesse tariffe)</td>
                                  </tr>
                                </tbody>
                              </table>
                              <div style={{marginTop:"8px",fontSize:"9px",color:T.dim}}>
                                u/plt: {(c.unitsPerPlt||0).toFixed(2)} · plt/cont: {r.pltPerContainer||"—"} · Temp: {r.temperature||"DRY"} · Area: {r.area||"NORD"}
                              </div>
                            </div>
                          </div>
                        );
                      })() : (() => {
                        const f4=(v:number)=>`€ ${v.toFixed(4)}`;
                        const row=(label:string,val:string,formula:string,color=T.text,bold=false)=>(
                          <tr key={label}>
                            <td style={{padding:"3px 12px 3px 0",fontSize:"11px",color:T.muted,whiteSpace:"nowrap"}}>{label}</td>
                            <td style={{padding:"3px 10px",fontSize:"11px",color,fontWeight:bold?"bold":"normal",fontFamily:"monospace",textAlign:"right"}}>{val}</td>
                            <td style={{padding:"3px 0 3px 14px",fontSize:"10px",color:T.dim,fontStyle:"italic"}}>{formula}</td>
                          </tr>
                        );
                        const sep=(label:string)=>(
                          <tr key={"sep"+label}><td colSpan={3} style={{padding:"6px 0 2px",borderTop:`1px solid ${T.border}`,fontSize:"9px",color:T.dim,letterSpacing:"1px",textTransform:"uppercase"}}>{label}</td></tr>
                        );
                        return (
                          <div>
                            <div style={{fontSize:"9px",color:T.gold,letterSpacing:"2px",textTransform:"uppercase",marginBottom:"8px"}}>Breakdown New Standard Cost · {r.ubicazione||"—"} · {r.temperature||"DRY"}</div>
                            <table style={{borderCollapse:"collapse"}}>
                              <thead><tr>
                                <th style={{padding:"2px 12px 4px 0",fontSize:"9px",color:T.dim,textAlign:"left",letterSpacing:"1px"}}>VOCE</th>
                                <th style={{padding:"2px 10px 4px",fontSize:"9px",color:T.dim,textAlign:"right",letterSpacing:"1px"}}>€ / UNIT</th>
                                <th style={{padding:"2px 0 4px 14px",fontSize:"9px",color:T.dim,textAlign:"left",letterSpacing:"1px"}}>FORMULA</th>
                              </tr></thead>
                              <tbody>
                                {sep("Costo acquisto")}
                                {row("Prezzo acquisto",f4(c.priceEur),"Da listino (DAP/FCA del mese)",T.text)}
                                {sep("Trasporto e dazi")}
                                {row("FOB / unit",f4(c.fob),"Freight On Board — da tabella COSTS",T.blue)}
                                {row("LIC / unit",f4(c.lic),"Local Import Charges — da tabella COSTS",T.blue)}
                                {row("VGM / unit",f4(c.vgm),"Verified Gross Mass — da tabella COSTS",T.blue)}
                                {c.hc>0&&row("Certificati / unit",f4(c.hc),"Health / import certificate — da tabella COSTS",T.blue)}
                                {row("Pallet / unit",f4(c.plt),`Costo pallet ÷ ${(c.unitsPerPlt||0).toFixed(2)} u/plt`,T.blue)}
                                {c.alc>0&&row("Alc. Tax / unit",f4(c.alc),"Tassa alcol — da anagrafica articolo",T.orange)}
                                {sep("Magazzino")}
                                {row("WH / unit",c.wh>0?f4(c.wh):"—",
                                  r.ubicazione==="MTO"?"MTO[temp] ÷ u/plt":
                                  r.ubicazione==="MTS"?"MTS-D + MTS-I ÷ u/plt + MTS-P ÷ collo":"—",T.purple)}
                                {sep("New Standard Cost")}
                                <tr>
                                  <td style={{padding:"4px 12px 4px 0",fontSize:"12px",color:T.green,fontWeight:"bold"}}>NEW SC €</td>
                                  <td style={{padding:"4px 10px",fontSize:"13px",color:T.green,fontWeight:"bold",fontFamily:"monospace",textAlign:"right"}}>{f4(c.step2Eur)}</td>
                                  <td style={{padding:"4px 0 4px 14px",fontSize:"10px",color:T.dim,fontStyle:"italic"}}>Prezzo + FOB + LIC + VGM + Cert. + Pallet + AlcTax + WH</td>
                                </tr>
                                <tr>
                                  <td style={{padding:"2px 12px 2px 0",fontSize:"11px",color:T.green}}>NEW SC HKD</td>
                                  <td style={{padding:"2px 10px",fontSize:"12px",color:T.green,fontWeight:"bold",fontFamily:"monospace",textAlign:"right"}}>{`HKD ${c.step2Hkd.toFixed(2)}`}</td>
                                  <td style={{padding:"2px 0 2px 14px",fontSize:"10px",color:T.dim,fontStyle:"italic"}}>{`New SC € × rate ${c.rate}`}</td>
                                </tr>
                              </tbody>
                            </table>
                            <div style={{marginTop:"8px",fontSize:"9px",color:T.dim}}>
                              u/plt: {(c.unitsPerPlt||0).toFixed(2)} · Rate HKD: {c.rate} · {r.ubicazione||"—"}
                            </div>
                          </div>
                        );
                      })()}
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



function InvoiceAndCosts({rows,setRows,branch,airList,products,xrefs,costRows,logistics,snapshots,setSnapshots,importLogs,setImportLogs,showToast,bumpImportTs}) {
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
      return{itemCode:code,description,date:dateStr,qty,unitPrice,isSample:qty>0&&(unitPrice===0||unitPrice===0.01),location,nHK,transport:isAirProd?"AIR":"SEA",_prodFound:!!prod};
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
        const newHkd=cr?.cost?.step2Hkd??null; // GC canonico (usato per filtri)
        const oldHkd=cr?.prevCost?.step2Hkd??null;
        const pct=newHkd!=null&&oldHkd!=null&&oldHkd>0?(newHkd-oldHkd)/oldHkd*100:null;
        const skipReason=isAir?"AIR":cr?.skipReason||(!prod?"NON IN ANAGRAFICA":"");
        const logEntry=prod?logistics?.find((l:any)=>l.productId===prod.id&&l.branch===branch):null;
        const logTransport=logEntry?.transport||"";
        // CAN: 4 isole
        const scGC  = cr?.cost?.step2GC  ?? null;
        const scTF  = cr?.cost?.step2TF  ?? null;
        const scLAN = cr?.cost?.step2LAN ?? null;
        const scFUE = cr?.cost?.step2FUE ?? null;
        return{...r,nHK:prod?.nHK||r.nHK||"",ifbNo:prod?.code||r.itemCode||"",
          description:r.description||prod?.description||"",ubicazione:cr?.ubicazione||"",logTransport,
          isAir,locationIsNCJ,mismatch,newHkd,oldHkd,pct,skipReason,scGC,scTF,scLAN,scFUE};
      });
  },[activeRows,costRows,products,xrefs,sortDir]);

  const mismatches  = enriched.filter((r:any)=>r.mismatch);
  const airCount    = enriched.filter((r:any)=>r.isAir).length;
  const uniqueNHK   = [...new Set(enriched.map((r:any)=>r.nHK).filter(Boolean))].sort() as string[];
  const uniqueIFBNo = [...new Set(enriched.map((r:any)=>r.ifbNo).filter(Boolean))].sort() as string[];

  let displayed=enriched as any[];
  if(excludeAir)               displayed=displayed.filter(r=>!r.isAir);
  if(filterTransport==="air")           displayed=displayed.filter(r=>r.isAir);
  else if(filterTransport==="sea")      displayed=displayed.filter(r=>!r.isAir);
  else if(filterTransport==="mismatch") displayed=displayed.filter(r=>r.mismatch);
  else if(filterTransport==="gomma")    displayed=displayed.filter(r=>r.logTransport==="GOMMA");
  else if(filterTransport==="mare")     displayed=displayed.filter(r=>r.logTransport==="MARE");
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
        {([
          [preview.length,"Totale",T.text],
          ...(branch!=="CAN"?[[preview.filter((r:any)=>r.transport==="AIR").length,"✈ AIR",T.orange]]:[] as any),
          ...(branch!=="CAN"?[[preview.filter((r:any)=>r.transport==="SEA").length,"⛴ SEA",T.blue]]:[] as any),
          [preview.filter((r:any)=>!r._prodFound).length,"⚠ Non in anagrafica",T.red],
        ] as [number,string,string][]).map(([n,l,c])=>(
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
            "Old SC":r.oldHkd!=null?roundN(r.oldHkd):"",
            ...(branch==="CAN"
              ? {"SC GC":r.isAir?"AIR":r.scGC!=null?roundN(r.scGC):"MANCANTE",
                 "SC TF":r.isAir?"AIR":r.scTF!=null?roundN(r.scTF):"MANCANTE",
                 "SC LAN":r.isAir?"AIR":r.scLAN!=null?roundN(r.scLAN):"MANCANTE",
                 "SC FUE":r.isAir?"AIR":r.scFUE!=null?roundN(r.scFUE):"MANCANTE"}
              : {"New SC":r.isAir?"AIR":(r.unitPrice===0||r.unitPrice===0.01)?"SAMPLE":r.newHkd!=null?roundN(r.newHkd):"MANCANTE"}),
            "Δ%":r.pct!=null?roundN(r.pct,1):"","Motivo":r.skipReason||"",
          })),
          "Fatture & Costi",`Fatture_${branch}.xlsx`
        )} style={{padding:"6px 14px",background:`${T.green}20`,border:`1px solid ${T.green}44`,borderRadius:"6px",color:T.green,cursor:"pointer",fontSize:"11px"}}>
          ⬇ Export Excel
        </button>
        {branch!=="CAN"&&<button onClick={()=>setExcludeAir(v=>!v)}
          style={{padding:"6px 14px",background:excludeAir?`${T.orange}20`:T.surface,color:excludeAir?T.orange:T.muted,border:`1px solid ${excludeAir?T.orange:T.border}`,borderRadius:"6px",cursor:"pointer",fontSize:"11px",fontWeight:excludeAir?"bold":"normal"}}>
          {excludeAir?`✓ AIR esclusi (${airCount})`:`✈ Escludi AIR (${airCount})`}
        </button>}
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
        {((branch==="CAN"
          ? [
              ["all",`Tutte (${enriched.length})`,T.text],
              ["gomma",`🚛 GOMMA (${enriched.filter((r:any)=>r.logTransport==="GOMMA").length})`,T.orange],
              ["mare",`🚢 MARE (${enriched.filter((r:any)=>r.logTransport==="MARE").length})`,T.blue],
            ]
          : [
              ["all",`Tutte (${enriched.length})`,T.text],
              ["air",`✈ AIR (${airCount})`,T.orange],
              ["sea",`⛴ SEA (${enriched.length-airCount})`,T.blue],
              ["mismatch",`⚠ Mismatch (${mismatches.length})`,T.purple],
            ]) as [string,string,string][]).map(([v,l,c])=>(
          <button key={v} onClick={()=>setFilterTransport(v)}
            style={{padding:"5px 12px",background:filterTransport===v?`${c}20`:T.surface,color:filterTransport===v?c:T.muted,border:`1px solid ${filterTransport===v?c:T.border}`,borderRadius:"6px",cursor:"pointer",fontSize:"11px",fontWeight:filterTransport===v?"bold":"normal"}}>
            {l}
          </button>
        ))}
      </div>

      <SearchBar value={search} onChange={setSearch} placeholder={`🔍 Cerca codice, ${branchN(branch)}, descrizione, location…`}/>

      <Section title={`${displayed.length} righe`}>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse"}}>
            <thead><tr>
              {["Data",branchN(branch)+" ▾","IFB No ▾","Descrizione","Qty","Prezzo","Location","Mag./Trasp.","Old SC",...(branch==="CAN"?["SC GC","SC TF","SC LAN","SC FUE"]:["New SC ▾"]),"Δ%","Motivo"].map((c,ci)=>{
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
                if(c==="New SC ▾") return(
                  <th key={c} style={{padding:"4px 8px",background:T.card,borderBottom:`1px solid ${T.border}`,position:"sticky",top:0,zIndex:10}}>
                    <select value={newHkdFilter} onChange={e=>setNewHkdFilter(e.target.value as any)}
                      style={{background:newHkdFilter!=="all"?`${T.gold}22`:T.card,color:newHkdFilter!=="all"?T.gold:T.muted,border:`1px solid ${newHkdFilter!=="all"?T.gold:T.border}`,borderRadius:"4px",padding:"3px 6px",fontSize:"10px",cursor:"pointer",fontFamily:"inherit",outline:"none"}}>
                      <option value="all">New SC ▾</option>
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
                    {branch==="CAN" ? (
                      <>
                        {([r.scGC,r.scTF,r.scLAN,r.scFUE] as (number|null)[]).map((v,i)=>(
                          <td key={i} style={{padding:"6px 10px",fontSize:"11px",fontFamily:"monospace",textAlign:"right"}}>
                            {r.isAir
                              ? <span style={{color:T.orange,fontWeight:"bold"}}>AIR</span>
                              : (r.unitPrice===0||r.unitPrice===0.01)
                                ? <span style={{color:T.purple,fontWeight:"bold"}}>SAMPLE</span>
                                : v!=null
                                  ? <span style={{color:T.gold,fontWeight:"bold"}}>{v.toFixed(2)}</span>
                                  : <span style={{color:T.red,fontWeight:"bold"}}>MANCANTE</span>
                            }
                          </td>
                        ))}
                      </>
                    ) : (
                      <td style={{padding:"6px 10px",fontSize:"11px",fontFamily:"monospace",textAlign:"right"}}>
                        {r.isAir
                          ? <span style={{color:T.orange,fontWeight:"bold"}}>AIR</span>
                          : (r.unitPrice===0||r.unitPrice===0.01)
                            ? <span style={{color:T.purple,fontWeight:"bold"}}>SAMPLE</span>
                            : r.newHkd!=null
                              ? <span style={{color:T.gold,fontWeight:"bold"}}>{r.newHkd.toFixed(2)}</span>
                              : <span style={{color:T.red,fontWeight:"bold"}}>MANCANTE</span>
                        }
                      </td>
                    )}
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


// ─── PRICE EXCEPTIONS ─────────────────────────────────────────────────────────
function PriceExceptions({branch, products, xrefs, priceExceptions, setPriceExceptions}) {
  const [search, setSearch] = useState("");
  const [selectedProd, setSelectedProd] = useState<any>(null);
  const [inputPrice, setInputPrice] = useState("");
  const [inputNote, setInputNote] = useState("");
  const [showSugg, setShowSugg] = useState(false);

  const isCAN = branch === "CAN";
  const branchExc = priceExceptions.filter((e:any) => e.branch === branch);

  // Build a lookup: xref productId → nHK/nCOMIT
  const xrefMap = useMemo(()=>{
    const m: Record<string,string> = {};
    xrefs.forEach((x:any)=>{ if(x.productId) m[String(x.productId)] = x.nHK || ""; });
    return m;
  }, [xrefs]);

  const activeProducts = products.filter(p => p.active);

  const suggestions = search.length >= 1
    ? activeProducts.filter(p => {
        const nVal = xrefMap[String(p.id)] || p.nHK || "";
        return (
          String(p.id).includes(search) ||
          (p.code||"").toLowerCase().includes(search.toLowerCase()) ||
          (p.description||"").toLowerCase().includes(search.toLowerCase()) ||
          nVal.toLowerCase().includes(search.toLowerCase())
        );
      }).slice(0,10)
    : [];

  function addException() {
    const price = parseFloat(inputPrice.replace(",","."));
    if(!selectedProd || isNaN(price) || price <= 0) return;
    const existing = priceExceptions.findIndex((e:any) =>
      e.branch === branch && (e.productId === selectedProd.id || (e.code && e.code === selectedProd.code))
    );
    const entry = {
      branch,
      productId: selectedProd.id,
      code: selectedProd.code,
      nHK: xrefMap[String(selectedProd.id)] || selectedProd.nHK || "",
      description: selectedProd.description,
      price,
      note: inputNote.trim(),
    };
    const updated = [...priceExceptions];
    if(existing >= 0) updated[existing] = entry;
    else updated.push(entry);
    setPriceExceptions(updated);
    setSelectedProd(null);
    setSearch("");
    setInputPrice("");
    setInputNote("");
  }

  function removeException(idx: number) {
    const exc = branchExc[idx];
    setPriceExceptions(priceExceptions.filter((e:any) =>
      !(e.branch === branch && e.productId === exc.productId)
    ));
  }

  return (
    <div style={{padding:"28px 32px",maxWidth:"900px"}}>
      <div style={{marginBottom:"24px"}}>
        <div style={{fontSize:"10px",letterSpacing:"3px",color:T.gold,textTransform:"uppercase",marginBottom:"4px"}}>Filiale · {branch}</div>
        <h2 style={{margin:0,fontSize:"20px",fontWeight:"bold"}}>⚡ Eccezioni Prezzi</h2>
        <div style={{color:T.muted,fontSize:"12px",marginTop:"6px"}}>
          Il prezzo qui inserito ha priorità assoluta su listino e listino carne per il calcolo del costo standard.
        </div>
      </div>

      {/* Form aggiunta */}
      <div style={{background:T.card,border:`1px solid ${T.border}`,borderRadius:"10px",padding:"18px 20px",marginBottom:"24px"}}>
        <div style={{fontSize:"11px",letterSpacing:"2px",color:T.gold,textTransform:"uppercase",marginBottom:"14px"}}>Aggiungi / Modifica eccezione</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 140px 1fr auto",gap:"10px",alignItems:"start"}}>

          {/* Ricerca articolo */}
          <div style={{position:"relative"}}>
            <div style={{fontSize:"10px",color:T.muted,marginBottom:"4px"}}>Articolo</div>
            <input
              value={selectedProd ? `[${selectedProd.id}] ${selectedProd.description}` : search}
              onChange={e=>{ setSearch(e.target.value); setSelectedProd(null); setShowSugg(true); }}
              onFocus={()=>setShowSugg(true)}
              placeholder={`Cerca per IFB No, ${isCAN?"N COMIT":"N HK"} o descrizione…`}
              style={{width:"100%",background:T.surface,border:`1px solid ${T.border}`,borderRadius:"6px",
                padding:"7px 10px",color:T.text,fontSize:"12px",fontFamily:"inherit",boxSizing:"border-box"}}
            />
            {showSugg && suggestions.length > 0 && !selectedProd && (
              <div style={{position:"absolute",top:"100%",left:0,right:0,background:T.card,
                border:`1px solid ${T.border}`,borderRadius:"6px",zIndex:100,maxHeight:"200px",overflowY:"auto"}}>
                {suggestions.map(p=>{
                  const nVal = xrefMap[String(p.id)] || p.nHK || "-";
                  return (
                    <div key={p.id}
                      onMouseDown={()=>{ setSelectedProd(p); setSearch(""); setShowSugg(false); }}
                      style={{padding:"7px 12px",cursor:"pointer",borderBottom:`1px solid ${T.border}`,fontSize:"12px"}}
                      onMouseEnter={e=>(e.currentTarget.style.background=T.surface)}
                      onMouseLeave={e=>(e.currentTarget.style.background="")}>
                      <span style={{color:T.gold,marginRight:"6px"}}>[{p.id}]</span>
                      <span>{p.description}</span>
                      <span style={{color:T.muted,marginLeft:"8px",fontSize:"11px"}}>{isCAN?"N COMIT":"N HK"}: {nVal}</span>
                    </div>
                  );
                })}
              </div>
            )}
          </div>

          {/* Prezzo */}
          <div>
            <div style={{fontSize:"10px",color:T.muted,marginBottom:"4px"}}>Prezzo (€/unit)</div>
            <input
              value={inputPrice}
              onChange={e=>setInputPrice(e.target.value)}
              placeholder="0.00"
              type="number"
              min="0"
              step="0.01"
              style={{width:"100%",background:T.surface,border:`1px solid ${T.border}`,borderRadius:"6px",
                padding:"7px 10px",color:T.text,fontSize:"12px",fontFamily:"inherit",boxSizing:"border-box"}}
            />
          </div>

          {/* Nota */}
          <div>
            <div style={{fontSize:"10px",color:T.muted,marginBottom:"4px"}}>Nota (opzionale)</div>
            <input
              value={inputNote}
              onChange={e=>setInputNote(e.target.value)}
              placeholder="Es: prezzo concordato fornitore…"
              style={{width:"100%",background:T.surface,border:`1px solid ${T.border}`,borderRadius:"6px",
                padding:"7px 10px",color:T.text,fontSize:"12px",fontFamily:"inherit",boxSizing:"border-box"}}
            />
          </div>

          {/* Bottone */}
          <div>
            <div style={{fontSize:"10px",color:"transparent",marginBottom:"4px"}}>·</div>
            <button
              onClick={addException}
              disabled={!selectedProd || !inputPrice}
              style={{padding:"7px 16px",background:selectedProd&&inputPrice?T.gold:"#333",
                color:selectedProd&&inputPrice?"#111":T.dim,border:"none",borderRadius:"6px",
                cursor:selectedProd&&inputPrice?"pointer":"default",fontFamily:"inherit",fontSize:"12px",
                fontWeight:"bold",whiteSpace:"nowrap"}}>
              {priceExceptions.some((e:any)=>e.branch===branch&&e.productId===selectedProd?.id) ? "Aggiorna" : "Aggiungi"}
            </button>
          </div>

        </div>
      </div>

      {/* Tabella eccezioni */}
      {branchExc.length === 0 ? (
        <div style={{color:T.dim,fontSize:"13px",textAlign:"center",padding:"32px",background:T.card,borderRadius:"10px",border:`1px solid ${T.border}`}}>
          Nessuna eccezione prezzo per {branch}.
        </div>
      ) : (
        <div style={{background:T.card,border:`1px solid ${T.border}`,borderRadius:"10px",overflow:"hidden"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:"12px"}}>
            <thead>
              <tr style={{background:T.surface}}>
                {["IFB No", isCAN?"N COMIT":"N HK", "Descrizione", "Prezzo (€/unit)", "Nota", "·"].map(h=>(
                  <th key={h} style={{padding:"8px 12px",textAlign:"left",fontSize:"10px",letterSpacing:"1px",
                    color:T.muted,textTransform:"uppercase",fontWeight:"normal",borderBottom:`1px solid ${T.border}`}}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {branchExc.map((exc:any, i:number)=>(
                <tr key={i} style={{borderTop:`1px solid ${T.border}`}}>
                  <td style={{padding:"8px 12px",color:T.gold}}>{exc.productId}</td>
                  <td style={{padding:"8px 12px",color:T.muted}}>{exc.nHK||"-"}</td>
                  <td style={{padding:"8px 12px"}}>{exc.description||exc.code||"-"}</td>
                  <td style={{padding:"8px 12px",color:T.green,fontWeight:"bold"}}>€ {Number(exc.price).toFixed(4)}</td>
                  <td style={{padding:"8px 12px",color:T.muted,fontStyle:"italic"}}>{exc.note||"-"}</td>
                  <td style={{padding:"8px 12px",textAlign:"center"}}>
                    <button onClick={()=>removeException(i)}
                      style={{background:"transparent",border:`1px solid ${T.red||"#c55"}`,color:T.red||"#c55",
                        borderRadius:"4px",padding:"2px 8px",cursor:"pointer",fontSize:"11px",fontFamily:"inherit"}}>
                      Rimuovi
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

// ─── MAIL GEN ─────────────────────────────────────────────────────────────────
// Only shows items with |delta| > 3% (point 7)
// ─── SC ATTUALI ───────────────────────────────────────────────────────────────
function ScAttualiPage({scAttuali, setScAttuali, branch, showToast}) {
  const [step, setStep] = useState("main");
  const [fileName, setFileName] = useState("");
  const [preview, setPreview] = useState<any[]>([]);
  const [search, setSearch] = useState("");

  function parseFile(file) {
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (e:any) => {
      try {
        const wb = XLSX.read(e.target.result, {type:"binary"});
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data: any[][] = XLSX.utils.sheet_to_json(ws, {header:1, defval:""});
        if (data.length < 2) { showToast("File vuoto", T.red); return; }
        let hi = 0;
        for (let i=0; i<Math.min(6, data.length); i++) {
          const rn = data[i].map((c:any)=>String(c||"").toLowerCase());
          if (rn.some((c:string)=>c.includes("item no")||c.includes("standard cost")||c.includes("last standard"))) { hi=i; break; }
        }
        const hdrs = data[hi].map((c:any)=>String(c||"").trim());
        const rows = data.slice(hi+1).filter((r:any[])=>r.some((c:any)=>c!==""));
        const fi = (aliases:string[]) => hdrs.findIndex(h=>aliases.some(a=>h.toLowerCase().replace(/[\s_%()/]/g,"").includes(a.replace(/[\s_%()/]/g,""))));
        const isHK = hdrs.some(h=>h.toLowerCase().includes("last standard cost"));
        const iCode     = fi(["itemno","item no"]);
        const iDesc     = fi(["description","descrizione"]);
        const iFifo     = isHK ? fi(["unitcost"]) : fi(["unitcost(fifo","unit cost"]);
        const iLastSC   = isHK ? fi(["laststandard","last standard"]) : fi(["standardcost","standard cost"]);
        const iSales3m  = fi(["saleslast","sales last","vendite"]);
        const iLastDate = fi(["lastpurchase","last purchase"]);
        const iStockQty = fi(["stockqty","stock quantity","stock"]);
        const iScGC     = !isHK ? fi(["scgrancanaria","gran canaria"]) : -1;
        const iScLan    = !isHK ? fi(["sclanzarote","lanzarote"]) : -1;

        const num = (v:any) => typeof v==="number" ? v : parseFloat(String(v||"").replace(/[€$,\s]/g,""))||0;
        const str = (v:any) => String(v||"").trim();

        const parsed = rows.map((row:any[])=>{
          const code = str(iCode>=0?row[iCode]:"");
          if (!code) return null;
          return {
            code,
            description: str(iDesc>=0?row[iDesc]:""),
            lastSC:   num(iLastSC>=0?row[iLastSC]:0),
            fifoUnit: num(iFifo>=0?row[iFifo]:0),
            salesLast3m: num(iSales3m>=0?row[iSales3m]:0),
            lastPurchaseDate: str(iLastDate>=0?row[iLastDate]:""),
            stockQty: num(iStockQty>=0?row[iStockQty]:0),
            scGC:  iScGC>=0  ? num(row[iScGC])  : 0,
            scLan: iScLan>=0 ? num(row[iScLan]) : 0,
          };
        }).filter(Boolean);

        if (!parsed.length) { showToast("Nessuna riga valida trovata", T.red); return; }
        setPreview(parsed);
        setStep("preview");
      } catch(err:any) { showToast("Errore lettura file: "+err.message, T.red); }
    };
    reader.readAsBinaryString(file);
  }

  function executeImport() {
    setScAttuali(preview);
    showToast(`SC Attuali: ${preview.length} articoli importati ✓`, T.gold);
    setStep("main");
    setPreview([]);
  }

  const displayed = (step==="main"?scAttuali:preview).filter((r:any)=>
    !search || r.code.toLowerCase().includes(search.toLowerCase()) || r.description.toLowerCase().includes(search.toLowerCase())
  );

  const isHKReport = branch !== "CAN";

  return (
    <div>
      <PageHeader title={`📊 SC Attuali · ${branch}`}
        sub={scAttuali.length>0 ? `${scAttuali.length} articoli in memoria` : "Nessun report caricato"}/>

      {step==="preview" ? (
        <Section title={`Preview — ${fileName} · ${preview.length} articoli`}>
          <div style={{display:"flex",gap:"10px",marginBottom:"14px"}}>
            <ActionBtn label="← Annulla" onClick={()=>{setStep("main");setPreview([]);}}/>
            <ActionBtn label={`✓ Importa ${preview.length} articoli`} onClick={executeImport} primary/>
          </div>
        </Section>
      ) : (
        <Section title="Carica report SC da BC / Navision">
          <div style={{fontSize:"12px",color:T.muted,marginBottom:"10px",lineHeight:"1.7"}}>
            {branch==="CAN"
              ? "Formato Navision: Item No · Description · STANDARD COST · SC GRANCANARIA · SC LANZAROTE · Last Purchase Date · Stock Quantity"
              : "Formato BC: Item No · Description · unitcost (FIFO) · Last Standard Cost · Sales Last 3 Months · Last Purchase Date · Stock Quantity"}
          </div>
          <label style={{display:"inline-block",padding:"8px 16px",background:`${T.gold}22`,border:`1px solid ${T.gold}44`,borderRadius:"6px",cursor:"pointer",fontSize:"12px",color:T.gold}}>
            ⇪ Carica Report SC ({branch})
            <input type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={e=>e.target.files?.[0]&&parseFile(e.target.files[0])}/>
          </label>
          {scAttuali.length>0&&<span style={{marginLeft:"14px",fontSize:"11px",color:T.muted}}>Ultima importazione: {scAttuali.length} articoli in memoria</span>}
        </Section>
      )}

      {displayed.length>0&&(
        <Section title={step==="preview"?"Anteprima dati":"Dati SC Attuali in memoria"}>
          <input placeholder="Cerca articolo..." value={search} onChange={e=>setSearch(e.target.value)}
            style={{...inputStyle(),marginBottom:"12px",width:"280px"}}/>
          <div style={{overflowX:"auto"}}>
            <table style={{borderCollapse:"collapse",width:"max-content",minWidth:"100%"}}>
              <thead><tr>
                {(isHKReport
                  ? ["Codice (N HK)","Descrizione","SC Attuale €","FIFO unit €","Vendite 3m","Last Purchase","Stock Qty"]
                  : ["Codice (N COMIT)","Descrizione","SC Standard €","FIFO unit €","SC Gran Can €","SC Lanzarote €","Last Purchase","Stock Qty"]
                ).map(h=><th key={h} style={{padding:"6px 10px",fontSize:"10px",color:T.gold,borderBottom:`1px solid ${T.border}`,textAlign:"right",whiteSpace:"nowrap"}}>{h}</th>)}
              </tr></thead>
              <tbody>
                {displayed.slice(0,300).map((r:any,i:number)=>(
                  <tr key={i} style={{borderBottom:`1px solid ${T.border}22`,background:i%2?"transparent":`${T.surface}33`}}>
                    <td style={{padding:"4px 10px",fontSize:"11px",color:T.text,fontFamily:"monospace"}}>{r.code}</td>
                    <td style={{padding:"4px 10px",fontSize:"11px",color:T.muted,maxWidth:"260px",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.description}</td>
                    <td style={{padding:"4px 10px",fontSize:"11px",color:T.gold,textAlign:"right",fontWeight:"bold"}}>{r.lastSC>0?`€ ${r.lastSC.toFixed(2)}`:"—"}</td>
                    <td style={{padding:"4px 10px",fontSize:"11px",color:T.muted,textAlign:"right"}}>{r.fifoUnit>0?r.fifoUnit.toFixed(4):"—"}</td>
                    {!isHKReport&&<td style={{padding:"4px 10px",fontSize:"11px",color:T.muted,textAlign:"right"}}>{r.scGC>0?`€ ${r.scGC.toFixed(2)}`:"—"}</td>}
                    {!isHKReport&&<td style={{padding:"4px 10px",fontSize:"11px",color:T.muted,textAlign:"right"}}>{r.scLan>0?`€ ${r.scLan.toFixed(2)}`:"—"}</td>}
                    {isHKReport&&<td style={{padding:"4px 10px",fontSize:"11px",color:T.muted,textAlign:"right"}}>{r.salesLast3m?r.salesLast3m.toFixed(0):"—"}</td>}
                    <td style={{padding:"4px 10px",fontSize:"11px",color:T.muted,textAlign:"right",whiteSpace:"nowrap"}}>{r.lastPurchaseDate||"—"}</td>
                    <td style={{padding:"4px 10px",fontSize:"11px",color:T.muted,textAlign:"right"}}>{r.stockQty!=null?r.stockQty:"—"}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {displayed.length>300&&<div style={{padding:"8px",fontSize:"11px",color:T.muted,textAlign:"center"}}>Mostrati 300 su {displayed.length}</div>}
        </Section>
      )}
    </div>
  );
}

// ─── CHECK MENSILE ────────────────────────────────────────────────────────────
function CheckMensile({costRows, branch, salesRows, xrefs, scAttuali, products}) {
  // Mesi disponibili dalle fatture (posting date)
  const availableMonths = useMemo(()=>{
    const s = new Set<string>();
    (salesRows||[]).forEach((r:any)=>{
      const d = r.postingDate ? String(r.postingDate).slice(0,7) : "";
      if(d&&d.length===7) s.add(d);
    });
    return [...s].sort().reverse();
  },[salesRows]);

  const [selectedMonth, setSelectedMonth] = useState<string>("");
  const [threshold, setThreshold] = useState(3);
  const [showOK, setShowOK] = useState(false);

  useEffect(()=>{
    if(!selectedMonth && availableMonths.length>0) setSelectedMonth(availableMonths[0]);
  },[availableMonths]);

  // mappe di lookup
  const xrefByIFB = useMemo(()=>{
    const m: Record<string,string>={};
    (xrefs||[]).forEach((x:any)=>{ if(x.ifbNo&&x.nHK) m[x.ifbNo]=x.nHK; });
    return m;
  },[xrefs]);

  const costMap = useMemo(()=>{
    const m: Record<string,any>={};
    (costRows||[]).forEach((r:any)=>{ m[r.id||r.code]=r; });
    return m;
  },[costRows]);

  const scMap = useMemo(()=>{
    const m: Record<string,any>={};
    (scAttuali||[]).forEach((r:any)=>{ if(r.code) m[r.code]=r; });
    return m;
  },[scAttuali]);

  // fatture del mese selezionato
  const monthRows = useMemo(()=>{
    if(!selectedMonth) return [];
    return (salesRows||[]).filter((r:any)=>{
      const d = r.postingDate ? String(r.postingDate).slice(0,7) : "";
      return d===selectedMonth;
    });
  },[salesRows, selectedMonth]);

  // righe analisi: un articolo per IFB No
  const analysisRows = useMemo(()=>{
    if(!selectedMonth||!monthRows.length) return [];
    const seen = new Set<string>();
    const rows: any[] = [];
    for (const inv of monthRows) {
      const ifbNo = inv.itemCode;
      if(!ifbNo||seen.has(ifbNo)) continue;
      seen.add(ifbNo);
      const nFiliale = xrefByIFB[ifbNo] || ifbNo;
      const scEntry  = scMap[nFiliale];
      const cr       = costMap[ifbNo];
      const oldSC    = scEntry?.lastSC || 0;
      // Step2: per CAN usa step2GC; per HK usa step2Eur
      const newSC    = branch==="CAN"
        ? (cr?.cost?.step2GC || cr?.cost?.step2Eur || 0)
        : (cr?.cost?.step2Eur || 0);
      const deltaAbs = oldSC>0 ? newSC-oldSC : 0;
      const deltaPct = oldSC>0 ? deltaAbs/oldSC*100 : 0;

      let azione:string, note="";
      if (!oldSC)                       { azione="NUOVO ARTICOLO"; }
      else if (Math.abs(deltaPct)>threshold) { azione="DA AGGIORNARE"; note=deltaPct>0?"aumentato listino":"calato listino"; }
      else                              { azione="OK"; }

      rows.push({
        codice:     nFiliale,
        ifbNo,
        description: cr?.description || inv.description || "",
        isNuovo:    !oldSC,
        oldSC,
        newSC,
        deltaPct,
        scFinale:   newSC,
        lastDate:   scEntry?.lastPurchaseDate || "",
        stockQty:   scEntry?.stockQty ?? "",
        azione,
        note,
        noCalc:     !newSC && !cr?.cost,
      });
    }
    return rows.sort((a,b)=>{
      const order = {DA_AGGIORNARE:0,"NUOVO ARTICOLO":1,OK:2};
      return (order[a.azione.replace(/ /g,"_")]??9)-(order[b.azione.replace(/ /g,"_")]??9);
    });
  },[monthRows, xrefByIFB, scMap, costMap, branch, threshold, selectedMonth]);

  const nuovi = analysisRows.filter(r=>r.azione==="NUOVO ARTICOLO");
  const daAgg = analysisRows.filter(r=>r.azione==="DA AGGIORNARE");
  const okRows= analysisRows.filter(r=>r.azione==="OK");
  const toAct = [...daAgg,...nuovi];
  const displayed = showOK ? analysisRows : toAct;

  function exportExcel() {
    const branchCode = branch==="CAN"?"COMIT":"HK";
    const monthFmt = selectedMonth.replace("-","_").slice(0,7);
    const mLabel = selectedMonth ? new Date(selectedMonth+"-01").toLocaleDateString("it-IT",{month:"short",year:"numeric"}) : "";
    const data = toAct.map((r:any)=>({
      "Codice":    r.codice,
      "IFB":       r.ifbNo,
      "Descrizione": r.description,
      "NUOVO":     r.isNuovo?"SI":"",
      "Old_SC":    r.oldSC>0 ? Number(r.oldSC.toFixed(2)) : "",
      "New_SC":    r.newSC>0 ? Number(r.newSC.toFixed(2)) : "",
      "Δ %":       r.oldSC>0 ? (r.deltaPct>0?"+":"")+r.deltaPct.toFixed(2)+"%" : "",
      "SC_FINALE": r.scFinale>0 ? Number(r.scFinale.toFixed(2)) : "",
      "Last_Date": r.lastDate,
      "Quantity":  r.stockQty,
      "AZIONE":    r.azione,
      "NOTE":      r.note,
    }));
    exportXLSX(data, "SC_Analisi", `STDC_Analisi_${branchCode}_${monthFmt}.xlsx`);
  }

  const AC:{[k:string]:string} = {
    "DA AGGIORNARE": T.orange,
    "NUOVO ARTICOLO": T.blue,
    "OK": T.green,
  };

  return (
    <div>
      <PageHeader title={`📅 Check Mensile · ${branch}`} sub="Confronto SC calcolato vs SC Attuali — soglia: > +3% o < -3%"/>

      {/* Selezione mese + threshold */}
      <Section title="Filtri">
        <div style={{display:"flex",gap:"16px",alignItems:"flex-end",flexWrap:"wrap"}}>
          <div>
            <label style={{fontSize:"10px",color:T.muted,display:"block",marginBottom:"4px",letterSpacing:"1px",textTransform:"uppercase"}}>Mese fatture</label>
            {availableMonths.length===0
              ? <div style={{fontSize:"12px",color:T.orange,padding:"7px 12px",border:`1px solid ${T.orange}44`,borderRadius:"6px"}}>⚠ Nessuna fattura caricata</div>
              : <select value={selectedMonth} onChange={e=>setSelectedMonth(e.target.value)} style={{...inputStyle(),minWidth:"160px",cursor:"pointer"}}>
                  {availableMonths.map(m=><option key={m} value={m}>{m}</option>)}
                </select>
            }
          </div>
          <div>
            <label style={{fontSize:"10px",color:T.muted,display:"block",marginBottom:"4px",letterSpacing:"1px",textTransform:"uppercase"}}>Soglia Δ%</label>
            <input type="number" value={threshold} onChange={e=>setThreshold(Number(e.target.value)||3)}
              min={0} max={50} step={0.5} style={{...inputStyle(),width:"80px"}}/>
          </div>
          {scAttuali.length===0&&(
            <div style={{fontSize:"12px",color:T.orange,padding:"7px 12px",border:`1px solid ${T.orange}44`,borderRadius:"6px"}}>
              ⚠ SC Attuali non caricati — vai alla pagina <strong>SC Attuali</strong>
            </div>
          )}
          {monthRows.length>0&&<div style={{fontSize:"11px",color:T.muted,paddingBottom:"2px"}}>
            {monthRows.length} righe fattura · {analysisRows.length} articoli univoci
          </div>}
        </div>
      </Section>

      {/* Riepilogo KPI */}
      {analysisRows.length>0&&(
        <Section title="Riepilogo">
          <div style={{display:"flex",gap:"20px",flexWrap:"wrap",marginBottom:"14px"}}>
            {[
              {label:"DA AGGIORNARE",n:daAgg.length,c:T.orange,icon:"⬆️"},
              {label:"NUOVI",n:nuovi.length,c:T.blue,icon:"🆕"},
              {label:"OK (nessuna azione)",n:okRows.length,c:T.green,icon:"✅"},
            ].map(({label,n,c,icon})=>(
              <div key={label} style={{background:`${c}11`,border:`1px solid ${c}44`,borderRadius:"8px",padding:"10px 18px",minWidth:"130px"}}>
                <div style={{fontSize:"9px",color:c,letterSpacing:"1px",textTransform:"uppercase",marginBottom:"2px"}}>{icon} {label}</div>
                <div style={{fontSize:"22px",fontWeight:"bold",color:c}}>{n}</div>
              </div>
            ))}
          </div>
          <div style={{display:"flex",gap:"10px",flexWrap:"wrap"}}>
            <ActionBtn label="📥 Esporta Excel" onClick={exportExcel} primary disabled={toAct.length===0}/>
            <ActionBtn label={showOK?"Nascondi OK ✓":"Mostra anche OK ✓"} onClick={()=>setShowOK(s=>!s)}/>
          </div>
        </Section>
      )}

      {/* Tabella */}
      {displayed.length>0 ? (
        <Section title={`Articoli · ${displayed.length} mostrati`}>
          <div style={{overflowX:"auto"}}>
            <table style={{borderCollapse:"collapse",width:"max-content",minWidth:"100%"}}>
              <thead><tr>
                {["Codice","IFB","Descrizione","NUOVO","Old SC €","New SC €","Δ %","SC FINALE €","Last Date","Stock","AZIONE","NOTE"].map(h=>(
                  <th key={h} style={{padding:"6px 10px",fontSize:"10px",color:T.gold,borderBottom:`1px solid ${T.border}`,textAlign:"right",whiteSpace:"nowrap"}}>{h}</th>
                ))}
              </tr></thead>
              <tbody>
                {displayed.map((r:any,i:number)=>{
                  const ac = AC[r.azione]||T.muted;
                  const isNuovo = r.azione==="NUOVO ARTICOLO";
                  const isDa    = r.azione==="DA AGGIORNARE";
                  return (
                    <tr key={i} style={{borderBottom:`1px solid ${T.border}22`,background:`${ac}09`}}>
                      <td style={{padding:"5px 10px",fontSize:"11px",color:T.text,fontFamily:"monospace",whiteSpace:"nowrap"}}>{r.codice}</td>
                      <td style={{padding:"5px 10px",fontSize:"11px",color:T.muted,fontFamily:"monospace"}}>{r.ifbNo}</td>
                      <td style={{padding:"5px 10px",fontSize:"11px",color:T.text,maxWidth:"230px",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.description}</td>
                      <td style={{padding:"5px 10px",fontSize:"11px",color:T.blue,textAlign:"center",fontWeight:"bold"}}>{isNuovo?"SI":""}</td>
                      <td style={{padding:"5px 10px",fontSize:"11px",color:T.muted,textAlign:"right"}}>{r.oldSC>0?`€ ${r.oldSC.toFixed(2)}`:"—"}</td>
                      <td style={{padding:"5px 10px",fontSize:"11px",color:T.gold,textAlign:"right",fontWeight:"bold"}}>{r.newSC>0?`€ ${r.newSC.toFixed(2)}`:r.noCalc?"NC":"—"}</td>
                      <td style={{padding:"5px 10px",fontSize:"11px",textAlign:"right",fontWeight:"bold",
                        color:r.deltaPct>0?T.orange:r.deltaPct<0?"#e05a5a":T.muted}}>
                        {r.oldSC>0?(r.deltaPct>0?"+":"")+r.deltaPct.toFixed(2)+"%":"—"}
                      </td>
                      <td style={{padding:"5px 10px",fontSize:"11px",color:T.green,textAlign:"right",fontWeight:"bold"}}>{r.scFinale>0?`€ ${r.scFinale.toFixed(2)}`:"—"}</td>
                      <td style={{padding:"5px 10px",fontSize:"11px",color:T.muted,textAlign:"right",whiteSpace:"nowrap"}}>{r.lastDate||"—"}</td>
                      <td style={{padding:"5px 10px",fontSize:"11px",color:T.muted,textAlign:"right"}}>{r.stockQty!==""?r.stockQty:"—"}</td>
                      <td style={{padding:"5px 10px",textAlign:"center"}}>
                        <span style={{padding:"2px 7px",borderRadius:"4px",fontSize:"10px",fontWeight:"bold",whiteSpace:"nowrap",
                          background:`${ac}22`,color:ac}}>{r.azione}</span>
                      </td>
                      <td style={{padding:"5px 10px",fontSize:"10px",color:T.muted,fontStyle:"italic",whiteSpace:"nowrap"}}>{r.note}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </Section>
      ) : selectedMonth&&!analysisRows.length ? (
        <div style={{padding:"32px",textAlign:"center",color:T.muted,fontSize:"13px"}}>
          {scAttuali.length===0
            ? "⚠️ Carica prima il report SC Attuali (pagina SC Attuali)."
            : `Nessun articolo nelle fatture di ${selectedMonth}.`}
        </div>
      ) : null}
    </div>
  );
}

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
function Storico({snapshots,setSnapshots,costHistory,setCostHistory,branch,showToast,macHkCostRows:_unused}) {
  const[sel,setSel]=useState<any>(null);
  const[sortDir,setSortDir]=useState("asc");
  const[deltaFilter,setDeltaFilter]=useState("all");
  const[showModified,setShowModified]=useState(false);
  const[showNew,setShowNew]=useState(false);
  const[selCostSnap,setSelCostSnap]=useState<any>(null);
  const[macProds,setMacProds]=useState<any[]>([]);

  // Per HK: carica prodotti MAC (servono HOFF flag e macToHkConv per derivare costi MAC affianco)
  useEffect(()=>{
    if(branch==="HK") IDB.get("ifb_products_MAC",[]).then((d:any[])=>setMacProds(d));
    else setMacProds([]);
  },[branch]);

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
          <THead cols={branch==="HK"&&macProds.length>0
            ? [branchN(branch),"IFB No","Descrizione","New SC HKD","New SC MAC (MOP)","HOFF","×UOM","Note"]
            : [branchN(branch),"IFB No","Descrizione","Costo HKD","Note"]} sticky />
          <tbody>{(selCostSnap.rows||[]).map((r:any,i:number)=>{
            const macProd = branch==="HK"&&macProds.length>0
              ? (macProds.find((p:any)=>p.id===r.id||p.code===r.code||p.nHK===r.nHK))
              : null;
            const macCost = macProd && r.cost!=null
              ? calcMAC({ hkCost:{step2Hkd:r.cost,step2Eur:0,priceEur:0,unitsPerPlt:0}, isHoff:macProd.isHoff??false, macToHkConv:macProd.macToHkConv>1?macProd.macToHkConv:1 })
              : null;
            return(
            <tr key={r.id||i} style={{borderBottom:`1px solid ${T.border}`,background:i%2===0?T.bg:T.surface}}>
              <TD mono><span style={{color:T.muted}}>{r.nHK||"—"}</span></TD>
              <TD mono><span style={{color:T.gold}}>{r.code}</span></TD>
              <TD>{r.description}</TD>
              <TD mono><span style={{color:T.gold,fontWeight:"bold"}}>{r.cost!=null?roundN(r.cost).toFixed(2):"—"}</span></TD>
              {branch==="HK"&&macProds.length>0&&<>
                <TD mono><span style={{color:macCost?T.green:T.dim,fontWeight:"bold"}}>{macCost?macCost.macNewSC.toFixed(2):"—"}</span></TD>
                <TD><span style={{color:macProd?(macProd.isHoff?T.orange:T.blue):T.dim,fontSize:"10px"}}>{macProd?(macProd.isHoff?"HOFF +3%":"non-HOFF +10%"):"—"}</span></TD>
                <TD mono><span style={{color:T.muted,fontSize:"10px"}}>{macProd?.macToHkConv>1?`×${macProd.macToHkConv}`:"—"}</span></TD>
              </>}
              <TD><span style={{color:T.dim,fontSize:"11px"}}>{r.skipReason||""}</span></TD>
            </tr>
            );
          })}</tbody>
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

  // Campi per branch: solo quelli rilevanti
  const FIELDS_HK  = ["nHK","code","description","category","uom","qtyPerBox","boxPerPallet","kgPerBox","kgxplt","temperature","active","vendorName","vendorName2"];
  const FIELDS_CAN = ["nHK","code","description","category","uom","qtyPerBox","boxPerPallet","kgPerBox","kgxplt","temperature","aiem","active","vendorName","vendorName2"];
  const FIELDS_MAC = ["nHK","code","description","isHoff","uom","hkUom","standardCostHkd","temperature","kgPerBox","qtyPerBox","active","vendorName"];
  const FIELDS = branch==="CAN" ? FIELDS_CAN : branch==="MAC" ? FIELDS_MAC : FIELDS_HK;

  const FLABELS: any = {
    nHK: branch==="MAC" ? "MACAO No (No_)" : branch==="CAN" ? "N COMIT (No_)" : "N HK (No_)",
    code: "IFB Item / BV No *",
    description: "Descrizione *",
    category: "Section",
    uom: branch==="MAC" ? "UOM vendita MACAO" : "UOM",
    qtyPerBox: "Qty/Cartone",
    boxPerPallet: "Cartoni/Pallet",
    kgPerBox: "Kg per Cartone",
    kgxplt: "Kg x PLT",
    temperature: "Product Type",
    aiem: "★ AIEM % (col. W anagrafica CAN)",
    isHoff: "HOFF Flag (1 = House of Fine Foods)",
    hkUom: "HK/BV UOM (per conversione automatica)",
    standardCostHkd: "★ Standard Cost HK (HKD) — base calcolo MAC",
    temperature: "Product Type (DRY/FRESH/FROZEN)",
    kgPerBox: "Kg per Cartone (per costi logistica)",
    qtyPerBox: "Qty per Cartone (per conversione UOM)",
    active: "Bloccato",
    vendorName: "Vendor Name",
    vendorName2: "Vendor Name 2"
  };

  const LOCAL_ALIASES: any = {
    nHK:         ["no","no_","macaono","macao no","macao_no","macaomastercode","macao mastercode","macaoitemno"],
    code:        ["ifbitem","ifb item","ifb no","ifb n","bvno","bv no","bvmastercode","bv mastercode"],
    description: ["description"],
    category:    ["sectiondescription","section description","section"],
    uom:         ["salesunitofmeasure","sales unit of measure","macaosalesunitofmeasure","macao salesunitofmeasure","macaouom","macao uom"],
    qtyPerBox:   ["quantityxpackaging","quantity x packaging"],
    boxPerPallet:["packagingxpallet","packaging x pallet"],
    kgPerBox:    ["netweight","net weight"],
    kgxplt:      ["kgxplt","kg x pallet","kg per pallet","kgperpallet","kgplt"],
    temperature: ["producttype","product type","product type rettificato"],
    aiem:        ["aiem","igic","alim","aiem%","aiem_perc","aiem_canarie","aiemperc"],
    isHoff:      ["ishoff","hoff","hofflag","hoff flag","hoff_flag","is hoff"],
    hkUom:            ["bvsalesunitofmeasure","bv salesunitofmeasure","bvuom","hk uom","hkuom"],
    standardCostHkd:  ["standardcost","standard cost","costo standard","costostandard","sc hkd","schkd","fob","fobprice","fob price"],
    temperature:      ["producttype","product type","producttype rettificato","tipoprodotto"],
    kgPerBox:         ["netweight","net weight","kgperbox","kg per box","kg/box","pesokg","peso netto"],
    qtyPerBox:        ["quantityxpackaging","quantity x packaging","qtxbox","qty per box","qtyperbox","pzperbox"],
    active:           ["blocked"],
    vendorName:  ["vendorname","vendor name"],
    vendorName2: ["vendorname2","vendor name 2"],
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
      aiem: parseFloat(r.aiem) || 0,
      isHoff: ["true","1","yes","hoff","si","sì","vero","x"].includes(String(r.isHoff||"").toLowerCase()),
      hkUom: r.hkUom ? String(r.hkUom).trim().toUpperCase() : "",
      standardCostHkd: parseFloat(String(r.standardCostHkd||"").replace(",",".")) || 0,
      temperature: mapBCVal("temperature", r.temperature) || "DRY",
      kgPerBox: parseFloat(String(r.kgPerBox||"").replace(",",".")) || 0,
      qtyPerBox: parseFloat(String(r.qtyPerBox||"").replace(",",".")) || 0,
      active: !["true", "1", "yes"].includes(String(r.active || "").toLowerCase()),
      vendorName: r.vendorName || "",
      vendorName2: r.vendorName2 || "",
    }));
  
    setProducts(newProds);
    IDB.set(`ifb_products_${branch}`, newProds);
    IDB.set(`ifb_anag_data_${now}`, newProds);
    const log = { id:now, type:"anagrafica", date:new Date(now).toISOString(), count:newProds.length, branch };
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
              IDB.set(`ifb_products_${branch}`, []);
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
            {FIELDS.map(f => {
              const isRequired = f==="code"||f==="description";
              const isAiem = f==="aiem";
              const isMacKey = ["hkUom","isHoff"].includes(f);
              const labelColor = isRequired?T.gold:isAiem?T.orange:isMacKey?T.purple:T.muted;
              return(
              <div key={f} style={isAiem?{border:`1px solid ${T.orange}33`,borderRadius:"6px",padding:"4px 6px",background:`${T.orange}08`}:{}}>
                <label style={{ fontSize: "10px", color: labelColor }}>{FLABELS[f]}</label>
                <select
                  value={map[f] || ""}
                  onChange={e => setMap((m: any) => ({ ...m, [f]: e.target.value }))}
                  style={{ ...inputStyle(), fontSize: "11px", padding: "4px 6px", borderColor: map[f]?T.gold:(isAiem&&!map[f])?T.orange:T.border }}
                >
                  <option value="">—</option>
                  {headers.map(h => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>
              );
            })}
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
                  <th style={{ textAlign: "left", padding: "4px" }}>{branchN(branch)}</th>
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
        <SyncScrollTable>
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
                {branch==="CAN" && <th style={{ padding: "7px 12px", background: T.card, color: T.orange, textAlign: "right", borderBottom: `1px solid ${T.border}`, fontSize: "11px" }}>AIEM%</th>}
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
                    {branch==="CAN" && <td style={{ padding: "7px 12px", fontSize: "12px", fontFamily: "monospace", textAlign: "right", color: p.aiem>0?T.orange:T.dim }}>{p.aiem>0?`${p.aiem}%`:"—"}</td>}
                    <td style={{ padding: "7px 12px", fontSize: "12px" }}>
                      <Chip label={p.active ? "Sì" : "No"} color={p.active ? T.green : T.red} />
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </SyncScrollTable>
      </Section>
    </div>
  );
}

// ─── BEVERAGE INFO (CAN — AIEM alcolici) ──────────────────────────────────────
function BeverageInfoPage({bevInfo, setBevInfo, products, showToast}: any) {
  const [step, setStep] = useState<"main"|"map"|"preview">("main");
  const [headers, setHeaders] = useState<string[]>([]);
  const [rawRows, setRawRows] = useState<any[]>([]);
  const [map, setMap] = useState<any>({});
  const [preview, setPreview] = useState<any[]>([]);
  const [fileName, setFileName] = useState("");
  const [search, setSearch] = useState("");

  const FIELDS = ["ifbNo","ltPerUnit","gradoAlcolico","eurPerLt","totaleBottiglia"];
  const FLABELS: any = {
    ifbNo: "IFB No * (codice articolo)",
    ltPerUnit: "LT (litri per unità)",
    gradoAlcolico: "Grado Alcolico (°)",
    eurPerLt: "€/LT (tariffa AIEM per litro)",
    totaleBottiglia: "Totale Bottiglia € (se già calcolato)",
  };
  const ALIASES: any = {
    ifbNo:          ["ifb n","ifb no","ifbno","bv no","codice","code","item no","ifb item","ifbitem"],
    ltPerUnit:      ["lt","litri","liters","volume","lt per unit","lt/unit","litri per unità"],
    gradoAlcolico:  ["grado alcolico","grado","gradi","abv","alcol","alcohol","alc %","degree","gradalcolico"],
    eurPerLt:       ["eur/lt","€/lt","euro/lt","tariffa","rate","eur lt","eurlt","tariffalit"],
    totaleBottiglia:["totale bottiglia","totale","total","totbottiglia","totalebottiglia","tot bott","totbott","totale bottiglia eur"],
  };

  function fi(aliases: string[], hdrs: string[]): string {
    const norm = (s: string) => s.toLowerCase().replace(/[€°\s_/]/g,"");
    return hdrs.find(h => aliases.some(a => norm(h)===norm(a))) ||
           hdrs.find(h => aliases.some(a => norm(h).includes(norm(a)))) || "";
  }

  function parseFile(file: File) {
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, {type:"binary"});
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data: any[][] = XLSX.utils.sheet_to_json(ws, {header:1, defval:""});
        // Cerca la riga header (quella con "IFB" o "LT")
        const hdrIdx = data.findIndex(row =>
          row.some((c:any) => String(c).toUpperCase().includes("IFB") || String(c).toUpperCase()==="LT")
        );
        if(hdrIdx < 0) { showToast("Intestazione non trovata", T.red); return; }
        const hdrs = data[hdrIdx].map((h:any) => String(h).trim()).filter((h:string) => h);
        const rows = data.slice(hdrIdx+1).filter(r => r.some((c:any) => c !== ""));
        setHeaders(hdrs); setRawRows(rows);
        const autoMap: any = {};
        FIELDS.forEach(f => { autoMap[f] = fi(ALIASES[f], hdrs); });
        setMap(autoMap); setStep("map");
      } catch(err:any) { showToast("Errore: "+err.message, T.red); }
    };
    reader.readAsBinaryString(file);
  }

  function buildPreview() {
    const idx: any = {};
    FIELDS.forEach(f => { idx[f] = headers.indexOf(map[f]); });
    const rows = rawRows.map(row => {
      const ifbNo = String(row[idx.ifbNo]||"").trim();
      if(!ifbNo) return null;
      const lt = parseFloat(String(row[idx.ltPerUnit]||"").replace(",",".")) || 0;
      const grado = parseFloat(String(row[idx.gradoAlcolico]||"").replace(",",".")) || 0;
      const eurLt = parseFloat(String(row[idx.eurPerLt]||"").replace(",",".")) || 0;
      const totRaw = idx.totaleBottiglia>=0 ? parseFloat(String(row[idx.totaleBottiglia]||"").replace(",",".")) : 0;
      const totCalc = lt > 0 && eurLt > 0 ? roundN(lt * eurLt, 4) : 0;
      const totaleBottiglia = totRaw > 0 ? totRaw : totCalc;
      if(totaleBottiglia <= 0 && lt <= 0) return null;
      const prod = products.find((p:any) => p.code === ifbNo);
      return { ifbNo, ltPerUnit:lt, gradoAlcolico:grado, eurPerLt:eurLt, totaleBottiglia, _found:!!prod, _desc:prod?.description||"—" };
    }).filter(Boolean);
    setPreview(rows); setStep("preview");
  }

  function executeImport() {
    const kept = bevInfo.filter((b:any) => !preview.find((p:any) => p.ifbNo===b.ifbNo));
    const next = [...preview.map((r:any) => ({ifbNo:r.ifbNo,ltPerUnit:r.ltPerUnit,gradoAlcolico:r.gradoAlcolico,eurPerLt:r.eurPerLt,totaleBottiglia:r.totaleBottiglia})), ...kept];
    setBevInfo(next); IDB.set("ifb_bevinfo", next);
    showToast(`Beverage Info: ${preview.length} articoli importati ✓`, T.gold);
    setStep("main"); setPreview([]); setRawRows([]); setHeaders([]);
  }

  const q = search.trim().toLowerCase();
  const displayed = q
    ? bevInfo.filter((b:any) => b.ifbNo?.toLowerCase().includes(q) || products.find((p:any)=>p.code===b.ifbNo)?.description?.toLowerCase().includes(q))
    : bevInfo;

  return (
    <div>
      <PageHeader title="🍷 Beverage Info · AIEM Alcolici (CAN)" sub="Importa dati alcolici: LT, Grado, €/LT → Totale AIEM fisso per unità"/>

      <div style={{display:"flex",gap:"10px",marginBottom:"16px",alignItems:"center",flexWrap:"wrap"}}>
        <label style={{display:"inline-block",padding:"8px 16px",background:T.gold,color:"#000",borderRadius:"6px",cursor:"pointer",fontWeight:"bold",fontSize:"12px"}}>
          📂 Carica file Beverage Info
          <input type="file" accept=".xlsx,.xls,.csv" onChange={e=>{const f=e.target.files?.[0];if(f)parseFile(f);e.target.value="";}} style={{display:"none"}}/>
        </label>
        {bevInfo.length>0&&<button onClick={()=>{if(window.confirm(`Eliminare tutti i ${bevInfo.length} dati beverage?`)){setBevInfo([]);IDB.set("ifb_bevinfo",[]);}}}
          style={{padding:"5px 12px",background:"none",border:`1px solid ${T.red}44`,borderRadius:"6px",color:T.red,cursor:"pointer",fontSize:"11px"}}>
          🗑 Svuota ({bevInfo.length})
        </button>}
      </div>

      {step==="map"&&(
        <Section title={`Mappatura — ${fileName}`}>
          <div style={{background:`${T.orange}10`,border:`1px solid ${T.orange}33`,borderRadius:"8px",padding:"10px 14px",marginBottom:"12px",fontSize:"11px",color:T.orange}}>
            ★ Se il file ha già la colonna "Totale Bottiglia", verrà usata. Altrimenti verrà calcolata come LT × €/LT.
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:"10px",marginBottom:"16px"}}>
            {FIELDS.map(f=>(
              <div key={f}>
                <label style={{fontSize:"10px",color:f==="ifbNo"?T.gold:T.muted}}>{FLABELS[f]}</label>
                <select value={map[f]||""} onChange={e=>setMap((m:any)=>({...m,[f]:e.target.value}))}
                  style={{...inputStyle(),fontSize:"11px",padding:"4px 6px",borderColor:map[f]?T.gold:T.border}}>
                  <option value="">—</option>
                  {headers.map(h=><option key={h} value={h}>{h}</option>)}
                </select>
              </div>
            ))}
          </div>
          <div style={{display:"flex",gap:"10px"}}>
            <ActionBtn label="Annulla" onClick={()=>setStep("main")}/>
            <ActionBtn label="Preview →" onClick={buildPreview} primary disabled={!map.ifbNo}/>
          </div>
        </Section>
      )}

      {step==="preview"&&(
        <Section title={`Preview · ${preview.length} articoli`}>
          <div style={{maxHeight:"200px",overflow:"auto",marginBottom:"12px",fontSize:"11px"}}>
            <table style={{width:"100%",borderCollapse:"collapse"}}>
              <thead><tr>
                {["IFB No","Descrizione","LT","Grado","€/LT","Totale €/unit"].map(h=>
                  <th key={h} style={{padding:"4px 8px",textAlign:"left",color:T.muted,borderBottom:`1px solid ${T.border}`}}>{h}</th>)}
              </tr></thead>
              <tbody>
                {preview.map((r:any,i:number)=>(
                  <tr key={i} style={{borderBottom:`1px solid ${T.border}`,background:r._found?undefined:`${T.orange}10`}}>
                    <td style={{padding:"3px 8px",color:T.gold,fontFamily:"monospace"}}>{r.ifbNo}</td>
                    <td style={{padding:"3px 8px",color:r._found?T.text:T.orange,fontSize:"11px"}}>{r._desc}</td>
                    <td style={{padding:"3px 8px",fontFamily:"monospace",textAlign:"right"}}>{r.ltPerUnit||"—"}</td>
                    <td style={{padding:"3px 8px",fontFamily:"monospace",textAlign:"right"}}>{r.gradoAlcolico||"—"}°</td>
                    <td style={{padding:"3px 8px",fontFamily:"monospace",textAlign:"right"}}>{r.eurPerLt>0?r.eurPerLt.toFixed(2):"—"}</td>
                    <td style={{padding:"3px 8px",fontFamily:"monospace",textAlign:"right",color:T.orange,fontWeight:"bold"}}>{r.totaleBottiglia>0?r.totaleBottiglia.toFixed(4):"—"}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div style={{display:"flex",gap:"10px"}}>
            <ActionBtn label="← Indietro" onClick={()=>setStep("map")}/>
            <ActionBtn label={`✓ Importa ${preview.length} articoli`} onClick={executeImport} primary/>
          </div>
        </Section>
      )}

      <SearchBar value={search} onChange={setSearch} placeholder="🔍 Cerca IFB No o descrizione…"/>

      {displayed.length>0 ? (
        <Section title={`${displayed.length} articoli con AIEM alcolico`}>
          <div style={{overflowX:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse"}}>
              <thead><tr>
                {["IFB No","Descrizione","LT/unit","Grado","€/LT","AIEM €/unit"].map(h=>
                  <th key={h} style={{padding:"7px 12px",background:T.card,color:T.muted,textAlign:"left",borderBottom:`1px solid ${T.border}`,fontSize:"11px"}}>{h}</th>)}
                <th style={{padding:"7px 12px",background:T.card,borderBottom:`1px solid ${T.border}`,fontSize:"11px"}}/>
              </tr></thead>
              <tbody>
                {displayed.map((b:any,i:number)=>{
                  const prod = products.find((p:any)=>p.code===b.ifbNo);
                  return(
                    <tr key={b.ifbNo} style={{borderBottom:`1px solid ${T.border}`,background:i%2===0?T.bg:T.surface}}>
                      <td style={{padding:"7px 12px",fontFamily:"monospace",color:T.gold}}>{b.ifbNo}</td>
                      <td style={{padding:"7px 12px",fontSize:"12px"}}>{prod?.description||<span style={{color:T.orange}}>⚠ non in anagrafica</span>}</td>
                      <td style={{padding:"7px 12px",fontFamily:"monospace",textAlign:"right"}}>{b.ltPerUnit||"—"}</td>
                      <td style={{padding:"7px 12px",fontFamily:"monospace",textAlign:"right"}}>{b.gradoAlcolico||"—"}°</td>
                      <td style={{padding:"7px 12px",fontFamily:"monospace",textAlign:"right"}}>{b.eurPerLt>0?b.eurPerLt.toFixed(2):"—"}</td>
                      <td style={{padding:"7px 12px",fontFamily:"monospace",textAlign:"right",color:T.orange,fontWeight:"bold"}}>{b.totaleBottiglia>0?b.totaleBottiglia.toFixed(4):"—"}</td>
                      <td style={{padding:"7px 12px"}}>
                        <MiniBtn label="✕" onClick={()=>{const n=bevInfo.filter((_:any,j:number)=>j!==bevInfo.indexOf(b));setBevInfo(n);IDB.set("ifb_bevinfo",n);}} color={T.red}/>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </Section>
      ) : (
        <div style={{color:T.muted,textAlign:"center",padding:"40px",fontSize:"13px"}}>
          Nessun dato beverage. Carica il file con LT, Grado Alcolico e €/LT.
        </div>
      )}
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
    IDB.set("ifb_meatprices", entries);
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
              setMeatPrices(data); IDB.set("ifb_meatprices", data);
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
          <button onClick={()=>{if(window.confirm(`Eliminare tutti i ${meatPrices.length} prezzi del listino carne?`)){setMeatPrices([]);IDB.set("ifb_meatprices",[]);}}}
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

// Scrollbar sincronizzata in cima e in fondo alla tabella
function SyncScrollTable({children}:any){
  const topRef = useRef<HTMLDivElement>(null);
  const botRef = useRef<HTMLDivElement>(null);
  const [w, setW] = useState(0);
  useEffect(()=>{
    const el = botRef.current?.querySelector("table") as HTMLElement;
    if(!el) return;
    const obs = new ResizeObserver(()=>setW(el.scrollWidth));
    obs.observe(el);
    return ()=>obs.disconnect();
  },[]);
  const syncTop = (e:any) => { if(botRef.current) botRef.current.scrollLeft = e.target.scrollLeft; };
  const syncBot = (e:any) => { if(topRef.current) topRef.current.scrollLeft = e.target.scrollLeft; };
  return(
    <div>
      <div ref={topRef} onScroll={syncTop} style={{overflowX:"auto",overflowY:"hidden",height:"12px",marginBottom:"2px"}}>
        <div style={{width:w||"100%",height:"1px"}}/>
      </div>
      <div ref={botRef} onScroll={syncBot} style={{overflowX:"auto"}}>{children}</div>
    </div>
  );
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