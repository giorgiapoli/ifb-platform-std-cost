// v2026-07-09x
import React, { useState, useMemo, useEffect, useRef, startTransition } from "react";
import * as XLSX from "xlsx";
import { supabase, IDB, CLOUD, getSession, getUserRole, listUsers, inviteUser, removeUser, signInWithOtp, signOut } from "./supabase";

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
// Persistent page-state helpers (search / filter survive navigation)
const psGet = (k: string, def: any) => { try { const v = localStorage.getItem(k); return v != null ? JSON.parse(v) : def; } catch { return def; } };
const psSet = (k: string, v: any) => { try { localStorage.setItem(k, JSON.stringify(v)); } catch {} };
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
  return val.includes("air");
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

// ─── HK CONV FACTORS (eccezioni per item) ────────────────────────────────────
const HK_CONV_DEFAULTS: {nHK:string; factor:number; description:string}[] = [
  { nHK:"GCMA-1015", factor:3,   description:"VFF08" },
  { nHK:"GCRI-1028", factor:0.5, description:"RST26" },
];

// ─── HK SPIRITS ALCOHOL TAX (>30°) ──────────────────────────────────────────
const HK_ALC_TAX_DEFAULTS: {ifbNo:string; hasAlcTax:boolean; nHK:string; ltPerUnit:number; gradoAlcolico:number; eurPerLt:number; totaleBottiglia:number}[] = [
  {ifbNo:"VZR08",     hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
  {ifbNo:"VZR03",     hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
  {ifbNo:"VZR01",     hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
  {ifbNo:"VZR02",     hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
  {ifbNo:"VZR04",     hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
  {ifbNo:"VZR05",     hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
  {ifbNo:"VZR06",     hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
  {ifbNo:"VZR07",     hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
  {ifbNo:"VZR09",     hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
  {ifbNo:"VZR10",     hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
  {ifbNo:"LFT02",     hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
  {ifbNo:"MGI01",     hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
  {ifbNo:"MGI02",     hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
  {ifbNo:"ELT01",     hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
  {ifbNo:"ELT02",     hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
  {ifbNo:"ELT03",     hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
  {ifbNo:"ELT04",     hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
  {ifbNo:"ELT05",     hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
  {ifbNo:"ELT06",     hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
  {ifbNo:"ELT07",     hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
  {ifbNo:"ELT08",     hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
  {ifbNo:"ELT09",     hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
  {ifbNo:"LEVGINTUS", hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
  {ifbNo:"LEVGINMED", hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
  {ifbNo:"Y1077",     hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
  {ifbNo:"LIQ63",     hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
  {ifbNo:"Z0430",     hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
  {ifbNo:"LIQ55",     hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
  {ifbNo:"FDG01",     hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
  {ifbNo:"DRMFIRE",   hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
  {ifbNo:"WMAT02-NV", hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
  {ifbNo:"DRM01",     hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
  {ifbNo:"MBGRAP-NV", hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
  {ifbNo:"NAR01",     hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
  {ifbNo:"NAR03",     hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
  {ifbNo:"NAR07",     hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
  {ifbNo:"NAR08",     hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
  {ifbNo:"NAR09",     hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
  {ifbNo:"NAR10",     hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
  {ifbNo:"NAR11",     hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
  {ifbNo:"PPR02",     hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
  {ifbNo:"PPR03",     hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
  {ifbNo:"PPR04",     hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
  {ifbNo:"PPR05",     hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
  {ifbNo:"PPR06",     hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
  {ifbNo:"PPR07",     hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
  {ifbNo:"FDG02",     hasAlcTax:true, nHK:"", ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0},
];

// ─── CANARIE CONV FACTORS (eccezioni per item) ───────────────────────────────
const CAN_CONV_DEFAULTS: {nComit:string; factor:number; description:string}[] = [
  { nComit:"17021", factor:2,           description:"LA DOLCISSIMA POLPA DI POMODORO BAG IN BOX 2*5 KG" },
  { nComit:"35530", factor:100,         description:"** KREADOC CANNOLO SICILIANO VUOTO 100*30 GR" },
  { nComit:"7220",  factor:0.02,        description:"CRODINO 48 X 10 CL" },
  { nComit:"7221",  factor:0.1,         description:"CAMPARI SODA 10% 10*9.8CL" },
  { nComit:"7231",  factor:0.33,        description:"PERONI N.AZZURRO BIRRA BOT. 33 CL" },
  { nComit:"7240",  factor:0.33,        description:"PERONI CHILL LEMON BIRRA CLUSTER 3x33CL" },
  { nComit:"7229",  factor:0.33,        description:"PERONI RED LABEL BIRRA 4.7 10.60PL CLUSTER 3x33CL" },
  { nComit:"7234",  factor:0.33,        description:"" },
  { nComit:"7235",  factor:0.33,        description:"ICHNUSA BIRRA VAP CLUSTER 3 x 33 CL" },
  { nComit:"15307", factor:40,          description:"** ARANCINO DI RISO AL RAGU BOLSA 40X220GR" },
  { nComit:"12234", factor:0.833333333, description:"** SOAVEGEL SUPPLI AL TELEFONO 1,2 KG" },
  { nComit:"5114",  factor:1,           description:"GRANA PADANO DOP SCAGLIE 250 GR X 2" },
  { nComit:"6741",  factor:1,           description:"" },
  { nComit:"11020", factor:8,           description:"FAGGETTO" },
  { nComit:"9650",  factor:0.2,         description:"RISO SCOTTI ARBORIO" },
  { nComit:"9651",  factor:0.2,         description:"RISO SCOTTI CARNAROLI" },
];

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

function calcCAN({ priceInput, ubicazione, product, logistic, bevData, priceMultiplier=1 }: any) {
  const { uom, qtyPerBox, boxPerPallet, kgPerBox, kgxplt, temperature, aiem: prodAiem, vendorName, vendorName2 } = product;
  const { pltPerContainer, area, hasAlcTax, alcTax, convFactor, transport, isLAN, isFUE } = logistic || {};

  const cf = Number(convFactor||1) || 1;
  const pm = Number(priceMultiplier||1) || 1;

  // unitsPerPlt (Y6 in modello): SE(J="PCS",Q*R,SE(J="BOX",R,...)) / CM (conv factor item)
  let unitsPerPlt: number;
  if (uom==="BOX")      unitsPerPlt = Number(boxPerPallet) / cf / pm;
  else if (uom==="KG")  unitsPerPlt = (Number(kgxplt)>0 ? Number(kgxplt) : 300) / cf / pm;
  else                  unitsPerPlt = (Number(qtyPerBox) * Number(boxPerPallet)) / cf / pm; // PCS

  // divisoreCollo (AC in modello): per MTS picking
  const divisoreCollo = uom==="BOX" ? 1 : uom==="KG" ? Number(kgPerBox||qtyPerBox) : Number(qtyPerBox);

  const plt_n = Math.max(Number(pltPerContainer)||1, 1);
  const totalUnits = unitsPerPlt * plt_n;
  if (!unitsPerPlt || !totalUnits) return null;

  const priceEur = Number(priceInput||0) * pm;
  if (!priceEur) return null;

  const temp: string = temperature || "DRY";
  const areaKey: string = area || "NORD";
  const isMARE = transport === "MARE";
  const isFF = temp === "FRESH" || temp === "FROZEN";

  // Pallet (BO): COSTS(LOG)!I1 / Y6 = 15 / unitsPerPlt
  const plt = COSTS_CAN.PLT / unitsPerPlt;

  // Tassa Alcolica: importo fisso/unit da Beverage Info (LT × €/LT) — SOLO SPIRITS.
  // Birre e vini non hanno tassa alcolica; l'AIEM% si applica a tutti (cumulativa).
  const isSpirits = (product?.category || "").toUpperCase() === "SPIRITS";
  const aiemFixed: number = isSpirits && (bevData?.totaleBottiglia ?? 0) > 0 ? Number(bevData.totaleBottiglia) : 0;
  const aiemPct = (Number(prodAiem)||0) > 0 ? Number(prodAiem) / 100
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

  // AIEM per isola: % su (prezzo + trasporto) — si cumula con Tassa Alcolica per gli alcolici
  const aiemGCTF   = (priceEur + transpPerIsland("GC"))  * aiemPct;
  const aiemLANFUE = (priceEur + transpPerIsland("LAN")) * aiemPct;
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

  // Arrotonda ogni componente foglia a 2 decimali
  const r2 = (x:number) => roundN(x,2);
  const pE  = r2(priceEur);
  const pL  = r2(plt);
  const whE = r2(wh);
  const taE = r2(aiemFixed);
  const vbE = r2(veronaBarcUnit);
  const asE = r2(assicUnit);
  const fGC = r2(freightPerIsland("GC")),  iGC = r2(inlandPerIsland("GC")),  bGC = r2(barcPerIsland("GC"));
  const fLAN = r2(freightPerIsland("LAN")), iLAN = r2(inlandPerIsland("LAN")), bLAN = r2(barcPerIsland("LAN"));
  const transpGC  = fGC  + iGC  + vbE + bGC  + asE;
  const transpLAN = fLAN + iLAN + vbE + bLAN + asE;
  // Base AIEM: AL + IF(MARE, AM+AU, BC+BE) — senza assicurazione
  // MARE: fGC+iGC>0, vbE=bGC=0; GOMMA: fGC=iGC=0, vbE+bGC>0
  const aGC  = r2((pE + fGC  + iGC  + vbE + bGC)  * aiemPct);
  const aLAN = r2((pE + fLAN + iLAN + vbE + bLAN) * aiemPct);

  const isTakochef = String(vendorName2||"").toUpperCase().includes("TAKOCHEF");
  for (const isl of CAN_ISLANDS) {
    const tr  = (isl==="LAN"||isl==="FUE") ? transpLAN : transpGC;
    const ai  = (isl==="LAN"||isl==="FUE") ? aLAN : aGC;
    step1[isl] = r2(pE + tr + pL + ai + taE);
    step2[isl] = isTakochef ? pE : r2(step1[isl] + whE);
  }
  // Excel: SE(isLAN=0; valore_GC; formula_LAN) — se prodotto non va a LAN/FUE → fallback = GC
  // isLAN/isFUE undefined (logistica non impostata) = comportamento precedente (usa tariffe LAN)
  if (isLAN === false) { step2.LAN = step2.GC; step1.LAN = step1.GC; }
  if (isFUE === false) { step2.FUE = step2.GC; step1.FUE = step1.GC; }

  return {
    priceEur: pE, plt: pL, aiemUnit: aGC, tassaAlcolica: taE, wh: whE,
    transport: transport||"GOMMA", unitsPerPlt,
    veronaBarcUnit: vbE, barcUnitGC: bGC, assicUnit: asE, freightGC: fGC, inlandGC: iGC,
    freightLAN: fLAN, barcUnitLAN: bLAN,
    aiemGCTF: aGC, aiemLANFUE: aLAN,
    isMARE,
    step1GC: step1.GC, step1TF: step1.TF, step1LAN: step1.LAN, step1FUE: step1.FUE,
    step2GC: step2.GC, step2TF: step2.TF, step2LAN: step2.LAN, step2FUE: step2.FUE,
    step1Eur: step1.GC, step1Hkd: step1.GC,
    step2Eur: step2.GC, step2Hkd: step2.GC, rate:1,
    fob:0, lic:0, vgm:0, hc:0, alc: aGC,
  };
}

function exportXLSX(rows: any[], sheetName: string, fileName: string, textCols?: string[]) {
  const textColSet = new Set(textCols || []);
  // Pre-processa le righe: le colonne testo vengono wrappate come oggetti cella SheetJS
  // prima di json_to_sheet, così SheetJS non può convertirle in numero
  const processed = rows.map(row => {
    if (!textColSet.size) return row;
    const out: any = {};
    for (const k of Object.keys(row)) {
      if (textColSet.has(k)) {
        out[k] = { t: 's', v: String(row[k] ?? "") };
      } else {
        out[k] = row[k];
      }
    }
    return out;
  });
  const ws = XLSX.utils.json_to_sheet(processed);
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
function calcHK({ priceInput, ubicazione, product, logistic, eurToHkd, priceMultiplier=1 }: any) {
  const { uom, qtyPerBox, boxPerPallet, kgPerBox, kgxplt, temperature } = product;
  const { pltPerContainer, area, hasCert, hasAlcTax, alcTax, convFactor } = logistic || {};
  const pm = Number(priceMultiplier||1) || 1;

  // ── Units per pallet ── (formula modello Excel) / conv factor item
  // PCS: qtyPerBox × boxPerPallet
  // BOX: boxPerPallet
  // KG:  kgxplt (kg per pallet = KgPerBox × qtyPerBox × boxPerPallet)
  let unitsPerPlt: number;
  if (uom==="BOX") unitsPerPlt = Number(boxPerPallet) / pm;
  else if (uom==="KG") {
    unitsPerPlt = (Number(kgxplt) > 0 ? Number(kgxplt) : 300) / pm;
  }
  else unitsPerPlt = Number(qtyPerBox) * Number(boxPerPallet) / pm; // PCS

  // ── Divisore collo per MTS picking ──
  const divisoreCollo =
    uom==="BOX" ? 1 :
    uom==="KG"  ? Number(kgPerBox||qtyPerBox) :
                  Number(qtyPerBox);

  // ── Total units per container ──
  const totalUnits = unitsPerPlt * Number(pltPerContainer);
  if (!totalUnits) return null;

  const priceEur = Number(priceInput||0) * Number(convFactor) * pm;

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

  // ── Tassa alcolica ── HK spirits >30°: 100% del prezzo acquisto (alcTax=0 → usa priceEur)
  const alc = hasAlcTax ? (Number(alcTax) > 0 ? Number(alcTax) : priceEur) : 0;

  // ── Warehouse ──
  let wh = 0;
  if (ubicazione==="MTO") {
    wh = (COSTS.MTO[temperature] ?? 0) / unitsPerPlt;
  } else if (ubicazione==="MTS") {
    wh = (COSTS.MTS_D[temperature] ?? 0) / unitsPerPlt
       + (COSTS.MTS_I[temperature] ?? 0) / unitsPerPlt
       + (COSTS.MTS_P[temperature] ?? 0) / divisoreCollo;
  }

  // Arrotonda ogni componente a 2 decimali prima di sommare
  const pE = roundN(priceEur,2), fE = roundN(fob,2), lE = roundN(lic,2), vE = roundN(vgm,2);
  const hE = roundN(hc,2), pL = roundN(plt,2), aE = roundN(alc,2), wE = roundN(wh,2);
  const step1Eur = roundN(pE + fE + lE + vE + hE + pL + aE, 2);
  const step2Eur = roundN(step1Eur + wE, 2);

  return {
    priceEur: pE, fob: fE, lic: lE, vgm: vE, hc: hE, plt: pL, alc: aE,
    step1Eur,
    step1Hkd: roundN(step1Eur * eurToHkd,2),
    wh: wE,
    step2Eur,
    step2Hkd: roundN(step2Eur * eurToHkd,2),
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
  const hkNewSC = roundN(hkCost.step2Hkd,2);
  const baseInMop = roundN(hkNewSC * conv * (1 + markup) * HKD_TO_MOP, 2);
  // Costo logistico ALL-IN per MAC UOM (MOP/kg × kg per MAC UOM)
  const logPerKg = MAC_LOG_PER_KG[String(temperature||"DRY").toUpperCase()] ?? 3;
  const logPerUom = roundN(kgPerMacUom > 0 ? logPerKg * kgPerMacUom : logPerKg, 2);
  const macNewSC = roundN(baseInMop + logPerUom, 2);
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
  const dap = pr.dapFinal || 0;
  const fca = pr.fcaDiscounted || pr.fcaPrice || 0;
  if(ubicazione==="FOR") return fca || dap;
  if(ubicazione==="MTO") return dap || fca;
  if(ubicazione==="MTS") { const m=pr.mtsPrice||0; return m!==0?m:dap; }
  return dap || fca;
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
  mtsPrice:      ["mts price","mts price (eur)"],
  fcaPrice:      ["fca price","fca price (eur)","fca"],
  fcaDiscount:   ["fca discount","fca disc","fca discount %"],
  fcaDiscounted: ["fca discounted","fca disc.","fca final"],
  dapPrice:      ["dap price","dap"],
  dapDiscount:   ["dap discount","dap disc"],
  dapDiscounted: ["dap discounted","dap final discounted"],
  dapFinalDirect:["dap final","dap final price","final price","prezzo acquisto"],
  carriageCost:  ["carriage cost","carriage","carriage cost (eur)","costo carriage","costo trasporto"],
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
  const isWine = sec.includes("WINE")||sec.includes("SPIRIT");
  const isX = isWine || FOR_VENDORS.has(vendorName||"");
  const pltCost = isWine ? 60 : (COSTS.VENDOR_CARRIAGE[vendorName]||0);
  const cu = unitsPerPlt>0 ? pltCost/unitsPerPlt : 0;
  const dd=dapDiscounted||0, fp=fcaPrice||0, fd=fcaDiscounted||0;
  if(dd!==0) return { dapFinal:dd, carriageUnit:cu, note:"DAP Disc." };
  if(!isX)   return { dapFinal:0,  carriageUnit:0,  note:"non-X" };
  if(isWine) return { dapFinal:fd!==0?fd+cu:0, carriageUnit:cu, note:"Wine FCA Disc+C" };
  return { dapFinal:fd!==0?fd+cu:0, carriageUnit:cu, note:"FCA Disc+C" };
}

const LS = {
  get: (k,def) => { try{ const v=localStorage.getItem(k); return v?JSON.parse(v):def; }catch{ return def; } },
  set: (k,v) => { try{ localStorage.setItem(k,JSON.stringify(v)); return true; }catch{ return false; } },
};

// Seed data (minimal)
// ─── IndexedDB per dati grandi ──────────────────────────────────────────────
// IDB e CLOUD importati da supabase.ts

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
  const[prices,setPrices]       = useState<any[]>([]);
  const[bcListini,setBcListini] = useState<any[]>([]); // prezzi BC listini — separati da prices per evitare re-render globali
  const[listiniReloadKey,setListiniReloadKey] = useState(0); // incrementa per forzare re-fetch listini
  const[listiniMode,setListiniMode] = useState<"bc"|"excel">(() => LS.get("ifb_listini_mode","bc"));
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
  const [bevInfo, setBevInfo] = useState<any[]>([]); // per branch: HK = spirits >30°; CAN = tassa alcolica propria
  const [priceExceptions, setPriceExceptions] = useState<any[]>(() => LS.get(`ifb_exceptions_${LS.get("ifb_branch","")}`, []));
  const [canConvFactors, setCanConvFactors] = useState<any[]>(() => LS.get("ifb_can_conv_factors", CAN_CONV_DEFAULTS));
  const [hkConvFactors,  setHkConvFactors]  = useState<any[]>(() => LS.get("ifb_hk_conv_factors",  HK_CONV_DEFAULTS));
  const [scAttuali, setScAttuali] = useState<any[]>([]);
  const [scHistory, setScHistory] = useState<any[]>([]); // storico SC Attuali per branch
  const [macHkCostRows, setMacHkCostRows] = useState<any[]>([]); // HK costs loaded for MAC derivation

  // ── Auth state ──────────────────────────────────────────────────────────────
  const supabaseEnabled = !!(import.meta.env.VITE_SUPABASE_URL && import.meta.env.VITE_SUPABASE_ANON_KEY);
  const localBypass = supabaseEnabled && typeof localStorage !== 'undefined' && localStorage.getItem('ifb_local_admin') === '1';
  const [authReady, setAuthReady] = useState(!supabaseEnabled || localBypass); // if no supabase, always ready
  const [authSession, setAuthSession] = useState<any>(null);
  const [authRole, setAuthRole] = useState<'admin'|'viewer'|null>(supabaseEnabled ? null : 'admin');
  const [authEmail, setAuthEmail] = useState("");
  const [authSent, setAuthSent] = useState(false);
  const [authError, setAuthError] = useState("");
  const [authLoading, setAuthLoading] = useState(false);
  const [showUserMgmt, setShowUserMgmt] = useState(false);
  const [userList, setUserList] = useState<any[]>([]);
  const [inviteEmail, setInviteEmail] = useState("");

  useEffect(()=>{
    if(!supabaseEnabled) return;
    (async()=>{
      const { supabase: sb } = await import("./supabase");
      if(!sb) { setAuthReady(true); setAuthRole('admin'); return; }
      const { data: { session } } = await sb.auth.getSession();
      setAuthSession(session);
      if(session?.user?.email) {
        const role = await getUserRole(session.user.email);
        setAuthRole(role);
      }
      setAuthReady(true);
      sb.auth.onAuthStateChange(async (_e, s) => {
        setAuthSession(s);
        if(s?.user?.email) {
          const role = await getUserRole(s.user.email);
          setAuthRole(role);
        } else {
          setAuthRole(null);
        }
      });
    })();
  // eslint-disable-next-line react-hooks/exhaustive-deps
  },[]);

  const navigate = (pageName, filter=null) => { setPageFilter(filter); setPage(pageName); };

  const branchRef = useRef(branch);
  useEffect(()=>{ branchRef.current = branch; },[branch]);
  const branchLoadedRef = useRef<string>("");
  const globalLoadedRef = useRef(false); // blocks global saves until IDB load completes

  // Load global data (logistics, meatPrices) from CLOUD (Supabase → IDB fallback)
  useEffect(()=>{
    (async()=>{
      setLogistics(await CLOUD.get("ifb_logistics", SEED_LOGISTIC));
      setMeatPrices(await CLOUD.get("ifb_meatprices", []));
      globalLoadedRef.current = true;
    })();
  },[]);

  // Reload price exceptions when branch changes
  useEffect(()=>{ if(branch) setPriceExceptions(LS.get(`ifb_exceptions_${branch}`,[])); },[branch]);
  // HK: pre-popola bevInfo con default spirits >30° se ancora vuoto
  useEffect(()=>{ if(branch==="HK" && bevInfo.length===0) setBevInfo(HK_ALC_TAX_DEFAULTS); },[branch, bevInfo.length]);
  // Save effects — only fire after load is complete
  useEffect(()=>{ if(branchRef.current) LS.set(`ifb_exceptions_${branchRef.current}`, priceExceptions); },[priceExceptions]);
  useEffect(()=>{ LS.set("ifb_can_conv_factors", canConvFactors); },[canConvFactors]);
  useEffect(()=>{ LS.set("ifb_hk_conv_factors",  hkConvFactors);  },[hkConvFactors]);
  useEffect(()=>{ if(branchRef.current&&branchLoadedRef.current===branchRef.current) CLOUD.set(`ifb_products_${branchRef.current}`, products); },[products]);
  useEffect(()=>{ if(globalLoadedRef.current) CLOUD.set("ifb_logistics", logistics); }, [logistics]);
  useEffect(()=>{ if(branchRef.current&&branchLoadedRef.current===branchRef.current) CLOUD.set(`ifb_airlist_${branchRef.current}`, airList); },[airList]);
  useEffect(()=>{ if(branchRef.current&&branchLoadedRef.current===branchRef.current) CLOUD.set(`ifb_xrefs_${branchRef.current}`, xrefs); },[xrefs]);
  useEffect(()=>{ if(branchRef.current&&branchLoadedRef.current===branchRef.current) CLOUD.set(`ifb_sales_invoice_${branchRef.current}`, salesRows); },[salesRows]);
  useEffect(()=>{ if(branchRef.current&&branchLoadedRef.current===branchRef.current) CLOUD.set(`ifb_scattuali_${branchRef.current}`, scAttuali); },[scAttuali]);
  useEffect(()=>{ if(branchRef.current&&branchLoadedRef.current===branchRef.current) IDB.set(`ifb_schistory_${branchRef.current}`, scHistory); },[scHistory]);
  // MAC: load saved HK costRows when switching to MAC branch
  useEffect(()=>{ if(branch==="MAC") IDB.get("ifb_hk_costrows_for_mac",[]).then((d:any[])=>setMacHkCostRows(d)); },[branch]);
  useEffect(()=>{ if(branchRef.current&&branchLoadedRef.current===branchRef.current) CLOUD.set(`ifb_prices_${branchRef.current}`, prices); },[prices]);
  useEffect(()=>{ if(branch) LS.set("ifb_branch",branch); },[branch]);
  useEffect(()=>{ LS.set("ifb_listini_mode",listiniMode); },[listiniMode]);
  useEffect(()=>{ if(globalLoadedRef.current) CLOUD.set("ifb_meatprices", meatPrices); }, [meatPrices]);
  useEffect(()=>{ if(branchRef.current&&branchLoadedRef.current===branchRef.current) CLOUD.set(`ifb_bevinfo_${branchRef.current}`, bevInfo); }, [bevInfo]);
  // Ricarica dati branch-specifici ad ogni cambio filiale
  useEffect(()=>{
    if(!branch) return;
    branchLoadedRef.current = ""; // reset — block saves while loading
    (async()=>{
      setProducts(await CLOUD.get(`ifb_products_${branch}`,[]));
      setXrefs(await CLOUD.get(`ifb_xrefs_${branch}`,[]));
      setAirList(await CLOUD.get(`ifb_airlist_${branch}`,[]));
      setSalesRows(await CLOUD.get(`ifb_sales_invoice_${branch}`,[]));
      setScAttuali(await CLOUD.get(`ifb_scattuali_${branch}`,[]));
      setScHistory(await IDB.get(`ifb_schistory_${branch}`,[]));
      setPrices(await CLOUD.get(`ifb_prices_${branch}`,[]));
      setBevInfo(await CLOUD.get(`ifb_bevinfo_${branch}`,[]));
      branchLoadedRef.current = branch; // unblock saves

      // Auto-fetch dati aggiornati da GitHub
      if(branch === "HK") {
        const base = import.meta.env.BASE_URL || "/ifb-platform-std-cost/";
        const t = Date.now();
        try {
          const [rxref, rsc, rana] = await Promise.all([
            fetch(`${base}data/hk_xref.json?t=${t}`),
            fetch(`${base}data/hk_sc.json?t=${t}`),
            fetch(`${base}data/hk_anagrafica.json?t=${t}`),
          ]);
          if(rxref.ok) { const d=await rxref.json(); if(Array.isArray(d)&&d.length>0){setXrefs(d);CLOUD.set(`ifb_xrefs_${branch}`,d);setDataSource(`xref_${branch}`,"bc");} }
          if(rsc.ok)  { const d=await rsc.json();  if(Array.isArray(d)&&d.length>0){setScAttuali(d);CLOUD.set(`ifb_scattuali_${branch}`,d);setDataSource(`scattuali_${branch}`,"bc");} }
          if(rana.ok) { const d=await rana.json(); if(Array.isArray(d)&&d.length>0){setProducts(d);CLOUD.set(`ifb_products_${branch}`,d);setDataSource(`anagrafica_${branch}`,"bc");} }
        } catch(_) { /* offline o errore fetch — usa dati IDB */ }
      }
      // CAN: dati da file NAV/COMIT committati in docs/data/ (rigenerati da sync_json_data.py)
      if(branch === "CAN") {
        const base = import.meta.env.BASE_URL || "/ifb-platform-std-cost/";
        const t = Date.now();
        try {
          const [rana, rxref, rwt] = await Promise.all([
            fetch(`${base}data/can_anagrafica.json?t=${t}`),
            fetch(`${base}data/can_xref.json?t=${t}`),
            fetch(`${base}data/can_worktab.json?t=${t}`),
          ]);
          let canProds: any[] = [];
          let canXrefs: any[] = [];
          if(rana.ok)  { canProds=await rana.json();  if(canProds.length>0){setProducts(canProds);CLOUD.set(`ifb_products_${branch}`,canProds);setDataSource(`anagrafica_${branch}`,"bc");} }
          if(rxref.ok) { canXrefs=await rxref.json(); if(canXrefs.length>0){setXrefs(canXrefs);CLOUD.set(`ifb_xrefs_${branch}`,canXrefs);setDataSource(`xref_${branch}`,"bc");} }
          // SC Attuali CAN: NON auto-caricata da JSON — gestita manualmente dalla pagina SC Attuali
          // Auto-build logistica CAN da worktab: converte nComit/ifbNo → productId
          if(rwt.ok && canProds.length > 0) {
            const wt: any[] = await rwt.json();
            const prevLog: any[] = await CLOUD.get("ifb_logistics", []);
            const otherBranchLog = prevLog.filter((l:any) => l.branch !== "CAN");
            // Indice delle righe CAN modificate manualmente: non vanno sovrascritte dal worktab
            const manualOverrides: Record<string, any> = {};
            prevLog.filter((l:any) => l.branch === "CAN" && l._manualOverride).forEach((l:any) => {
              manualOverrides[l.productId] = l;
            });
            const canLog = wt.flatMap((row: any) => {
              const prod = canProds.find((p:any) => p.nHK === row.nComit || p.code === row.ifbNo)
                        || canProds.find((p:any) => canXrefs.find((x:any) => x.nHK === row.nComit && x.ifbNo === p.code));
              if(!prod) return [];
              // Preserva modifiche manuali dell'utente
              if(manualOverrides[prod.id]) return [manualOverrides[prod.id]];
              const transport = (row.transport||"").toUpperCase().includes("MARE") ? "MARE" : "GOMMA";
              const ubicazione = row.mtsUb === "MTS" ? "MTS" : row.mtsUb === "FOR" ? "FOR" : "MTO";
              return [{
                productId: prod.id,
                nHK: prod.nHK,
                branch: "CAN",
                area: row.area || "NORD",
                ubicazione,
                transport,
                pltPerContainer: 0,
                hasCert: false,
                hasAlcTax: false,
                alcTax: 0,
                convFactor: 1,
                carriage: row.carriage || 0,
                temperatureOverride: null,
                fromImport: true,
                isGC: !!row.isGC,
                isTF: !!row.isTF,
                isFUE: !!row.isFUE,
                isLAN: !!row.isLAN,
              }];
            });
            const merged = [...otherBranchLog, ...canLog];
            setLogistics(merged);
            CLOUD.set("ifb_logistics", merged);
          }
        } catch(_) { /* offline — usa dati IDB */ }
      }

      // Auto-fetch fatture IFB da GitHub (HK + CAN + MAC, filtrate per branch/customer)
      if(["HK","CAN","MAC"].includes(branch)) {
        const base = import.meta.env.BASE_URL || "/ifb-platform-std-cost/";
        try {
          const r = await fetch(`${base}data/ifb_fatture.json?t=${Date.now()}`);
          if(r.ok) {
            const all = await r.json();
            if(Array.isArray(all) && all.length > 0) {
              const prods: any[] = await IDB.get(`ifb_products_${branch}`, []);
              const xrs: any[]   = await IDB.get(`ifb_xrefs_${branch}`, []);
              const airl: any[]  = await IDB.get(`ifb_airlist_${branch}`, []);
              // Lookup maps O(1) per fatture
              const fByCode: Record<string,any> = {};
              const fByNHK:  Record<string,any> = {};
              prods.forEach((p: any) => { if(p.code) fByCode[p.code]=p; if(p.nHK) fByNHK[p.nHK]=p; });
              const fXrByIfb: Record<string,string> = {};
              xrs.forEach((x: any) => { if(x.ifbNo && x.nHK) fXrByIfb[x.ifbNo]=x.nHK; });
              const airSet = new Set(airl.flatMap((a: any) => [a.productId, a.code, a.nHK].filter(Boolean)));
              const branchRows = all
                .filter((row: any) => row.Branch === branch)
                .map((row: any) => {
                  const code = String(row["No_"] || "").trim();
                  const prod = fByCode[code] || fByNHK[code] || (fXrByIfb[code] ? fByNHK[fXrByIfb[code]] : null);
                  const nHK  = prod?.nHK || fXrByIfb[code] || "";
                  const isAirProd = prod && (
                    airSet.has(prod.id) || airSet.has(prod.code) || airSet.has(prod.nHK) ||
                    isAirTransport(prod.bcTransportation)
                  );
                  return {
                    itemCode:   code,
                    description: String(row["Description"] || "").trim(),
                    date:       String(row["Last Posting Date"] || ""),
                    qty:        Number(row["Quantity"] || 0),
                    unitPrice:  Number(row["Price"] || 0),
                    location:   String(row["Location Code"] || "").trim(),
                    nHK,
                    transport:  isAirProd ? "AIR" : "SEA",
                    _prodFound: !!prod,
                    branch,
                    _fromBC:    true,
                  };
                });
              if(branchRows.length > 0) {
                setSalesRows(branchRows);
                CLOUD.set(`ifb_sales_invoice_${branch}`, branchRows);
                setDataSource(`fatture_${branch}`,"bc");
              }
            }
          }
        } catch(_) { /* offline — usa dati IDB */ }
      }

    })();
  },[branch]);

  // Fetch listini separato — si riattiva su cambio branch O su ricarica manuale (listiniReloadKey)
  useEffect(()=>{
    if(!branch || !["HK","CAN","MAC"].includes(branch)) return;
    const IDB_KEY = `ifb_listini_entries_${branch}`;
    const base = import.meta.env.BASE_URL || "/ifb-platform-std-cost/";
    // Pulisce subito i dati del branch precedente per evitare di mostrare dati sbagliati
    setBcListini([]);
    (async()=>{
      // 1) Carica subito da IDB (cache locale) — solo se i dati appartengono al branch corrente
      const cached: any[] = (await IDB.get(IDB_KEY, []) as any[]).filter((e:any) => !e.branch || e.branch === branch);
      if(cached.length > 0) {
        startTransition(() => { setBcListini(cached); setDataSource(`listini_${branch}`, "bc"); });
      }
      // 2) Fetch JSON aggiornato da GitHub
      try {
        const nc = {cache:"no-store"} as RequestInit;
        let resp = await fetch(`${base}data/ifb_listini_${branch}.json?t=${Date.now()}`, nc);
        if(!resp.ok) resp = await fetch(`${base}data/ifb_listini.json?t=${Date.now()}`, nc);
        if(resp.ok) {
          const raw = await resp.json();
          const all = Array.isArray(raw) ? raw.filter((r:any) => (r.b || r.Branch) === branch) : [];
          if(all.length > 0) {
            const prods: any[] = await IDB.get(`ifb_products_${branch}`, []);
            const xrs: any[]   = await IDB.get(`ifb_xrefs_${branch}`, []);
            const byCode: Record<string,any> = {};
            const byNHK:  Record<string,any> = {};
            prods.forEach((p: any) => { if(p.code) byCode[String(p.code)]=p; if(p.nHK) byNHK[String(p.nHK)]=p; });
            const xrByIfb: Record<string,string> = {};
            xrs.forEach((x: any) => { if(x.ifbNo && x.nHK) xrByIfb[String(x.ifbNo)]=String(x.nHK); });
            const nowMonth = new Date().toISOString().slice(0,7);
            const newEntries: any[] = [];
            all.forEach((row: any) => {
              const code = String(row["n"] || row["No_"] || "").trim();
              if(!code) return;
              const nhkField = String(row["nhk"] || "").trim();
              const prod = (nhkField ? byNHK[nhkField] : null) || byCode[code] || byNHK[code] || (xrByIfb[code] ? byNHK[xrByIfb[code]] : null);
              const purchUom = String(row["pu"] || "").trim().toUpperCase();
              const scriptCf = Number(row["cf"] || 1);
              // Regola generale (uguale a PBI con fatt_conv):
              //   cf > 1 → script ha già diviso il prezzo per cf (conv_qty da qtyperunitofmeasure).
              //            Se pu == hkUom la divisione era inutile: re-moltiplica.
              //            Se pu != hkUom il prezzo è già nella base UoM: nessuna azione.
              //   cf == 1 → script non ha applicato conversioni (fatt_conv=1 in PBI).
              //             Il prezzo è già corretto nell'unità pu: nessuna conversione aggiuntiva.
              let convFactor = 1;
              const hkUom: string = prod?.uom || "PCS";
              let displayUom: string;
              if (scriptCf > 1 && purchUom && prod) {
                if (purchUom !== "PCS" && purchUom === hkUom) {
                  // Script ha diviso inutilmente (pu non-PCS già in hkUom): re-moltiplica
                  convFactor = 1 / scriptCf;
                  displayUom = hkUom;
                } else {
                  // Script ha diviso correttamente: prezzo per base UoM (= purchUom se PCS, altrimenti hkUom)
                  displayUom = purchUom || hkUom;
                }
              } else {
                // cf=1: prezzo per purchUom (PBI usa Base Unit of Measure dell'item)
                displayUom = purchUom || hkUom;
              }
              const div = (p: number) => convFactor !== 1 ? p / convFactor : p;
              newEntries.push({
                productId:     nhkField || prod?.nHK || prod?.id || `BC_${code}`,
                itemCode:      code,
                nHK:           nhkField || prod?.nHK || "",
                bcDesc:        String(row["d"] || row["Description"] || "").trim(),
                pu:            displayUom,
                branch, month: nowMonth,
                fcaPrice:      div(Number(row["fp"] ?? row["FCA_Price"]      ?? 0)),
                fcaDiscounted: div(Number(row["fc"] ?? row["FCA_Discounted"] ?? 0)),
                dapPrice:      div(Number(row["dp"] ?? row["DAP_Price"]      ?? 0)),
                dapDiscounted: div(Number(row["dc"] ?? row["DAP_Discounted"] ?? 0)),
                carriageCost:  div(Number(row["cr"] ?? row["Carriage"]       ?? 0)),
                dapFinal: (()=>{
                  const dc=div(Number(row["dc"]??row["DAP_Final"]??row["DAP_Discounted"]??0));
                  if(dc>0) return dc;
                  const fc=div(Number(row["fc"]??row["FCA_Discounted"]??0));
                  const fp=div(Number(row["fp"]??row["FCA_Price"]??0));
                  const cr=div(Number(row["cr"]??row["Carriage"]??0));
                  if(cr>0&&(fc>0||fp>0)) return (fc||fp)+cr;
                  return 0;
                })(),
                mtsPrice:      div(Number(row["mp"] ?? row["MTS_Price"]      ?? 0)),
              });
            });
            // Dedup per productId+branch: tieni quello con sconto se disponibile, altrimenti quello col prezzo più alto
            const dedupMap = new Map<string, any>();
            newEntries.forEach(e => {
              const key = `${e.branch || branch}_${e.productId}`;
              const prev = dedupMap.get(key);
              if (!prev) { dedupMap.set(key, e); return; }
              const hasDisc = (x:any) => (x.fcaDiscounted||0) > 0 && (x.fcaDiscounted||0) < (x.fcaPrice||0);
              const eDisc = hasDisc(e), prevDisc = hasDisc(prev);
              if (eDisc && !prevDisc) { dedupMap.set(key, e); return; } // nuovo ha sconto, vecchio no → sostituisci
              if (!eDisc && !prevDisc && (e.fcaPrice||0) > (prev.fcaPrice||0)) dedupMap.set(key, e); // nessuno ha sconto → tieni il più alto
            });
            const dedupEntries = [...dedupMap.values()];
            IDB.set(IDB_KEY, dedupEntries);
            startTransition(() => { setBcListini(dedupEntries); setDataSource(`listini_${branch}`, "bc"); });
          }
        }
      } catch(_) { /* offline — usa dati IDB già caricati */ }
    })();
  },[branch, listiniReloadKey]);

  const showToast = (msg,color=T.green) => { setToast({msg,color}); setTimeout(()=>setToast(null),3500); };
  const bumpImportTs = () => { const ts=Date.now(); setLastImportTs(ts); LS.set("ifb_last_import_ts",ts); return ts; };

  // Arricchisce bcListini con carriage/DAP derivati dalla logistica (wine=60€/plt, FOR vendors)
  // Eseguito come useMemo per garantire che products sia già caricato
  const bcListiniEnriched = useMemo(()=>{
    if(!bcListini.length || !products.length) return bcListini;
    return bcListini.map((p:any)=>{
      if((p.carriageCost||0)>0 || (p.dapFinal||0)>0) return p;
      const prod=products.find((pr:any)=>pr.id===p.productId||pr.code===(p.itemCode||p.n||""));
      if(!prod) return p;
      const r=calcDAPFinal({
        dapDiscounted:p.dapDiscounted||0, fcaPrice:p.fcaPrice||0, fcaDiscounted:p.fcaDiscounted||0,
        vendorName:prod.vendorName||prod.vendorName2||"",
        section:prod.category||"", products, code:prod.code||"",
      });
      if(r.dapFinal>0||r.carriageUnit>0) return {...p, carriageCost:r.carriageUnit, dapFinal:r.dapFinal};
      return p;
    });
  },[bcListini,products]);

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
    // CAN: l'anagrafica è già filtrata (solo prodotti filiale), non filtrare per vendor
    const eligible = branch === "CAN"
      ? products.filter((p:any) => p.active !== false)
      : products.filter(p => isIFBVendor(p.vendorName));

    return eligible.map(prod => {
      // Eccezione prezzo manuale: ha priorità assoluta su listino e carne
      const exc = priceExceptions.find((e:any) =>
        e.branch === branch && (
          e.productId === prod.id ||
          (e.code && e.code === prod.code) ||
          (e.nHK && prod.nHK && e.nHK === prod.nHK)
        )
      );

      const airEntry = branch!=="CAN" ? airList.find((a:any)=>
          a.productId === prod.id ||
          (a.code && a.code === prod.code) ||
          (a.nHK && prod.nHK && a.nHK === prod.nHK)
        ) : null;
      if(branch!=="CAN" && ((airEntry && isAirTransport(airEntry.transportation)) || isAirTransport(prod.bcTransportation)))
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

      // CAN: i prezzi vengono SOLO dal file listino caricato dall'utente (prices).
      //      Il BC listino non è usato per il calcolo SC CAN (solo per fatture/costi).
      // HK/MAC: match per N COMIT (namespace BC), fallback IFB solo se nHK assente.
      const pr     = prices.find(p=>p.productId===prod.id&&p.branch===branch&&p.month===month)
                  || (branch === "CAN" ? null
                    : (prod.nHK && bcListiniEnriched.find((p:any)=>(p.itemCode||p.n)===prod.nHK&&(p.branch||p.b)===branch))
                      || (!prod.nHK && prod.code && bcListiniEnriched.find((p:any)=>(p.itemCode||p.n)===prod.code&&(p.branch||p.b)===branch)));
      const prPrev = prices.find(p=>p.productId===prod.id&&p.branch===branch&&p.month===prevM);

      const ub = log.ubicazione;
      const effectiveProd = log.temperatureOverride ? { ...prod, temperature: log.temperatureOverride } : prod;

      // Branch-agnostic calc helper
      const isCAN_b = branch === "CAN";
      const bevData = bevInfo.find((b:any) => b.ifbNo === prod.code || (prod.nHK && b.nHK === prod.nHK)) || null;
      // Fattore di conversione item: CAN → lookup per N COMIT via xrefs; HK → lookup per nHK
      const nComit = isCAN_b ? (prod.nHK || xrefs.find((x:any)=>x.ifbNo===prod.code)?.nHK || "") : "";
      const configuredCf = isCAN_b
        ? (nComit ? (canConvFactors.find((c:any)=>c.nComit===nComit)?.factor||1) : 1)
        : (hkConvFactors.find((c:any)=>c.nHK===prod.nHK)?.factor||1);
      const calcCost = (pi: number, cf = configuredCf) =>
        isCAN_b
          ? calcCAN({ priceInput:pi, ubicazione:ub, product:effectiveProd, logistic:log, bevData, priceMultiplier:cf })
          : calcHK({ priceInput:pi, ubicazione:ub, product:effectiveProd, logistic:{...log, category:prod.category, hasAlcTax: log.hasAlcTax||(bevData?.hasAlcTax===true)||(bevData?.totaleBottiglia>0), alcTax: log.alcTax||(bevData?.totaleBottiglia||0)}, eurToHkd:fxRate, priceMultiplier:cf });

      // Eccezione prezzo: bypassa listino e carne
      // Eccezione prezzo: +2% intercompany markup come listino BC
      if(exc && exc.price > 0) {
        const costE = calcCost(exc.price * (100 / 98));
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
        return { pi: meat.pricePerKg * kgPerUnit * (100 / 98) }; // +2% intercompany
      };

      if(!pr) {
        const mf = meatFallback();
        if(!mf) {
          if(prPrev) {
            // Listino chiuso: usa SC da scAttuali come KEEP OLD
            const scEntry = scAttuali.find((s:any) =>
              s.code === prod.code ||
              s.code === prod.nHK ||
              s.ifbCode === prod.code ||
              s.ifbCode === prod.nHK ||
              s.code === String(prod.id)
            );
            // CAN: usa scGC come riferimento keep old
            const keepOldVal = branch === "CAN"
              ? (scEntry?.scGC || scEntry?.scLan || scEntry?.lastSC || 0)
              : (scEntry?.lastSC || 0);
            if(scEntry && keepOldVal > 0) {
              const isCAN_b2 = branch === "CAN";
              const keepCost = isCAN_b2
                ? { step2Eur: keepOldVal, step2Hkd: keepOldVal * fxRate, _keepOld: true }
                : { step2Hkd: keepOldVal, step2Eur: keepOldVal / fxRate, _keepOld: true };
              return { ...prod, cost:keepCost, prevCost:null, delta:null, priceInput:null,
                ubicazione:ub, skipReason:"KEEP OLD", _keepOld:true };
            }
          }
          const skipReason = prPrev
            ? `LISTINO CHIUSO E NON RIAPERTO (${branch}/${month})`
            : `NO PREZZO (${branch}/${month})`;
          return { ...prod, cost:null, prevCost:null, priceInput:null, ubicazione:ub, skipReason };
        }
        const cost2 = calcCost(mf.pi);
        return { ...prod, cost:cost2, prevCost:null, delta:null, priceInput:mf.pi, isNew:true,
          flagged:false, ubicazione:ub, pltUsed:plt, area:log.area||"NORD", pltPerContainer:plt,
          temperatureOverride:log.temperatureOverride||null,
          skipReason: cost2 ? undefined : "CALC=0", _fromMeatList:true };
      }

      const uomConvFactor = 1;
      // Il fattore di conversione si applica SOLO se il prezzo è stato trovato
      // tramite cross-reference (codice del listino ≠ nHK del prodotto).
      // Es: 15307 (nHK) → prezzo trovato come KDC01 → applica ×40.
      //     KDC01 (nHK) → prezzo trovato come KDC01 → fattore = 1 (match diretto).
      const prCode = isCAN_b ? String((pr as any)?.itemCode || (pr as any)?.n || "") : "";
      const foundViaXref = isCAN_b && !!prCode && !!prod.nHK && prCode !== prod.nHK;
      const itemCf = foundViaXref ? configuredCf : (isCAN_b ? 1 : configuredCf);
      // Se il prezzo è trovato via xref e il prodotto xref ha UoM diversa, dividi per la sua qtyPerBox
      // Es: 15307(PCS) trova KDC01(BOX, qpb=40) → rawPi/40 prima di applicare itemCf=40
      const xrefProd = foundViaXref ? products.find((p:any) => p.nHK === prCode || p.code === prCode) : null;
      // Divide per qtyPerBox del prodotto xref (BOX→PCS), o per il conv factor se xrefProd non trovato
      // (quando il listino BC ha il prezzo per-BOX ma il target è PCS)
      const xrefQpbDiv = !foundViaXref ? 1
        : xrefProd && xrefProd.uom === "BOX" && prod.uom === "PCS"
          ? (Number(xrefProd.qtyPerBox) || 1)
          : (prod.uom === "PCS" && configuredCf > 1 ? configuredCf : 1);
      // Per HK MTO: se dapFinal=0 ma fcaDiscounted>0, aggiungi carriage da work tab
      const enrichPriceWithCarriage = (p: any) => {
        if(!p) return 0;
        if(isCAN_b) {
          const dap = p.dapFinal || 0;
          const fca = p.fcaDiscounted || p.fcaPrice || 0;
          if(ub === "FOR") return fca || dap;
          if(dap === 0 && fca > 0) {
            const sec2 = (prod.category || "").toUpperCase();
            if(sec2.includes("WINE") || sec2.includes("SPIRIT")) {
              const uom2 = prod.uom || "PCS";
              const div2 = uom2 === "BOX" ? (Number(prod.boxPerPallet)||1)
                         : uom2 === "KG"  ? (Number(prod.kgxplt)||300)
                         : (Number(prod.qtyPerBox)||1) * (Number(prod.boxPerPallet)||1);
              return fca + (div2 > 0 ? 60 / div2 : 0);
            }
          }
          return selectPrice(p, ub);
        }
        const sel = selectPrice(p, ub);
        if(sel > 0) return sel;
        // fallback: fca + carriage da logistica / unitsPerPlt
        const fca = p.fcaDiscounted || p.fcaPrice || 0;
        if(!fca) return 0;
        const upm = prod.uom==="BOX" ? Number(prod.boxPerPallet)
                  : prod.uom==="KG"  ? (Number(prod.kgxplt)||300)
                  : Number(prod.qtyPerBox)*Number(prod.boxPerPallet);
        const cu = upm > 0 ? (log.carriage||0) / upm : 0;
        return fca + cu;
      };
      const rawPi  = (enrichPriceWithCarriage(pr)  || 0) * uomConvFactor;
      const rawPiP = prPrev ? (enrichPriceWithCarriage(prPrev) || 0) * uomConvFactor : null;
      // CAN: il BC listino può avere pu="BOX" con prezzo per-BOX (display divide per qtyPerBox).
      // calcCAN si aspetta priceInput per-base-UoM (PCS) e moltiplica × itemCf per avere per-BOX.
      // Senza questa normalizzazione si avrebbe: 18.8/BOX × 100 = 1880 invece di 0.188/PCS × 100 = 18.8.
      const prPu = (pr as any)?.pu || "";
      const needsBoxToPCS = isCAN_b && prPu === "BOX" && prod.uom === "PCS" && configuredCf > 1;
      const qpbDiv = needsBoxToPCS ? (Number(prod.qtyPerBox) || 1) : xrefQpbDiv;
      const pi  = rawPi  / qpbDiv;
      const piP = rawPiP != null ? rawPiP / qpbDiv : null;

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
  }, [products,logistics,prices,fx,airList,meatPrices,priceExceptions,branch,month,bevInfo,scAttuali,xrefs,canConvFactors,hkConvFactors,bcListiniEnriched]);

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
    ...(!isMAC ? [{id:"pricecompare", icon:"⚖", label:"🔬 Confronto Listini"}] : []),
    ...(!isMAC ? [{id:"meatlist",  icon:"🥩", label:"Listino Carne"}] : []),
    ...(!isMAC ? [{id:"bevinfo", icon:"🍷", label: isCAN ? "Beverage Info (Alcol Tax CAN)" : "Beverage Info (Alcol Tax)"}] : []),
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

  // ── Auth gate ───────────────────────────────────────────────────────────────
  if(!authReady) return (
    <div style={{display:"flex",height:"100vh",alignItems:"center",justifyContent:"center",background:T.bg,color:T.muted,fontFamily:"inherit",fontSize:"13px"}}>
      Verifica accesso…
    </div>
  );

  if(supabaseEnabled && !localBypass && !authSession) return (
    <div style={{display:"flex",height:"100vh",width:"100vw",background:T.bg,alignItems:"center",justifyContent:"center",fontFamily:"'Palatino Linotype','Book Antiqua',Palatino,serif"}}>
      <div style={{textAlign:"center",maxWidth:"420px",padding:"40px",background:T.card,border:`1px solid ${T.border}`,borderRadius:"20px"}}>
        <div style={{fontSize:"10px",letterSpacing:"4px",color:T.gold,textTransform:"uppercase",marginBottom:"8px"}}>IFB Platform</div>
        <h2 style={{color:T.text,margin:"0 0 6px",fontSize:"24px"}}>Cost Intelligence</h2>
        <div style={{color:T.muted,fontSize:"12px",marginBottom:"32px"}}>Accesso riservato al personale autorizzato</div>
        {!authSent ? (
          <>
            <input
              type="email" placeholder="La tua email aziendale" value={authEmail}
              onChange={e=>setAuthEmail(e.target.value)}
              onKeyDown={e=>{ if(e.key==="Enter") {
                setAuthLoading(true); setAuthError("");
                signInWithOtp(authEmail).then(()=>setAuthSent(true)).catch(err=>setAuthError(err.message)).finally(()=>setAuthLoading(false));
              }}}
              style={{width:"100%",padding:"12px 14px",background:T.bg,border:`1px solid ${T.border}`,borderRadius:"8px",color:T.text,fontSize:"14px",outline:"none",boxSizing:"border-box",marginBottom:"12px"}}
            />
            <button
              disabled={authLoading||!authEmail}
              onClick={()=>{ setAuthLoading(true); setAuthError(""); signInWithOtp(authEmail).then(()=>setAuthSent(true)).catch(err=>setAuthError(err.message)).finally(()=>setAuthLoading(false)); }}
              style={{width:"100%",padding:"12px",background:T.gold,border:"none",borderRadius:"8px",color:"#000",fontSize:"14px",fontWeight:"bold",cursor:"pointer",opacity:authLoading||!authEmail?0.6:1}}>
              {authLoading ? "Invio…" : "Invia link di accesso"}
            </button>
            {authError && <div style={{marginTop:"10px",color:T.red,fontSize:"12px"}}>{authError}</div>}
          </>
        ) : (
          <div style={{color:T.green,fontSize:"14px",lineHeight:"1.6"}}>
            ✉ Link inviato a <strong>{authEmail}</strong><br/>
            <span style={{color:T.muted,fontSize:"12px"}}>Controlla la tua email e clicca il link per accedere.</span>
            <br/><br/>
            <button onClick={()=>{ setAuthSent(false); setAuthEmail(""); }} style={{background:"none",border:"none",color:T.muted,cursor:"pointer",fontSize:"12px",textDecoration:"underline"}}>
              Usa un'altra email
            </button>
          </div>
        )}
        <div style={{marginTop:"24px",borderTop:`1px solid ${T.border}`,paddingTop:"16px"}}>
          <button onClick={()=>{ localStorage.setItem('ifb_local_admin','1'); window.location.reload(); }}
            style={{background:"none",border:"none",color:T.border,cursor:"pointer",fontSize:"10px",letterSpacing:"1px"}}>
            accesso dispositivo
          </button>
        </div>
      </div>
    </div>
  );

  if(supabaseEnabled && !localBypass && authSession && authRole===null) return (
    <div style={{display:"flex",height:"100vh",alignItems:"center",justifyContent:"center",background:T.bg,color:T.red,fontFamily:"inherit",fontSize:"13px",flexDirection:"column",gap:"12px"}}>
      <div>⛔ Accesso non autorizzato per <strong>{authSession.user.email}</strong></div>
      <button onClick={signOut} style={{padding:"8px 16px",background:"none",border:`1px solid ${T.border}`,borderRadius:"6px",color:T.muted,cursor:"pointer",fontSize:"12px"}}>Esci</button>
    </div>
  );

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
  xrefs={xrefs}
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
  bcListini={bcListiniEnriched}
  setBcListini={setBcListini}
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
  reloadListini={()=>{ setBcListini([]); setListiniReloadKey(k=>k+1); }}
  listiniMode={listiniMode} setListiniMode={setListiniMode}
/>,


    
    fx:          <FxRates fx={fx} setFx={setFx} branch={branch} month={month}/>,
    air:         <AirListPage airList={airList} setAirList={setAirList} products={products} xrefs={xrefs} branch={branch} snapshots={snapshots} setSnapshots={setSnapshots} importLogs={importLogs} setImportLogs={setImportLogs} showToast={showToast} bumpImportTs={bumpImportTs}/>,
    pricecompare: <PriceComparePage bcListini={bcListiniEnriched} prices={prices} products={products} xrefs={xrefs} branch={branch} month={month}/>,
    meatlist: <MeatPriceListPage meatPrices={meatPrices} setMeatPrices={setMeatPrices} products={products} xrefs={xrefs} importLogs={importLogs} setImportLogs={setImportLogs} snapshots={snapshots} setSnapshots={setSnapshots} showToast={showToast} bumpImportTs={bumpImportTs}/>,
    bevinfo: <BeverageInfoPage bevInfo={bevInfo} setBevInfo={setBevInfo} products={products} xrefs={xrefs} showToast={showToast} branch={branch}/>,
    exceptions:  <PriceExceptions branch={branch} products={products} xrefs={xrefs} priceExceptions={priceExceptions} setPriceExceptions={setPriceExceptions} canConvFactors={canConvFactors} setCanConvFactors={setCanConvFactors} hkConvFactors={hkConvFactors} setHkConvFactors={setHkConvFactors}/>,
    costs:       <CostTable costRows={costRows} branch={branch} month={month} logistics={logistics} lastImportTs={lastImportTs} lastCalcTs={lastCalcTs} setLastCalcTs={setLastCalcTs} setCostHistory={setCostHistory} initFilter={pageFilter} salesRows={salesRows} products={products} xrefs={xrefs} listiniMode={listiniMode} setListiniMode={setListiniMode} reloadListini={()=>{ setBcListini([]); setListiniReloadKey(k=>k+1); }}/>,
    invoice: <InvoiceAndCosts rows={salesRows} setRows={setSalesRows} branch={branch} airList={airList} products={products} xrefs={xrefs} costRows={costRows} logistics={logistics} snapshots={snapshots} setSnapshots={setSnapshots} importLogs={importLogs} setImportLogs={setImportLogs} showToast={showToast} bumpImportTs={bumpImportTs} scAttuali={scAttuali}/>,
    scattuali: <ScAttualiPage scAttuali={scAttuali} setScAttuali={setScAttuali} scHistory={scHistory} setScHistory={setScHistory} branch={branch} showToast={showToast} xrefs={xrefs}/>,
    storico: <Storico
      snapshots={snapshots}
      setSnapshots={setSnapshots}
      costHistory={costHistory}
      setCostHistory={setCostHistory}
      branch={branch}
      showToast={showToast}
      macHkCostRows={macHkCostRows}
    />,
    check: <CheckMensile costRows={costRows} branch={branch} salesRows={salesRows} xrefs={xrefs} scAttuali={scAttuali} products={products} logistics={logistics}/>,
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
          <div style={{marginLeft:"auto",display:"flex",gap:"6px",alignItems:"center"}}>
            {supabaseEnabled && authSession && (
              <span style={{fontSize:"10px",color:T.muted,maxWidth:"140px",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={authSession.user.email}>{authSession.user.email}</span>
            )}
            {supabaseEnabled && authRole==="admin" && (
              <button onClick={async()=>{ setUserList(await listUsers()); setShowUserMgmt(true); }}
                style={{padding:"5px 10px",background:"none",border:`1px solid ${T.border}`,borderRadius:"5px",color:T.muted,cursor:"pointer",fontFamily:"inherit",fontSize:"10px"}}>
                👥 Utenti
              </button>
            )}
            {supabaseEnabled && authSession && (
              <button onClick={()=>{ if(window.confirm("Esci dall'account?")) signOut(); }}
                style={{padding:"5px 10px",background:"none",border:`1px solid ${T.border}`,borderRadius:"5px",color:T.muted,cursor:"pointer",fontFamily:"inherit",fontSize:"10px"}}>
                Esci
              </button>
            )}
            <button onClick={()=>setPage("mail")} style={{padding:"5px 12px",background:T.gold,border:"none",borderRadius:"5px",color:T.bg,cursor:"pointer",fontFamily:"inherit",fontSize:"10px",fontWeight:"bold"}}>✉ Mail</button>
          </div>
        </div>
        <div style={{flex:1,paddingTop:"20px",paddingLeft:"28px",paddingBottom:"20px",paddingRight:"24px",overflow:"auto",minWidth:0,boxSizing:"border-box"}}>
  {isCAN && (page==="air" || page==="fx") ? pages["dashboard"] : pages[page]}
</div>
      </div>
      {toast&&<div style={{position:"fixed",bottom:"24px",right:"24px",padding:"10px 18px",background:toast.color,borderRadius:"8px",color:"#fff",fontSize:"12px",fontWeight:"bold",boxShadow:"0 8px 24px rgba(0,0,0,0.4)",zIndex:1000}}>{toast.msg}</div>}
      {showUserMgmt&&(
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.6)",zIndex:2000,display:"flex",alignItems:"center",justifyContent:"center"}} onClick={()=>setShowUserMgmt(false)}>
          <div style={{background:T.card,border:`1px solid ${T.border}`,borderRadius:"16px",padding:"28px",minWidth:"420px",maxWidth:"560px",width:"90%"}} onClick={e=>e.stopPropagation()}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"20px"}}>
              <h3 style={{margin:0,color:T.text,fontSize:"16px"}}>Gestione Accessi</h3>
              <button onClick={()=>setShowUserMgmt(false)} style={{background:"none",border:"none",color:T.muted,cursor:"pointer",fontSize:"18px"}}>✕</button>
            </div>
            <div style={{marginBottom:"16px"}}>
              <div style={{fontSize:"11px",color:T.muted,marginBottom:"6px"}}>Invita utente</div>
              <div style={{display:"flex",gap:"8px"}}>
                <input type="email" placeholder="email@inalcafb.com" value={inviteEmail} onChange={e=>setInviteEmail(e.target.value)}
                  style={{flex:1,padding:"8px 10px",background:T.bg,border:`1px solid ${T.border}`,borderRadius:"6px",color:T.text,fontSize:"13px",outline:"none"}}/>
                <button onClick={async()=>{
                    if(!inviteEmail) return;
                    await inviteUser(inviteEmail,"viewer");
                    setUserList(await listUsers());
                    setInviteEmail("");
                    showToast(`${inviteEmail} aggiunto come viewer`, T.green);
                  }}
                  style={{padding:"8px 14px",background:T.gold,border:"none",borderRadius:"6px",color:"#000",fontSize:"12px",fontWeight:"bold",cursor:"pointer"}}>
                  Aggiungi
                </button>
              </div>
            </div>
            <table style={{width:"100%",borderCollapse:"collapse"}}>
              <thead><tr style={{borderBottom:`1px solid ${T.border}`}}>
                <th style={{textAlign:"left",padding:"6px",color:T.muted,fontSize:"11px",fontWeight:"normal"}}>Email</th>
                <th style={{textAlign:"left",padding:"6px",color:T.muted,fontSize:"11px",fontWeight:"normal"}}>Ruolo</th>
                <th/>
              </tr></thead>
              <tbody>{userList.map((u:any)=>(
                <tr key={u.email} style={{borderBottom:`1px solid ${T.border}22`}}>
                  <td style={{padding:"8px 6px",color:T.text,fontSize:"13px"}}>{u.email}</td>
                  <td style={{padding:"8px 6px"}}>
                    <span style={{padding:"2px 8px",borderRadius:"10px",fontSize:"11px",background:u.role==="admin"?`${T.gold}22`:`${T.green}22`,color:u.role==="admin"?T.gold:T.green}}>{u.role}</span>
                  </td>
                  <td style={{padding:"8px 6px",textAlign:"right"}}>
                    {u.email!==authSession?.user?.email&&(
                      <button onClick={async()=>{
                          if(!window.confirm(`Rimuovere accesso per ${u.email}?`)) return;
                          await removeUser(u.email);
                          setUserList(await listUsers());
                          showToast(`${u.email} rimosso`, T.red);
                        }}
                        style={{background:"none",border:"none",color:T.red,cursor:"pointer",fontSize:"12px"}}>✕ Rimuovi</button>
                    )}
                  </td>
                </tr>
              ))}</tbody>
            </table>
          </div>
        </div>
      )}
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
  const[search,setSearchRaw]=useState(()=>psGet(`pg_${branch}_xref_search`,""));
  const setSearch=(v:string)=>{setSearchRaw(v);psSet(`pg_${branch}_xref_search`,v);};

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
          ? ["n comit","ncomit","comit","canarie","can no","can n","n°","numero comit","codice comit","cod comit","codcan","n_comit","no_"]
          : ["n hk","nhk","hk","n_hk","gc code","gc no","hk code","hk no","hong kong"];
        // Alias per colonna IFB — "no_" e "code" tolti (troppo generici, matchano N COMIT)
        const ifbA=["ifb n","ifb no","ifb no.","ifb item","bv no","bv n","no_ifb","ifb","item no"];
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
    setXrefs(next);CLOUD.set(`ifb_xrefs_${branch}`,next);setDataSource(`xref_${branch}`,"manual");
    const log={id,type:"xref",fileName,date:new Date(id).toISOString(),count:incoming.length,diffs,branch};
    const newLogs=[log,...importLogs];setImportLogs(newLogs);LS.set("ifb_importlogs",newLogs);
    const newSnaps=[log,...snapshots].slice(0,50);setSnapshots(newSnaps);LS.set("ifb_snapshots",newSnaps);
    bumpImportTs();showToast(`XRef: ${incoming.length} voci · ${diffs.filter(d=>d.isNew).length} nuove ✓`,T.gold);
    setStep("main");setPreview([]);setRawRows([]);setHeaders([]);
  }

  const displayed=xrefs.filter(x=>!search||x.nHK?.toLowerCase().includes(search.toLowerCase())||x.ifbNo?.toLowerCase().includes(search.toLowerCase()));

  return(
    <div>
      <PageHeader title={`⇄ XRef ${branchCode} / IFB N · ${branch}`} sub="Codici filiale ↔ IFB N — ogni filiale ha la propria tabella" srcKey={`xref_${branch}`}/>
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
              <button onClick={()=>{if(window.confirm(`Eliminare tutte le ${xrefs.length} XRef di ${branch}?`)){setXrefs([]);CLOUD.set(`ifb_xrefs_${branch}`,[]);}}}
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
                    <TD><MiniBtn label="✕" onClick={()=>{const n=xrefs.filter((_,j)=>j!==xrefs.indexOf(x));setXrefs(n);CLOUD.set(`ifb_xrefs_${branch}`,n);}} color={T.red}/></TD>
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
    { icon:"⚡", label:"Eccezioni Prezzi", color:T.text,
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
        <span style={{flex:1,fontSize:"13px",fontWeight:"bold",color:color}}>{label}</span>
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
    <div>
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

        // Mappa il codice filiale (N HK / N COMIT) → "code"
        const branchCodeAliases = ["n hk", "nhk", "n comit", "ncomit", "comit", "n canarie", "ncanarie", "n. comit"];
        // Mappa il codice IFB → "ifbCode" (secondo campo opzionale per fallback lookup)
        const ifbCodeAliases = ["ifb item", "ifb no", "ifb n", "no_", "item no.", "codice ifb", "ifb code"];

        for(const h of hdrs) {
          const hl = h.toLowerCase().trim();
          if(branchCodeAliases.some(a => hl === a || hl.includes(a))) {
            am["code"] = h;
            break;
          }
        }
        // Sempre cercare anche il codice IFB come campo separato
        for(const h of hdrs) {
          const hl = h.toLowerCase().trim();
          if(ifbCodeAliases.some(a => hl === a || hl.includes(a))) {
            am["ifbCode"] = h;
            break;
          }
        }
        // Fallback: se non trovato né N filiale né IFB separato, usa IFB come "code"
        if(!am["code"] && !am["ifbCode"]) {
          const genericAliases = ["no.", "no", "code", "codice"];
          for(const h of hdrs) {
            const hl = h.toLowerCase().trim();
            if(genericAliases.some(a => hl === a || hl.includes(a))) {
              am["code"] = h;
              break;
            }
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
        const priceFields = ["mtsPrice", "fcaPrice", "fcaDiscount", "fcaDiscounted", "dapPrice", "dapDiscount", "dapDiscounted", "dapFinalDirect", "carriageCost"];
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
      const rawIfbCode = String(get(row, "ifbCode") || "").trim();
      const rawDescription = String(get(row, "description") || get(row, "code") || "").trim();

      // Almeno uno dei due codici deve essere presente e valido
      if(!rawCode && !rawIfbCode) { skipped++; return null; }
      if(rawCode && !isValidCode(rawCode) && rawIfbCode && !isValidCode(rawIfbCode)) { skipped++; return null; }

      const prod = (rawCode && isValidCode(rawCode) && findProduct(rawCode, products, xrefs))
                || (rawIfbCode && isValidCode(rawIfbCode) && findProduct(rawIfbCode, products, xrefs))
                || null;
      
      const mtsPrice = parseFloat(get(row, "mtsPrice")) || 0;
      const fcaPrice = parseFloat(get(row, "fcaPrice")) || 0;
      const fcaDiscount = parseFloat(get(row, "fcaDiscount")) || 0;
      const fcaDiscounted = parseFloat(get(row, "fcaDiscounted")) || (fcaPrice - (fcaDiscount * fcaPrice / 100)) || 0;
      const dapPrice = parseFloat(get(row, "dapPrice")) || 0;
      const dapDiscount = parseFloat(get(row, "dapDiscount")) || 0;
      const dapDiscounted = parseFloat(get(row, "dapDiscounted")) || (dapPrice - (dapDiscount * dapPrice / 100)) || 0;
      const dapFinalDirect = parseFloat(get(row, "dapFinalDirect")) || 0;
      const carriageCost = parseFloat(get(row, "carriageCost")) || 0;

      let dapFinal = 0;
      let dapNote = "";
      if(dapFinalDirect !== 0) {
        dapFinal = dapFinalDirect;
        dapNote = "da file";
      } else if(dapDiscounted !== 0) {
        dapFinal = dapDiscounted;
        dapNote = "da DAP Disc.";
      } else if(carriageCost > 0 && (fcaDiscounted > 0 || fcaPrice > 0)) {
        dapFinal = (fcaDiscounted || fcaPrice) + carriageCost;
        dapNote = "FCA+Carriage";
      }
      
      const existing = prod ? prices.find(p => p.productId === prod.id && p.branch === branch && p.month === importMonth) : null;
      
      const displayCode = rawCode || rawIfbCode;
      return {
        _idx: idx,
        rawCode: displayCode,
        ifbNo_from_file: displayCode,
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
        carriageCost: roundN(carriageCost),
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
        fcaPrice: r.fcaPrice,
        carriageCost: r.carriageCost||0
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
    CLOUD.set(`ifb_prices_${branch}`, updated);
    setListiniMode("excel");
    
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
              <label style={{display:"block", fontSize:"11px", color:T.gold, marginBottom:"5px"}}>📌 N {branchN(branch)} (N COMIT / N HK)</label>
              <select
                value={mapping["code"] || ""}
                onChange={e => setMapping(m => ({...m, code: e.target.value || null}))}
                style={{...inputStyle(), cursor:"pointer", borderColor:(!mapping["code"] && !mapping["ifbCode"]) ? T.red+"88" : T.border}}
              >
                <option value="">— non mappato —</option>
                {headers.map(h => <option key={h} value={h}>{h}</option>)}
              </select>
            </div>

            <div>
              <label style={{display:"block", fontSize:"11px", color:T.gold, marginBottom:"5px"}}>📌 IFB Item (codice IFB)</label>
              <select
                value={mapping["ifbCode"] || ""}
                onChange={e => setMapping(m => ({...m, ifbCode: e.target.value || null}))}
                style={{...inputStyle(), cursor:"pointer", borderColor:(!mapping["code"] && !mapping["ifbCode"]) ? T.red+"88" : T.border}}
              >
                <option value="">— non mappato —</option>
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
            ⚡ Mappa almeno uno tra N {branchN(branch)} e IFB Item. Se entrambi mappati, il lookup usa prima N {branchN(branch)}, poi IFB Item come fallback. I campi prezzi vengono rilevati automaticamente.
          </div>

          <div style={{display:"flex", gap:"10px", marginTop:"16px"}}>
            <ActionBtn label="← Ricarica" onClick={reset}/>
            <ActionBtn label="Preview →" onClick={buildPreview} primary disabled={!mapping["code"] && !mapping["ifbCode"]}/>
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
    // Dedup per id (stesso articolo su più righe → ultima occorrenza vince)
    const dedupMap=new Map<string,any>(); newProds.forEach(p=>dedupMap.set(p.id,p)); const dedupedProds=[...dedupMap.values()];
    const prevMap=Object.fromEntries(products.map(p=>[p.id,p]));
    const diffs=[];
    for(const p of dedupedProds){
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
    const snap={id:now,type:"anagrafica",date:new Date(now).toISOString(),count:dedupedProds.length,diffs,products:dedupedProds,branch};
    const newSnaps=[snap,...snapshots].slice(0,50);setSnapshots(newSnaps);LS.set("ifb_snapshots",newSnaps);
    setProducts(dedupedProds);
    const savedProd = LS.set(`ifb_products_${branch}`, dedupedProds);
    if (!savedProd) showToast("⚠ LocalStorage piena: anagrafica NON salvata. Esporta i dati.", T.red);
    const log={id:now,type:"anagrafica",date:new Date(now).toISOString(),msg:`Importati ${dedupedProds.length} articoli`};
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
      <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:"12px",marginBottom:"20px"}}>
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
  const[search,setSearchRaw]=useState(()=>psGet(`pg_${branch}_airlist_search`,""));
  const setSearch=(v:string)=>{setSearchRaw(v);psSet(`pg_${branch}_airlist_search`,v);};

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
    setAirList(next); CLOUD.set(`ifb_airlist_${branch}`, next);
    const now = Date.now();
    IDB.set(`ifb_air_data_${now}`, next);
    const log = {id:now,type:"air",date:new Date(now).toISOString(),count:valid.length,diffs:[],branch};
    const newLogs = [log,...importLogs]; setImportLogs(newLogs); LS.set("ifb_importlogs",newLogs);
    const newSnaps = [log,...snapshots].slice(0,50); setSnapshots(newSnaps); LS.set("ifb_snapshots",newSnaps);
    bumpImportTs(); showToast(`AIR ${branch}: lista sostituita con ${valid.length} articoli ✓`, T.gold);
    setStep("main"); setPreview([]); setRawRows([]); setHeaders([]);
  }

  const branchAir = airList;

  // Item classificati AIR automaticamente da BC (bcTransportation contiene "air")
  const bcAirItems = useMemo(()=>
    products
      .filter((p:any)=>isAirTransport(p.bcTransportation))
      .map((p:any)=>({
        productId: p.id, code: p.code, nHK: p.nHK,
        description: p.description, transportation: p.bcTransportation, _fromBC: true,
      }))
  , [products]);
  const bcAirCount = bcAirItems.length;

  // Lista unificata: BC auto + manuali (senza duplicati)
  const manualOnlyAir = branchAir.filter((a:any)=>
    !bcAirItems.some((b:any)=>b.productId===a.productId||(b.code&&b.code===a.code)||(b.nHK&&b.nHK===a.nHK))
  );
  const allAirItems = [...bcAirItems, ...manualOnlyAir];

  const _sq=search.toLowerCase();
  const displayed=allAirItems.filter((a:any)=>!search
    ||a.description?.toLowerCase().includes(_sq)
    ||a.code?.toLowerCase().includes(_sq)
    ||a.nHK?.toLowerCase().includes(_sq));

  return(
    <div>
      <PageHeader title="✈ AIR Transport" sub="Articoli trasportati via aerea — esclusi da Standard Cost (calcolo solo SEA)"/>
      <BcBanner title="Classificazione automatica da BC Brightview">
        Gli articoli con campo <b style={{color:T.text}}>Transportation</b> impostato a <b style={{color:T.orange}}>CH AIR</b>, <b style={{color:T.orange}}>DRY AIR</b> o <b style={{color:T.orange}}>FR AIR</b> nell'item card di BC Brightview vengono classificati automaticamente come AIR ({bcAirCount} articoli da BC).
        La lista manuale qui sotto integra o sovrascrive per gli articoli non presenti in BC.
      </BcBanner>

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
  Lista AIR ({allAirItems.length})
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
          {allAirItems.length>0&&(
            <>
              <SearchBar value={search} onChange={setSearch} placeholder="🔍 Cerca articolo AIR…"/>
              <Section title={`${displayed.length} articoli AIR · ${bcAirCount} da BC · ${manualOnlyAir.length} manuali`}>
                <table style={{width:"100%",borderCollapse:"collapse"}}>
                  <THead cols={["Codice",branchN(branch),"Descrizione","Sorgente","Azioni"]}sticky/>
                  <tbody>{displayed.map((a:any,i:number)=>(
                    <tr key={a.productId||i} style={{borderBottom:`1px solid ${T.border}`}}>
                      <TD mono><span style={{color:T.gold}}>{a.code}</span></TD>
                      <TD mono><span style={{color:T.muted}}>{a.nHK||"—"}</span></TD>
                      <TD>{a.description}</TD>
                      <TD>
                        {a._fromBC
                          ? <Chip label={`BC: ${a.transportation}`} color={T.blue}/>
                          : <Chip label="Manuale" color={T.orange}/>}
                      </TD>
                      <TD>{!a._fromBC&&<MiniBtn label="✕ Rimuovi" onClick={()=>{const n=airList.filter((_,j)=>j!==airList.indexOf(a));setAirList(n);IDB.set(`ifb_airlist_${branch}`,n);}} color={T.red}/>}</TD>
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

  const isCAN = branch === "CAN";
  const costKey = isCAN ? "step2GC" : "step2Hkd";
  const costLabel = isCAN ? "New SC GC (€)" : "Step2 HKD";
  const prevLabel = isCAN ? "SC prec. GC (€)" : "Prec. HKD";

  const calcOk   = costRows.filter((r:any)=>r.cost?.[costKey]!=null);
  const flagged  = costRows.filter((r:any)=>r.cost?.[costKey]!=null&&r.prevCost?.[costKey]!=null&&r.prevCost[costKey]>0&&Math.abs((r.cost[costKey]-r.prevCost[costKey])/r.prevCost[costKey])>=0.03);
  const air      = costRows.filter((r:any)=>r.isAir);
  const noPrice  = costRows.filter((r:any)=>!r.cost&&!r.isAir&&(r.skipReason?.includes("NO PREZZO")||r.skipReason?.includes("LISTINO CHIUSO")));
  const noLog    = costRows.filter((r:any)=>!r.cost&&!r.isAir&&r.skipReason==="NO LOGISTICA");
  const calcZero = costRows.filter((r:any)=>!r.cost&&!r.isAir&&r.skipReason?.includes("CALC=0"));

  const STATS = [
    { id:"ok",      n:calcOk.length,   label:"Costi calcolati",       color:T.green,  rows:calcOk   },
    { id:"flagged", n:flagged.length,  label:"Variazioni ≥3%",        color:T.orange, rows:flagged  },
    ...(!isCAN ? [{ id:"air", n:air.length, label:"AIR (esclusi)", color:T.blue, rows:air }] : []),
    { id:"noPrice", n:noPrice.length,  label:"Senza prezzo",          color:T.red,    rows:noPrice  },
    { id:"noLog",   n:noLog.length,    label:"No logistica",          color:T.red,    rows:noLog    },
    { id:"calc0",   n:calcZero.length, label:"Calc=0 (UOM/qty)",      color:T.orange, rows:calcZero },
  ];

  const panel = STATS.find(s=>s.id===activePanel);

  function renderPanel() {
    if(!panel||panel.rows.length===0)
      return <div style={{padding:"24px",textAlign:"center",color:T.dim,fontSize:"13px"}}>Nessun articolo in questa categoria.</div>;

    if(activePanel==="ok"||activePanel==="flagged") return (
      <table style={{width:"100%",borderCollapse:"collapse"}}>
        <THead cols={[branchN(branch),"IFB No","Descrizione","Ubicaz.",isCAN?"Area":"",costLabel,prevLabel,"Δ%"]}sticky/>
        <tbody>{panel.rows.map((r:any,i:number)=>{
          const cv = r.cost?.[costKey];
          const pv = r.prevCost?.[costKey];
          const pct = cv!=null&&pv!=null&&pv>0?(cv-pv)/pv*100:null;
          return(
            <tr key={r.id} style={{borderBottom:`1px solid ${T.border}`,background:i%2===0?T.bg:T.surface}}>
              <TD mono><span style={{color:T.muted}}>{r.nHK||"—"}</span></TD>
              <TD mono><span style={{color:T.gold}}>{r.code}</span></TD>
              <TD>{r.description}</TD>
              <TD><Chip label={r.ubicazione||"—"} color={r.ubicazione==="FOR"?T.purple:r.ubicazione==="MTS"?T.blue:T.green}/></TD>
              {isCAN ? <TD><span style={{color:T.muted,fontSize:"11px"}}>{r.area||"NORD"}</span></TD> : <TD/>}
              <TD mono><span style={{color:T.gold,fontWeight:"bold"}}>{cv!=null?cv.toFixed(2):"—"}</span></TD>
              <TD mono><span style={{color:T.muted}}>{pv!=null?pv.toFixed(2):"—"}</span></TD>
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
  const[search,setSearchRaw]=useState(()=>psGet(`pg_${branch}_logistics_search`,""));
  const setSearch=(v:string)=>{setSearchRaw(v);psSet(`pg_${branch}_logistics_search`,v);};
  const[showOnlyMissing,setShowOnlyMissing]=useState(initFilter==="missing");
  const[mapStep,setMapStep]=useState("idle");
  const[logHeaders,setLogHeaders]=useState([]);
  const[logRawRows,setLogRawRows]=useState([]);
  const[colIdx,setColIdx]=useState({});

  const allIFBProducts = products.filter(p=>isIFBVendor(p.vendorName));

  const [editingRows, setEditingRows] = useState<Set<string>>(new Set());
  function toggleEdit(id:string) { setEditingRows(prev=>{ const s=new Set(prev); s.has(id)?s.delete(id):s.add(id); return s; }); }
  function saveRow(id:string) {
    const existing = getLog(id);
    let next = logistics;
    if(!existing) {
      // Entry never existed: create it with current defaults
      next = [...logistics, getOrDefault(id)];
      setLogistics(next);
    }
    CLOUD.set("ifb_logistics", next);
    showToast("Salvato ✓", T.green);
    setEditingRows(prev=>{ const s=new Set(prev); s.delete(id); return s; });
  }

  function getLog(productId) {
    return logistics.find(l=>l.productId===productId && l.branch===branch) || null;
  }

  function getOrDefault(productId) {
    return getLog(productId) || {productId, branch, area:"NORD", ubicazione:"MTO", pltPerContainer:20, hasCert:false, hasAlcTax:false, alcTax:0, convFactor:1, carriage:0};
  }
  
  function update(productId, field, rawVal) {
    const existing = getLog(productId);
    const val = ["ubicazione","area","transport"].includes(field) ? rawVal :
                ["hasCert","hasAlcTax"].includes(field) ? (rawVal === "true") :
                (parseFloat(rawVal) || 0);
    const next = existing
      ? logistics.map(l => l.productId===productId&&l.branch===branch ? {...l, [field]:val, _manualOverride:true} : l)
      : [...logistics, {...getOrDefault(productId), [field]: val, _manualOverride:true}];
    setLogistics(next);
    CLOUD.set("ifb_logistics", next);
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
          iGC:  fi(["gc","gran canaria","grancanaria"]),
          iTF:  fi(["tf","tenerife","fuerteventura tf"]),
          iFUE: fi(["fue","fuerteventura","fuerte"]),
          iLAN: fi(["lan","lanzarote"]),
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
    const { iNHK, iIFB, iUb, iArea, iPlt, iCert, iTemp, iCarriage, iAirSea, iTransport, iAlcTax, iGC, iTF, iFUE, iLAN } = colIdx;
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
  
      // Isole destinazione CAN (GC / TF / FUE / LAN) — colonna con "1" o "x" = true
      const parseIsland = (idx:number) => {
        if(idx < 0) return undefined;
        const v = String(row[idx]||"").trim().toLowerCase();
        return v==="1"||v==="x"||v==="si"||v==="sì"||v==="yes"||v==="true";
      };
      const isGC_v  = parseIsland(iGC);
      const isTF_v  = parseIsland(iTF);
      const isFUE_v = parseIsland(iFUE);
      const isLAN_v = parseIsland(iLAN);

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
        fromImport: true,
        ...(finalTransport ? { transport: finalTransport } : {}),
        ...(isGC_v  !== undefined ? { isGC:  isGC_v  } : {}),
        ...(isTF_v  !== undefined ? { isTF:  isTF_v  } : {}),
        ...(isFUE_v !== undefined ? { isFUE: isFUE_v } : {}),
        ...(isLAN_v !== undefined ? { isLAN: isLAN_v } : {}),
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
    CLOUD.set("ifb_logistics", next);
  
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
      <PageHeader title={`Logistica · ${branch}`} sub={`${withCount} con logistica · ${missingCount} senza logistica — totale ${allIFBProducts.length} IFB`}/>

      <div style={{fontSize:"11px", color:T.muted, marginBottom:"10px", padding:"6px 10px", background:`${T.gold}08`, borderRadius:"6px", border:`1px solid ${T.gold}22`}}>
        🟡 Righe <strong style={{color:T.gold}}>dorate</strong> = importate da Work_tab (modificabili) &nbsp;·&nbsp;
        🟠 Righe arancioni = senza logistica
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
          CLOUD.set("ifb_logistics", newLogistics);
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
      Verranno importati per <strong style={{color:T.gold}}>{branch}</strong>: Ubicazione, Area, Plt/Container, Health Certificate, Carriage, Tassa Alcolica{branch==="CAN" ? ", Isole (GC/TF/FUE/LAN)" : ""}
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


      {displayed.length === 0 && showOnlyMissing && (
        <div style={{padding:"32px", textAlign:"center", background:`${T.green}11`, borderRadius:"8px", color:T.green, fontSize:"13px"}}>
          ✅ PERFETTO! Tutti i {allIFBProducts.length} prodotti IFB hanno parametri logistici per {branch}!
        </div>
      )}

      <div style={{marginBottom:"12px"}}>
        <button
          onClick={()=>{
            CLOUD.set("ifb_logistics", logistics);
            showToast("Salvato ✓", T.green);
          }}
          style={{padding:"8px 18px",background:`${T.green}20`,border:`1px solid ${T.green}66`,borderRadius:"6px",color:T.green,cursor:"pointer",fontSize:"12px",fontWeight:"bold"}}>
          💾 Salva dati logistici
        </button>
      </div>

      {displayed.length > 0 && (
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%", borderCollapse:"collapse", fontSize:"12px"}}>
            <thead>
              <tr>
              {["IFB No",branchN(branch),"Descrizione","Ubicaz.","Area",...(branch==="CAN"?["Trasporto","GC","TF","FUE","LAN"]:[]),...(branch!=="CAN"?["Cert."]:[]),"Alcol >30°",...(branch!=="CAN"?["Carriage"]:[]),"Conv.",""].map(c=>(
                <th key={c} style={{padding:"7px 12px",background:T.card,color:T.muted,textAlign:"left",borderBottom:`1px solid ${T.border}`,fontSize:"11px",fontWeight:"normal",position:"sticky",top:0,zIndex:10}}>{c}</th>
              ))}
              </tr>
            </thead>
            <tbody key={_sq}>
              {displayed.map((prod, i) => {
                const l = getOrDefault(prod.id);
                const hasEntry = !!getLog(prod.id);
                const isEditing = !hasEntry || editingRows.has(prod.id);
                const btnS = (c:string):React.CSSProperties => ({padding:"3px 9px",fontSize:"11px",background:`${c}18`,border:`1px solid ${c}55`,borderRadius:"5px",color:c,cursor:"pointer",fontWeight:"bold"});
                return (
                  <tr key={prod.id} style={{borderBottom:`1px solid ${T.border}`, background:!hasEntry ? `${T.orange}08` : (i%2===0 ? T.bg : T.surface)}}>
                    <td style={{padding:"7px 12px", fontSize:"12px", fontFamily:"monospace"}}><span style={{color:T.gold}}>{prod.code}</span></td>
                    <td style={{padding:"7px 12px", fontSize:"12px", fontFamily:"monospace"}}><span style={{color:T.muted}}>{prod.nHK||"—"}</span></td>
                    <td style={{padding:"7px 12px", fontSize:"12px"}}>
                      {prod.description}
                      {!hasEntry && <span style={{marginLeft:"6px", fontSize:"9px", color:T.orange, fontWeight:"bold"}}>⚠ MANCANTE</span>}
                    </td>
                    {!isEditing ? (
                      <>
                        <td style={{padding:"7px 12px"}}><Chip label={l.ubicazione||"—"} color={l.ubicazione==="FOR"?T.purple:l.ubicazione==="MTS"?T.blue:T.green}/></td>
                        <td style={{padding:"7px 12px", fontSize:"12px", color:T.muted}}>{l.area||"—"}</td>
                        {branch==="CAN"&&<td style={{padding:"7px 12px"}}><Chip label={l.transport||"GOMMA"} color={l.transport==="MARE"?T.blue:T.muted}/></td>}
                        {branch==="CAN"&&<>
                          {(["isGC","isTF","isFUE","isLAN"] as const).map(k=>(
                            <td key={k} style={{padding:"4px 8px",textAlign:"center",fontSize:"13px"}}>{l[k]?"✓":"—"}</td>
                          ))}
                        </>}
                        {branch!=="CAN"&&<td style={{padding:"7px 12px", fontSize:"12px", color:T.muted}}>{l.hasCert?"Sì":"No"}</td>}
                        <td style={{padding:"7px 12px", fontSize:"12px", color:T.muted}}>{l.hasAlcTax?"Sì":"No"}</td>
                        {branch!=="CAN"&&<td style={{padding:"7px 12px", fontSize:"12px", fontFamily:"monospace", color:T.muted}}>{l.carriage||0}</td>}
                        <td style={{padding:"7px 12px", fontSize:"12px", fontFamily:"monospace", color:T.dim}}>{l.convFactor||1}</td>
                        <td style={{padding:"7px 12px"}}><button style={btnS(T.gold)} onClick={()=>toggleEdit(prod.id)}>✏️ Modifica</button></td>
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
                        {branch==="CAN"&&<td style={{padding:"7px 12px"}}>
                          <select value={l.transport||"GOMMA"} onChange={e=>update(prod.id,"transport",e.target.value)}
                            style={{background:T.card,color:T.blue,border:`1px solid ${T.border}`,borderRadius:"4px",padding:"3px 6px",fontSize:"11px",width:"80px"}}>
                            {["GOMMA","MARE"].map(v=><option key={v} value={v}>{v}</option>)}
                          </select>
                        </td>}
                        {branch==="CAN"&&(["isGC","isTF","isFUE","isLAN"] as const).map(k=>(
                          <td key={k} style={{padding:"4px 8px",textAlign:"center"}}>
                            <input type="checkbox" checked={!!l[k]}
                              onChange={e=>update(prod.id,k,e.target.checked)}
                              style={{width:"16px",height:"16px",cursor:"pointer",accentColor:T.gold}}/>
                          </td>
                        ))}
                        {branch!=="CAN"&&<td style={{padding:"7px 12px"}}>
                          <select value={String(l.hasCert||false)} onChange={e=>update(prod.id,"hasCert",e.target.value)}
                            style={{background:T.card,color:T.text,border:`1px solid ${T.border}`,borderRadius:"4px",padding:"3px 6px",fontSize:"11px",width:"60px"}}>
                            <option value="false">No</option><option value="true">Sì</option>
                          </select>
                        </td>}
                        <td style={{padding:"7px 12px"}}>
                          <select value={String(l.hasAlcTax||false)} onChange={e=>update(prod.id,"hasAlcTax",e.target.value)}
                            style={{background:T.card,color:T.text,border:`1px solid ${T.border}`,borderRadius:"4px",padding:"3px 6px",fontSize:"11px",width:"60px"}}>
                            <option value="false">No</option><option value="true">Sì</option>
                          </select>
                        </td>
                        {branch!=="CAN"&&<td style={{padding:"7px 12px"}}>
                          <input type="number" defaultValue={l.carriage||0}
                            onBlur={e=>update(prod.id,"carriage",e.target.value)}
                            style={{width:"55px",background:"transparent",color:T.gold,border:"none",textAlign:"right",fontSize:"12px",borderBottom:`1px solid ${T.border}`}}/>
                        </td>}
                        <td style={{padding:"7px 12px"}}>
                          <input type="number" defaultValue={l.convFactor||1} step="0.01"
                            onBlur={e=>update(prod.id,"convFactor",e.target.value)}
                            style={{width:"50px",background:"transparent",color:T.muted,border:"none",textAlign:"right",fontSize:"11px",borderBottom:`1px solid ${T.border}`}}/>
                        </td>
                        <td style={{padding:"7px 12px",whiteSpace:"nowrap"}}>
                          <span style={{display:"flex",gap:"5px"}}>
                            <button style={btnS(T.green)} onClick={()=>saveRow(prod.id)}>💾 Salva</button>
                            {hasEntry&&<button style={btnS(T.muted)} onClick={()=>toggleEdit(prod.id)}>Annulla</button>}
                          </span>
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

// ─── PRICE COMPARE (debug temporaneo BC vs Excel) ────────────────────────────
function PriceComparePage({ bcListini, prices, products, xrefs, branch, month }: any) {
  const FIELDS = [
    { key: "fcaPrice",      label: "FCA Price",   bc: "fcaPrice",      xl: "fcaPrice" },
    { key: "fcaDiscounted", label: "FCA Disc.",    bc: "fcaDiscounted", xl: "fcaDiscounted" },
    { key: "dapPrice",      label: "DAP Price",   bc: "dapPrice",      xl: "dapPrice" },
    { key: "dapFinal",      label: "DAP Final",   bc: "dapFinal",      xl: "dapFinal" },
    { key: "mtsPrice",      label: "MTS Price",   bc: "mtsPrice",      xl: "mtsPrice" },
  ];
  const [filter, setFilterRaw] = useState<"all"|"diff"|"mts"|"real">(()=>psGet(`pg_${branch}_pricecmp_filter`,"real"));
  const setFilter=(v:"all"|"diff"|"mts"|"real")=>{setFilterRaw(v);psSet(`pg_${branch}_pricecmp_filter`,v);};
  const [search, setSearchRaw] = useState(()=>psGet(`pg_${branch}_pricecmp_search`,""));
  const setSearch=(v:string)=>{setSearchRaw(v);psSet(`pg_${branch}_pricecmp_search`,v);};
  const [reasonFilter, setReasonFilterRaw] = useState<string>(()=>psGet(`pg_${branch}_pricecmp_reason`,"all"));
  const setReasonFilter=(v:string)=>{setReasonFilterRaw(v);psSet(`pg_${branch}_pricecmp_reason`,v);};
  const [hideAbsentXl, setHideAbsentXl] = useState(true);

  const xlByProductId = useMemo(() => {
    const m: Record<string, any> = {};
    prices.filter((p: any) => p.branch === branch && p.month === month).forEach((p: any) => { m[String(p.productId)] = p; });
    return m;
  }, [prices, branch, month]);

  const bcByProductId = useMemo(() => {
    const m: Record<string, any> = {};
    bcListini.filter((p: any) => (p.branch || p.b) === branch).forEach((p: any) => { m[String(p.productId)] = p; });
    return m;
  }, [bcListini, branch]);

  const bcByItemCode = useMemo(() => {
    const m: Record<string, any> = {};
    bcListini.filter((p: any) => (p.branch || p.b) === branch).forEach((p: any) => { if(p.itemCode) m[String(p.itemCode)] = p; });
    return m;
  }, [bcListini, branch]);

  const allProductIds = useMemo(() => {
    const s = new Set<string>([...Object.keys(xlByProductId)]);
    Object.keys(bcByProductId).forEach(k => { if(!k.startsWith("BC_")) s.add(k); });
    return [...s];
  }, [xlByProductId, bcByProductId]);

  const rows = useMemo(() => {
    return allProductIds.map(pid => {
      const xl = xlByProductId[pid];
      const prod = products.find((p: any) => String(p.nHK) === pid) || products.find((p: any) => String(p.id) === pid);
      const bc = bcByProductId[pid] || (prod?.code ? bcByItemCode[prod.code] : null) || bcByItemCode[pid];

      const qpb = Number(prod?.qtyPerBox || prod?.pcsPerBox || 1) || 1;
      const bcPu = String(bc?.pu || bc?.purchaseUom || "").toUpperCase();
      const bcCf2 = Number(bc?.cf2 || 1) || 1;
      const bcRawUom = String(bc?.pur || bcPu).toUpperCase();
      const uf: Record<string, number> = bc?.uf || {};
      const diffs: { field: string; label: string; bc: number; xl: number; bcNorm: number; delta: number; reason: string; uomNote?: string; bcUom: string; xlUom: string }[] = [];

      FIELDS.forEach(f => {
        const bcVal = Number(bc?.[f.bc] || 0);
        const xlVal = Number(xl?.[f.xl] || 0);
        if (bcVal === 0 && xlVal === 0) return;

        let reason = "";
        let uomNote: string | undefined;
        let bcNorm = bcVal;
        let delta = xlVal - bcVal;
        let bcUom = bcPu || "?";
        // UoM Excel: dall'anagrafica prodotto (hkUom), fallback su xl o prod.uom
        const xlUomRaw = String(prod?.hkUom || xl?.hkUom || xl?.uom || prod?.uom || "").toUpperCase().trim();
        let xlUom = xlUomRaw || "?";

        if (bcVal === 0 && xlVal > 0) {
          reason = "🟡 assente in BC";
        } else if (bcVal > 0 && xlVal === 0) {
          reason = "🔴 assente in Excel";
        } else {
          const UOM_TOL = 0.04;

          // 1. Confronto diretto (stesso UoM o prezzi già allineati)
          const rawPct = Math.abs((xlVal - bcVal) / bcVal);
          if (rawPct < UOM_TOL) {
            if (Math.abs(xlVal - bcVal) < 0.01) return;
            reason = "≈ diff < 4%";
            bcNorm = bcVal;
            delta = xlVal - bcVal;
          } else if (bcPu && xlUom && bcPu !== xlUom && Object.keys(uf).length > 0 && uf[bcPu] != null && uf[xlUom] != null) {
            // 2. UoM diversa: converti usando tabella UoM (uf) — valido per KG, LT, BOX, PCS, ecc.
            // formula: prezzo_per_xlUom = prezzo_per_bcPu * uf[xlUom] / uf[bcPu]
            const puFactor = uf[bcPu];
            const xlFactor = uf[xlUom];
            const factor = xlFactor / puFactor;
            const bcConverted = bcVal * factor;
            const pctConverted = Math.abs((xlVal - bcConverted) / bcConverted);
            if (pctConverted < UOM_TOL) {
              if (pctConverted < 0.01) return;
              bcNorm = bcConverted;
              delta = xlVal - bcConverted;
              reason = `📦 UoM: BC per-${bcPu}, Excel per-${xlUom}`;
              uomNote = `BC €${bcVal.toFixed(2)}×${factor.toFixed(4)}=€${bcConverted.toFixed(2)} ≈ Excel €${xlVal.toFixed(2)}`;
              bcUom = bcPu;
            } else {
              bcNorm = bcConverted;
              delta = xlVal - bcConverted;
              const pct = delta / bcConverted;
              if (Math.abs(pct) < 0.025) reason = "≈ diff < 2.5%";
              else if (delta > 0) reason = `📈 Excel +${(pct*100).toFixed(1)}%`;
              else reason = `📉 Excel ${(pct*100).toFixed(1)}%`;
              uomNote = `BC €${bcVal.toFixed(2)} per ${bcPu} → €${bcConverted.toFixed(2)} per ${xlUom}`;
            }
          } else if (qpb > 1) {
            // 3. Fallback qpb: BC per-PCS × qpb ≈ xl (Excel per-BOX)
            const bcBoxed = bcVal * qpb;
            const pctBoxed = Math.abs((xlVal - bcBoxed) / bcBoxed);
            // 4. xl × qpb ≈ BC (BC per-BOX, Excel per-PCS)
            const xlBoxed = xlVal * qpb;
            const pctUnboxed = Math.abs((xlBoxed - bcVal) / bcVal);

            if (pctBoxed < UOM_TOL) {
              if (pctBoxed < 0.01) return;
              bcNorm = bcBoxed;
              delta = xlVal - bcBoxed;
              reason = "📦 UoM: BC per-PCS, Excel per-BOX";
              uomNote = `BC €${bcVal.toFixed(2)}×${qpb}=€${bcBoxed.toFixed(2)} ≈ Excel €${xlVal.toFixed(2)}`;
              bcUom = "PCS";
            } else if (pctUnboxed < UOM_TOL) {
              if (pctUnboxed < 0.01) return;
              bcNorm = bcVal;
              delta = xlBoxed - bcVal;
              reason = "📦 UoM: BC per-BOX, Excel per-PCS";
              uomNote = `Excel €${xlVal.toFixed(2)}×${qpb}=€${xlBoxed.toFixed(2)} ≈ BC €${bcVal.toFixed(2)}`;
              bcUom = "BOX";
            } else {
              bcNorm = bcVal;
              delta = xlVal - bcVal;
              const pct = delta / bcVal;
              if (Math.abs(pct) < 0.025) reason = "≈ diff < 2.5%";
              else if (delta > 0) reason = `📈 Excel +${(pct*100).toFixed(1)}%`;
              else reason = `📉 Excel ${(pct*100).toFixed(1)}%`;
            }
          } else {
            bcNorm = bcVal;
            delta = xlVal - bcVal;
            const pct = delta / bcVal;
            if (Math.abs(pct) < 0.025) reason = "≈ diff < 2.5%";
            else if (delta > 0) reason = `📈 Excel +${(pct*100).toFixed(1)}%`;
            else reason = `📉 Excel ${(pct*100).toFixed(1)}%`;
          }
        }
        const bcRaw = bcVal / bcCf2;
        diffs.push({ field: f.key, label: f.label, bc: bcVal, xl: xlVal, bcNorm, delta, reason, uomNote, bcUom, xlUom, bcRaw, bcRawUom });
      });

      const hasDiff = diffs.length > 0;
      const hasMtsDiff = diffs.some(d => d.field === "mtsPrice");
      return { pid, prod, xl, bc, diffs, hasDiff, hasMtsDiff, qpb };
    });
  }, [allProductIds, xlByProductId, bcByProductId, products]);

  // Pivot: campo × tipo-motivo → conteggio articoli
  const REASON_CATS = [
    { key: "missing_bc",   label: "🟡 assente BC",    match: (r: string) => r.includes("assente in BC") },
    { key: "missing_xl",   label: "🔴 assente Excel",  match: (r: string) => r.includes("assente in Excel") },
    { key: "uom",          label: "📦 UoM mismatch",  match: (r: string) => r.startsWith("📦") },
    { key: "diff_pct",     label: "📈📉 diff >2.5%",   match: (r: string) => r.startsWith("📈") || r.startsWith("📉") },
    { key: "approx",       label: "≈ diff <2.5%",      match: (r: string) => r.startsWith("≈") },
  ];
  const pivot = useMemo(() => {
    const p: Record<string, Record<string, number>> = {};
    FIELDS.forEach(f => { p[f.key] = {}; REASON_CATS.forEach(c => { p[f.key][c.key] = 0; }); });
    rows.forEach(row => row.diffs.forEach((d: any) => {
      const cat = REASON_CATS.find(c => c.match(d.reason));
      if (cat && p[d.field]) p[d.field][cat.key]++;
    }));
    return p;
  }, [rows]);

  const displayed = useMemo(() => {
    let r = rows;
    if (hideAbsentXl) r = r.map(row => ({ ...row, diffs: row.diffs.filter((d: any) => !d.reason.includes("assente in Excel")) })).filter(row => row.diffs.length > 0);
    if (filter === "real") r = r.map(row => ({ ...row, diffs: row.diffs.filter((d: any) => !d.reason.startsWith("≈") && !d.reason.startsWith("📦")) })).filter(row => row.diffs.length > 0);
    if (filter === "diff") r = r.filter(r => r.hasDiff);
    if (filter === "mts")  r = r.filter(r => r.hasMtsDiff);
    if (reasonFilter !== "all") {
      const cat = REASON_CATS.find(c => c.key === reasonFilter);
      if (cat) r = r.filter(r => r.diffs.some((d: any) => cat.match(d.reason)));
    }
    if (search) {
      const q = search.toLowerCase();
      r = r.filter(r => r.prod?.code?.toLowerCase().includes(q) || r.prod?.description?.toLowerCase().includes(q) || r.prod?.nHK?.toLowerCase().includes(q));
    }
    const maxAbsDelta = (row: any) => Math.max(...row.diffs.map((d: any) => Math.abs(d.delta)), 0);
    r = [...r].sort((a, b) => maxAbsDelta(b) - maxAbsDelta(a));
    return r;
  }, [rows, filter, reasonFilter, search]);

  const hasXl = Object.keys(xlByProductId).length > 0;
  const hasBc = Object.keys(bcByProductId).length > 0;

  return (
    <div>
      <PageHeader title={`🔬 Confronto Listini · ${branch}`} sub={`BC: ${Object.keys(bcByProductId).length} articoli · Excel ${month}: ${Object.keys(xlByProductId).length} articoli`} />
      <div style={{ background: `${T.gold}11`, border: `1px solid ${T.gold}33`, borderRadius: "8px", padding: "10px 14px", marginBottom: "14px", fontSize: "12px", color: T.muted }}>
        <b style={{ color: T.gold }}>Sezione temporanea di debug.</b> Confronta prezzi BC vs Excel {month}. I prezzi BC sono già convertiti alla base UoM.
        {!hasXl && <span style={{ color: T.orange }}> — nessun prezzo Excel per {branch}/{month}: vai in Listini e importa il pricelist.</span>}
        {!hasBc && <span style={{ color: T.orange }}> — nessun dato BC: ricarica i listini BC dalla pagina Listini.</span>}
      </div>

      {/* Pivot riepilogo */}
      {hasXl && hasBc && (
      <div style={{ marginBottom: "16px", overflowX: "auto" }}>
        <table style={{ borderCollapse: "collapse", fontSize: "11px", width: "100%" }}>
          <thead>
            <tr style={{ background: T.surface }}>
              <th style={{ padding: "6px 10px", textAlign: "left", color: T.muted, fontWeight: "normal", borderBottom: `1px solid ${T.border}` }}>Campo</th>
              {REASON_CATS.map(c => <th key={c.key} style={{ padding: "6px 14px", textAlign: "center", color: T.muted, fontWeight: "normal", borderBottom: `1px solid ${T.border}`, cursor: "pointer", whiteSpace: "nowrap" }} onClick={() => { setFilter("diff"); setReasonFilter(c.key); }}>{c.label}</th>)}
            </tr>
          </thead>
          <tbody>
            {FIELDS.map(f => (
              <tr key={f.key} style={{ borderBottom: `1px solid ${T.border}22` }}>
                <td style={{ padding: "5px 10px", color: T.dim, fontFamily: "monospace" }}>{f.label}</td>
                {REASON_CATS.map(c => {
                  const n = pivot[f.key]?.[c.key] || 0;
                  const isHot = c.key !== "approx" && n > 0;
                  return (
                    <td key={c.key} style={{ padding: "5px 14px", textAlign: "center", cursor: n > 0 ? "pointer" : "default" }}
                      onClick={() => { if (n > 0) { setFilter("diff"); setReasonFilter(c.key); }}}>
                      <span style={{ color: isHot ? (c.key === "missing_bc" ? T.orange : c.key === "missing_xl" ? T.red : T.gold) : T.dim, fontWeight: isHot ? "bold" : "normal" }}>
                        {n > 0 ? n : "—"}
                      </span>
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
        <div style={{ fontSize: "10px", color: T.dim, marginTop: "4px" }}>Clicca su un'intestazione o su un numero per filtrare</div>
      </div>
      )}

      <div style={{ display: "flex", gap: "8px", marginBottom: "14px", alignItems: "center", flexWrap: "wrap" }}>
        {([
          ["all",  `Tutti (${rows.length})`],
          ["real", `🔬 Problemi reali (${rows.filter(r => r.diffs.some((d:any) => !d.reason.startsWith("≈"))).length})`],
          ["mts",  `MTS diff (${rows.filter(r => r.hasMtsDiff).length})`],
        ] as [string,string][]).map(([f, l]) => (
          <button key={f} onClick={() => { setFilter(f as any); setReasonFilter("all"); }}
            style={{ padding: "5px 12px", background: filter === f && reasonFilter === "all" ? `${T.gold}22` : T.surface, border: `1px solid ${filter === f && reasonFilter === "all" ? T.gold : T.border}`, borderRadius: "6px", color: filter === f && reasonFilter === "all" ? T.gold : T.muted, cursor: "pointer", fontSize: "11px" }}>
            {l}
          </button>
        ))}
        <button onClick={() => setHideAbsentXl(v => !v)}
          style={{ padding: "5px 12px", background: hideAbsentXl ? `${T.red}22` : T.surface, border: `1px solid ${hideAbsentXl ? T.red : T.border}`, borderRadius: "6px", color: hideAbsentXl ? T.red : T.muted, cursor: "pointer", fontSize: "11px" }}>
          {hideAbsentXl ? "✕ Nascondi assenti Excel" : "Mostra assenti Excel"}
        </button>
        {reasonFilter !== "all" && (
          <button onClick={() => setReasonFilter("all")} style={{ padding: "5px 10px", background: `${T.orange}22`, border: `1px solid ${T.orange}`, borderRadius: "6px", color: T.orange, cursor: "pointer", fontSize: "11px" }}>
            ✕ {REASON_CATS.find(c => c.key === reasonFilter)?.label}
          </button>
        )}
        <SearchBar value={search} onChange={setSearch} placeholder="🔍 Cerca…" style={{ marginBottom: 0, maxWidth: "220px" }} />
      </div>

      {displayed.length === 0 ? (
        <div style={{ padding: "40px", textAlign: "center", color: T.muted, fontSize: "13px" }}>
          {!hasXl || !hasBc ? "Carica entrambe le sorgenti per confrontarle." : "Nessuna differenza trovata."}
        </div>
      ) : (
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "12px" }}>
            <thead>
              <tr style={{ background: T.surface, borderBottom: `1px solid ${T.border}` }}>
                <th style={{ padding: "8px 10px", textAlign: "left", color: T.muted, fontWeight: "normal" }}>Codice</th>
                <th style={{ padding: "8px 10px", textAlign: "left", color: T.muted, fontWeight: "normal" }}>Descrizione</th>
                <th style={{ padding: "8px 10px", textAlign: "center", color: T.muted, fontWeight: "normal" }}>qpb</th>
                <th style={{ padding: "8px 10px", textAlign: "left", color: T.muted, fontWeight: "normal" }}>Campo</th>
                <th style={{ padding: "8px 10px", textAlign: "center", color: T.dim, fontWeight: "normal", fontSize: "10px" }}>UoM IFB</th>
                <th style={{ padding: "8px 10px", textAlign: "right", color: T.dim, fontWeight: "normal", fontSize: "11px" }}>Prezzo IFB</th>
                <th style={{ padding: "8px 10px", textAlign: "center", color: T.blue, fontWeight: "normal", fontSize: "10px" }}>UoM HK</th>
                <th style={{ padding: "8px 10px", textAlign: "right", color: T.blue, fontWeight: "normal" }}>Prezzo HK</th>
                <th style={{ padding: "8px 10px", textAlign: "center", color: T.green, fontWeight: "normal", fontSize: "10px" }}>UoM Excel</th>
                <th style={{ padding: "8px 10px", textAlign: "right", color: T.green, fontWeight: "normal" }}>Excel</th>
                <th style={{ padding: "8px 10px", textAlign: "right", color: T.muted, fontWeight: "normal" }}>Δ</th>
                <th style={{ padding: "8px 10px", textAlign: "left", color: T.muted, fontWeight: "normal" }}>Motivo</th>
              </tr>
            </thead>
            <tbody>
              {displayed.slice(0, 500).map(({ pid, prod, diffs, hasDiff, qpb }) => {
                if (!hasDiff && filter !== "all") return null;
                if (!hasDiff) {
                  return (
                    <tr key={pid} style={{ borderBottom: `1px solid ${T.border}22` }}>
                      <td style={{ padding: "6px 10px", color: T.gold, fontFamily: "monospace" }}>{prod?.code || prod?.nHK || pid}</td>
                      <td style={{ padding: "6px 10px", color: T.muted, fontSize: "11px" }} colSpan={9}>{prod?.description || "—"} <span style={{ color: T.dim }}>· nessuna differenza</span></td>
                    </tr>
                  );
                }
                return diffs.map((d: any, i: number) => (
                  <tr key={`${pid}-${d.field}`} style={{ borderBottom: i === diffs.length - 1 ? `1px solid ${T.border}` : `1px solid ${T.border}11`, background: d.reason.startsWith("📦") ? `${T.blue}06` : d.field === "mtsPrice" ? `${T.gold}08` : "transparent" }}>
                    {i === 0 && (
                      <>
                        <td rowSpan={diffs.length} style={{ padding: "6px 10px", color: T.gold, fontFamily: "monospace", verticalAlign: "top" }}>{prod?.code || prod?.nHK || pid}</td>
                        <td rowSpan={diffs.length} style={{ padding: "6px 10px", color: T.muted, fontSize: "11px", verticalAlign: "top", maxWidth: "200px" }}>{prod?.description || "—"}</td>
                        <td rowSpan={diffs.length} style={{ padding: "6px 10px", textAlign: "center", color: T.dim, fontFamily: "monospace", fontSize: "11px", verticalAlign: "top" }}>{qpb > 1 ? qpb : "—"}</td>
                      </>
                    )}
                    <td style={{ padding: "4px 10px", color: T.dim, fontFamily: "monospace" }}>{d.label}</td>
                    <td style={{ padding: "4px 10px", textAlign: "center", color: T.dim, fontFamily: "monospace", fontSize: "10px" }}>{d.bcRawUom || "?"}</td>
                    <td style={{ padding: "4px 10px", textAlign: "right", color: T.dim, fontFamily: "monospace", fontSize: "11px" }}>
                      {d.bcRaw > 0 ? `€ ${d.bcRaw.toFixed(2)}` : <span style={{ color: T.dim }}>—</span>}
                    </td>
                    <td style={{ padding: "4px 10px", textAlign: "center", color: T.blue, fontFamily: "monospace", fontSize: "10px" }}>{d.bcUom || "?"}</td>
                    <td style={{ padding: "4px 10px", textAlign: "right", color: T.blue, fontFamily: "monospace" }}>
                      {d.bc > 0 ? (
                        d.bcNorm !== d.bc
                          ? <><span style={{ color: T.dim, fontSize: "10px" }}>€{d.bc.toFixed(2)}×{Math.round(d.bcNorm/d.bc)} </span>€{d.bcNorm.toFixed(2)}</>
                          : `€ ${d.bc.toFixed(2)}`
                      ) : <span style={{ color: T.dim }}>—</span>}
                    </td>
                    <td style={{ padding: "4px 10px", textAlign: "center", color: T.green, fontFamily: "monospace", fontSize: "10px" }}>{d.xlUom}</td>
                    <td style={{ padding: "4px 10px", textAlign: "right", color: T.green, fontFamily: "monospace" }}>{d.xl > 0 ? `€ ${d.xl.toFixed(2)}` : <span style={{ color: T.dim }}>—</span>}</td>
                    <td style={{ padding: "4px 10px", textAlign: "right", fontFamily: "monospace", color: d.delta > 0 ? T.orange : d.delta < 0 ? T.red : T.dim }}>
                      {d.bcNorm > 0 && d.xl > 0 ? `${d.delta > 0 ? "+" : ""}${d.delta.toFixed(2)}` : "—"}
                    </td>
                    <td style={{ padding: "4px 10px", color: T.muted, fontSize: "11px" }}>
                      {d.reason}
                      {d.uomNote && <div style={{ color: T.dim, fontSize: "10px", marginTop: "2px" }}>{d.uomNote}</div>}
                    </td>
                  </tr>
                ));
              })}
            </tbody>
          </table>
          {displayed.length > 500 && <div style={{ padding: "10px", textAlign: "center", color: T.dim, fontSize: "11px" }}>Mostrati 500/{displayed.length}</div>}
        </div>
      )}
    </div>
  );
}

// ─── PRICES (con import integrato e storico) ─────────────────────────────────
function Prices({ prices, setPrices, bcListini = [], setBcListini, products, branch, month, setPrices: setPricesParent, salesRows = [], xrefs = [],
  importLogs, setImportLogs, snapshots, setSnapshots, showToast, bumpImportTs, reloadListini, listiniMode = "bc", setListiniMode = (_:any)=>{} }) {
const [search, setSearchRaw] = useState(()=>psGet(`pg_${branch}_prices_search`,""));
const setSearch=(v:string)=>{setSearchRaw(v);psSet(`pg_${branch}_prices_search`,v);};
const [invoiceOnly, setInvoiceOnly] = useState(false);
const [importStep, setImportStep] = useState<"idle"|"map"|"preview"|"done">("idle");
const [headers, setHeaders] = useState<string[]>([]);
const [rawRows, setRawRows] = useState<any[]>([]);
const [mapping, setMapping] = useState<any>({});
const [preview, setPreview] = useState<any[]>([]);
const [fileName, setFileName] = useState("");
const [importMonth, setImportMonth] = useState(month);
const [doneInfo, setDoneInfo] = useState<any>(null);
// listiniMode e setListiniMode dal parent (App.tsx)
const [lastExcelData, setLastExcelData] = useState<{rawRows:any[];headers:string[];mapping:any;month:string;fileName:string}|null>(null);
const excelInputRef = useRef<HTMLInputElement>(null);

// Storico import listini
const priceSnaps = snapshots.filter((s: any) => s.type === "prices" && s.branch === branch);

function exportToExcel() {
  const rows = displayed.map((p: any) => {
    const prod = prodById[String(p.productId)];
    return {
      "N HK":         prod?.nHK    || "",
      "IFB No":       prod?.code   || p.itemCode  || "",
      "Descrizione":  prod?.description || p.bcDesc || "",
      "Base UoM":     prod?.uom    || "",
      "Purchase UoM": p.pu         || "",
      "MTS Price":    p.mtsPrice   || "",
      "FCA Price":    p.fcaPrice   || "",
      "FCA Disc.":    p.fcaDiscounted || "",
      "Carriage":     p.carriageCost || "",
      "DAP Price":    p.dapPrice   || "",
      "DAP Disc.":    p.dapDiscounted || "",
      "DAP Final":    p.dapFinal   || "",
      "Mese":         p.month      || month,
    };
  });
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, `Listini ${branch}`);
  XLSX.writeFile(wb, `Listini_${branch}_${month}.xlsx`);
}

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

// Auto-mapping dei campi — priorità al codice filiale (nHK/nComit) sull'IFB code
const am: any = {};
const branchCodeAliases2 = ["n hk", "nhk", "n comit", "ncomit", "comit"];
const ifbCodeAliases2 = ["ifb item", "ifb no", "ifb n", "no_", "item no.", "codice", "code"];
let foundBranchCode2 = false;
for (const h of hdrs) {
const hl = h.toLowerCase().trim();
if (branchCodeAliases2.some(a => hl === a || hl.includes(a))) {
am["code"] = h;
foundBranchCode2 = true;
break;
}
}
for (const h of hdrs) {
const hl = h.toLowerCase().trim();
if (ifbCodeAliases2.some(a => hl === a || hl.includes(a))) {
if (foundBranchCode2) { am["ifbCode"] = h; } else { am["code"] = h; }
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

const priceFields = ["mtsPrice", "fcaPrice", "fcaDiscount", "fcaDiscounted", "dapPrice", "dapDiscount", "dapDiscounted", "dapFinalDirect", "carriageCost", "section", "vendorName"];
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

function buildPreview(ovr?: {rawRows?:any[];headers?:string[];mapping?:any;month?:string}) {
const _rows    = ovr?.rawRows   ?? rawRows;
const _headers = ovr?.headers   ?? headers;
const _mapping = ovr?.mapping   ?? mapping;
const _month   = ovr?.month     ?? importMonth;
const get = (row: any, field: string) => {
const col = _mapping[field];
if (!col) return null;
const i = _headers.indexOf(col);
return i >= 0 ? row[i] : null;
};

if (_month !== importMonth) setImportMonth(_month);
const mapped = _rows.map((row, idx) => {
const rawCode    = String(get(row, "code")    || "").trim();
const rawIfbCode = String(get(row, "ifbCode") || "").trim();
const rawDescription = String(get(row, "description") || get(row, "code") || "").trim();

if (!rawCode && !rawIfbCode) return null;
if (rawCode && !isValidCode(rawCode) && !rawIfbCode) return null;

// rawCode = N COMIT (colonna "No_"): cerca per nHK prima, poi fallback generico
// rawIfbCode = colonna IFB No: cerca per code/xref
const prodByNComit = rawCode
  ? (products.find((pr:any) => pr.nHK && pr.nHK === rawCode) || null)
  : null;
const prod = prodByNComit
  || findProduct(rawIfbCode, products, xrefs)
  || findProduct(rawCode, products, xrefs);

const mtsPrice = parseFloat(get(row, "mtsPrice")) || 0;
const fcaPrice = parseFloat(get(row, "fcaPrice")) || 0;
const fcaDiscount = parseFloat(get(row, "fcaDiscount")) || 0;
// FCA Discounted: calcolato da FCA Price × (1 - FCA Discount% / 100)
// Per CAN: il file ha FCA Price + FCA Discount (%), non la colonna FCA Discounted precalcolata
// Legge la colonna pre-calcolata se presente, altrimenti calcola e arrotonda a 2dp (uguale all'Excel utente)
const fcaDiscounted = parseFloat(get(row, "fcaDiscounted")) || (fcaPrice > 0 ? roundN(fcaPrice * (1 - fcaDiscount / 100), 2) : 0);
const dapPrice = parseFloat(get(row, "dapPrice")) || 0;
const dapDiscount = parseFloat(get(row, "dapDiscount")) || 0;
const dapDiscounted = parseFloat(get(row, "dapDiscounted")) || (dapPrice > 0 ? roundN(dapPrice * (1 - dapDiscount / 100), 2) : 0);
const dapFinalDirect = parseFloat(get(row, "dapFinalDirect")) || 0;

let dapFinal = 0;
let dapNote = "";
if (dapFinalDirect !== 0) {
  dapFinal = dapFinalDirect;
  dapNote = "da file";
} else if (prod) {
  dapFinal = dapDiscounted || 0;
  dapNote = dapDiscounted ? (fcaDiscount > 0 || dapDiscount > 0 ? `DAP scontato (-${dapDiscount}%)` : "DAP") : "";

  // CAN Wine/Spirits: se DAP = 0 ma FCA NET > 0, calcola carriage = 60€/plt ÷ unità per plt
  // (su NAV il report non riesce a calcolare il DAP perché manca il carriage, lo ricostruiamo internamente)
  const _ub = String(get(row, "ubicazione") || "").toUpperCase();
  if (branch === "CAN" && dapFinal === 0 && fcaDiscounted > 0 && _ub !== "FOR") {
    const sec = (String(get(row, "section") || prod.category || "")).toUpperCase();
    if (sec.includes("WINE") || sec.includes("SPIRIT")) {
      const uom = prod.uom || "PCS";
      const qpb = Number(prod.qtyPerBox) || 1;
      const bpp = Number(prod.boxPerPallet) || 1;
      const kgp = Number(prod.kgxplt) || 300;
      const divisoreUom = uom === "BOX" ? bpp : uom === "KG" ? kgp : qpb * bpp;
      const carriageUnit = divisoreUom > 0 ? 60 / divisoreUom : 0;
      dapFinal = roundN(fcaDiscounted + carriageUnit, 6);
      dapNote = `Wine/Spirits: FCA NET + PLT(60/${divisoreUom})`;
    }
  }
}

const existing = prod ? prices.find(p => p.productId === prod.id && p.branch === branch && p.month === _month) : null;

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
dapDiscounted: roundN(dapDiscounted),
fcaDiscount: roundN(fcaDiscount),
dapDiscount: roundN(dapDiscount),
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
dapDiscounted: r.dapDiscounted || r.dapFinal || 0,
dapPrice: r.dapPrice,
fcaPrice: r.fcaPrice,
fcaDiscount: r.fcaDiscount || 0,
dapDiscount: r.dapDiscount || 0,
};
const prev = idx >= 0 ? updated[idx] : null;
const diffFields = [];

["dapFinal", "mtsPrice", "fcaDiscounted", "dapDiscounted", "dapPrice", "fcaPrice"].forEach(f => {
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
CLOUD.set(`ifb_prices_${branch}`, updated);
// Salva snapshot completo in IDB per poter ripristinare in seguito
IDB.set(`ifb_price_snap_${snId}`, updated.filter((p: any) => p.branch === branch && p.month === importMonth));

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
setListiniMode("excel");
setLastExcelData({ rawRows, headers, mapping, month: importMonth, fileName });
}

function loadFromSnapshot(snap: any) {
  IDB.get(`ifb_price_snap_${snap.id}`, null).then((data: any) => {
    if (!data || !Array.isArray(data) || data.length === 0) {
      showToast("Nessun dato trovato per questo snapshot (importato prima di questa versione)", T.orange);
      return;
    }
    const d = new Date(snap.id).toLocaleDateString("it-IT");
    if (!window.confirm(`Ripristinare ${data.length} prezzi del ${d} (${snap.month})? Sostituirà i prezzi attuali per ${snap.branch}/${snap.month}.`)) return;
    const others = prices.filter((p: any) => !(p.branch === snap.branch && p.month === snap.month));
    const restored = [...others, ...data];
    setPrices(restored);
    CLOUD.set(`ifb_prices_${branch}`, restored);
    setListiniMode("excel");
    showToast(`${data.length} prezzi ripristinati da ${d}`, T.green);
  });
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

// Lookup map prodotti O(1) — costruita una volta sola
const prodById = useMemo(() => {
  const m: Record<string,any> = {};
  products.forEach((p: any) => { if(p.id) m[String(p.id)] = p; });
  return m;
}, [products]);

const filtered = useMemo(() => {
  const baseList = listiniMode === "excel"
    ? prices.filter((p: any) => p.branch === branch && p.month === month)
    : bcListini.filter((p: any) => p.branch === branch && (p.fcaPrice > 0 || p.dapPrice > 0 || p.fcaDiscounted > 0 || p.dapDiscounted > 0));
  return baseList.filter((p: any) => {
    if (invoiceOnly && !invoiceProductIds.has(p.productId)) return false;
    return true;
  });
}, [listiniMode, bcListini, prices, branch, month, invoiceOnly, invoiceProductIds]);

const displayed = useMemo(() => {
  if (!search) return filtered;
  const q = search.toLowerCase();
  return filtered.filter((p: any) => {
    const prod = prodById[String(p.productId)];
    return prod?.description?.toLowerCase().includes(q) ||
      prod?.code?.toLowerCase().includes(q) ||
      prod?.nHK?.toLowerCase().includes(q) ||
      (p.itemCode || "").toLowerCase().includes(q) ||
      (p.nHK     || "").toLowerCase().includes(q) ||
      (p.bcDesc  || "").toLowerCase().includes(q);
  });
}, [filtered, search, prodById]);

const COLS = ["fcaPrice", "fcaDiscounted", "carriageCost", "dapPrice", "mtsPrice", "dapDiscounted", "dapFinal"];
const LABELS = ["FCA Price", "FCA Disc.", "Carriage", "DAP Price", "MTS Price", "DAP Disc.", "DAP Final"];

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
<div style={{ marginTop: "16px", display:"flex", gap:"10px", justifyContent:"center" }}>
<button onClick={async () => { setListiniMode("bc"); await IDB.del(`ifb_listini_entries_${branch}`); showToast("Caricamento da BC…", T.gold); reloadListini?.(); }}
  style={{ padding:"10px 20px", background:`${T.gold}22`, border:`1px solid ${T.gold}`, borderRadius:"8px", color:T.gold, cursor:"pointer", fontWeight:"bold", fontSize:"13px" }}>
  🏢 Carica da BC
</button>
<label style={{ padding:"10px 20px", background:`${T.green}22`, border:`1px solid ${T.green}`, borderRadius:"8px", color:T.green, cursor:"pointer", fontWeight:"bold", fontSize:"13px" }}>
  📂 Carica da Excel
  <input type="file" accept=".xlsx,.xls,.csv" onChange={e => { const f=e.target.files?.[0]; if(f) parseFile(f); e.target.value=""; }} style={{ display:"none" }} />
</label>
</div>
</div>
);
}

return (
<div>
<PageHeader title={`Listini · ${branch} · ${month}`} sub={`${filtered.length} prezzi caricati`} />
{listiniMode === "bc" ? (
  <BcBanner title="Dati aggiornati automaticamente da BC IFB Italia">
    Listini prezzi FCA / DAP / MTS caricati ogni giorno alle 07:00 dal listino acquisto e vendita di <b style={{color:T.text}}>Business Central IFB Italia</b>. Il campo <b style={{color:T.text}}>DAP</b> viene calcolato dalla tabella costi trasporto BC quando non è presente un prezzo DAP esplicito (Pallet1 ÷ pz/pallet).
  </BcBanner>
) : (
  <div style={{ background: `${T.green}11`, border: `1px solid ${T.green}33`, borderRadius: "8px", padding: "10px 14px", marginBottom: "14px", fontSize: "12px", color: T.muted, display: "flex", alignItems: "center", gap: "10px" }}>
    <span style={{ fontSize: "16px" }}>📂</span>
    <span>
      <b style={{ color: T.green }}>Dati da Excel</b>
      {lastExcelData ? <> · <b style={{ color: T.text }}>{lastExcelData.fileName.replace(/^.*[\\/]/,"")}</b> · mese <b style={{ color: T.text }}>{lastExcelData.month}</b></> : null}
      <span style={{ color: T.dim }}> — per tornare ai prezzi BC clicca 🏢 BC in basso</span>
    </span>
  </div>
)}

{/* Toolbar import */}
<div style={{ display: "flex", gap: "10px", marginBottom: "14px", alignItems: "center", flexWrap: "wrap" }}>

{/* Sorgente BC */}
<div style={{ display:"flex", gap:"2px", background:T.surface, border:`1px solid ${listiniMode==="bc"?T.gold:T.border}`, borderRadius:"8px", padding:"3px", alignItems:"center" }}>
  <button onClick={() => setListiniMode("bc")}
    style={{ padding:"5px 12px", background: listiniMode==="bc" ? `${T.gold}22` : "none", border:`1px solid ${listiniMode==="bc"?T.gold:"transparent"}`, borderRadius:"6px", color: listiniMode==="bc"?T.gold:T.muted, cursor:"pointer", fontSize:"12px", fontWeight: listiniMode==="bc"?"bold":"normal", whiteSpace:"nowrap" }}>
    🏢 BC
  </button>
  <button onClick={async () => { setListiniMode("bc"); await IDB.del(`ifb_listini_entries_${branch}`); showToast("Ricaricamento da BC…", T.gold); reloadListini?.(); }}
    style={{ padding:"5px 8px", background:"none", border:"none", color:T.muted, cursor:"pointer", fontSize:"13px" }} title="Ricarica da BC">
    🔄
  </button>
</div>

{/* Sorgente Excel */}
<div style={{ display:"flex", gap:"2px", background:T.surface, border:`1px solid ${listiniMode==="excel"?T.green:T.border}`, borderRadius:"8px", padding:"3px", alignItems:"center" }}>
  <label style={{ padding:"5px 12px", background: listiniMode==="excel" ? `${T.green}22` : "none", border:`1px solid ${listiniMode==="excel"?T.green:"transparent"}`, borderRadius:"6px", color: listiniMode==="excel"?T.green:T.muted, cursor:"pointer", fontSize:"12px", fontWeight: listiniMode==="excel"?"bold":"normal", whiteSpace:"nowrap" }}>
    📂 Excel{lastExcelData ? ` · ${lastExcelData.fileName.replace(/^.*[\\/]/,"")}` : ""}
    <input ref={excelInputRef} type="file" accept=".xlsx,.xls,.csv" onChange={e => { const f=e.target.files?.[0]; if(f) parseFile(f); e.target.value=""; }} style={{ display:"none" }} />
  </label>
  <button onClick={() => { if(lastExcelData) buildPreview({ rawRows:lastExcelData.rawRows, headers:lastExcelData.headers, mapping:lastExcelData.mapping, month:lastExcelData.month }); else excelInputRef.current?.click(); }}
    style={{ padding:"5px 8px", background:"none", border:"none", color:lastExcelData?T.muted:T.dim, cursor:lastExcelData?"pointer":"default", fontSize:"13px" }} title="Ricarica da ultimo Excel">
    🔄
  </button>
</div>

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

<button onClick={exportToExcel} disabled={displayed.length === 0} style={{ padding: "5px 12px", background: "none", border: `1px solid ${T.green}66`, borderRadius: "6px", color: T.green, cursor: "pointer", fontSize: "11px" }}>
⬇ Excel
</button>

{(setPricesParent || (listiniMode === "bc" && setBcListini)) && (
<button onClick={async () => {
  if (listiniMode === "bc") {
    if (!window.confirm(`Svuotare i dati BC per ${branch}?`)) return;
    await IDB.del(`ifb_listini_entries_${branch}`);
    setBcListini((prev: any[]) => prev.filter((p: any) => (p.branch || p.b) !== branch));
    showToast("Dati BC svuotati", T.red);
  } else {
    if (!window.confirm(`Eliminare tutti i prezzi Excel ${branch}/${month}?`)) return;
    setPricesParent(prices.filter((p: any) => !(p.branch === branch && p.month === month)));
    showToast("Prezzi Excel eliminati", T.red);
  }
}} style={{ padding: "5px 12px", background: "none", border: `1px solid ${T.red}44`, borderRadius: "6px", color: T.red, cursor: "pointer", fontSize: "11px" }}>
✕ Svuota {branch}{listiniMode === "excel" ? `/${month}` : " BC"}
</button>
)}
</div>

{/* Step di import - Mappa */}
{importStep === "map" && (
<div style={{ background: T.card, border: `1px solid ${T.gold}`, borderRadius: "8px", padding: "16px", marginBottom: "16px" }}>
<div style={{ color: T.gold, fontWeight: "bold", marginBottom: "12px" }}>Mappatura colonne · {fileName}</div>
{/* Campi base */}
<div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: "12px", marginBottom: "12px" }}>
<div>
<label style={{ fontSize: "11px", color: T.gold }}>📌 {branchN(branch)} (N COMIT / N HK)</label>
<select value={mapping["code"] || ""} onChange={e => setMapping((m: any) => ({ ...m, code: e.target.value || undefined }))} style={{ ...inputStyle(), fontSize: "12px", borderColor: (!mapping["code"] && !mapping["ifbCode"]) ? T.red+"88" : T.border }}>
<option value="">— non mappato —</option>
{headers.map(h => <option key={h} value={h}>{h}</option>)}
</select>
</div>
<div>
<label style={{ fontSize: "11px", color: T.gold }}>📌 IFB Item (codice IFB)</label>
<select value={mapping["ifbCode"] || ""} onChange={e => setMapping((m: any) => ({ ...m, ifbCode: e.target.value || undefined }))} style={{ ...inputStyle(), fontSize: "12px", borderColor: (!mapping["code"] && !mapping["ifbCode"]) ? T.red+"88" : T.border }}>
<option value="">— non mappato —</option>
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
</div>
<div style={{ display: "grid", gridTemplateColumns: "1fr", gap: "12px", marginBottom: "12px" }}>
<div>
<label style={{ fontSize: "11px", color: T.muted }}>📅 Mese listino</label>
<input type="month" value={importMonth} onChange={e => setImportMonth(e.target.value)} style={{ ...inputStyle(), fontSize: "12px", width: "160px" }} />
</div>
</div>
{/* Campi prezzo — auto-rilevati, modificabili manualmente */}
<div style={{ fontSize: "11px", color: T.muted, marginBottom: "6px", borderTop: `1px solid ${T.border}`, paddingTop: "10px" }}>Prezzi (auto-rilevati — modifica se necessario)</div>
<div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: "10px", marginBottom: "16px" }}>
{([
["mtsPrice",       "💲 MTS Price"],
["fcaPrice",       "💲 FCA Price"],
["fcaDiscount",    "% FCA Discount"],
["fcaDiscounted",  "💲 FCA Discounted"],
["dapPrice",       "💲 DAP Price"],
["dapDiscount",    "% DAP Discount"],
["dapDiscounted",  "💲 DAP Discounted"],
["dapFinalDirect", "💲 DAP Final"],
["carriageCost",   "💲 Carriage Cost"],
] as [string,string][]).map(([field, label]) => (
<div key={field}>
<label style={{ fontSize: "11px", color: mapping[field] ? T.green : T.dim }}>{label}</label>
<select value={mapping[field] || ""} onChange={e => setMapping((m: any) => ({ ...m, [field]: e.target.value || undefined }))} style={{ ...inputStyle(), fontSize: "12px", borderColor: mapping[field] ? `${T.green}66` : undefined }}>
<option value="">— non mappato —</option>
{headers.map(h => <option key={h} value={h}>{h}</option>)}
</select>
</div>
))}
</div>
<div style={{ display: "flex", gap: "10px" }}>
<ActionBtn label="Annulla" onClick={resetImport} />
<ActionBtn label="Preview →" onClick={buildPreview} primary disabled={!mapping["code"] && !mapping["ifbCode"]} />
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
<thead><tr><th>Codice</th><th>Descrizione</th><th>Match</th>{branch==="CAN"?<><th>FCA Disc.</th><th>DAP Disc.</th><th>MTS Price</th></>:<th>DAP Final</th>}<th>Stato</th></tr></thead>
<tbody>{preview.slice(0, 20).map(r => (
<tr key={r._idx} style={{ borderBottom: `1px solid ${T.border}` }}>
  <td style={{ fontFamily: "monospace", color: T.gold }}>{r.ifbNo_from_file}</td>
  <td>{r.description_from_file}</td>
  <td>{r._hasProduct ? <span style={{ color: T.green }}>✓ {r.ifbNo_from_anag}</span> : <span style={{ color: T.red }}>✗</span>}</td>
  {branch==="CAN"?<>
    <td style={{ fontFamily: "monospace", color: T.muted }}>{r.fcaDiscounted > 0 ? `€ ${r.fcaDiscounted.toFixed(4)}${r.fcaDiscount > 0 ? ` (-${r.fcaDiscount}%)` : ""}` : "—"}</td>
    <td style={{ fontFamily: "monospace", color: T.orange }}>{r.dapDiscounted > 0 ? `€ ${r.dapDiscounted.toFixed(4)}${r.dapDiscount > 0 ? ` (-${r.dapDiscount}%)` : ""}` : "—"}</td>
    <td style={{ fontFamily: "monospace" }}>{r.mtsPrice > 0 ? `€ ${r.mtsPrice.toFixed(4)}` : "—"}</td>
  </>:<td style={{ fontFamily: "monospace" }}>{r.dapFinal > 0 ? `€ ${r.dapFinal.toFixed(2)}` : "—"}</td>}
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
<THead cols={[branchN(branch),"IFB No","Descrizione (Base UoM)","Purchase UoM","FCA Price","FCA Disc.","Carriage","DAP Price","MTS Price","DAP Disc.","DAP Final"]} sticky />
<tbody>
{displayed.slice(0, 150).map((p: any, i: number) => {
const prod = prodById[String(p.productId)];
const inInvoice = invoiceProductIds.has(p.productId);
const puom = p.pu || "";
const baseUom = prod?.uom || "";
// Fattore di conversione da Purchase UOM a Base UOM
const _pr = prod || products.find((pr:any) => pr.code === (p.n || p.itemCode));
const _qpb = Number(_pr?.qtyPerBox) || 1;
const _kpb = Number(_pr?.kgPerBox)  || 0;
let convFactor = 1;
// Fallback: se pu non è già stata aggiornata a baseUom nel loading, converti qui
if (puom && baseUom && puom !== baseUom) {
  if      (puom === "BOX" && baseUom === "PCS") convFactor = 1 / _qpb;           // ÷qpb
  else if (puom === "BOX" && baseUom === "KG")  convFactor = 1 / (_kpb || 1);    // ÷kpb
  else if (puom === "KG"  && baseUom === "PCS") convFactor = _kpb > 0 ? _kpb : 1; // ×kpb
  else if (puom === "PCS" && baseUom === "KG")  convFactor = _kpb > 0 ? _kpb : 1; // ×kpb
}
const needsConv = convFactor !== 1;
return (
  <tr key={p.productId} style={{ borderBottom: `1px solid ${T.border}`, background: i % 2 === 0 ? T.bg : T.surface }}>
    <TD mono><span style={{ color: T.muted }}>{prod?.nHK || (xrefs.find((x:any)=>x.ifbNo===(prod?.code||p.n||p.itemCode))?.nHK) || "—"}</span></TD>
    <TD mono>
      <span style={{ color: T.gold }}>{prod?.code || p.itemCode || p.productId}</span>
      {inInvoice && <span style={{ marginLeft: "5px", fontSize: "9px", color: T.blue }}>📋</span>}
    </TD>
    <TD>
      <div style={{display:"flex",alignItems:"center",gap:"6px"}}>
        <span>{prod?.description || p.bcDesc || <span style={{ color: T.dim, fontSize: "11px" }}>{p.itemCode}</span>}</span>
        {baseUom && <span style={{fontSize:"9px",color:T.gold,background:`${T.gold}18`,padding:"1px 5px",borderRadius:"4px",fontFamily:"monospace",whiteSpace:"nowrap"}}>{baseUom}</span>}
        {needsConv && <span style={{fontSize:"8px",color:T.muted,whiteSpace:"nowrap"}}>{puom}→{baseUom}</span>}
      </div>
    </TD>
    <TD mono><span style={{ color: puom ? T.muted : T.dim, fontSize: "10px" }}>{puom || "—"}</span></TD>
    {COLS.map(f => {
      const raw = p[f] || 0;
      const val = raw * convFactor;
      return (
      <TD key={f} mono>
        <span style={{ color: val > 0 ? T.text : T.dim }}>
          {val > 0 ? `€ ${roundN(val).toFixed(4).replace(/\.?0+$/,"")}` : "—"}
        </span>
      </TD>
    );
    })}
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
  setCostHistory,initFilter,salesRows=[],products=[],xrefs=[],listiniMode="bc",setListiniMode=(_:any)=>{},reloadListini=()=>{}}: any) {

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

  const searchLow = search.toLowerCase();
  let filtered: any[] = costRows.filter((r:any)=>
    !search
    ||r.description?.toLowerCase().includes(searchLow)
    ||r.code?.toLowerCase().includes(searchLow)
    ||r.nHK?.toLowerCase().includes(searchLow));

// APPLICA FILTRI MULTIPLI (include = mostra solo questi; exclude = nascondi questi)
  const applyFlag = (flag:false|"include"|"exclude", test:(r:any)=>boolean) => {
    if(flag==="include") filtered=filtered.filter(test);
    else if(flag==="exclude") filtered=filtered.filter(r=>!test(r));
  };
  applyFlag(filterFlags.costCalculated, r=> r.cost?.step2Hkd!=null);
  applyFlag(filterFlags.flagged,        r=> r.flagged===true);
  applyFlag(filterFlags.air,            r=> r.isAir===true);
  applyFlag(filterFlags.noPrice,        r=> !r.cost&&!r.isAir&&!!(r.skipReason?.includes("NO PREZZO")||r.skipReason?.includes("LISTINO CHIUSO")));
  applyFlag(filterFlags.noLog,          r=> !r.cost&&!r.isAir&&r.skipReason==="NO LOGISTICA");
  applyFlag(filterFlags.calcZero,       r=> !r.cost&&!r.isAir&&!!r.skipReason?.includes("CALC=0"));
  applyFlag(filterFlags.keepOld,        r=> { const d=lastOrderDate[r.id]; return !!(d&&d<sixMonthsAgo); });
if (invoiceOnly) {
  filtered = filtered.filter((r:any) => invoiceIds.has(r.id));
}

if(initFilter==="flagged") filtered=filtered.filter((r:any)=>r.flagged===true);
else if(initFilter==="errors") filtered=filtered.filter((r:any)=>!r.cost&&!r.isAir&&r.skipReason?.includes("CALC=0"));

  const calc    = filtered.filter((r:any)=>r.cost?.step2Hkd!=null);
  const noPrice = filtered.filter((r:any)=>!r.cost&&!r.isAir&&(r.skipReason?.includes("NO PREZZO")||r.skipReason?.includes("LISTINO CHIUSO")));
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

  const f4=(v:number|undefined)=>v!=null&&v!==0?v.toFixed(2):"—";
  const f2=(v:number|undefined)=>v!=null&&v!==0?v.toFixed(2):"—";

  return(
    <div style={{paddingRight:"20px"}}>
      {/* ── toolbar ── */}
      <div style={{display:"flex",alignItems:"center",gap:"12px",marginBottom:"12px",flexWrap:"wrap"}}>
        <PageHeader title={`Standard Cost · ${branch} · ${month}`}
          sub={branch==="MAC"?`${calc.length} articoli · SC HKD × markup × ${HKD_TO_MOP} HKD/MOP`:`${filtered.length} articoli · ${calc.length} calcolati · INALCA F&B · SEA`}/>
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
        <button onClick={()=>{ setListiniMode("bc"); reloadListini(); }}
          title="Ricarica listini prezzi da BC IFB e aggiorna tutti i calcoli"
          style={{padding:"7px 14px",background:listiniMode==="excel"?`${T.blue}22`:`${T.surface}`,
            border:`1px solid ${listiniMode==="excel"?T.blue:T.border}`,borderRadius:"6px",
            color:listiniMode==="excel"?T.blue:T.muted,cursor:"pointer",fontSize:"12px",marginTop:"-8px",whiteSpace:"nowrap"}}>
          {listiniMode==="excel" ? "📊 Excel attivo — 🔄 Aggiorna da BC" : "🔄 Aggiorna Listini da BC"}
        </button>
        <button onClick={()=>exportXLSX(
          filtered.filter((r:any)=>r.cost).map((r:any)=> branch==="CAN" ? (()=>{
            const nComitVal = (xrefs||[]).find((x:any)=>x.ifbNo===r.code||x.ifbNo===r.nHK)?.nHK || r.nHK || "";
            const ifbVal = r.code || r.nHK || "";
            return {
            "N COMIT":nComitVal,"IFB No":ifbVal,"Descrizione":r.description||"",
            "UOM":r.uom||"","Ubicazione":r.ubicazione||"","Trasporto":r.cost?.transport||"",
            "Temp.":r.temperature||"","Temp. Rettif.":r.temperatureOverride||"",
            "Prezzo EUR":roundN(r.cost?.priceEur),
            "Verona-Barc":roundN(r.cost?.veronaBarcUnit),"Barc-GC":roundN(r.cost?.barcUnitGC),
            "Freight GC":roundN(r.cost?.freightGC),"Inland GC":roundN(r.cost?.inlandGC),
            "Assic.":roundN(r.cost?.assicUnit),"Pallet":roundN(r.cost?.plt),
            "AIEM GC/TF":roundN(r.cost?.aiemGCTF),"AIEM LAN/FUE":roundN(r.cost?.aiemLANFUE),
            "Tassa Alcolica":roundN(r.cost?.tassaAlcolica),"WH EUR":roundN(r.cost?.wh),
            "Step1 GC":roundN(r.cost?.step1GC),"Step1 TF":roundN(r.cost?.step1TF),
            "Step1 LAN":roundN(r.cost?.step1LAN),"Step1 FUE":roundN(r.cost?.step1FUE),
            "Step2 GC":roundN(r.cost?.step2GC),"Step2 TF":roundN(r.cost?.step2TF),
            "Step2 LAN":roundN(r.cost?.step2LAN),"Step2 FUE":roundN(r.cost?.step2FUE),
            "Δ%":r.delta!=null?roundN(r.delta,1):"",
          }})() : ({
            "N HK":r.nHK||"","IFB No":r.code||"","Descrizione":r.description||"",
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
          "Standard Cost",`SC_${branch}_${month}.xlsx`,
          branch==="CAN" ? ["N COMIT","IFB No"] : ["N HK","IFB No"]
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
              <TH align="left" sticky w={55}>{branchN(branch)}</TH>
              <TH align="left" w={55}>IFB No</TH>
              <TH align="left" w={150}>Descrizione</TH>
              <TH w={42}>UOM</TH>
              <TH w={42}>Ubicaz.</TH>
              {branch==="CAN"&&<TH w={50} align="center">Tratta</TH>}
              <TH w={42} align="center">Temp.</TH>
              <TH w={42} align="center">Rettif.</TH>
              <TH accent={T.blue} w={58}>Prezzo €</TH>
              {branch==="CAN" ? <>
                <TH accent={T.blue} w={52}>Trasp.</TH>
                <TH accent={T.blue} w={48}>Pallet</TH>
                <TH accent={T.blue} w={48}>AIEM</TH>
              </> : <>
                <TH accent={T.blue} w={55}>FOB</TH>
                <TH accent={T.blue} w={55}>LIC</TH>
                <TH accent={T.blue} w={48}>VGM</TH>
                <TH accent={T.blue} w={48}>Cert.</TH>
                <TH accent={T.blue} w={52}>Pallet</TH>
                <TH accent={T.blue} w={52}>Alc.Tax</TH>
              </>}
              <TH accent={T.purple} w={52}>WH €</TH>
              {branch==="CAN" ? <>
                <TH accent={T.gold} w={68}>New SC GC/TF</TH>
                <TH accent={T.gold} w={72}>New SC FUE/LAN ✓</TH>
              </> : <>
                <TH accent={T.green} w={65}>New SC €</TH>
                <TH accent={T.green} w={78}>New SC HKD ✓</TH>
              </>}
              <TH w={48}>Δ%</TH>
              <TH w={72}>Ultimo ordine</TH>
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
                  <td style={{...cellL(),maxWidth:"150px",whiteSpace:"normal",wordBreak:"break-word",lineHeight:"1.3"}}>
                    {r.description}
                    {r.isAir&&<span style={{marginLeft:"4px",color:T.orange,fontSize:"9px"}}>✈</span>}
                    {r._keepOld&&<span style={{marginLeft:"4px",color:T.orange,fontSize:"9px",fontWeight:"bold"}}>🔒 KEEP OLD</span>}
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
                    <td style={cell(T.gold,true)}>{c?`€${c.step2GC.toFixed(4)}`:"—"}</td>
                    <td style={cell(T.gold,true)}>
                      <span style={{fontSize:"11px",fontWeight:"bold"}}>
                        {c?`€${c.step2FUE.toFixed(4)}`:<span style={{color:T.dim,fontWeight:"normal",fontSize:"9px"}}>{r.skipReason||"—"}</span>}
                      </span>
                    </td>
                  </> : <>
                    <td style={cell(T.green,true)}>{c?`€${c.step2Eur.toFixed(4)}`:"—"}</td>
                    <td style={{...cell(undefined,true),background:hkd!=null?`${T.gold}33`:undefined}}>
                      <span style={{fontSize:"11px",fontWeight:"bold",color:hkd!=null?T.gold:T.dim}}>
                        {hkd!=null?`${hkd.toFixed(2)}`:<span style={{fontWeight:"normal",fontSize:"9px"}}>{r.skipReason||"—"}</span>}
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
                        const f4=(v:number)=>`€ ${v.toFixed(2)}`;
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
                                  {sep("Pallet & AIEM / Tassa Alcolica")}
                                  {row("Pallet",f4(c.plt),f4(c.plt),"15 € ÷ u/plt (COSTS LOG!I1)",T.blue)}
                                  {c.aiemGCTF>0&&row("AIEM",f4(c.aiemGCTF),f4(c.aiemLANFUE||0),"(Prezzo + Trasporto isola) × AIEM%",T.orange)}
                                  {c.tassaAlcolica>0&&row("Tassa Alcolica",f4(c.tassaAlcolica),f4(c.tassaAlcolica),"LT × €/LT (Beverage Info — fisso/unit)",T.orange)}
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
                        const f4=(v:number)=>`€ ${v.toFixed(2)}`;
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



function InvoiceAndCosts({rows,setRows,branch,airList,products,xrefs,costRows,logistics,snapshots,setSnapshots,importLogs,setImportLogs,showToast,bumpImportTs,scAttuali}) {
  const [step,setStep]         = useState(()=>rows?.length?"view":"upload");
  const [preview,setPreview]   = useState<any[]>([]);
  const [headers,setHeaders]   = useState<string[]>([]);
  const [mapping,setMapping]   = useState<any>({});
  const [rawRows,setRawRows]   = useState<any[]>([]);
  const [fileName,setFileName] = useState("");
  const FK = (k:string) => `inv_filter_${branch}_${k}`;
  const lsGet = <T,>(k:string,def:T):T=>{ try{ const v=localStorage.getItem(FK(k)); return v!=null?JSON.parse(v):def; }catch{ return def; } };
  const lsSet = (k:string,v:any)=>{ try{ localStorage.setItem(FK(k),JSON.stringify(v)); }catch{} };
  const [excludeAir,setExcludeAirRaw]     = useState(()=>lsGet("excludeAir",false));
  const [last30,setLast30Raw]             = useState(()=>lsGet("last30",false));
  const [dedup,setDedupRaw]               = useState(()=>lsGet("dedup",false));
  const [excludeKeepOld,setExcludeKeepOldRaw] = useState(()=>lsGet("excludeKeepOld",false));
  const [excludeSample,setExcludeSampleRaw]   = useState(()=>lsGet("excludeSample",false));
  const [newHkdFilter,setNewHkdFilterRaw] = useState<"all"|"ok"|"mancante"|"air">(()=>lsGet("newHkdFilter","all"));
  const [scFilter,setScFilterRaw]         = useState<"all"|"ok"|"mancante"|"sample">(()=>lsGet("scFilter","all"));
  const [filterTransport,setFilterTransportRaw] = useState(()=>lsGet("filterTransport","all"));
  const [filterNHK,setFilterNHKRaw]   = useState(()=>lsGet("filterNHK",""));
  const [filterIFBNo,setFilterIFBNoRaw] = useState(()=>lsGet("filterIFBNo",""));
  const [filterLocation,setFilterLocationRaw] = useState<"all"|"ncj"|"non-ncj">(()=>lsGet("filterLocation","all"));
  const [filterScBC,setFilterScBCRaw] = useState<"all"|"assente">(()=>lsGet("filterScBC","all"));
  const [filterMotivo,setFilterMotivoRaw] = useState<"all"|"no-log"|"no-price"|"anagrafica"|"sample"|"keep-old">(()=>lsGet("filterMotivo","all"));
  const [filterScNavGC,setFilterScNavGCRaw] = useState<"all"|"assente">(()=>lsGet("filterScNavGC","all"));
  const [filterDeltaPct,setFilterDeltaPctRaw] = useState<"all"|"pos3"|"neg3">(()=>lsGet("filterDeltaPct","all"));
  const [showNoAna,setShowNoAnaRaw] = useState(()=>lsGet("showNoAna",false));
  const [search,setSearchRaw]     = useState(()=>lsGet("search",""));
  const setExcludeAir=(v:any)=>{ const nv=typeof v==="function"?v(excludeAir):v; setExcludeAirRaw(nv); lsSet("excludeAir",nv); };
  const setLast30=(v:any)=>{ const nv=typeof v==="function"?v(last30):v; setLast30Raw(nv); lsSet("last30",nv); };
  const setDedup=(v:any)=>{ const nv=typeof v==="function"?v(dedup):v; setDedupRaw(nv); lsSet("dedup",nv); };
  const setExcludeKeepOld=(v:any)=>{ const nv=typeof v==="function"?v(excludeKeepOld):v; setExcludeKeepOldRaw(nv); lsSet("excludeKeepOld",nv); };
  const setExcludeSample=(v:any)=>{ const nv=typeof v==="function"?v(excludeSample):v; setExcludeSampleRaw(nv); lsSet("excludeSample",nv); };
  const setNewHkdFilter=(v:any)=>{ setNewHkdFilterRaw(v); lsSet("newHkdFilter",v); };
  const setScFilter=(v:any)=>{ setScFilterRaw(v); lsSet("scFilter",v); };
  const setFilterTransport=(v:any)=>{ setFilterTransportRaw(v); lsSet("filterTransport",v); };
  const setFilterNHK=(v:any)=>{ setFilterNHKRaw(v); lsSet("filterNHK",v); };
  const setFilterIFBNo=(v:any)=>{ setFilterIFBNoRaw(v); lsSet("filterIFBNo",v); };
  const setFilterLocation=(v:any)=>{ setFilterLocationRaw(v); lsSet("filterLocation",v); };
  const setFilterScBC=(v:any)=>{ setFilterScBCRaw(v); lsSet("filterScBC",v); };
  const setFilterMotivo=(v:any)=>{ setFilterMotivoRaw(v); lsSet("filterMotivo",v); };
  const setFilterScNavGC=(v:any)=>{ setFilterScNavGCRaw(v); lsSet("filterScNavGC",v); };
  const setFilterDeltaPct=(v:any)=>{ setFilterDeltaPctRaw(v); lsSet("filterDeltaPct",v); };
  const setShowNoAna=(v:any)=>{ const nv=typeof v==="function"?v(showNoAna):v; setShowNoAnaRaw(nv); lsSet("showNoAna",nv); };
  const setSearch=(v:string)=>{ setSearchRaw(v); lsSet("search",v); };
  const [sortDir,setSortDir]   = useState<"desc"|"asc">("desc");
  const [sortBy,setSortBy]     = useState<"date"|"delta">("date");

  useEffect(()=>{ if(rows?.length&&step==="upload") setStep("view"); },[rows]);

  function saveRows(data:any[]) {
    setRows(data);
    CLOUD.set(`ifb_sales_invoice_${branch}`, data);
    setDataSource(`fatture_${branch}`,"manual");
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

  // Trova xref per ifbNo preferendo nHK puramente numerico (N COMIT reale vs codice BC alfanumerico)
  const xrefByIfbNoPreferNumeric = useMemo(()=>{
    const fn = (ifbNo:string) => {
      const matches=(xrefs||[]).filter((x:any)=>String(x.ifbNo)===ifbNo);
      return matches.find((x:any)=>/^\d+$/.test(String(x.nHK)))||matches[0]||null;
    };
    return fn;
  },[xrefs]);

  const scAttualiMap=useMemo(()=>{
    const m: Record<string,any>={};
    // Indice primario: rec.code (N COMIT per CAN, N HK per HK)
    (scAttuali||[]).forEach((r:any)=>{ if(r.code!=null) m[String(r.code)]=r; });
    // CAN: ifbCode salvato direttamente nel record — aggiungilo come chiave diretta
    (scAttuali||[]).forEach((r:any)=>{ if(r.ifbCode) { const k=String(r.ifbCode); if(!m[k]) m[k]=r; } });
    // Alias bidirezionale tramite xref
    (scAttuali||[]).forEach((rec:any)=>{
      if(rec.code==null) return;
      const k = String(rec.code);
      const xrByNHK=(xrefs||[]).find((x:any)=>String(x.nHK)===k);
      if(xrByNHK?.ifbNo){ const ak=String(xrByNHK.ifbNo); if(!m[ak]) m[ak]=rec; }
      const xrByIfb=xrefByIfbNoPreferNumeric(k);
      if(xrByIfb?.nHK){ const ak=String(xrByIfb.nHK); if(!m[ak]) m[ak]=rec; }
    });
    return m;
  },[scAttuali,xrefs,xrefByIfbNoPreferNumeric]);

  // Items che appaiono in fattura → se hanno NO PREZZO ora, il listino era aperto e poi chiuso
  const soldItemCodes = useMemo(()=>new Set(activeRows.map((r:any)=>String(r.itemCode||""))), [activeRows]);

  const enriched=useMemo(()=>{
    return [...activeRows]
      .sort((a:any,b:any)=>{
        if(!a.date&&!b.date)return 0;if(!a.date)return 1;if(!b.date)return -1;
        return sortDir==="desc"?b.date.localeCompare(a.date):a.date.localeCompare(b.date);
      })
      .map((r:any)=>{
        // CAN: r.itemCode = N COMIT → xref → prod. Fallback: prova IFB code diretto
        const prod=findProduct(r.itemCode,products,xrefs)
          ||(branch==="CAN"?products?.find((p:any)=>p.code===r.itemCode)||null:null);
        const cr=prod?costRows.find((c:any)=>c.id===prod.id):null;
        const isAir=branch!=="CAN"&&(r.transport==="AIR"||cr?.isAir===true||cr?.skipReason==="AIR");
        const locationIsNCJ=branch!=="CAN"&&String(r.location||"").toUpperCase().includes("NCJ");
        const mismatch=branch!=="CAN"&&((isAir&&!locationIsNCJ)||(!isAir&&locationIsNCJ));
        const newHkd=cr?.cost?.step2Hkd??null;
        const oldHkd=cr?.prevCost?.step2Hkd??null;
        const pct=newHkd!=null&&oldHkd!=null&&oldHkd>0?(newHkd-oldHkd)/oldHkd*100:null;
        const isNonFoodCAN=branch==="CAN"&&/^HO\.RE\.CA\./i.test(String(prod?.category||""));
        const isSampleRow=r.unitPrice===0||r.unitPrice===0.01;
        const logEntry=prod?logistics?.find((l:any)=>l.productId===prod.id&&l.branch===branch):null;
        const skipReasonRaw=isAir?"AIR":isNonFoodCAN?"NON FOOD":isSampleRow?"SAMPLE":
          cr?.skipReason||(!prod?"NON IN ANAGRAFICA":!cr&&!logEntry?"NO LOGISTICA":"");
        // Se "NO PREZZO" ma l'articolo appare in fattura → listino era aperto e poi chiuso
        const skipReason = skipReasonRaw?.includes("NO PREZZO") && soldItemCodes.has(String(r.itemCode||""))
          ? skipReasonRaw.replace("NO PREZZO","LISTINO CHIUSO E NON RIAPERTO")
          : skipReasonRaw;
        const logTransport=logEntry?.transport||"";
        // Usa island flags dalla logistica: default true se entry mancante
        const destGCTF   = !logEntry || logEntry.isGC || logEntry.isTF;
        const destLANFUE = !logEntry || logEntry.isLAN || logEntry.isFUE;
        const scGC  = destGCTF   ? (cr?.cost?.step2GC  ?? null) : null;
        const scFUE = destLANFUE ? (cr?.cost?.step2FUE ?? null) : null;
        // CAN: scAttuali keyed by N COMIT (e alias IFB via xref aggiunto nella mappa)
        //   1. r.itemCode (N COMIT o IFB — mappa copre entrambi)
        //   2. N COMIT ricavato dall'xref se r.itemCode è IFB
        //   3. prod.code (IFB) — già in mappa come alias
        // HK:  scAttuali keyed by N HK  → usa prod.nHK o r.nHK
        // CAN: itemCode = N COMIT (fatture Comit), scAttuali keyed by N COMIT (Item No)
        // Fallback: cerca anche per IFB code (tramite xref o prod.code)
        const sItemCode = String(r.itemCode||"");
        const nComitFromIfb = branch==="CAN"
          ? (String(xrefByIfbNoPreferNumeric(sItemCode)?.nHK||""))
          : "";
        const ifbFromNComit = branch==="CAN"
          ? (String((xrefs||[]).find((x:any)=>String(x.nHK)===sItemCode)?.ifbNo||""))
          : "";
        const scaRec = branch==="CAN"
          ? (scAttualiMap[sItemCode]           // 1. N COMIT diretto
             || scAttualiMap[nComitFromIfb]    // 2. N COMIT via xref (se itemCode era IFB)
             || scAttualiMap[ifbFromNComit]    // 3. IFB code via xref (se itemCode era N COMIT)
             || scAttualiMap[prod?.code||""])  // 4. IFB code da prodotto
          : (scAttualiMap[prod?.nHK||""] || scAttualiMap[r.nHK||""] || scAttualiMap[prod?.code||""] || scAttualiMap[sItemCode]);
        const bcStdCost = branch==="CAN" ? null : (prod?.standardCostHkd || null);
        const deltaSC = branch==="CAN" ? null
          : (newHkd != null && bcStdCost != null && bcStdCost > 0 ? newHkd - bcStdCost : null);
        const scBcGcTf  = branch==="CAN" ? (scaRec?.scGC  || null) : null;
        const scBcFueLan= branch==="CAN" ? (scaRec?.scLan || null) : null;
        // Delta CAN: SC proposto (scGC/scFUE) vs SC Attuali (scBcGcTf/scBcFueLan)
        const deltaGC  = branch==="CAN" && scGC!=null && scBcGcTf!=null && scBcGcTf>0 ? scGC - scBcGcTf : null;
        const deltaFUE = branch==="CAN" && scFUE!=null && scBcFueLan!=null && scBcFueLan>0 ? scFUE - scBcFueLan : null;
        // Δ% CAN: usa GC come riferimento; HK: usa prevCost se disponibile, altrimenti bcStdCost
        // canPct: usa GC/TF come riferimento se prodotto va lì, altrimenti LAN/FUE
        const canPct  = branch==="CAN"
          ? (destGCTF && scBcGcTf!=null && scBcGcTf>0 && scGC!=null
              ? (scGC-scBcGcTf)/scBcGcTf*100
              : destLANFUE && scBcFueLan!=null && scBcFueLan>0 && scFUE!=null
                ? (scFUE-scBcFueLan)/scBcFueLan*100
                : null)
          : null;
        const refHkd  = oldHkd ?? bcStdCost;
        const hkPct   = branch!=="CAN" && newHkd!=null && refHkd!=null && refHkd>0 ? (newHkd-refHkd)/refHkd*100 : null;
        const finalPct= branch==="CAN" ? canPct : (pct ?? hkPct);
        const lastOrderRaw = logEntry?.lastOrderDate || r.date;
        const lastOrderD = lastOrderRaw ? new Date(String(lastOrderRaw).slice(0,10)) : null;
        const isKeepOld = lastOrderD ? ((Date.now()-lastOrderD.getTime())/(86400000))>180 : false;
        return{...r,nHK:prod?.nHK||r.nHK||"",ifbNo:prod?.code||r.itemCode||"",
          description:r.description||prod?.description||"",ubicazione:cr?.ubicazione||"",logTransport,
          isAir,locationIsNCJ,mismatch,newHkd,oldHkd,pct:finalPct,skipReason,scGC,scFUE,
          bcStdCost,deltaSC,scBcGcTf,scBcFueLan,deltaGC,deltaFUE,isKeepOld,
          uomVendita:prod?.uom||"",uomAcquisto:prod?.baseUom||""};
      });
  },[activeRows,costRows,products,xrefs,sortDir,scAttualiMap]);

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
  const isSample = (r:any) => r.unitPrice===0||r.unitPrice===0.01;
  if(scFilter==="ok")            displayed=displayed.filter(r=>r.scGC!=null&&!isSample(r));
  else if(scFilter==="mancante") displayed=displayed.filter(r=>r.scGC===null&&!isSample(r));
  else if(scFilter==="sample")   displayed=displayed.filter(r=>isSample(r));
  if(last30){ const cutStr=new Date(Date.now()-30*86400000).toISOString().slice(0,10); displayed=displayed.filter(r=>String(r.date||r.postingDate||"").slice(0,10)>=cutStr); }
  if(dedup){
    const best=new Map<string,any>();
    displayed.forEach(r=>{
      const key=String(r.itemCode||r.ifbNo||"");
      const prev=best.get(key);
      if(!prev||new Date(r.date)>new Date(prev.date)) best.set(key,r);
    });
    displayed=[...best.values()];
  }
  if(filterNHK==="__MISSING__") displayed=displayed.filter(r=>!r.nHK);
  else if(filterNHK)            displayed=displayed.filter(r=>r.nHK===filterNHK);
  if(filterLocation==="ncj")     displayed=displayed.filter(r=>String(r.location||"").toUpperCase().includes("NCJ"));
  else if(filterLocation==="non-ncj") displayed=displayed.filter(r=>!String(r.location||"").toUpperCase().includes("NCJ"));
  if(filterIFBNo) displayed=displayed.filter(r=>r.ifbNo===filterIFBNo);
  if(filterScBC==="assente") displayed=displayed.filter(r=>
    branch==="CAN" ? (!r.scBcGcTf&&!r.scBcFueLan) : (!r.bcStdCost||r.bcStdCost===0));
  if(filterScNavGC==="assente") displayed=displayed.filter(r=>!r.scBcGcTf||r.scBcGcTf===0);
  if(filterMotivo==="no-log") displayed=displayed.filter(r=>r.skipReason==="NO LOGISTICA");
  else if(filterMotivo==="no-price") displayed=displayed.filter(r=>r.skipReason?.includes("NO PREZZO")||r.skipReason?.includes("LISTINO CHIUSO"));
  else if(filterMotivo==="anagrafica") displayed=displayed.filter(r=>r.skipReason==="NON IN ANAGRAFICA");
  else if(filterMotivo==="sample") displayed=displayed.filter(r=>r.skipReason==="SAMPLE");
  else if(filterMotivo==="non-sample") displayed=displayed.filter(r=>r.skipReason!=="SAMPLE");
  else if(filterMotivo==="keep-old") displayed=displayed.filter(r=>r.isKeepOld===true);
  else if(filterMotivo==="non-food") displayed=displayed.filter(r=>r.skipReason==="NON FOOD");
  if(search){const q=search.toLowerCase();displayed=displayed.filter(r=>r.description?.toLowerCase().includes(q)||r.itemCode?.toLowerCase().includes(q)||r.nHK?.toLowerCase().includes(q)||r.location?.toLowerCase().includes(q));}
  displayed=displayed.filter(r=>!r.description?.toUpperCase().includes("FREIGHT"));
  displayed=displayed.filter(r=>r.qty>0||r.isSample);
  if(!showNoAna&&filterMotivo!=="anagrafica") displayed=displayed.filter(r=>r.skipReason!=="NON IN ANAGRAFICA");
  if(excludeSample) displayed=displayed.filter(r=>r.skipReason!=="SAMPLE");
  if(filterMotivo!=="non-food") displayed=displayed.filter(r=>r.skipReason!=="NON FOOD");
  displayed=displayed.filter(r=>r.itemCode!=="ITEM"&&r.itemCode!=="item");
  if(excludeKeepOld) displayed=displayed.filter(r=>!r.isKeepOld);
  if(filterDeltaPct!=="all"){
    if(filterDeltaPct==="pos3") displayed=displayed.filter(r=>r.pct!=null&&r.pct>=3);
    else displayed=displayed.filter(r=>r.pct!=null&&r.pct<=-3);
  }
  if(sortBy==="delta"){
    displayed=[...displayed].sort((a,b)=>sortDir==="desc"?(b.pct??0)-(a.pct??0):(a.pct??0)-(b.pct??0));
  }

  // ── STEPS IMPORT ──────────────────────────────────────────────────────────
  if(step==="map") return(
    <div>
      <PageHeader title="📋 Fatture & Costi · Mappatura" sub={`${fileName} · ${rawRows.length} righe`}/>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:"12px",marginBottom:"20px"}}>
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
      <PageHeader title={`Fatture & Costi · ${branch}`} sub={`${enriched.length} righe · ${fileName||"dati caricati"}`} srcKey={`fatture_${branch}`}/>
      {activeRows.some((r:any)=>r._fromBC) && (
        <BcBanner title="Dati caricati automaticamente da BC IFB Italia">
          Le righe fattura sono importate direttamente da <b style={{color:T.text}}>Business Central IFB Italia</b> (Item Ledger Entry). La colonna <b style={{color:T.text}}>SC BC</b> mostra lo Standard Cost a sistema su BC Brightview; <b style={{color:T.text}}>Δ SC</b> è la differenza rispetto al costo calcolato dall'app.
        </BcBanner>
      )}

      {/* Mismatch banner — compact */}
      {mismatches.length>0&&(
        <div style={{display:"flex",alignItems:"center",gap:"8px",marginBottom:"8px",fontSize:"10px",color:T.orange,opacity:0.7}}>
          <span>⚠ {mismatches.filter((r:any)=>r.isAir&&!r.locationIsNCJ).length} AIR senza NCJ · {mismatches.filter((r:any)=>!r.isAir&&r.locationIsNCJ).length} NCJ ma SEA</span>
          <button onClick={()=>setFilterTransport(v=>v==="mismatch"?"all":"mismatch")}
            style={{padding:"1px 8px",background:"none",color:T.purple,border:`1px solid ${T.purple}44`,borderRadius:"3px",cursor:"pointer",fontSize:"9px"}}>
            {filterTransport==="mismatch"?"Tutte":"Filtra"}
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
            "Qty":r.qty||"","Prezzo Unit.":r.unitPrice||"",
            ...(branch==="CAN"?{"UOM IFB":r.uomVendita||"","UOM Acq.":r.uomAcquisto||""}:{}),
            "Location":r.location||"",
            "Mismatch":r.mismatch?"⚠ "+( r.isAir&&!r.locationIsNCJ?"AIR senza NCJ":"NCJ ma SEA"):"",
            "Mag./Trasp.":r.isAir?"AIR":r.ubicazione||"",
            "Old SC":r.oldHkd!=null?roundN(r.oldHkd):"",
            ...(branch==="CAN"
              ? {"SC GC/TF":r.scGC!=null?roundN(r.scGC):"MANCANTE",
                 "SC FUE/LAN":r.scFUE!=null?roundN(r.scFUE):"MANCANTE",
                 "SC NAV GC/TF":r.scBcGcTf!=null&&r.scBcGcTf>0?roundN(r.scBcGcTf):"",
                 "SC NAV FUE/LAN":r.scBcFueLan!=null&&r.scBcFueLan>0?roundN(r.scBcFueLan):""}
              : {"New SC":r.isAir?"AIR":(r.unitPrice===0||r.unitPrice===0.01)?"SAMPLE":r.newHkd!=null?roundN(r.newHkd):"MANCANTE",
                 "SC BC":r.bcStdCost>0?roundN(r.bcStdCost):"",
                 "Δ SC":r.deltaSC!=null?roundN(r.deltaSC):""}),
            "Δ%":r.pct!=null?roundN(r.pct,1):"","Motivo":r.skipReason||"",
          })),
          "Fatture & Costi",`Fatture_${branch}.xlsx`,
          branch==="CAN" ? ["N COMIT","IFB No"] : [branchN(branch),"IFB No"]
        )} style={{padding:"6px 14px",background:`${T.green}20`,border:`1px solid ${T.green}44`,borderRadius:"6px",color:T.green,cursor:"pointer",fontSize:"11px"}}>
          ⬇ Export Excel
        </button>
        {branch!=="CAN"&&<button onClick={()=>setExcludeAir(v=>!v)}
          style={{padding:"6px 14px",background:excludeAir?`${T.orange}20`:T.surface,color:excludeAir?T.orange:T.muted,border:`1px solid ${excludeAir?T.orange:T.border}`,borderRadius:"6px",cursor:"pointer",fontSize:"11px",fontWeight:excludeAir?"bold":"normal"}}>
          {excludeAir?`✓ AIR esclusi (${airCount})`:`✈ Escludi AIR (${airCount})`}
        </button>}
        <button onClick={()=>setLast30(v=>!v)}
          style={{padding:"6px 14px",background:last30?`${T.blue}25`:T.surface,color:last30?T.blue:T.muted,border:`1px solid ${last30?T.blue:T.border}`,borderRadius:"6px",cursor:"pointer",fontSize:"11px",fontWeight:last30?"bold":"normal"}}>
          {last30?"✓ Ultimi 30gg":"📅 Ultimi 30gg"}
        </button>
        <button onClick={()=>setDedup(v=>!v)}
          style={{padding:"6px 14px",background:dedup?`${T.purple}25`:T.surface,color:dedup?T.purple:T.muted,border:`1px solid ${dedup?T.purple:T.border}`,borderRadius:"6px",cursor:"pointer",fontSize:"11px",fontWeight:dedup?"bold":"normal"}}>
          {dedup?"✓ Senza duplicati":"⧉ Senza duplicati"}
        </button>
        <button onClick={()=>setShowNoAna(v=>!v)}
          style={{padding:"6px 14px",background:showNoAna?`${T.red}20`:T.surface,color:showNoAna?T.red:T.muted,border:`1px solid ${showNoAna?T.red+"66":T.border}`,borderRadius:"6px",cursor:"pointer",fontSize:"11px",fontWeight:showNoAna?"bold":"normal"}}>
          {showNoAna?"✓ Non in Ana.":"👁 Non in Ana."}
        </button>
        <button onClick={()=>setExcludeKeepOld(v=>!v)}
          style={{padding:"6px 14px",background:excludeKeepOld?`${T.orange}20`:T.surface,color:excludeKeepOld?T.orange:T.muted,border:`1px solid ${excludeKeepOld?T.orange:T.border}`,borderRadius:"6px",cursor:"pointer",fontSize:"11px",fontWeight:excludeKeepOld?"bold":"normal"}}>
          {excludeKeepOld?"✓ Keep Old esclusi":"⏰ Escludi Keep Old"}
        </button>
        <button onClick={()=>setExcludeSample(v=>!v)}
          style={{padding:"6px 14px",background:excludeSample?`${T.purple}20`:T.surface,color:excludeSample?T.purple:T.muted,border:`1px solid ${excludeSample?T.purple:T.border}`,borderRadius:"6px",cursor:"pointer",fontSize:"11px",fontWeight:excludeSample?"bold":"normal"}}>
          {excludeSample?"✓ Sample esclusi":"🧪 Escludi Sample"}
        </button>
        <button onClick={()=>window.location.reload()}
          style={{padding:"6px 14px",background:T.surface,color:T.muted,border:`1px solid ${T.border}`,borderRadius:"6px",cursor:"pointer",fontSize:"11px"}}>
          🔄 Ricarica
        </button>
        <button onClick={()=>{ if(sortBy==="date") setSortDir(d=>d==="desc"?"asc":"desc"); else{ setSortBy("date"); setSortDir("desc"); } }}
          style={{padding:"6px 14px",background:sortBy==="date"?`${T.blue}22`:T.surface,color:sortBy==="date"?T.blue:T.muted,border:`1px solid ${sortBy==="date"?T.blue:T.border}`,borderRadius:"6px",cursor:"pointer",fontSize:"11px"}}>
          Data {sortBy==="date"?(sortDir==="desc"?"↓":"↑"):""}
        </button>
        <button onClick={()=>{ if(sortBy==="delta") setSortDir(d=>d==="desc"?"asc":"desc"); else{ setSortBy("delta"); setSortDir("desc"); } }}
          style={{padding:"6px 14px",background:sortBy==="delta"?`${T.gold}22`:T.surface,color:sortBy==="delta"?T.gold:T.muted,border:`1px solid ${sortBy==="delta"?T.gold:T.border}`,borderRadius:"6px",cursor:"pointer",fontSize:"11px"}}>
          Δ% {sortBy==="delta"?(sortDir==="desc"?"↓":"↑"):""}
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
              ["mismatch",`⚠ Mismatch (${mismatches.length})`,T.purple],
            ]) as [string,string,string][]).map(([v,l,c])=>(
          <button key={v} onClick={()=>setFilterTransport(v)}
            style={{padding:"5px 12px",background:filterTransport===v?`${c}20`:T.surface,color:filterTransport===v?c:T.muted,border:`1px solid ${filterTransport===v?c:T.border}`,borderRadius:"6px",cursor:"pointer",fontSize:"11px",fontWeight:filterTransport===v?"bold":"normal"}}>
            {l}
          </button>
        ))}
      </div>

      <SearchBar value={search} onChange={setSearch} placeholder={`🔍 Cerca codice, ${branchN(branch)}, descrizione, location…`}/>

      <div style={{display:"flex",gap:"12px",marginBottom:"10px",alignItems:"center",flexWrap:"wrap"}}>
        <span style={{fontSize:"11px",color:T.muted}}>{branch==="CAN"?"SC NAV:":"SC BC:"}</span>
        <select value={filterScBC} onChange={e=>setFilterScBC(e.target.value as any)}
          style={{background:filterScBC!=="all"?`${T.gold}22`:T.surface,color:filterScBC!=="all"?T.gold:T.muted,border:`1px solid ${filterScBC!=="all"?T.gold:T.border}`,borderRadius:"6px",padding:"5px 10px",fontSize:"11px",cursor:"pointer",outline:"none"}}>
          <option value="all">{branch==="CAN"?"SC NAV: Tutte":"SC BC: Tutte"}</option>
          <option value="assente">{branch==="CAN"?"SC NAV: vuoti (—)":"SC BC: assente (—)"}</option>
        </select>
        <span style={{fontSize:"11px",color:T.muted}}>Motivo:</span>
        <select value={filterMotivo} onChange={e=>setFilterMotivo(e.target.value as any)}
          style={{background:filterMotivo!=="all"?`${T.orange}22`:T.surface,color:filterMotivo!=="all"?T.orange:T.muted,border:`1px solid ${filterMotivo!=="all"?T.orange:T.border}`,borderRadius:"6px",padding:"5px 10px",fontSize:"11px",cursor:"pointer",outline:"none"}}>
          <option value="all">Motivo: Tutti</option>
          <option value="no-log">Senza Logistica</option>
          <option value="no-price">Senza Prezzo</option>
          <option value="anagrafica">Non in Anagrafica</option>
          <option value="sample">Sample</option>
          <option value="keep-old">Keep Old</option>
        </select>
        <span style={{fontSize:"11px",color:T.muted}}>Δ%:</span>
        {(["all","pos3","neg3"] as const).map(v=>{
          const active=filterDeltaPct===v;
          const col=v==="pos3"?T.red:v==="neg3"?T.blue:T.text;
          const label=v==="all"?"Tutti":v==="pos3"?">+3%":"<-3%";
          return(
            <button key={v} onClick={()=>setFilterDeltaPct(v)}
              style={{padding:"5px 10px",background:active?`${col}22`:T.surface,color:active?col:T.muted,border:`1px solid ${active?col:T.border}`,borderRadius:"6px",cursor:"pointer",fontSize:"11px",fontWeight:active?"bold":"normal"}}>
              {label}
            </button>
          );
        })}
      </div>

      <Section title={`${displayed.length} righe`}>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse"}}>
            <thead><tr>
              {["Data",branchN(branch)+" ▾","IFB No ▾","Descrizione","Qty","Prezzo",...(branch==="CAN"?["UOM IFB","UOM Acq."]:[] as any),"Location ▾","Mag./Trasp.","Old SC",...(branch==="CAN"?["SC GC/TF ▾","SC FUE/LAN","SC NAV GC/TF ▾","SC NAV FUE/LAN","Δ GC/TF","Δ FUE/LAN"]:["New SC ▾","SC BC","Δ SC"]),"Δ%","Motivo"].map((c,ci)=>{
                const narrowW = ci===0?"80px":ci===1?"80px":ci===2?"70px":undefined;
                if(c===branchN(branch)+" ▾") return(
                  <th key={c} style={{padding:"4px 4px",background:T.card,borderBottom:`1px solid ${T.border}`,position:"sticky",top:0,zIndex:10,width:"62px",maxWidth:"62px"}}>
                    <select value={filterNHK} onChange={e=>setFilterNHK(e.target.value)}
                      style={{background:filterNHK?`${T.gold}22`:T.card,color:filterNHK?T.gold:T.muted,border:`1px solid ${filterNHK?T.gold:T.border}`,borderRadius:"4px",padding:"2px 4px",fontSize:"10px",cursor:"pointer",fontFamily:"inherit",outline:"none",maxWidth:"60px",width:"60px"}}>
                      <option value="">{branchN(branch)} ▾</option>
                      <option value="__MISSING__">❌ Senza {branchN(branch)}</option>
                      {uniqueNHK.map(v=><option key={v} value={v}>{v}</option>)}
                    </select>
                  </th>
                );
                if(c==="IFB No ▾") return(
                  <th key={c} style={{padding:"4px 4px",background:T.card,borderBottom:`1px solid ${T.border}`,position:"sticky",top:0,zIndex:10,width:"54px",maxWidth:"54px"}}>
                    <select value={filterIFBNo} onChange={e=>setFilterIFBNo(e.target.value)}
                      style={{background:filterIFBNo?`${T.gold}22`:T.card,color:filterIFBNo?T.gold:T.muted,border:`1px solid ${filterIFBNo?T.gold:T.border}`,borderRadius:"4px",padding:"2px 4px",fontSize:"10px",cursor:"pointer",fontFamily:"inherit",outline:"none",maxWidth:"52px",width:"52px"}}>
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
                if(c==="Location ▾") return(
                  <th key={c} style={{padding:"4px 8px",background:T.card,borderBottom:`1px solid ${T.border}`,position:"sticky",top:0,zIndex:10}}>
                    <select value={filterLocation} onChange={e=>setFilterLocation(e.target.value as any)}
                      style={{background:filterLocation!=="all"?`${T.gold}22`:T.card,color:filterLocation!=="all"?T.gold:T.muted,border:`1px solid ${filterLocation!=="all"?T.gold:T.border}`,borderRadius:"4px",padding:"3px 6px",fontSize:"10px",cursor:"pointer",fontFamily:"inherit",outline:"none"}}>
                      <option value="all">Location ▾</option>
                      <option value="ncj">NCJ</option>
                      <option value="non-ncj">Non NCJ</option>
                    </select>
                  </th>
                );
                if(c==="SC GC/TF ▾") return(
                  <th key={c} style={{padding:"4px 8px",background:T.card,borderBottom:`1px solid ${T.border}`,position:"sticky",top:0,zIndex:10}}>
                    <select value={scFilter} onChange={e=>setScFilter(e.target.value as any)}
                      style={{background:scFilter!=="all"?`${T.gold}22`:T.card,color:scFilter!=="all"?T.gold:T.muted,border:`1px solid ${scFilter!=="all"?T.gold:T.border}`,borderRadius:"4px",padding:"3px 6px",fontSize:"10px",cursor:"pointer",fontFamily:"inherit",outline:"none"}}>
                      <option value="all">SC GC/TF ▾</option>
                      <option value="ok">✅ Con costo</option>
                      <option value="mancante">❌ MANCANTE</option>
                      <option value="sample">📦 SAMPLE</option>
                    </select>
                  </th>
                );
                if(c==="SC NAV GC/TF ▾") return(
                  <th key={c} style={{padding:"4px 8px",background:T.card,borderBottom:`1px solid ${T.border}`,position:"sticky",top:0,zIndex:10}}>
                    <select value={filterScNavGC} onChange={e=>setFilterScNavGC(e.target.value as any)}
                      style={{background:filterScNavGC!=="all"?`${T.gold}22`:T.card,color:filterScNavGC!=="all"?T.gold:T.muted,border:`1px solid ${filterScNavGC!=="all"?T.gold:T.border}`,borderRadius:"4px",padding:"3px 6px",fontSize:"10px",cursor:"pointer",fontFamily:"inherit",outline:"none"}}>
                      <option value="all">SC NAV GC/TF ▾</option>
                      <option value="assente">❌ Mancanti (—)</option>
                    </select>
                  </th>
                );
                return <th key={c} style={{padding:"3px 6px",background:T.card,color:T.muted,textAlign:"left",borderBottom:`1px solid ${T.border}`,fontSize:"10px",fontWeight:"normal",whiteSpace:"nowrap",position:"sticky",top:0,zIndex:10,width:narrowW,maxWidth:narrowW,minWidth:ci===3?"140px":undefined}}>{c}</th>;
              })}
            </tr></thead>
            <tbody>
              {displayed.slice(0,1000).map((r:any,i:number)=>{
                const mismatchType=r.mismatch?(r.isAir&&!r.locationIsNCJ?"AIR senza NCJ":"NCJ ma SEA"):"";
                return(
                  <tr key={i} style={{borderBottom:`1px solid ${T.border}`,background:r.mismatch?`${T.purple}10`:i%2===0?T.bg:T.surface}}>
                    <td style={{padding:"2px 4px",fontSize:"9px",fontFamily:"monospace",whiteSpace:"nowrap",width:"74px",maxWidth:"74px"}}><span style={{color:T.gold,fontWeight:"bold"}}>{r.date||"—"}</span></td>
                    <td style={{padding:"2px 4px",fontSize:"9px",fontFamily:"monospace",whiteSpace:"nowrap",width:"60px",maxWidth:"60px",overflow:"hidden",textOverflow:"ellipsis"}}><span style={{color:T.muted}}>{r.nHK||"—"}</span></td>
                    <td style={{padding:"2px 4px",fontSize:"9px",fontFamily:"monospace",whiteSpace:"nowrap",width:"52px",maxWidth:"52px",overflow:"hidden",textOverflow:"ellipsis"}}><span style={{color:T.gold}}>{r.ifbNo||r.itemCode||"—"}</span></td>
                    <td style={{padding:"2px 4px",fontSize:"10px",minWidth:"140px",maxWidth:"160px",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={r.description}>
                      <span style={{color:r._prodFound===false?T.orange:T.text}}>{r.description}</span>
                    </td>
                    <td style={{padding:"2px 4px",fontSize:"9px",fontFamily:"monospace",textAlign:"right",width:"44px",maxWidth:"44px"}}><span style={{color:T.muted}}>{r.qty||"—"}</span></td>
                    <td style={{padding:"2px 4px",fontSize:"9px",fontFamily:"monospace",textAlign:"right",width:"50px",maxWidth:"50px"}}>
                      {r.isSample?<Chip label="S" color={T.purple}/>:<span style={{color:T.muted}}>{r.unitPrice>0?r.unitPrice.toFixed(2):"—"}</span>}
                    </td>
                    {branch==="CAN"&&<td style={{padding:"2px 4px",fontSize:"9px",fontFamily:"monospace",textAlign:"center",width:"40px",maxWidth:"40px"}}><span style={{color:T.blue}}>{r.uomVendita||"—"}</span></td>}
                    {branch==="CAN"&&<td style={{padding:"2px 4px",fontSize:"9px",fontFamily:"monospace",textAlign:"center",width:"40px",maxWidth:"40px"}}><span style={{color:T.gold}}>{r.uomAcquisto||"—"}</span></td>}
                    <td style={{padding:"2px 4px",fontSize:"9px",fontFamily:"monospace",width:"50px",maxWidth:"50px",overflow:"hidden",textOverflow:"ellipsis"}}><span style={{color:r.mismatch?T.purple:T.muted}}>{r.location||"—"}</span></td>
                    <td style={{padding:"2px 4px",fontSize:"9px",width:"44px",maxWidth:"44px"}}>
                      {r.isAir
                        ? <Chip label="AIR" color={r.locationIsNCJ?T.green:T.orange}/>
                        : <Chip label={r.ubicazione||"—"} color={r.ubicazione==="FOR"?T.purple:r.ubicazione==="MTS"?T.blue:T.green}/>
                      }
                    </td>
                    <td style={{padding:"2px 4px",fontSize:"9px",fontFamily:"monospace",textAlign:"right",width:"50px",maxWidth:"50px"}}><span style={{color:T.muted}}>{r.oldHkd!=null?r.oldHkd.toFixed(2):"—"}</span></td>
                    {branch==="CAN" ? (
                      <>
                        {([r.scGC,r.scFUE] as (number|null)[]).map((v,i)=>(
                          <td key={i} style={{padding:"2px 4px",fontSize:"9px",fontFamily:"monospace",textAlign:"right",width:"56px",maxWidth:"56px"}}>
                            {(r.unitPrice===0||r.unitPrice===0.01)
                              ? <span style={{color:T.purple,fontWeight:"bold"}}>SAMP.</span>
                              : v!=null
                                ? <span style={{color:T.gold,fontWeight:"bold"}}>{v.toFixed(2)}</span>
                                : <span style={{color:T.red,fontWeight:"bold"}}>MANC.</span>
                            }
                          </td>
                        ))}
                        {([r.scBcGcTf,r.scBcFueLan] as (number|null)[]).map((v,i)=>(
                          <td key={`bc${i}`} style={{padding:"2px 4px",fontSize:"9px",fontFamily:"monospace",textAlign:"right",width:"50px",maxWidth:"50px"}}>
                            <span style={{color:T.muted}}>{v!=null&&v>0?v.toFixed(2):"—"}</span>
                          </td>
                        ))}
                        {([r.deltaGC,r.deltaFUE] as (number|null)[]).map((v,i)=>(
                          <td key={`dcan${i}`} style={{padding:"2px 4px",fontSize:"9px",fontFamily:"monospace",textAlign:"right",width:"46px",maxWidth:"46px"}}>
                            {v!=null
                              ? <span style={{color:v>0.5?T.red:v<-0.5?T.green:T.text,fontWeight:Math.abs(v)>0.5?"bold":"normal"}}>
                                  {v>0?"+":""}{v.toFixed(2)}
                                </span>
                              : <span style={{color:T.dim}}>—</span>}
                          </td>
                        ))}
                      </>
                    ) : (
                      <>
                        <td style={{padding:"2px 4px",fontSize:"9px",fontFamily:"monospace",textAlign:"right",width:"62px",maxWidth:"62px"}}>
                          {r.isAir
                            ? <span style={{color:T.orange,fontWeight:"bold"}}>AIR</span>
                            : (r.unitPrice===0||r.unitPrice===0.01)
                              ? <span style={{color:T.purple,fontWeight:"bold"}}>SAMP.</span>
                              : r.newHkd!=null
                                ? <span style={{color:T.gold,fontWeight:"bold"}}>HKD {r.newHkd.toFixed(2)}</span>
                                : <span style={{color:T.red,fontWeight:"bold"}}>MANC.</span>
                          }
                        </td>
                        <td style={{padding:"2px 4px",fontSize:"9px",fontFamily:"monospace",textAlign:"right",width:"56px",maxWidth:"56px"}}>
                          <span style={{color:T.muted}}>{r.bcStdCost>0?r.bcStdCost.toFixed(2):"—"}</span>
                        </td>
                        <td style={{padding:"2px 4px",fontSize:"9px",fontFamily:"monospace",textAlign:"right",width:"46px",maxWidth:"46px"}}>
                          {r.deltaSC!=null
                            ? <span style={{color:r.deltaSC>1?T.red:r.deltaSC<-1?T.green:T.text,fontWeight:Math.abs(r.deltaSC)>1?"bold":"normal"}}>
                                {r.deltaSC>0?"+":""}{r.deltaSC.toFixed(2)}
                              </span>
                            : <span style={{color:T.dim}}>—</span>}
                        </td>
                      </>
                    )}
                    <td style={{padding:"2px 4px",fontSize:"9px",textAlign:"right",width:"42px",maxWidth:"42px"}}>
                      {r.pct!=null?<span style={{color:r.pct>3?T.red:r.pct<-3?T.green:T.text,fontWeight:Math.abs(r.pct)>3?"bold":"normal"}}>{r.pct>0?"+":""}{r.pct.toFixed(1)}%</span>:<span style={{color:T.dim}}>—</span>}
                    </td>
                    <td style={{padding:"2px 4px",fontSize:"9px",maxWidth:"90px",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                      {r.isKeepOld?<span style={{color:T.orange,fontWeight:"bold"}}>KEEP OLD</span>:r.skipReason?<span style={{color:T.orange,fontStyle:"italic"}}>{r.skipReason}</span>:<span style={{color:T.dim}}>—</span>}
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
function PriceExceptions({branch, products, xrefs, priceExceptions, setPriceExceptions, canConvFactors=[], setCanConvFactors=(_:any)=>{}, hkConvFactors=[], setHkConvFactors=(_:any)=>{}}) {
  const [search, setSearchRaw] = useState(()=>psGet(`pg_${branch}_excepts_search`,""));
  const setSearch=(v:string)=>{setSearchRaw(v);psSet(`pg_${branch}_excepts_search`,v);};
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
    <div style={{padding:"0"}}>
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
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:"10px"}}>
            <thead>
              <tr style={{background:T.surface}}>
                {["IFB No", isCAN?"N COMIT":"N HK", "Descrizione", "Prezzo (€/unit)", "Nota", "·"].map(h=>(
                  <th key={h} style={{padding:"3px 6px",textAlign:"left",fontSize:"9px",letterSpacing:"1px",
                    color:T.muted,textTransform:"uppercase",fontWeight:"normal",borderBottom:`1px solid ${T.border}`}}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {branchExc.map((exc:any, i:number)=>(
                <tr key={i} style={{borderTop:`1px solid ${T.border}`}}>
                  <td style={{padding:"3px 6px",color:T.gold,fontFamily:"monospace"}}>{exc.productId}</td>
                  <td style={{padding:"3px 6px",color:T.muted,fontFamily:"monospace"}}>{exc.nHK||"-"}</td>
                  <td style={{padding:"3px 6px",maxWidth:"180px",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{exc.description||exc.code||"-"}</td>
                  <td style={{padding:"3px 6px",color:T.green,fontWeight:"bold",fontFamily:"monospace"}}>€ {Number(exc.price).toFixed(4)}</td>
                  <td style={{padding:"3px 6px",color:T.muted,fontStyle:"italic"}}>{exc.note||"-"}</td>
                  <td style={{padding:"3px 6px",textAlign:"center"}}>
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

      {/* ── Sezione HK: Fattori di Conversione ── */}
      {!isCAN && (
        <div style={{marginTop:"36px"}}>
          <div style={{fontSize:"10px",letterSpacing:"3px",color:T.gold,textTransform:"uppercase",marginBottom:"4px"}}>Hong Kong · Solo</div>
          <h2 style={{margin:"0 0 6px",fontSize:"18px",fontWeight:"bold"}}>⚖️ Fattori di Conversione</h2>
          <div style={{color:T.muted,fontSize:"12px",marginBottom:"18px"}}>
            Moltiplicatore applicato al prezzo listino <em>e</em> divisore del calcolo unità/pallet.
            Usare per articoli dove il fornitore esprime PCS/BOX in modo diverso dall'anagrafica.
          </div>
          <HkConvFactorForm hkConvFactors={hkConvFactors} setHkConvFactors={setHkConvFactors} />
          <div style={{background:T.card,border:`1px solid ${T.border}`,borderRadius:"10px",overflow:"hidden",marginTop:"16px"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:"11px"}}>
              <thead>
                <tr style={{background:T.surface}}>
                  {["N HK","IFB N","Fattore","Descrizione","·"].map(h=>(
                    <th key={h} style={{padding:"6px 10px",textAlign:"left",fontSize:"9px",letterSpacing:"1px",
                      color:T.muted,textTransform:"uppercase",fontWeight:"normal",borderBottom:`1px solid ${T.border}`}}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {hkConvFactors.map((row:any, i:number)=>{
                  const ifbN = xrefs.find((x:any)=>x.nHK===row.nHK)?.ifbNo || "-";
                  return (
                  <tr key={i} style={{borderTop:`1px solid ${T.border}`}}>
                    <td style={{padding:"5px 10px",color:T.gold,fontFamily:"monospace",fontWeight:"bold"}}>{row.nHK}</td>
                    <td style={{padding:"5px 10px",color:T.muted,fontFamily:"monospace"}}>{ifbN}</td>
                    <td style={{padding:"5px 10px",color:row.factor===1?T.dim:T.green,fontFamily:"monospace",fontWeight:"bold"}}>
                      ×{Number(row.factor).toLocaleString("it-IT",{maximumFractionDigits:6})}
                    </td>
                    <td style={{padding:"5px 10px",color:T.text}}>{row.description||"-"}</td>
                    <td style={{padding:"5px 10px",textAlign:"center"}}>
                      <button
                        onClick={()=>setHkConvFactors((prev:any)=>prev.filter((_:any,j:number)=>j!==i))}
                        style={{background:"transparent",border:`1px solid ${T.red||"#c55"}`,color:T.red||"#c55",
                          borderRadius:"4px",padding:"2px 8px",cursor:"pointer",fontSize:"11px",fontFamily:"inherit"}}>
                        Rimuovi
                      </button>
                    </td>
                  </tr>
                  );
                })}
                {hkConvFactors.length===0 && (
                  <tr><td colSpan={5} style={{padding:"24px",textAlign:"center",color:T.dim}}>Nessun fattore di conversione.</td></tr>
                )}
              </tbody>
            </table>
          </div>
          <button
            onClick={()=>setHkConvFactors(HK_CONV_DEFAULTS)}
            style={{marginTop:"10px",padding:"5px 12px",background:"transparent",border:`1px solid ${T.border}`,
              color:T.muted,borderRadius:"6px",cursor:"pointer",fontSize:"11px",fontFamily:"inherit"}}>
            Ripristina default
          </button>
        </div>
      )}

      {/* ── Sezione CAN: Fattori di Conversione ── */}
      {isCAN && (
        <div style={{marginTop:"36px"}}>
          <div style={{fontSize:"10px",letterSpacing:"3px",color:T.gold,textTransform:"uppercase",marginBottom:"4px"}}>Canarie · Solo</div>
          <h2 style={{margin:"0 0 6px",fontSize:"18px",fontWeight:"bold"}}>⚖️ Fattori di Conversione</h2>
          <div style={{color:T.muted,fontSize:"12px",marginBottom:"18px"}}>
            Moltiplicatore applicato al prezzo listino <em>e</em> divisore del calcolo unità/pallet (colonna CM del modello Excel).
            Usare per articoli dove il fornitore esprime PCS/BOX in modo diverso dalla anagrafica.
          </div>

          {/* Form aggiunta/modifica */}
          <CanConvFactorForm
            canConvFactors={canConvFactors}
            setCanConvFactors={setCanConvFactors}
            xrefs={xrefs}
            products={products}
          />

          {/* Tabella */}
          <div style={{background:T.card,border:`1px solid ${T.border}`,borderRadius:"10px",overflow:"hidden",marginTop:"16px"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:"11px"}}>
              <thead>
                <tr style={{background:T.surface}}>
                  {["N COMIT","IFB N","Fattore","Descrizione","·"].map(h=>(
                    <th key={h} style={{padding:"6px 10px",textAlign:"left",fontSize:"9px",letterSpacing:"1px",
                      color:T.muted,textTransform:"uppercase",fontWeight:"normal",borderBottom:`1px solid ${T.border}`}}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {canConvFactors.map((row:any, i:number)=>{
                  const ifbN = xrefs.find((x:any)=>x.nHK===row.nComit)?.ifbNo || "-";
                  return (
                  <tr key={i} style={{borderTop:`1px solid ${T.border}`}}>
                    <td style={{padding:"5px 10px",color:T.gold,fontFamily:"monospace",fontWeight:"bold"}}>{row.nComit}</td>
                    <td style={{padding:"5px 10px",color:T.muted,fontFamily:"monospace"}}>{ifbN}</td>
                    <td style={{padding:"5px 10px",color:row.factor===1?T.dim:T.green,fontFamily:"monospace",fontWeight:"bold"}}>
                      ×{Number(row.factor).toLocaleString("it-IT",{maximumFractionDigits:6})}
                    </td>
                    <td style={{padding:"5px 10px",color:T.text,maxWidth:"240px",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{row.description||"-"}</td>
                    <td style={{padding:"5px 10px",textAlign:"center"}}>
                      <button
                        onClick={()=>setCanConvFactors((prev:any)=>prev.filter((_:any,j:number)=>j!==i))}
                        style={{background:"transparent",border:`1px solid ${T.red||"#c55"}`,color:T.red||"#c55",
                          borderRadius:"4px",padding:"2px 8px",cursor:"pointer",fontSize:"11px",fontFamily:"inherit"}}>
                        Rimuovi
                      </button>
                    </td>
                  </tr>
                  );
                })}
                {canConvFactors.length===0 && (
                  <tr><td colSpan={5} style={{padding:"24px",textAlign:"center",color:T.dim}}>Nessun fattore di conversione.</td></tr>
                )}
              </tbody>
            </table>
          </div>

          <button
            onClick={()=>setCanConvFactors(CAN_CONV_DEFAULTS)}
            style={{marginTop:"10px",padding:"5px 12px",background:"transparent",border:`1px solid ${T.border}`,
              color:T.muted,borderRadius:"6px",cursor:"pointer",fontSize:"11px",fontFamily:"inherit"}}>
            Ripristina default
          </button>
        </div>
      )}
    </div>
  );
}

function HkConvFactorForm({hkConvFactors, setHkConvFactors}:any) {
  const [nHK,   setNHK]   = useState("");
  const [factor, setFactor] = useState("");
  const [desc,   setDesc]   = useState("");

  function save() {
    const nk = nHK.trim();
    const f  = parseFloat(factor.replace(",","."));
    if(!nk || isNaN(f) || f <= 0) return;
    setHkConvFactors((prev:any)=>{
      const idx = prev.findIndex((r:any)=>r.nHK===nk);
      const entry = {nHK:nk, factor:f, description:desc.trim()};
      if(idx>=0){ const u=[...prev]; u[idx]=entry; return u; }
      return [...prev, entry];
    });
    setNHK(""); setFactor(""); setDesc("");
  }

  const exists = hkConvFactors.some((r:any)=>r.nHK===nHK.trim());

  return (
    <div style={{background:T.card,border:`1px solid ${T.border}`,borderRadius:"10px",padding:"16px 18px"}}>
      <div style={{fontSize:"11px",letterSpacing:"2px",color:T.gold,textTransform:"uppercase",marginBottom:"12px"}}>Aggiungi / Modifica fattore</div>
      <div style={{display:"grid",gridTemplateColumns:"160px 140px 1fr auto",gap:"10px",alignItems:"end"}}>
        <div>
          <div style={{fontSize:"10px",color:T.muted,marginBottom:"4px"}}>N HK (es. GCMA-1015)</div>
          <input value={nHK} onChange={e=>setNHK(e.target.value)} placeholder="es. GCMA-1015"
            style={{width:"100%",background:T.surface,border:`1px solid ${T.border}`,borderRadius:"6px",
              padding:"7px 10px",color:T.text,fontSize:"12px",fontFamily:"inherit",boxSizing:"border-box"}} />
        </div>
        <div>
          <div style={{fontSize:"10px",color:T.muted,marginBottom:"4px"}}>Fattore (×)</div>
          <input value={factor} onChange={e=>setFactor(e.target.value)} placeholder="es. 3"
            type="number" min="0" step="any"
            style={{width:"100%",background:T.surface,border:`1px solid ${T.border}`,borderRadius:"6px",
              padding:"7px 10px",color:T.text,fontSize:"12px",fontFamily:"inherit",boxSizing:"border-box"}} />
        </div>
        <div>
          <div style={{fontSize:"10px",color:T.muted,marginBottom:"4px"}}>Descrizione (opzionale)</div>
          <input value={desc} onChange={e=>setDesc(e.target.value)} placeholder="es. VFF08"
            style={{width:"100%",background:T.surface,border:`1px solid ${T.border}`,borderRadius:"6px",
              padding:"7px 10px",color:T.text,fontSize:"12px",fontFamily:"inherit",boxSizing:"border-box"}} />
        </div>
        <button onClick={save} disabled={!nHK||!factor}
          style={{padding:"7px 16px",background:nHK&&factor?T.gold:"#333",
            color:nHK&&factor?"#111":T.dim,border:"none",borderRadius:"6px",
            cursor:nHK&&factor?"pointer":"default",fontFamily:"inherit",fontSize:"12px",fontWeight:"bold"}}>
          {exists?"Aggiorna":"Aggiungi"}
        </button>
      </div>
    </div>
  );
}

function CanConvFactorForm({canConvFactors, setCanConvFactors, xrefs, products}:any) {
  const [nComit, setNComit]       = useState("");
  const [factor, setFactor]       = useState("");
  const [desc, setDesc]           = useState("");

  function save() {
    const nc = nComit.trim();
    const f  = parseFloat(factor.replace(",","."));
    if(!nc || isNaN(f) || f <= 0) return;
    setCanConvFactors((prev:any)=>{
      const idx = prev.findIndex((r:any)=>r.nComit===nc);
      const entry = {nComit:nc, factor:f, description:desc.trim()};
      if(idx>=0){ const u=[...prev]; u[idx]=entry; return u; }
      return [...prev, entry];
    });
    setNComit(""); setFactor(""); setDesc("");
  }

  // Auto-fill desc from xref when nComit matches
  function onNComitChange(val:string) {
    setNComit(val);
    const xr = xrefs.find((x:any)=>x.nHK===val.trim());
    if(xr){
      const prod = products.find((p:any)=>p.id===xr.productId||p.code===xr.ifbNo);
      if(prod) setDesc(prod.description||"");
    }
  }

  const exists = canConvFactors.some((r:any)=>r.nComit===nComit.trim());

  return (
    <div style={{background:T.card,border:`1px solid ${T.border}`,borderRadius:"10px",padding:"16px 18px"}}>
      <div style={{fontSize:"11px",letterSpacing:"2px",color:T.gold,textTransform:"uppercase",marginBottom:"12px"}}>Aggiungi / Modifica fattore</div>
      <div style={{display:"grid",gridTemplateColumns:"120px 140px 1fr auto",gap:"10px",alignItems:"end"}}>
        <div>
          <div style={{fontSize:"10px",color:T.muted,marginBottom:"4px"}}>N COMIT</div>
          <input value={nComit} onChange={e=>onNComitChange(e.target.value)} placeholder="es. 7231"
            style={{width:"100%",background:T.surface,border:`1px solid ${T.border}`,borderRadius:"6px",
              padding:"7px 10px",color:T.text,fontSize:"12px",fontFamily:"inherit",boxSizing:"border-box"}} />
        </div>
        <div>
          <div style={{fontSize:"10px",color:T.muted,marginBottom:"4px"}}>Fattore (×)</div>
          <input value={factor} onChange={e=>setFactor(e.target.value)} placeholder="es. 0.33"
            type="number" min="0" step="any"
            style={{width:"100%",background:T.surface,border:`1px solid ${T.border}`,borderRadius:"6px",
              padding:"7px 10px",color:T.text,fontSize:"12px",fontFamily:"inherit",boxSizing:"border-box"}} />
        </div>
        <div>
          <div style={{fontSize:"10px",color:T.muted,marginBottom:"4px"}}>Descrizione (opzionale)</div>
          <input value={desc} onChange={e=>setDesc(e.target.value)} placeholder="Descrizione articolo"
            style={{width:"100%",background:T.surface,border:`1px solid ${T.border}`,borderRadius:"6px",
              padding:"7px 10px",color:T.text,fontSize:"12px",fontFamily:"inherit",boxSizing:"border-box"}} />
        </div>
        <button onClick={save} disabled={!nComit||!factor}
          style={{padding:"7px 16px",background:nComit&&factor?T.gold:"#333",
            color:nComit&&factor?"#111":T.dim,border:"none",borderRadius:"6px",
            cursor:nComit&&factor?"pointer":"default",fontFamily:"inherit",fontSize:"12px",fontWeight:"bold"}}>
          {exists?"Aggiorna":"Aggiungi"}
        </button>
      </div>
    </div>
  );
}

// ─── MAIL GEN ─────────────────────────────────────────────────────────────────
// Only shows items with |delta| > 3% (point 7)
// ─── SC ATTUALI ───────────────────────────────────────────────────────────────
function ScAttualiPage({scAttuali, setScAttuali, scHistory, setScHistory, branch, showToast}) {
  const isCAN = branch === "CAN";
  const [step, setStep] = useState<"main"|"map"|"preview">("main");
  const [fileName, setFileName] = useState("");
  const [rawHeaders, setRawHeaders] = useState<string[]>([]);
  const [rawRows, setRawRows] = useState<any[][]>([]);
  const [colMap, setColMap] = useState<Record<string,string>>({});
  const [preview, setPreview] = useState<any[]>([]);
  const [search, setSearchRaw] = useState(()=>psGet(`pg_${branch}_scatt_search`,""));
  const setSearch=(v:string)=>{setSearchRaw(v);psSet(`pg_${branch}_scatt_search`,v);};
  const [importMonth, setImportMonth] = useState(()=>{
    const d = new Date(); return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`;
  });
  const [selEntry, setSelEntry] = useState<any>(null);
  const [histOpen, setHistOpen] = useState(true);

  // Campi da mappare
  const MAP_FIELDS: {key:string; label:string; required?:boolean; canOnly?:boolean; hkOnly?:boolean}[] = [
    {key:"code",            label:"N COMIT (codice articolo)",required:true},
    {key:"ifbCode",         label:"IFB N",                   canOnly:true},
    {key:"description",     label:"Descrizione"},
    {key:"executionDate",   label:"Execution Date",          canOnly:true},
    {key:"lastSC",          label:isCAN?"Standard Cost corrente":"Last Standard Cost"},
    {key:"fifoUnit",        label:"FIFO Unit Cost"},
    {key:"scGC",            label:"SC GC / TF",              canOnly:true},
    {key:"scLan",           label:"SC FUE / LAN",            canOnly:true},
    {key:"salesLast3m",     label:"Vendite ultimi 3 mesi",   hkOnly:true},
    {key:"lastPurchaseDate",label:"Last Purchase Date"},
    {key:"stockQty",        label:"Stock Qty"},
  ].filter(f => isCAN ? !f.hkOnly : !f.canOnly);

  function loadFile(file: File) {
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (e:any) => {
      try {
        const wb = XLSX.read(e.target.result, {type:"binary"});
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data: any[][] = XLSX.utils.sheet_to_json(ws, {header:1, defval:""});
        if (data.length < 2) { showToast("File vuoto", T.red); return; }
        // Trova riga intestazione
        let hi = 0;
        for (let i=0; i<Math.min(8, data.length); i++) {
          const rn = data[i].map((c:any)=>String(c||"").toLowerCase());
          if (rn.some((c:string)=>c.includes("item no")||c.includes("item no.")||c.includes("no_")||c.includes("standard cost")||c.includes("last standard")||c.includes("codice"))) { hi=i; break; }
        }
        const hdrs = data[hi].map((c:any)=>String(c||"").trim()).filter(h=>h!=="");
        const rows = data.slice(hi+1).filter((r:any[])=>r.some((c:any)=>c!==""));
        setRawHeaders(hdrs);
        setRawRows(rows);
        // Auto-detect suggerimenti
        const norm = (s:string) => s.toLowerCase().replace(/[\s_%()/]/g,"");
        const suggest: Record<string,string> = {};
        const HINTS: Record<string,string[]> = {
          code:             ["ncomit","n comit","no_","itemno","item no","codice","code"],
          ifbCode:          ["ifbn","ifb n","ifbno","ifb no","ifbitem","ifb item"],
          description:      ["description","descrizione","desc"],
          executionDate:    ["executiondate","execution date","execdate","data"],
          lastSC:           ["standardcost","standard cost","laststandard","last standard"],
          fifoUnit:         ["unitcost","unit cost","fifoda item","fifo da item"],
          scGC:             ["scgrancanaria","scgrancan","grancanaria","gran canaria","grancan","scgc","gctf","gc"],
          scLan:            ["sclanzarote","sclanza","lanzarote","lanza","sclan","scfue","fue","lantf"],
          salesLast3m:      ["saleslast","sales last","vendite"],
          lastPurchaseDate: ["lastpurchase","last purchase"],
          stockQty:         ["stockqty","stock quantity","stock","giacenza"],
        };
        MAP_FIELDS.forEach(f=>{
          const hints = HINTS[f.key]||[];
          const found = hdrs.find(h=>hints.some(hint=>norm(h).includes(hint)));
          if(found) suggest[f.key]=found;
        });
        setColMap(suggest);
        setStep("map");
      } catch(err:any) { showToast("Errore lettura file: "+err.message, T.red); }
    };
    reader.readAsBinaryString(file);
  }

  function buildPreview() {
    const num = (v:any) => typeof v==="number" ? v : parseFloat(String(v||"").replace(/[€$,\s]/g,""))||0;
    const str = (v:any) => String(v||"").trim();
    const idx = (field:string) => rawHeaders.indexOf(colMap[field]||"");
    const get = (row:any[], field:string) => { const i=idx(field); return i>=0?row[i]:""; };
    // Converte execution date in valore comparabile (Excel serial number o stringa)
    const parseDate = (v:any): number => {
      if(!v) return 0;
      if(typeof v==="number") return v; // Excel serial date
      const s = String(v).trim();
      // Prova formato DD/MM/YYYY o DD-MM-YYYY
      const m = s.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
      if(m) return parseInt(m[3])*10000 + parseInt(m[2])*100 + parseInt(m[1]);
      const d = new Date(s); return isNaN(d.getTime()) ? 0 : d.getTime();
    };
    const parsed = rawRows.map((row:any[])=>{
      const code = str(get(row,"code"));
      if(!code) return null;
      return {
        code,
        ifbCode:          str(get(row,"ifbCode")),
        description:      str(get(row,"description")),
        executionDate:    str(get(row,"executionDate")),
        _execDateVal:     parseDate(get(row,"executionDate")),
        lastSC:           num(get(row,"lastSC")),
        fifoUnit:         num(get(row,"fifoUnit")),
        scGC:             num(get(row,"scGC")),
        scLan:            num(get(row,"scLan")),
        salesLast3m:      num(get(row,"salesLast3m")),
        lastPurchaseDate: str(get(row,"lastPurchaseDate")),
        stockQty:         num(get(row,"stockQty")),
      };
    }).filter(Boolean);
    if(!parsed.length) { showToast("Nessuna riga valida", T.red); return; }
    // CAN: dedup per codice — tieni solo la riga con execution date più recente
    let deduped = parsed;
    if(isCAN && idx("executionDate") >= 0) {
      const byCode: Record<string,any> = {};
      parsed.forEach((r:any) => {
        const existing = byCode[r.code];
        if(!existing || r._execDateVal > existing._execDateVal) byCode[r.code] = r;
      });
      deduped = Object.values(byCode);
    }
    setPreview(deduped);
    setStep("preview");
  }

  function executeImport() {
    const cleanRows = preview.map(({_execDateVal:_, ...r}:any)=>r);
    const entry = {
      id: Date.now(),
      month: importMonth,
      fileName,
      date: new Date().toISOString(),
      branch,
      count: cleanRows.length,
      rows: cleanRows,
    };
    setScAttuali(cleanRows);
    setDataSource(`scattuali_${branch}`,"manual");
    const newHist = [entry, ...scHistory].slice(0, 24);
    setScHistory(newHist);
    IDB.set(`ifb_schistory_${branch}`, newHist);
    setSelEntry(entry);
    showToast(`SC Attuali ${importMonth}: ${cleanRows.length} articoli importati ✓`, T.gold);
    setStep("main");
    setPreview([]);
  }

  function clearAll() {
    if(!window.confirm(`Svuotare tutti i dati SC Attuali per ${branch}?`)) return;
    setScAttuali([]);
    CLOUD.set(`ifb_scattuali_${branch}`,[]);
    setSelEntry(null);
    showToast("SC Attuali svuotati", T.orange);
  }

  function executeImport() {
    const entry = {
      id: Date.now(),
      month: importMonth,
      fileName,
      date: new Date().toISOString(),
      branch,
      count: preview.length,
      rows: preview,
    };
    setScAttuali(preview);
    setDataSource(`scattuali_${branch}`,"manual");
    const newHist = [entry, ...scHistory].slice(0, 24); // max 24 snapshot
    setScHistory(newHist);
    IDB.set(`ifb_schistory_${branch}`, newHist);
    setSelEntry(entry);
    showToast(`SC Attuali ${importMonth}: ${preview.length} articoli importati ✓`, T.gold);
    setStep("main");
    setPreview([]);
  }

  const isHKReport = branch !== "CAN";
  const viewRows = selEntry ? selEntry.rows : (step==="preview" ? preview : scAttuali);
  const displayed = viewRows.filter((r:any)=>{
    if(!search) return true;
    const q = search.toLowerCase();
    return String(r.code||"").toLowerCase().includes(q)
      || String(r.ifbCode||"").toLowerCase().includes(q)
      || String(r.description||"").toLowerCase().includes(q);
  });

  function renderTable(rows: any[]) {
    return (
      <div style={{overflowX:"auto"}}>
        <table style={{borderCollapse:"collapse",width:"100%"}}>
          <THead cols={isCAN
            ? ["N COMIT","IFB N","Descrizione","SC Standard €","FIFO unit €","SC GC/TF €","SC LAN/FUE €","Since"]
            : ["Codice","Descrizione","SC Attuale HK$","FIFO unit HK$","Vendite 3m","Last Purchase","Stock Qty"]}
          />
          <tbody>
            {rows.slice(0,400).map((r:any,i:number)=>(
              <tr key={i} style={{borderBottom:`1px solid ${T.border}22`,background:i%2?"transparent":`${T.surface}33`}}>
                <td style={{padding:"3px 6px",fontSize:"10px",color:T.text,fontFamily:"monospace",whiteSpace:"nowrap"}}>{r.code}</td>
                {isCAN&&<td style={{padding:"3px 6px",fontSize:"10px",color:T.muted,fontFamily:"monospace",whiteSpace:"nowrap"}}>{r.ifbCode||"—"}</td>}
                <td style={{padding:"3px 6px",fontSize:"10px",color:T.muted,maxWidth:"200px",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.description}</td>
                <td style={{padding:"3px 6px",fontSize:"10px",color:T.gold,textAlign:"right",fontWeight:"bold",whiteSpace:"nowrap"}}>{r.lastSC>0?`${isCAN?"€":"HK$"} ${r.lastSC.toFixed(2)}`:"—"}</td>
                <td style={{padding:"3px 6px",fontSize:"10px",color:T.muted,textAlign:"right",whiteSpace:"nowrap"}}>{r.fifoUnit>0?`${isCAN?"€":"HK$"} ${r.fifoUnit.toFixed(4)}`:"—"}</td>
                {isCAN&&<td style={{padding:"3px 6px",fontSize:"10px",color:T.blue,textAlign:"right",fontWeight:"bold",whiteSpace:"nowrap"}}>{r.scGC>0?`€ ${r.scGC.toFixed(2)}`:"—"}</td>}
                {isCAN&&<td style={{padding:"3px 6px",fontSize:"10px",color:T.blue,textAlign:"right",fontWeight:"bold",whiteSpace:"nowrap"}}>{r.scLan>0?`€ ${r.scLan.toFixed(2)}`:"—"}</td>}
                {isCAN&&<td style={{padding:"3px 6px",fontSize:"10px",color:T.dim,textAlign:"right",whiteSpace:"nowrap"}}>{r.executionDate||"—"}</td>}
                {!isCAN&&<td style={{padding:"3px 6px",fontSize:"10px",color:T.muted,textAlign:"right",whiteSpace:"nowrap"}}>{r.salesLast3m?r.salesLast3m.toFixed(0):"—"}</td>}
                {!isCAN&&<td style={{padding:"3px 6px",fontSize:"10px",color:T.muted,textAlign:"right",whiteSpace:"nowrap"}}>{r.lastPurchaseDate||"—"}</td>}
                {!isCAN&&<td style={{padding:"3px 6px",fontSize:"10px",color:T.muted,textAlign:"right",whiteSpace:"nowrap"}}>{r.stockQty!=null?r.stockQty:"—"}</td>}
              </tr>
            ))}
          </tbody>
        </table>
        {rows.length>400&&<div style={{padding:"6px",fontSize:"11px",color:T.muted,textAlign:"center"}}>Mostrati 400 su {rows.length}</div>}
      </div>
    );
  }

  return (
    <div>
      <PageHeader title={`📊 SC Attuali · ${branch}`}
        sub={scHistory.length>0 ? `${scHistory.length} import salvati` : "Nessun report salvato"}
        srcKey={`scattuali_${branch}`}/>

      {/* ── IMPORT ── */}
      {step==="map" ? (
        <Section title={`Mappa colonne — ${fileName} (${rawHeaders.length} colonne trovate)`}>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(280px,1fr))",gap:"10px",marginBottom:"14px"}}>
            {MAP_FIELDS.map(f=>(
              <div key={f.key} style={{display:"flex",flexDirection:"column",gap:"3px"}}>
                <label style={{fontSize:"10px",color:f.required?T.gold:T.muted,fontWeight:f.required?"bold":"normal"}}>
                  {f.label}{f.required?" *":""}
                </label>
                <select value={colMap[f.key]||""} onChange={e=>setColMap(m=>({...m,[f.key]:e.target.value}))}
                  style={{...inputStyle(),fontFamily:"monospace",fontSize:"11px"}}>
                  <option value="">— non mappare —</option>
                  {rawHeaders.map((h,i)=><option key={i} value={h}>{h}</option>)}
                </select>
              </div>
            ))}
          </div>
          <div style={{marginBottom:"10px",fontSize:"11px",color:T.dim}}>
            Anteprima prime 3 righe del file:
            <div style={{overflowX:"auto",marginTop:"6px"}}>
              <table style={{borderCollapse:"collapse",fontSize:"10px"}}>
                <thead><tr>{rawHeaders.map((h,i)=><th key={i} style={{padding:"2px 6px",background:T.surface,color:T.muted,border:`1px solid ${T.border}`,whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
                <tbody>{rawRows.slice(0,3).map((row,ri)=><tr key={ri}>{rawHeaders.map((_,ci)=><td key={ci} style={{padding:"2px 6px",color:T.text,border:`1px solid ${T.border}22`,whiteSpace:"nowrap",maxWidth:"120px",overflow:"hidden",textOverflow:"ellipsis"}}>{String(row[ci]??"")} </td>)}</tr>)}</tbody>
              </table>
            </div>
          </div>
          <div style={{display:"flex",gap:"10px"}}>
            <ActionBtn label="← Annulla" onClick={()=>{setStep("main");setRawHeaders([]);setRawRows([]);}}/>
            <ActionBtn label={`Anteprima →`} onClick={buildPreview} primary disabled={!colMap["code"]}/>
          </div>
        </Section>
      ) : step==="preview" ? (
        <Section title={`Anteprima — ${fileName} · ${preview.length} articoli`}>
          <div style={{display:"flex",gap:"12px",alignItems:"center",flexWrap:"wrap",marginBottom:"14px"}}>
            <div style={{display:"flex",flexDirection:"column",gap:"4px"}}>
              <label style={{fontSize:"10px",color:T.muted}}>Mese di valenza</label>
              <input value={importMonth} onChange={e=>setImportMonth(e.target.value)}
                placeholder="es. 2026-06"
                style={{...inputStyle(),width:"140px",fontFamily:"monospace"}}/>
            </div>
            <div style={{display:"flex",flexDirection:"column",gap:"4px"}}>
              <label style={{fontSize:"10px",color:T.muted}}>File</label>
              <span style={{fontSize:"12px",color:T.muted,fontFamily:"monospace"}}>{fileName}</span>
            </div>
          </div>
          <div style={{display:"flex",gap:"10px"}}>
            <ActionBtn label="← Mappa" onClick={()=>setStep("map")}/>
            <ActionBtn label={`✓ Salva come ${importMonth} (${preview.length} art.)`} onClick={executeImport} primary/>
          </div>
        </Section>
      ) : (
        <Section title="Carica report SC da NAV / BC">
          <div style={{display:"flex",gap:"12px",alignItems:"center",flexWrap:"wrap"}}>
            <label style={{display:"inline-block",padding:"8px 16px",background:`${T.gold}22`,border:`1px solid ${T.gold}44`,borderRadius:"6px",cursor:"pointer",fontSize:"12px",color:T.gold}}>
              📂 Carica Report SC ({branch})
              <input type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={e=>e.target.files?.[0]&&loadFile(e.target.files[0])}/>
            </label>
            {scAttuali.length>0&&(
              <button onClick={clearAll} style={{padding:"8px 14px",background:`${T.red}18`,border:`1px solid ${T.red}44`,borderRadius:"6px",cursor:"pointer",fontSize:"12px",color:T.red}}>
                🗑 Svuota SC Attuali
              </button>
            )}
            <div style={{fontSize:"11px",color:T.muted,lineHeight:"1.6"}}>
              {branch==="CAN"
                ? "NAV: Item No · STANDARD COST · SC GRANCANARIA · SC LANZAROTE (poi mappi le colonne)"
                : "BC: Item No · Last Standard Cost · FIFO · Sales 3m · Stock Qty (poi mappi le colonne)"}
            </div>
          </div>
        </Section>
      )}

      {/* ── STORICO LIST ── */}
      {scHistory.length>0&&(
        <Section title="">
          <button onClick={()=>setHistOpen(v=>!v)}
            style={{display:"flex",alignItems:"center",gap:"8px",background:"none",border:"none",cursor:"pointer",padding:"0",marginBottom:histOpen?"12px":"0"}}>
            <span style={{fontSize:"11px",fontWeight:"bold",color:T.muted,letterSpacing:"0.08em",textTransform:"uppercase"}}>
              {histOpen?"▾":"▸"} Storico import ({scHistory.length})
            </span>
          </button>
          {histOpen&&(
            <div style={{display:"flex",flexDirection:"column",gap:"4px"}}>
              {scHistory.map((e:any)=>{
                const isSel = selEntry?.id===e.id;
                return (
                  <div key={e.id} style={{display:"flex",alignItems:"center",gap:"8px",padding:"6px 10px",
                    background:isSel?`${T.gold}18`:T.card,
                    border:`1px solid ${isSel?T.gold:T.border}`,
                    borderRadius:"6px",cursor:"pointer"}}
                    onClick={()=>{ setSelEntry(isSel?null:e); setSearch(""); }}>
                    <span style={{fontFamily:"monospace",fontSize:"12px",color:isSel?T.gold:T.text,fontWeight:"bold",minWidth:"70px"}}>{e.month}</span>
                    <span style={{fontSize:"11px",color:T.muted,flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={e.fileName}>
                      📄 {e.fileName}
                    </span>
                    <span style={{fontSize:"10px",color:T.dim,whiteSpace:"nowrap"}}>
                      {new Date(e.date).toLocaleDateString("it-IT")} · {e.count} art.
                    </span>
                    <button onClick={ev=>{ev.stopPropagation();
                      if(!window.confirm(`Eliminare import ${e.month}?`)) return;
                      const nh=scHistory.filter((x:any)=>x.id!==e.id);
                      setScHistory(nh); IDB.set(`ifb_schistory_${branch}`,nh);
                      if(isSel) setSelEntry(null);
                    }} style={{background:"none",border:"none",color:T.muted,cursor:"pointer",fontSize:"12px",padding:"2px 4px"}}>✕</button>
                  </div>
                );
              })}
            </div>
          )}
        </Section>
      )}

      {/* ── TABELLA DATI ── */}
      {viewRows.length>0&&(
        <Section title={selEntry
          ? `${selEntry.month} · ${selEntry.fileName} · ${selEntry.count} articoli`
          : step==="preview"?"Anteprima":"SC Attuali correnti"}>
          <div style={{display:"flex",alignItems:"center",gap:"10px",marginBottom:"10px"}}>
            <input placeholder="🔍 Cerca articolo..." value={search} onChange={e=>setSearch(e.target.value)}
              style={{...inputStyle(),width:"260px"}}/>
            {search&&<span style={{fontSize:"11px",color:T.muted}}>{displayed.length} risultati</span>}
            {search&&<button onClick={()=>setSearch("")} style={{fontSize:"11px",color:T.muted,background:"none",border:"none",cursor:"pointer"}}>✕</button>}
          </div>
          {renderTable(displayed)}
        </Section>
      )}
    </div>
  );
}

// ─── CHECK MENSILE ────────────────────────────────────────────────────────────
function CheckMensile({costRows, branch, salesRows, xrefs, scAttuali, products, logistics=[]}) {
  const isCAN = branch === "CAN";
  const cur = isCAN ? "€" : "HK$";
  // Converte seriale Excel o stringa data → "YYYY-MM-DD"
  const toIsoDate = (v:any): string => {
    if(!v && v!==0) return "";
    if(typeof v==="number") return new Date((v-25569)*86400000).toISOString().slice(0,10);
    const s = String(v).trim();
    if(/^\d{5,6}$/.test(s)) return new Date((parseInt(s)-25569)*86400000).toISOString().slice(0,10);
    return s.slice(0,10);
  };
  const xrefByIfbNoPreferNumeric = (ifbNo:string) => {
    const matches=(xrefs||[]).filter((x:any)=>String(x.ifbNo)===ifbNo);
    return matches.find((x:any)=>/^\d+$/.test(String(x.nHK)))||matches[0]||null;
  };

  const availableMonths = useMemo(()=>{
    const s = new Set<string>();
    (salesRows||[]).forEach((r:any)=>{
      const d = toIsoDate(r.date||r.postingDate||"").slice(0,7);
      if(d&&d.length===7) s.add(d);
    });
    return [...s].sort().reverse();
  },[salesRows]);

  const [selectedMonth, setSelectedMonth] = useState<string>("");
  const [threshold, setThreshold] = useState(3);

  useEffect(()=>{
    if(!selectedMonth && availableMonths.length>0) setSelectedMonth(availableMonths[0]);
  },[availableMonths]);

  // xref: nHK (=N COMIT per CAN) → ifbNo
  const xrefByNFiliale = useMemo(()=>{
    const m: Record<string,string>={};
    (xrefs||[]).forEach((x:any)=>{ if(x.nHK&&x.ifbNo) m[x.nHK]=x.ifbNo; });
    return m;
  },[xrefs]);

  // costMap: keyed by IFB code (priorità IFB per listini)
  const costMap = useMemo(()=>{
    const m: Record<string,any>={};
    (costRows||[]).forEach((r:any)=>{
      if(r.code) m[r.code]=r;
      if(r.id) m[String(r.id)]=r;
    });
    return m;
  },[costRows]);

  // scMap con alias bidirezionali IFB↔N COMIT via xref + ifbCode diretto
  const scMap = useMemo(()=>{
    const m: Record<string,any>={};
    (scAttuali||[]).forEach((r:any)=>{ if(r.code!=null) m[String(r.code)]=r; });
    // CAN: ifbCode salvato direttamente nel record
    (scAttuali||[]).forEach((r:any)=>{ if(r.ifbCode) { const k=String(r.ifbCode); if(!m[k]) m[k]=r; } });
    (scAttuali||[]).forEach((rec:any)=>{
      if(rec.code==null) return;
      const k=String(rec.code);
      const xrByNHK=(xrefs||[]).find((x:any)=>String(x.nHK)===k);
      if(xrByNHK?.ifbNo){ const ak=String(xrByNHK.ifbNo); if(!m[ak]) m[ak]=rec; }
      const matches=(xrefs||[]).filter((x:any)=>String(x.ifbNo)===k);
      const xrByIfb=matches.find((x:any)=>/^\d+$/.test(String(x.nHK)))||matches[0];
      if(xrByIfb?.nHK){ const ak=String(xrByIfb.nHK); if(!m[ak]) m[ak]=rec; }
    });
    return m;
  },[scAttuali, xrefs]);

  // logMap keyed by productId
  const logMap = useMemo(()=>{
    const m: Record<string,any>={};
    (logistics||[]).forEach((l:any)=>{ if(l.productId) m[String(l.productId)]=l; });
    return m;
  },[logistics]);

  const [rollingDays, setRollingDays] = useState(30);
  const monthRows = useMemo(()=>{
    const cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - rollingDays);
    const cutoffStr = cutoff.toISOString().slice(0,10);
    return (salesRows||[]).filter((r:any)=>{
      const d = toIsoDate(r.date||r.postingDate||"");
      return d >= cutoffStr;
    });
  },[salesRows, rollingDays]);

  const analysisRows = useMemo(()=>{
    if(!monthRows.length) return [];
    const seen = new Set<string>();
    const rows: any[] = [];
    for (const inv of monthRows) {
      const nFiliale = inv.itemCode;
      if(!nFiliale) continue;
      // Escludi SAMPLE (price=0), AIR
      if(inv.unitPrice===0||inv.unitPrice===0.01) continue;
      if(inv.transport==="AIR") continue;
      const sCode = String(nFiliale);
      // Risolvi xref bidirezionale
      const nComitFromIfb2 = isCAN ? String(xrefByIfbNoPreferNumeric(sCode)?.nHK||"") : "";
      const ifbFromNComit2 = isCAN ? String((xrefs||[]).find((x:any)=>String(x.nHK)===sCode)?.ifbNo||"") : "";
      // Per CAN: nFiliale=IFB code, ifbNo=N COMIT (da xref ifbNo→nHK)
      // Per HK:  nFiliale=N HK,    ifbNo=IFB code (da xref nHK→ifbNo)
      const ifbNo = isCAN
        ? (nComitFromIfb2 || nFiliale)
        : (xrefByNFiliale[nFiliale] || ifbFromNComit2 || nFiliale);
      // CAN: ifbFromNComit2 = IFB code quando sCode è N COMIT numerico (es. "4141" → "MT310")
      const resolvedIFB = isCAN ? (ifbFromNComit2 || (nComitFromIfb2 ? nFiliale : "")) : "";
      const sameCode = isCAN ? (ifbNo === (ifbFromNComit2 || nFiliale)) : (ifbNo===nFiliale);
      // Chiave canonica = N COMIT numerico se disponibile, altrimenti IFB code
      // Questo risolve doppia codifica (es. MT314 e 4144 = stesso item): sempre il numero vince
      const isNumeric = (s:string) => /^\d+$/.test(s);
      const numericNComit = isCAN && isNumeric(sCode) ? sCode : (isCAN && isNumeric(nComitFromIfb2) ? nComitFromIfb2 : "");
      const canonicalKey = numericNComit || ifbNo || nFiliale;
      if(seen.has(canonicalKey)) continue;
      seen.add(canonicalKey);
      // CAN: sCode può essere N COMIT (es. "4141") → cerca anche via findProduct (usa p.nHK) e via IFB risolto
      const prod = findProduct(sCode, products, xrefs)
        || findProduct(resolvedIFB, products, xrefs)
        || (products||[]).find((p:any)=>p.code===ifbNo||p.code===nFiliale||p.code===resolvedIFB);
      // Salta righe non-articolo (FREIGHT, servizi, ecc.) — non in anagrafica e senza cost row
      if(!prod && !costMap[ifbNo] && !costMap[nFiliale] && !costMap[resolvedIFB]) continue;
      const scEntry = isCAN
        ? (scMap[sCode] || scMap[nComitFromIfb2] || scMap[ifbFromNComit2] || scMap[prod?.code||""] || scMap[ifbNo])
        : (scMap[prod?.nHK||""] || scMap[sCode] || scMap[prod?.code||""] || scMap[ifbNo]);
      const cr = costMap[ifbNo] || costMap[nFiliale] || (isCAN ? (costMap[resolvedIFB] || costMap[prod?.code||""]) : null);
      // Escludi NON FOOD CAN (HO.RE.CA. SUPPLY)
      if(isCAN && /^HO\.RE\.CA\./i.test(String(prod?.category||""))) continue;
      const logEntry = prod ? logMap[String(prod.id)] : null;
      const lastOrderRaw2 = logEntry?.lastOrderDate || inv.date || inv.postingDate;
      const lastOrderD = lastOrderRaw2 ? new Date(toIsoDate(lastOrderRaw2)) : null;
      const isKeepOld = lastOrderD ? ((Date.now()-lastOrderD.getTime())/86400000)>180 : false;
      // Usa island flags dalla logistica: default GC/TF se entry mancante (isola principale)
      const destGCTF_cm   = !logEntry || logEntry.isGC || logEntry.isTF;
      const destLANFUE_cm = !logEntry || logEntry.isLAN || logEntry.isFUE;
      const oldSC    = isCAN
        ? (destGCTF_cm
            ? (scEntry?.scGC || scEntry?.lastSC || 0)
            : (scEntry?.scLan || scEntry?.lastSC || 0))
        : (scEntry?.lastSC || 0);
      const newSC    = isCAN
        ? (destGCTF_cm
            ? (cr?.cost?.step2GC || cr?.cost?.step2Eur || 0)
            : (cr?.cost?.step2LAN || cr?.cost?.step2Eur || 0))
        : (cr?.cost?.step2Hkd || 0);
      const deltaAbs = oldSC>0 ? newSC-oldSC : 0;
      const deltaPct = oldSC>0 ? deltaAbs/oldSC*100 : 0;
      // noCalc = DA INSERIRE: newSC non calcolabile (manca logistica, prezzo zero, ecc.) — indipendentemente da oldSC
      const noCalc   = newSC===0 && !isKeepOld;
      // isNuovo = NUOVI ARTICOLI: newSC calcolato presente, ma oldSC (SC BC/NAV) assente
      const isNuovo  = newSC>0 && oldSC===0 && !isKeepOld;
      rows.push({
        nFiliale: (isCAN && resolvedIFB && resolvedIFB!==nFiliale) ? resolvedIFB : nFiliale,
        ifbNo, sameCode,
        description: cr?.description || inv.description || "",
        oldSC, newSC, deltaPct, absDelta:Math.abs(deltaPct), noCalc,
        skipReason: cr?.skipReason || "",
        lastDate: toIsoDate(scEntry?.lastPurchaseDate || inv.date || inv.postingDate || ""),
        stockQty: scEntry?.stockQty ?? "",
        isKeepOld,
        isNuovo,
        isDelta:  oldSC>0 && newSC>0 && Math.abs(deltaPct)>threshold && !isKeepOld,
        isKeepOldOrdered: isKeepOld,
      });
    }
    return rows;
  },[monthRows, xrefByNFiliale, scMap, costMap, products, logMap, isCAN, threshold]);

  const alert1 = analysisRows.filter(r=>r.isNuovo);          // NUOVI ARTICOLI: newSC calcolato, oldSC assente
  const alert2 = analysisRows.filter(r=>r.isDelta);          // TO UPDATE Δ%
  const alert3 = analysisRows.filter(r=>r.isKeepOldOrdered && !r.isNuovo);
  const alert4 = analysisRows.filter(r=>r.noCalc && !r.isKeepOld); // DA INSERIRE: mancanti ma NON keep old

  function exportExcel() {
    const branchCode = isCAN?"COMIT":"HK";
    const today = new Date(); const monthFmt = `${today.getFullYear()}_${String(today.getMonth()+1).padStart(2,"0")}`;
    const all = [
      ...alert1.map((r:any)=>({...r,tipo:"NUOVI ARTICOLI"})),
      ...alert4.map((r:any)=>({...r,tipo:"DA INSERIRE"})),
      ...alert2.map((r:any)=>({...r,tipo:"TO UPDATE (Delta%)"})),
      ...alert3.map((r:any)=>({...r,tipo:"TO UPDATE (Keep Old)"})),
    ];
    const data = all.map((r:any)=>({
      "Tipo":        r.tipo,
      "N COMIT":     isCAN ? (r.sameCode ? (isNumCode(r.nFiliale) ? r.nFiliale : "") : r.ifbNo) : r.nFiliale,
      "IFB No":      isCAN ? (r.sameCode ? (isNumCode(r.nFiliale) ? "" : r.nFiliale) : r.nFiliale) : r.ifbNo,
      "Descrizione": r.description,
      "Old SC":      r.oldSC>0 ? Number(r.oldSC.toFixed(2)) : "",
      "New SC":      r.newSC>0 ? Number(r.newSC.toFixed(2)) : "",
      "Delta %":     r.oldSC>0 ? (r.deltaPct>0?"+":"")+r.deltaPct.toFixed(2)+"%" : "",
      "Last Date":   r.lastDate,
      "Stock Qty":   r.stockQty,
    }));
    exportXLSX(data, "SC_Analisi", `STDC_Analisi_${branchCode}_${monthFmt}.xlsx`, ["N Filiale","IFB"]);
  }

  const TH = ({h}:{h:string}) => <th style={{padding:"3px 8px",fontSize:"9px",color:T.gold,borderBottom:`1px solid ${T.border}`,whiteSpace:"nowrap",textAlign:"left",letterSpacing:"0.5px",textTransform:"uppercase"}}>{h}</th>;
  const tdC  = (v:string) => <td style={{padding:"3px 8px",fontSize:"10px",color:T.text,fontFamily:"monospace",whiteSpace:"nowrap"}}>{v||"—"}</td>;
  const tdM  = (v:any)    => <td style={{padding:"3px 8px",fontSize:"10px",color:T.muted,whiteSpace:"nowrap"}}>{v!=null&&v!==""?v:"—"}</td>;
  const tdD  = (v:string) => <td style={{padding:"3px 8px",fontSize:"10px",color:T.text,maxWidth:"220px",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{v}</td>;
  const tdSC = (v:number) => <td style={{padding:"3px 8px",fontSize:"10px",color:T.gold,textAlign:"right",fontWeight:"bold",whiteSpace:"nowrap"}}>{v>0?`${cur} ${v.toFixed(2)}`:"—"}</td>;
  const tdDp = (pct:number,old:number) => <td style={{padding:"3px 8px",fontSize:"10px",textAlign:"right",fontWeight:"bold",whiteSpace:"nowrap",color:pct>0?T.orange:T.red}}>{old>0?(pct>0?"+":"")+pct.toFixed(2)+"%":"—"}</td>;
  const fCode = isCAN ? "IFB No" : "N HK";
  const hasDualCode = analysisRows.some((r:any)=>!r.sameCode);
  const isNumCode = (s:string) => /^\d+$/.test(s||"");
  // CAN: nFiliale=IFB code (da BC Italia), ifbNo=N COMIT (da xref)
  // Ordine colonne: N COMIT prima, IFB No dopo (coerente con Righe Fatture)
  // sameCode=true: un solo codice — se alfanumerico (es. M4839) → IFB No; se numerico → N COMIT
  const tdCodes = (r:any) => hasDualCode
    ? isCAN
      ? r.sameCode
        ? isNumCode(r.nFiliale)
          ? <>{tdC(r.nFiliale)}{tdC("")}</>
          : <>{tdC("")}{tdC(r.nFiliale)}</>
        : <>{tdC(r.ifbNo)}{tdC(r.nFiliale)}</>
      : <>{tdC(r.nFiliale)}{tdC(r.sameCode?"":r.ifbNo)}</>
    : tdC(r.nFiliale);
  const thCodes = hasDualCode
    ? isCAN
      ? <><TH h="N COMIT"/><TH h="IFB No"/></>
      : <><TH h="N HK"/><TH h="IFB No"/></>
    : <TH h={fCode}/>;

  return (
    <div>
      <PageHeader title={`📅 Check Mensile · ${branch}`} sub="TODO list aggiornamento Standard Cost"/>

      <Section title="Periodo di riferimento">
        <div style={{display:"flex",gap:"16px",alignItems:"flex-end",flexWrap:"wrap"}}>
          <div>
            <label style={{fontSize:"10px",color:T.muted,display:"block",marginBottom:"4px",letterSpacing:"1px",textTransform:"uppercase"}}>Ultimi giorni</label>
            <select value={rollingDays} onChange={e=>setRollingDays(Number(e.target.value))} style={{...inputStyle(),minWidth:"140px",cursor:"pointer"}}>
              {[15,30,45,60,90].map(d=><option key={d} value={d}>Ultimi {d} giorni</option>)}
            </select>
          </div>
          <div>
            <label style={{fontSize:"10px",color:T.muted,display:"block",marginBottom:"4px",letterSpacing:"1px",textTransform:"uppercase"}}>Soglia Δ%</label>
            <input type="number" value={threshold} onChange={e=>setThreshold(Number(e.target.value)||3)}
              min={0} max={50} step={0.5} style={{...inputStyle(),width:"80px"}}/>
          </div>
          {scAttuali.length===0&&(
            <div style={{fontSize:"12px",color:T.orange,padding:"7px 12px",border:`1px solid ${T.orange}44`,borderRadius:"6px"}}>
              ⚠ SC Attuali non caricati
            </div>
          )}
          {analysisRows.length>0&&<div style={{display:"flex",gap:"10px",alignItems:"center"}}>
            <span style={{fontSize:"11px",color:T.muted}}>{monthRows.length} righe · {analysisRows.length} articoli univoci</span>
            <span style={{fontSize:"10px",color:T.dim}}>(ultimi {rollingDays}gg)</span>
            <ActionBtn label="📥 Esporta Excel" onClick={exportExcel} primary disabled={alert1.length+alert2.length+alert3.length+alert4.length===0}/>
          </div>}
        </div>
      </Section>

      {analysisRows.length>0&&<>
        {/* KPI */}
        <div style={{display:"flex",gap:"14px",flexWrap:"wrap",padding:"0 0 18px"}}>
          {([
            {label:"NUOVI ARTICOLI",n:alert1.length,c:T.blue,icon:"🆕"},
            {label:"DA INSERIRE",n:alert4.length,c:T.red,icon:"❌"},
            {label:"TO UPDATE (Δ%)",n:alert2.length,c:T.orange,icon:"⬆"},
            {label:"TO UPDATE (Keep Old)",n:alert3.length,c:T.purple,icon:"♻"},
          ] as {label:string,n:number,c:string,icon:string}[]).map(({label,n,c,icon})=>(
            <div key={label} style={{background:`${c}11`,border:`1px solid ${c}44`,borderRadius:"8px",padding:"10px 18px",minWidth:"140px"}}>
              <div style={{fontSize:"9px",color:c,letterSpacing:"1px",textTransform:"uppercase",marginBottom:"2px"}}>{icon} {label}</div>
              <div style={{fontSize:"24px",fontWeight:"bold",color:n>0?c:T.dim}}>{n}</div>
            </div>
          ))}
        </div>

        {/* ALERT 1 — NUOVI ARTICOLI */}
        <Section title={`🆕 1. NUOVI ARTICOLI — SC calcolato, non ancora in sistema (${alert1.length})`} accent={T.blue}>
          <div style={{fontSize:"11px",color:T.muted,marginBottom:"8px"}}>SC BC/NAV assente ma Standard Cost calcolabile dal listino — da inserire in Business Central.</div>
          <div style={{overflowX:"auto"}}>
            <table style={{borderCollapse:"collapse",width:"max-content",minWidth:"100%"}}>
              <thead><tr>{thCodes}<TH h="Descrizione"/><TH h="New SC Calc"/><TH h="Stock"/><TH h="Last Date"/></tr></thead>
              <tbody>
                {alert1.length===0
                  ? <tr><td colSpan={5+(hasDualCode?1:0)} style={{padding:"10px",fontSize:"11px",color:T.dim,textAlign:"center"}}>Nessun articolo ✓</td></tr>
                  : alert1.map((r:any,i:number)=>(
                    <tr key={i} style={{borderBottom:`1px solid ${T.border}22`,background:`${T.blue}07`}}>
                      {tdCodes(r)}{tdD(r.description)}
                      {tdSC(r.newSC)}{tdM(r.stockQty)}{tdM(r.lastDate)}
                    </tr>
                  ))}
              </tbody>
            </table>
          </div>
        </Section>

        {/* ALERT 4 — DA INSERIRE (newSC non calcolabile) */}
        <Section title={`❌ 2. DA INSERIRE — New SC non calcolabile (${alert4.length})`} accent={T.red}>
          <div style={{fontSize:"11px",color:T.muted,marginBottom:"8px"}}>Standard Cost non calcolabile: manca logistica, prezzo zero, o listino assente. Acquistati nell'ultimo mese — da verificare.</div>
          <div style={{overflowX:"auto"}}>
            <table style={{borderCollapse:"collapse",width:"max-content",minWidth:"100%"}}>
              <thead><tr>{thCodes}<TH h="Descrizione"/><TH h="SC in macchina"/><TH h="Skip Reason"/><TH h="Last Date"/></tr></thead>
              <tbody>
                {alert4.length===0
                  ? <tr><td colSpan={5+(hasDualCode?1:0)} style={{padding:"10px",fontSize:"11px",color:T.dim,textAlign:"center"}}>Nessun articolo ✓</td></tr>
                  : alert4.map((r:any,i:number)=>(
                    <tr key={i} style={{borderBottom:`1px solid ${T.border}22`,background:`${T.red}07`}}>
                      {tdCodes(r)}{tdD(r.description)}
                      <td style={{padding:"3px 8px",fontSize:"10px",color:T.muted,textAlign:"right",whiteSpace:"nowrap"}}>{r.oldSC>0?`${cur} ${r.oldSC.toFixed(2)}`:"—"}</td>
                      <td style={{padding:"3px 8px",fontSize:"10px",color:T.muted,textAlign:"left"}}>{r.skipReason||"—"}</td>
                      {tdM(r.lastDate)}
                    </tr>
                  ))}
              </tbody>
            </table>
          </div>
        </Section>

        {/* ALERT 2 */}
        <Section title={`⬆ 3. TO UPDATE — delta >${threshold}% (${alert2.length})`} accent={T.orange}>
          <div style={{fontSize:"11px",color:T.muted,marginBottom:"8px"}}>SC in macchina vs SC calcolato: variazione oltre la soglia.</div>
          <div style={{overflowX:"auto"}}>
            <table style={{borderCollapse:"collapse",width:"max-content",minWidth:"100%"}}>
              <thead><tr>{thCodes}<TH h="Descrizione"/><TH h="Old SC"/><TH h="New SC"/><TH h="Δ %"/><TH h="Ult. Fattura"/></tr></thead>
              <tbody>
                {alert2.length===0
                  ? <tr><td colSpan={6+(hasDualCode?1:0)} style={{padding:"10px",fontSize:"11px",color:T.dim,textAlign:"center"}}>Nessun articolo ✓</td></tr>
                  : alert2.map((r:any,i:number)=>(
                    <tr key={i} style={{borderBottom:`1px solid ${T.border}22`,background:r.deltaPct>0?`${T.orange}07`:`${T.red}07`}}>
                      {tdCodes(r)}{tdD(r.description)}
                      <td style={{padding:"3px 8px",fontSize:"10px",color:T.muted,textAlign:"right",whiteSpace:"nowrap"}}>{r.oldSC>0?`${cur} ${r.oldSC.toFixed(2)}`:"—"}</td>
                      {tdSC(r.newSC)}{tdDp(r.deltaPct,r.oldSC)}{tdM(r.lastDate)}
                    </tr>
                  ))}
              </tbody>
            </table>
          </div>
        </Section>

        {/* ALERT 3 */}
        <Section title={`♻ 4. TO UPDATE — Keep Old tornato a ordine (${alert3.length})`} accent={T.purple}>
          <div style={{fontSize:"11px",color:T.muted,marginBottom:"8px"}}>Non ordinati da &gt;180 giorni ma presenti nelle fatture di questo mese — SC probabilmente da aggiornare.</div>
          <div style={{overflowX:"auto"}}>
            <table style={{borderCollapse:"collapse",width:"max-content",minWidth:"100%"}}>
              <thead><tr>{thCodes}<TH h="Descrizione"/><TH h="Old SC"/><TH h="New SC"/><TH h="Δ %"/><TH h="Stock"/><TH h="Last Date"/></tr></thead>
              <tbody>
                {alert3.length===0
                  ? <tr><td colSpan={7+(hasDualCode?1:0)} style={{padding:"10px",fontSize:"11px",color:T.dim,textAlign:"center"}}>Nessun articolo ✓</td></tr>
                  : alert3.map((r:any,i:number)=>(
                    <tr key={i} style={{borderBottom:`1px solid ${T.border}22`,background:`${T.purple}07`}}>
                      {tdCodes(r)}{tdD(r.description)}
                      <td style={{padding:"3px 8px",fontSize:"10px",color:T.muted,textAlign:"right",whiteSpace:"nowrap"}}>{r.oldSC>0?`${cur} ${r.oldSC.toFixed(2)}`:"—"}</td>
                      {tdSC(r.newSC)}{tdDp(r.deltaPct,r.oldSC)}{tdM(r.stockQty)}{tdM(r.lastDate)}
                    </tr>
                  ))}
              </tbody>
            </table>
          </div>
        </Section>
      </>}

      {!analysisRows.length&&(
        <div style={{padding:"32px",textAlign:"center",color:T.muted,fontSize:"13px"}}>
          {scAttuali.length===0
            ? "Carica prima il report SC Attuali (pagina SC Attuali)."
            : salesRows.length===0
            ? "Nessuna fattura caricata."
            : `Nessun articolo negli ultimi ${rollingDays} giorni.`}
        </div>
      )}
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
  const[typeFilter,setTypeFilter]=useState("all");

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
        {/* Filtro per tipo */}
        <div style={{display:"flex",gap:"6px",flexWrap:"wrap",marginBottom:"12px"}}>
          {[["all","Tutti"],["anagrafica","Anagrafica"],["prices","Listini"],["xref","XRef"],["air","AIR"],["sales","Fatture"]].map(([v,l])=>{
            const count = v==="all" ? branchSnaps.length : branchSnaps.filter((s:any)=>s.type===v).length;
            if(count===0 && v!=="all") return null;
            return <button key={v} onClick={()=>{setTypeFilter(v);setSel(null);}}
              style={{padding:"4px 10px",background:typeFilter===v?T.gold:T.surface,color:typeFilter===v?"#000":T.muted,
                border:`1px solid ${typeFilter===v?T.gold:T.border}`,borderRadius:"4px",cursor:"pointer",fontSize:"10px",fontWeight:typeFilter===v?"bold":"normal"}}>
              {l} ({count})
            </button>;
          })}
        </div>
        {snapshots.length===0
          ? <div style={{padding:"24px",textAlign:"center",color:T.dim,fontSize:"13px"}}>Nessuno snapshot ancora.</div>
          : <div style={{display:"flex",flexDirection:"column",gap:"6px",maxHeight:"320px",overflowY:"auto"}}>
              {branchSnaps.filter((s:any)=>typeFilter==="all"||s.type===typeFilter).map((s:any)=>(
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

        const isXRef = sel.type==="xref";
        const realModified = isXRef
          ? diffs.filter((d:any)=>!d.isNew&&d.changed)
          : diffs
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
              {showModified&&!isXRef&&<>
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
                  : sel.type==="xref"
                    ? <table style={{width:"100%",borderCollapse:"collapse"}}>
                        <THead cols={[branchN(branch),"IFB No"]} sticky />
                        <tbody>{newItems.map((d:any,i:number)=>(
                          <tr key={i} style={{borderBottom:`1px solid ${T.border}`}}>
                            <TD mono><span style={{color:T.muted}}>{d.nHK}</span></TD>
                            <TD mono><span style={{color:T.gold}}>{d.ifbNo}</span></TD>
                          </tr>
                        ))}</tbody>
                      </table>
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
                  : sel.type==="xref"
                    ? <table style={{width:"100%",borderCollapse:"collapse"}}>
                        <THead cols={[branchN(branch),"IFB No vecchio","IFB No nuovo"]} sticky />
                        <tbody>{diffs.filter((d:any)=>!d.isNew&&d.changed).map((d:any,i:number)=>(
                          <tr key={i} style={{borderBottom:`1px solid ${T.border}`,background:i%2===0?T.bg:T.surface}}>
                            <TD mono><span style={{color:T.muted}}>{d.nHK}</span></TD>
                            <TD mono><span style={{color:T.red,textDecoration:"line-through"}}>{d.oldIFB}</span></TD>
                            <TD mono><span style={{color:T.green}}>{d.ifbNo}</span></TD>
                          </tr>
                        ))}</tbody>
                      </table>
                    : <div style={{overflowX:"auto"}}>
                        <table style={{width:"100%",borderCollapse:"collapse"}}>
                          <THead cols={[`IFB No / ${branchN(branch)}`,"Descrizione","Campo",`Vecchio (${prevDate})`,`Nuovo (${thisDate})`,"Δ%"]} sticky />
                          <tbody>{shownDiffs.map((d:any,i:number)=>
                            d.fields.map((f:any,j:number)=>{
                              const oldR=roundN(f.old||0),newR=roundN(f.new||0);
                              const isPriceF=["fcaPrice","fcaDiscounted","dapPrice","dapDiscounted","mtsPrice","dapFinal"].includes(f.field);
                              const pct=isPriceF&&oldR!==0?(newR-oldR)/Math.abs(oldR)*100:null;
                              return(
                                <tr key={`${i}-${j}`} style={{
                                  borderBottom:j===d.fields.length-1?`1px solid ${T.border}`:`1px solid ${T.border}44`,
                                  background:i%2===0?T.bg:T.surface}}>
                                  {j===0&&<>
                                    <td rowSpan={d.fields.length} style={{padding:"3px 6px",borderBottom:`1px solid ${T.border}`,verticalAlign:"top",fontFamily:"monospace",fontSize:"11px",color:T.gold}}>
                                      {d.ifbNo||d.id}<br/>
                                      <span style={{color:T.muted,fontSize:"10px"}}>{d.nHK||""}</span>
                                    </td>
                                    <td rowSpan={d.fields.length} style={{padding:"3px 6px",borderBottom:`1px solid ${T.border}`,verticalAlign:"top",fontSize:"11px",color:T.text}}>
                                      {d.description}
                                    </td>
                                  </>}
                                  <TD><span style={{color:T.muted,fontSize:"10px"}}>{f.field}</span></TD>
                                  <TD mono><span style={{color:T.red}}>{isPriceF?roundN(f.old||0).toFixed(2):String(f.old??"")}</span></TD>
                                  <TD mono><span style={{color:T.green}}>{isPriceF?roundN(f.new||0).toFixed(2):String(f.new??"")}</span></TD>
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
function Products({ products, setProducts, branch, xrefs=[], importLogs, setImportLogs, snapshots, setSnapshots, showToast, bumpImportTs }) {
  const [search, setSearch] = useState("");
  const [onlyIFB, setOnlyIFB] = useState(true);
  const [sortAna, setSortAna] = useState<"default"|"az"|"za">("default");
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
  const FIELDS_CAN = ["nHK","code","description","category","uom","baseUom","qtyPerBox","boxPerPallet","netWeightPcs","kgPerBox","kgxplt","temperature","aiem","active","vendorName","vendorName2"];
  const FIELDS_MAC = ["nHK","code","description","isHoff","uom","hkUom","standardCostHkd","temperature","kgPerBox","qtyPerBox","active","vendorName"];
  const FIELDS = branch==="CAN" ? FIELDS_CAN : branch==="MAC" ? FIELDS_MAC : FIELDS_HK;

  const FLABELS: any = {
    nHK: branch==="MAC" ? "MACAO No (No_)" : branch==="CAN" ? "N COMIT (No_)" : "N HK (No_)",
    code: "IFB Item / BV No *",
    description: "Descrizione *",
    category: "Section",
    uom: branch==="MAC" ? "UOM vendita MACAO" : "UOM vendita IFB (Sales UOM)",
    baseUom: "UOM acquisto filiale SC NAV GC/TF (Base UOM)",
    qtyPerBox: "Qty/Cartone (per conversione UOM)",
    boxPerPallet: "Cartoni/Pallet",
    netWeightPcs: "Peso Netto PCS (kg/pz) — per conversione €/KG→€/PCS",
    kgPerBox: "Kg per Cartone (per costi logistica)",
    kgxplt: "Kg x PLT",
    temperature: "Product Type (DRY/FRESH/FROZEN)",
    aiem: "★ AIEM % (col. W anagrafica CAN)",
    isHoff: "HOFF Flag (1 = House of Fine Foods)",
    hkUom: "HK/BV UOM (per conversione automatica)",
    standardCostHkd: "★ Standard Cost HK (HKD) — base calcolo MAC",
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
    baseUom:     ["baseunitofmeasure","base unit of measure","baseuom","base uom","uomacquisto","uom acquisto","uombase"],
    qtyPerBox:   ["quantityxpackaging","quantity x packaging"],
    boxPerPallet:["packagingxpallet","packaging x pallet"],
    netWeightPcs:["pesonettopcs","peso netto pcs","netweightpcs","net weight pcs","pesonettopz","peso netto pz","pcsnetweight","pcs net weight","netweightpcs","pcsnettweight"],
    kgPerBox:    ["netweightbox","net weight box","pesonetto","peso netto","pesonetto box","peso netto box","netweight","net weight"],
    kgxplt:      ["kgxplt","kg x pallet","kg per pallet","kgperpallet","kgplt","kgxplt netto","kg x plt netto","kg plt netto"],
    temperature: ["producttype","product type","product type rettificato"],
    aiem:        ["aiem","igic","alim","aiem%","aiem_perc","aiem_canarie","aiemperc"],
    isHoff:      ["ishoff","hoff","hofflag","hoff flag","hoff_flag","is hoff"],
    hkUom:            ["bvsalesunitofmeasure","bv salesunitofmeasure","bvuom","hk uom","hkuom"],
    standardCostHkd:  ["standardcost","standard cost","costo standard","costostandard","sc hkd","schkd","fob","fobprice","fob price"],
    temperature:      ["producttype","product type","producttype rettificato","tipoprodotto"],
    kgPerBox:         ["netweight","net weight","kgperbox","kg per box","kg/box","pesokg","peso netto","pesonetto","peso netto box","net weight box","netweightbox"],
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
    // Se la colonna mappata è "blocked/bloccato" il valore va invertito; se è "active/attivo" va usato direttamente
    const activeColName = (map["active"] || "").toLowerCase();
    const isBlockedCol = activeColName.includes("blocked") || activeColName.includes("bloccato");
    const truthyVals = ["true","1","yes","si","sì","x","y"];
    const parseActive = (v: any) => {
      const s = String(v || "").toLowerCase().trim();
      if (!s) return true; // campo vuoto → attivo per default
      const truthy = truthyVals.includes(s);
      return isBlockedCol ? !truthy : truthy;
    };
    const newProds = preview.map((r: any) => ({
      id: r.code || r.nHK,
      code: r.code,
      nHK: r.nHK,
      description: r.description,
      category: mapBCVal("category", r.category),
      uom: mapBCVal("uom", r.uom),
      baseUom: r.baseUom ? String(r.baseUom).trim().toUpperCase() : "",
      qtyPerBox: parseFloat(String(r.qtyPerBox||"").replace(",",".")) || 0,
      boxPerPallet: parseFloat(r.boxPerPallet) || 0,
      netWeightPcs: parseFloat(String(r.netWeightPcs||"").replace(",",".")) || 0,
      kgPerBox: (()=>{ const nwp=parseFloat(String(r.netWeightPcs||"").replace(",","."));const qty=parseFloat(String(r.qtyPerBox||"").replace(",","."));return (nwp>0&&qty>0)?roundN(nwp*qty,2):(parseFloat(String(r.kgPerBox||"").replace(",","."))||0); })(),
      temperature: mapBCVal("temperature", r.temperature) || "DRY",
      kgxplt: (()=>{ const nwp=parseFloat(String(r.netWeightPcs||"").replace(",","."));const qty=parseFloat(String(r.qtyPerBox||"").replace(",","."));const plt=parseFloat(r.boxPerPallet)||0;const kpb=(nwp>0&&qty>0)?roundN(nwp*qty,2):(parseFloat(String(r.kgPerBox||"").replace(",","."))||0);return parseFloat(r.kgxplt)>0?parseFloat(r.kgxplt):roundN(kpb*qty*plt,0); })(),
      aiem: parseFloat(r.aiem) || 0,
      isHoff: ["true","1","yes","hoff","si","sì","vero","x"].includes(String(r.isHoff||"").toLowerCase()),
      hkUom: r.hkUom ? String(r.hkUom).trim().toUpperCase() : "",
      standardCostHkd: parseFloat(String(r.standardCostHkd||"").replace(",",".")) || 0,
      active: parseActive(r.active),
      vendorName: r.vendorName || "",
      vendorName2: r.vendorName2 || "",
    }));
  
    // Calcola diff rispetto all'anagrafica attuale
    const prevMap = Object.fromEntries(products.map((p:any)=>[p.id,p]));
    const diffs:any[] = [];
    newProds.forEach((p:any)=>{
      const old = prevMap[p.id];
      if(!old){ diffs.push({id:p.id,isNew:true,description:p.description,fields:[]}); return; }
      const fields:any[]=[];
      for(const k of["description","category","uom","qtyPerBox","boxPerPallet","kgPerBox","temperature","aiem","active","vendorName"]){
        const ov=String(old[k]??""), nv=String(p[k]??"");
        if(ov!==nv) fields.push({field:k,old:ov,new:nv});
      }
      if(fields.length) diffs.push({id:p.id,isNew:false,description:p.description,fields});
    });

    setProducts(newProds);
    CLOUD.set(`ifb_products_${branch}`, newProds);
    IDB.set(`ifb_anag_data_${now}`, newProds);
    setDataSource(`anagrafica_${branch}`,"manual");
    const log = { id:now, type:"anagrafica", date:new Date(now).toISOString(), count:newProds.length, branch, diffs };
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

  const q = search.trim().toLowerCase();
  const baseList = onlyIFB ? products.filter((p: any) => isIFBVendor(p.vendorName)) : products;
  const xrefIfbByNHK: Record<string,string> = {};
  (xrefs||[]).forEach((x:any)=>{ if(x.nHK&&x.ifbNo) xrefIfbByNHK[String(x.nHK)]=String(x.ifbNo); });
  let filtered = q
    ? products.filter((p: any) => {
        const ifbNo = p.code || xrefIfbByNHK[p.nHK] || "";
        return String(p.description||"").toLowerCase().includes(q) ||
          String(ifbNo).toLowerCase().includes(q) ||
          String(p.nHK||"").toLowerCase().includes(q) ||
          String(p.vendorName||"").toLowerCase().includes(q);
      })
    : baseList;
  if(sortAna==="az") filtered=[...filtered].sort((a:any,b:any)=>String(a.code||"").localeCompare(String(b.code||"")));
  else if(sortAna==="za") filtered=[...filtered].sort((a:any,b:any)=>String(b.code||"").localeCompare(String(a.code||"")));

  return (
    <div>
      <PageHeader title="Anagrafica Articoli" sub={`${products.length} articoli · ${products.filter((p: any) => isIFBVendor(p.vendorName)).length} INALCA F&B`} srcKey={`anagrafica_${branch}`}/>
      <BcBanner title={branch==="CAN" ? "Anagrafica manuale — sistema gestionale NAV" : "Dati aggiornati automaticamente da BC Brightview (HK)"}>
        {branch==="CAN"
          ? <>Le Canarie sono gestite su <b style={{color:T.text}}>NAV</b> (non su BC), senza accesso API diretto. L'anagrafica va caricata <b style={{color:T.orange}}>manualmente da file export NAV</b>: codice articolo, descrizione, UoM, kg/box, pz/box, temperatura. I <b style={{color:T.text}}>prezzi e listini</b> vengono dalla pagina Listini (import separato da BC IFB Italia).</>
          : <>Anagrafica articoli caricata ogni giorno alle 07:00 dall'item card di <b style={{color:T.text}}>Business Central Brightview</b>: descrizione, categoria, UoM, kg/box, pz/box, box/pallet, temperatura, fornitore, <b style={{color:T.text}}>Transportation</b> (AIR/SEA) e <b style={{color:T.text}}>Standard Cost</b> a sistema. È possibile importare manualmente da file per sovrascrivere.</>
        }
      </BcBanner>

      {/* Toolbar import */}
      <div style={{ display: "flex", gap: "10px", marginBottom: "16px", alignItems: "center", flexWrap: "wrap" }}>
        <label style={{ display: "inline-block", padding: "8px 16px", background: T.gold, color: "#000", borderRadius: "6px", cursor: "pointer", fontWeight: "bold", fontSize: "12px" }}>
          📂 Carica anagrafica ({branch==="CAN"?"NAV export":"BC export"})
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
          {onlyIFB ? `✓ Solo IF&B (${baseList.length})` : `Mostra tutti (${products.length})`}
        </button>

        <button
          onClick={()=>setSortAna((s:any)=>s==="az"?"za":s==="za"?"default":"az")}
          style={{padding:"5px 12px",background:sortAna!=="default"?`${T.blue}20`:T.surface,color:sortAna!=="default"?T.blue:T.muted,border:`1px solid ${sortAna!=="default"?T.blue:T.border}`,borderRadius:"6px",cursor:"pointer",fontSize:"11px"}}
        >
          {sortAna==="az"?"A→Z cod.":sortAna==="za"?"Z→A cod.":"Ordine cod. ▾"}
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
          <div style={{ overflowX: "auto", marginBottom: "12px", fontSize: "11px" }}>
            <table style={{ borderCollapse: "collapse", whiteSpace: "nowrap" }}>
              <thead>
                <tr style={{ background: T.card }}>
                  {FIELDS.map(f => (
                    <th key={f} style={{ textAlign: "left", padding: "3px 8px", borderBottom: `1px solid ${T.border}`, color: map[f] ? T.gold : T.dim, fontSize: "10px", fontWeight: "normal" }}>
                      {FLABELS[f] || f}
                      {map[f] && <div style={{ color: T.muted, fontSize: "9px", fontWeight: "normal" }}>← {map[f]}</div>}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {preview.slice(0, 30).map((r: any, i: number) => (
                  <tr key={i} style={{ borderBottom: `1px solid ${T.border}`, background: i % 2 === 0 ? T.bg : T.surface }}>
                    {FIELDS.map(f => {
                      const v = r[f];
                      const isEmpty = v === "" || v == null;
                      const isKey = f === "code" || f === "nHK";
                      const isNum = ["qtyPerBox","boxPerPallet","netWeightPcs","kgPerBox","kgxplt","aiem"].includes(f);
                      return (
                        <td key={f} style={{ padding: "2px 8px", color: isEmpty ? T.dim : isKey ? T.gold : isNum && Number(v) > 0 ? T.text : T.muted, fontFamily: isNum || isKey ? "monospace" : "inherit" }}>
                          {isEmpty ? "—" : String(v)}
                        </td>
                      );
                    })}
                  </tr>
                ))}
              </tbody>
            </table>
            {preview.length > 30 && <div style={{ padding: "6px", color: T.dim, fontSize: "10px" }}>…e altri {preview.length - 30} articoli</div>}
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
                {([
                  [branchN(branch),"left","muted"],["IFB No","left","muted"],["Descrizione","left","muted"],
                  ["Vendor","left","muted"],["Cat.","left","muted"],["UOM","left","muted"],
                  ["Qty/Box","right","muted"],["Box/Plt","right","muted"],
                  ...(branch==="CAN"?[["Peso Netto PCS","right","blue"]]:[] as any),
                  ["Kg/Box","right","muted"],["Kg/Plt","right","muted"],["Temp","left","muted"],
                  ...(branch==="CAN"?[["AIEM%","right","orange"]]:[]),
                  ["Attivo","left","muted"],
                ] as [string,string,string][]).map(([label,align,col])=>(
                  <th key={label} style={{padding:"3px 6px",background:T.card,color:(T as any)[col]||T.muted,textAlign:align as any,borderBottom:`1px solid ${T.border}`,fontSize:"10px",fontWeight:"normal",whiteSpace:"nowrap"}}>{label}</th>
                ))}
              </tr>
            </thead>
            <tbody key={q}>
              {filtered.map((p: any, i: number) => {
                const kgxplt = p.kgxplt || roundN((parseFloat(p.kgPerBox) || 0) * (parseFloat(p.boxPerPallet) || 0));
                const tdS: React.CSSProperties = {padding:"3px 6px",fontSize:"10px",whiteSpace:"nowrap"};
                const tdM: React.CSSProperties = {...tdS,fontFamily:"monospace"};
                return (
                  <tr key={p.id} style={{ borderBottom: `1px solid ${T.border}`, background: i % 2 === 0 ? T.bg : T.surface }}>
                    <td style={tdM}><span style={{ color: T.muted }}>{p.nHK || "—"}</span></td>
                    <td style={tdM}><span style={{ color: T.gold }}>{p.code || xrefIfbByNHK[p.nHK] || "—"}</span></td>
                    <td style={{...tdS,maxWidth:"180px",overflow:"hidden",textOverflow:"ellipsis"}} title={p.description}>{p.description}</td>
                    <td style={{...tdS,maxWidth:"110px",overflow:"hidden",textOverflow:"ellipsis"}}>
                      <span style={{ color: isIFBVendor(p.vendorName) ? T.gold : T.muted }}>{p.vendorName || "—"}</span>
                    </td>
                    <td style={tdS}><Chip label={p.category || "—"} color={p.category === "WINE" ? T.purple : p.category === "MEAT" ? T.red : T.blue} /></td>
                    <td style={tdS}><Chip label={p.uom || "—"} color={T.muted} /></td>
                    <td style={{...tdM,textAlign:"right"}}>{p.qtyPerBox || "—"}</td>
                    <td style={{...tdM,textAlign:"right"}}>{p.boxPerPallet || "—"}</td>
                    {branch==="CAN" && <td style={{...tdM,textAlign:"right",color:p.netWeightPcs>0?T.blue:T.red}}>{p.netWeightPcs>0?p.netWeightPcs:"⚠ mancante"}</td>}
                    <td style={{...tdM,textAlign:"right"}}>{p.kgPerBox || "—"}</td>
                    <td style={{...tdM,textAlign:"right"}}><span style={{ color: kgxplt > 0 ? T.text : T.dim }}>{kgxplt > 0 ? kgxplt : "—"}</span></td>
                    <td style={tdS}><Chip label={p.temperature || "—"} color={p.temperature === "FROZEN" ? T.blue : p.temperature === "FRESH" ? T.green : T.muted} /></td>
                    {branch==="CAN" && <td style={{...tdM,textAlign:"right",color:p.aiem>0?T.orange:T.dim}}>{p.aiem>0?`${p.aiem}%`:"—"}</td>}
                    <td style={tdS}><Chip label={p.active ? "Sì" : "No"} color={p.active ? T.green : T.red} /></td>
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

// ─── BEVERAGE INFO (per branch: HK = spirits >30°; CAN = tassa alcolica propria) ──
function BeverageInfoPage({bevInfo, setBevInfo, products, xrefs=[], showToast, branch}: any) {
  const isHK = branch === "HK";
  const [step, setStep] = useState<"main"|"map"|"preview">("main");
  const [headers, setHeaders] = useState<string[]>([]);
  const [rawRows, setRawRows] = useState<any[]>([]);
  const [map, setMap] = useState<any>({});
  const [preview, setPreview] = useState<any[]>([]);
  const [fileName, setFileName] = useState("");
  const [search, setSearchRaw] = useState(()=>psGet(`pg_${branch}_bevinfo_search`,""));
  const setSearch=(v:string)=>{setSearchRaw(v);psSet(`pg_${branch}_bevinfo_search`,v);};
  const [editingIdx, setEditingIdx] = useState<number|null>(null);
  const [editRow, setEditRow] = useState<any>({});
  const isHKRef = useRef(isHK);
  isHKRef.current = isHK;

  // Tassa alcolica CAN: 750,36 €/HL di alcol puro → 7,5036 €/L
  const CAN_ALC_EUR_PER_LT = 7.5036;

  function saveEdit(idx: number) {
    let updated: any;
    if(isHK) {
      const raw = String(editRow.codeInput||editRow.nHK||editRow.ifbNo||"").trim();
      if(!raw) { showToast("Inserisci N HK o IFB No", T.red); return; }
      const prod = products.find((p:any) => p.nHK===raw || p.code===raw)
                || products.find((p:any) => xrefs.some((x:any)=>x.nHK===raw && x.ifbNo===p.code));
      const resolvedNHK = prod?.nHK || (products.find((p:any)=>p.code===raw)?.nHK) || (xrefs.find((x:any)=>x.ifbNo===raw)?.nHK) || "";
      const resolvedIfb = prod?.code || (xrefs.find((x:any)=>x.nHK===raw)?.ifbNo) || raw;
      updated = { nHK: resolvedNHK||"", ifbNo: resolvedIfb, hasAlcTax: !!editRow.hasAlcTax, ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0 };
    } else {
      const rawCode = String(editRow.codeInput || editRow.ifbNo || "").trim();
      if(!rawCode) { showToast("Inserisci codice articolo", T.red); return; }
      const prod2 = findProduct(rawCode, products, xrefs) || products.find((p:any)=>p.code===rawCode);
      const resolvedIfb = prod2?.code || rawCode;
      const isSpirits2 = (prod2?.category||"").toUpperCase() === "SPIRITS";
      const lt = parseFloat(String(editRow.ltPerUnit||"").replace(",",".")) || 0;
      const gradoRaw2 = parseFloat(String(editRow.gradoAlcolico||"0").replace(",",".")) || 0;
      const grado = gradoRaw2 > 0 && gradoRaw2 < 1 ? gradoRaw2 * 100 : gradoRaw2;
      // Tassa alcolica si applica SOLO a SPIRITS — per vini/altri: totaleBottiglia=0
      const totaleBottiglia = isSpirits2 && lt > 0 && grado > 0 ? roundN((grado / 100) * CAN_ALC_EUR_PER_LT * lt, 2) : 0;
      updated = { ifbNo: resolvedIfb, ltPerUnit: lt, gradoAlcolico: grado, eurPerLt: isSpirits2 ? roundN((grado / 100) * CAN_ALC_EUR_PER_LT, 4) : 0, totaleBottiglia };
    }
    const next = idx === -1
      ? [...bevInfo, updated]
      : bevInfo.map((b:any, i:number) => i===idx ? updated : b);
    setBevInfo(next); IDB.set(`ifb_bevinfo_${branch}`, next);
    setEditingIdx(null); setEditRow({});
    showToast("Salvato ✓", T.gold);
  }

  function startAdd() {
    setEditingIdx(-1);
    setEditRow(isHK ? { codeInput:"", hasAlcTax:true } : { codeInput:"", ltPerUnit:"", gradoAlcolico:"" });
  }

  // Campi diversi per HK (flag >30°) vs CAN (tassa alcolica: litri/grado → calcolato)
  const FIELDS = isHK
    ? ["nHK","ifbNo","hasAlcTax"]
    : ["ifbNo","ltPerUnit","gradoAlcolico","totaleBottiglia"];
  const FLABELS: any = isHK ? {
    nHK:       "N HK (codice filiale)",
    ifbNo:     "IFB No (codice articolo)",
    hasAlcTax: "Tassato (>30° / True / Sì)",
  } : {
    ifbNo:          "IFB No * (codice articolo)",
    ltPerUnit:      "LT (litri per unità/bottiglia)",
    gradoAlcolico:  "Grado Alcolico (°)",
    totaleBottiglia:"Totale Bottiglia € (se già calcolato — altrimenti calcolato da LT × Grado%  × 7,5036)",
  };
  const ALIASES: any = isHK ? {
    nHK:       ["n hk","nhk","hk code","hk no","gc code","no hk","codice hong kong","hong kong","hk"],
    ifbNo:     ["ifb n","ifb no","ifbno","codice","code","item no","ifb item"],
    hasAlcTax: [">30","gradi>30","alcolico","tassato","tax","spirits","alc tax","has alc"],
  } : {
    ifbNo:          ["ifb n","ifb no","ifbno","bv no","codice","code","item no","ifb item","ifbitem"],
    ltPerUnit:      ["quantità x plt","quantity x plt","lt","litri","liters","volume","lt per unit","lt/unit","litri per unità","litriper"],
    gradoAlcolico:  ["grado alcolico","gradoalcolico","grado","gradi","abv","alcohol degree","degree","alc degree"],
    eurPerLt:       ["eur/lt","€/lt","euro/lt","eur/l","eur lt","eurlt","tariffa","rate","price per lt"],
    totaleBottiglia:["totale bottiglia","totale bottiglia eur","tot bott","totbott","totalebottiglia","totale","total","total bottle"],
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
        // Cerca la riga header: deve contenere "IFB" (codice articolo) oppure
        // avere almeno 3 celle valorizzate tra cui "LT", "GRADO", "EUR" (evita righe calc singole)
        const norm = (s:string) => String(s).trim().toUpperCase();
        const hdrIdx = data.findIndex(row => {
          const cells = row.map((c:any) => norm(c)).filter(Boolean);
          if(cells.length < 2) return false;
          const hasIFB = cells.some(c => c.includes("IFB") || c==="NO_" || c==="ITEM");
          const hasNHK = cells.some(c => c.includes("N HK") || c==="NHK" || c.includes("HK") || c.includes("HONG KONG"));
          const hasLT  = cells.some(c => c==="LT");
          const hasAlc = cells.some(c => c.includes("ALCOLICO") || c.includes("GRADO") || c.includes("EUR/LT") || c.includes("EUR/L") || c.includes(">30") || c.includes("SPIRITS") || c.includes("TAX"));
          return hasIFB || hasNHK || (hasLT && hasAlc);
        });
        if(hdrIdx < 0) { showToast("Intestazione non trovata — il file deve avere una colonna 'IFB Item' o 'IFB No'", T.red); return; }
        const hdrs = data[hdrIdx].map((h:any) => String(h).trim()).filter((h:string) => h);
        const rows = data.slice(hdrIdx+1).filter(r => r.some((c:any) => c !== ""));
        setHeaders(hdrs); setRawRows(rows);
        const autoMap: any = {};
        FIELDS.forEach(f => { autoMap[f] = fi(ALIASES[f], hdrs); });
        setMap(autoMap);
        // HK e CAN: salta mapping, vai diretto alla preview con auto-detect
        if(isHKRef.current) {
          const idx2: any = {};
          FIELDS.forEach(f => { idx2[f] = hdrs.indexOf(autoMap[f]); });
          const rows2 = rows.map((row:any) => {
            const nHK = idx2.nHK >= 0 ? String(row[idx2.nHK]||"").trim() : "";
            const ifbNoRaw = idx2.ifbNo >= 0 ? String(row[idx2.ifbNo]||"").trim() : "";
            const alcRaw = idx2.hasAlcTax >= 0 ? String(row[idx2.hasAlcTax]||"").trim().toLowerCase() : "";
            const hasAlcTax = ["true","sì","si","yes","1","x","vero"].includes(alcRaw);
            const key = nHK || ifbNoRaw;
            if(!key) return null;
            const prod = products.find((p:any) => p.nHK === nHK || p.code === ifbNoRaw || p.code === nHK);
            const ifbNo = ifbNoRaw || prod?.code || nHK;
            return { nHK, ifbNo, hasAlcTax, _found:!!prod, _desc:prod?.description||"—" };
          }).filter(Boolean);
          setPreview(rows2); setStep("preview");
          return;
        }
        // CAN: auto-detect e vai diretto alla preview senza step mapping
        {
          const idx2: any = {};
          FIELDS.forEach(f => { idx2[f] = hdrs.indexOf(autoMap[f]); });
          // Fallback: cerca "IFB Item" / "No_" / "LT" / "GRADO ALCOLICO" per nome diretto
          if(idx2.ifbNo < 0) idx2.ifbNo = hdrs.findIndex(h => /ifb/i.test(h) && /item|n[°o]?_?/i.test(h)) >= 0
            ? hdrs.findIndex(h => /ifb/i.test(h)) : hdrs.findIndex(h => /^ifb/i.test(h));
          if(idx2.ltPerUnit < 0) idx2.ltPerUnit = hdrs.findIndex(h => h.trim().toUpperCase() === "LT");
          if(idx2.gradoAlcolico < 0) idx2.gradoAlcolico = hdrs.findIndex(h => /grado/i.test(h) && /alcolico/i.test(h) && !/in\s*%/i.test(h));
          const rows2 = rows.map((row:any) => {
            const ifbNo = idx2.ifbNo >= 0 ? String(row[idx2.ifbNo]||"").trim() : "";
            if(!ifbNo) return null;
            const lt = parseFloat(String(row[idx2.ltPerUnit >= 0 ? idx2.ltPerUnit : -1]||"").replace(",",".")) || 0;
            const gradoRaw = parseFloat(String(row[idx2.gradoAlcolico >= 0 ? idx2.gradoAlcolico : -1]||"").replace(",",".")) || 0;
            const grado = gradoRaw > 0 && gradoRaw < 1 ? gradoRaw * 100 : gradoRaw;
            const totaleBottiglia = lt > 0 && grado > 0 ? roundN((grado / 100) * CAN_ALC_EUR_PER_LT * lt, 2) : 0;
            if(totaleBottiglia <= 0 && lt <= 0) return null;
            const prod = findProduct(ifbNo, products, xrefs);
            return { ifbNo, ltPerUnit:lt, gradoAlcolico:grado, eurPerLt: roundN((grado/100)*CAN_ALC_EUR_PER_LT,4), totaleBottiglia, _found:!!prod, _desc:prod?.description||"—" };
          }).filter(Boolean);
          setPreview(rows2); setStep("preview");
          return;
        }
      } catch(err:any) { showToast("Errore: "+err.message, T.red); }
    };
    reader.readAsBinaryString(file);
  }

  function buildPreview() {
    const idx: any = {};
    FIELDS.forEach(f => { idx[f] = headers.indexOf(map[f]); });

    if(isHK) {
      const rows = rawRows.map(row => {
        const nHK = idx.nHK >= 0 ? String(row[idx.nHK]||"").trim() : "";
        const ifbNoRaw = idx.ifbNo >= 0 ? String(row[idx.ifbNo]||"").trim() : "";
        const alcRaw = idx.hasAlcTax >= 0 ? String(row[idx.hasAlcTax]||"").trim().toLowerCase() : "";
        const hasAlcTax = alcRaw === "true" || alcRaw === "sì" || alcRaw === "si" || alcRaw === "yes" || alcRaw === "1" || alcRaw === "x" || alcRaw === "vero";
        const key = nHK || ifbNoRaw;
        if(!key) return null;
        const prod = products.find((p:any) => p.nHK === nHK || p.code === ifbNoRaw || p.code === nHK);
        const ifbNo = ifbNoRaw || prod?.code || nHK;
        return { nHK, ifbNo, hasAlcTax, _found:!!prod, _desc:prod?.description||"—" };
      }).filter(Boolean);
      setPreview(rows); setStep("preview");
      return;
    }

    const rows = rawRows.map(row => {
      const ifbNo = String(row[idx.ifbNo]||"").trim();
      if(!ifbNo) return null;
      const lt = parseFloat(String(row[idx.ltPerUnit]||"").replace(",",".")) || 0;
      const gradoRaw = parseFloat(String(row[idx.gradoAlcolico]||"").replace(",",".")) || 0;
      // Normalize: Excel % cells store as decimal (0.20 = 20°), convert to degrees
      const grado = gradoRaw > 0 && gradoRaw < 1 ? gradoRaw * 100 : gradoRaw;
      // Always calculate from formula: LT × (Grado/100) × 7.5036 — do not trust file's TOTALE column
      const totaleBottiglia = lt > 0 && grado > 0 ? roundN((grado / 100) * CAN_ALC_EUR_PER_LT * lt, 2) : 0;
      if(totaleBottiglia <= 0 && lt <= 0) return null;
      const prod = findProduct(ifbNo, products, xrefs);
      return { ifbNo, ltPerUnit:lt, gradoAlcolico:grado, eurPerLt: roundN((grado/100)*CAN_ALC_EUR_PER_LT,4), totaleBottiglia, _found:!!prod, _desc:prod?.description||"—" };
    }).filter(Boolean);
    setPreview(rows); setStep("preview");
  }

  function executeImport() {
    if(isHK) {
      const kept = bevInfo.filter((b:any) => !preview.find((p:any) => (p.nHK && b.nHK===p.nHK) || b.ifbNo===p.ifbNo));
      const next = [...preview.map((r:any) => ({nHK:r.nHK, ifbNo:r.ifbNo, hasAlcTax:r.hasAlcTax, ltPerUnit:0, gradoAlcolico:0, eurPerLt:0, totaleBottiglia:0})), ...kept];
      setBevInfo(next); IDB.set(`ifb_bevinfo_${branch}`, next);
      showToast(`Beverage Info HK: ${preview.length} articoli importati ✓`, T.gold);
      setStep("main"); setPreview([]); setRawRows([]); setHeaders([]);
      return;
    }
    const kept = bevInfo.filter((b:any) => !preview.find((p:any) => p.ifbNo===b.ifbNo));
    const next = [...preview.map((r:any) => ({ifbNo:r.ifbNo,ltPerUnit:r.ltPerUnit,gradoAlcolico:r.gradoAlcolico,eurPerLt:r.eurPerLt,totaleBottiglia:r.totaleBottiglia})), ...kept];
    setBevInfo(next); IDB.set(`ifb_bevinfo_${branch}`, next);
    showToast(`Beverage Info: ${preview.length} articoli importati ✓`, T.gold);
    setStep("main"); setPreview([]); setRawRows([]); setHeaders([]);
  }

  const getNComit = (ifbNo: string) => xrefs.find((x:any)=>x.ifbNo===ifbNo)?.nHK || "";
  const q = search.trim().toLowerCase();
  const displayed = q
    ? bevInfo.filter((b:any) => b.ifbNo?.toLowerCase().includes(q)
        || products.find((p:any)=>p.code===b.ifbNo)?.description?.toLowerCase().includes(q)
        || getNComit(b.ifbNo).toLowerCase().includes(q))
    : bevInfo;

  return (
    <div>
      <PageHeader title={isHK ? "🍷 Beverage Info · Tassa Alcolica (HK)" : "🍷 Beverage Info · Tassa Alcolica (CAN)"} sub={isHK ? "Importa lista articoli con Spirits >30° — tassa = 100% del prezzo acquisto" : "Importa dati alcolici: LT, Grado Alcolico, €/LT → Totale tassa per unità"}/>

      <div style={{display:"flex",gap:"10px",marginBottom:"16px",alignItems:"center",flexWrap:"wrap"}}>
        <label style={{display:"inline-block",padding:"8px 16px",background:T.gold,color:"#000",borderRadius:"6px",cursor:"pointer",fontWeight:"bold",fontSize:"12px"}}>
          📂 Carica file Beverage Info
          <input type="file" accept=".xlsx,.xls,.csv" onChange={e=>{const f=e.target.files?.[0];if(f)parseFile(f);e.target.value="";}} style={{display:"none"}}/>
        </label>
        {isHK&&<button onClick={()=>{setBevInfo(HK_ALC_TAX_DEFAULTS);IDB.set(`ifb_bevinfo_${branch}`,HK_ALC_TAX_DEFAULTS);showToast("Default HK ripristinati ✓",T.gold);}}
          style={{padding:"5px 12px",background:"none",border:`1px solid ${T.gold}44`,borderRadius:"6px",color:T.gold,cursor:"pointer",fontSize:"11px"}}>
          ↺ Ripristina Default
        </button>}
        {bevInfo.length>0&&<button onClick={()=>{if(window.confirm(`Eliminare tutti i ${bevInfo.length} dati beverage?`)){setBevInfo([]);IDB.set(`ifb_bevinfo_${branch}`,[]);}}}
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
                {(isHK
                  ? ["N HK","IFB No","Descrizione","Tassa Alcol (>30°)"]
                  : ["IFB No","Descrizione","LT","Grado","€/LT","Totale €/unit"]
                ).map(h=><th key={h} style={{padding:"4px 8px",textAlign:"left",color:T.muted,borderBottom:`1px solid ${T.border}`}}>{h}</th>)}
              </tr></thead>
              <tbody>
                {preview.map((r:any,i:number)=>(
                  <tr key={i} style={{borderBottom:`1px solid ${T.border}`,background:r._found?undefined:`${T.orange}10`}}>
                    {isHK ? <>
                      <td style={{padding:"3px 8px",color:T.gold,fontFamily:"monospace"}}>{r.nHK||"—"}</td>
                      <td style={{padding:"3px 8px",color:T.gold,fontFamily:"monospace"}}>{r.ifbNo}</td>
                      <td style={{padding:"3px 8px",color:r._found?T.text:T.orange,fontSize:"11px"}}>{r._desc}</td>
                      <td style={{padding:"3px 8px",textAlign:"center",fontWeight:"bold",color:r.hasAlcTax?T.orange:T.muted}}>{r.hasAlcTax?"✓ SÌ":"—"}</td>
                    </> : <>
                      <td style={{padding:"3px 8px",color:T.gold,fontFamily:"monospace"}}>{r.ifbNo}</td>
                      <td style={{padding:"3px 8px",color:r._found?T.text:T.orange,fontSize:"11px"}}>{r._desc}</td>
                      <td style={{padding:"3px 8px",fontFamily:"monospace",textAlign:"right"}}>{r.ltPerUnit||"—"}</td>
                      <td style={{padding:"3px 8px",fontFamily:"monospace",textAlign:"right"}}>{r.gradoAlcolico||"—"}°</td>
                      <td style={{padding:"3px 8px",fontFamily:"monospace",textAlign:"right"}}>{r.eurPerLt>0?r.eurPerLt.toFixed(2):"—"}</td>
                      <td style={{padding:"3px 8px",fontFamily:"monospace",textAlign:"right",color:T.orange,fontWeight:"bold"}}>{r.totaleBottiglia>0?r.totaleBottiglia.toFixed(4):"—"}</td>
                    </>}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div style={{display:"flex",gap:"10px"}}>
            {!isHK && <ActionBtn label="← Indietro" onClick={()=>setStep("map")}/>}
            <ActionBtn label={`✓ Importa ${preview.length} articoli`} onClick={executeImport} primary/>
          </div>
        </Section>
      )}

      <div style={{display:"flex",gap:"8px",alignItems:"center",marginBottom:"10px"}}>
        <SearchBar value={search} onChange={setSearch} placeholder="🔍 Cerca IFB No o descrizione…"/>
        <button onClick={startAdd} style={{padding:"6px 14px",background:`${T.gold}20`,border:`1px solid ${T.gold}44`,borderRadius:"6px",color:T.gold,cursor:"pointer",fontSize:"12px",whiteSpace:"nowrap"}}>+ Aggiungi</button>
      </div>

      {(displayed.length>0 || editingIdx===-1) ? (
        <Section title={isHK ? `${displayed.length} articoli con tassa alcolica (>30°) — HK` : `${displayed.length} articoli con tassa alcolica — CAN`}>
          <div style={{overflowX:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse"}}>
              <thead><tr>
                {(isHK
                  ? ["N HK","IFB No","Descrizione","Tassa Alcol",""]
                  : ["N COMIT","IFB No","Descrizione","LT/unit","Grado","€/LT","Tassa Alcolica €/unit",""]
                ).map(h=><th key={h} style={{padding:"3px 6px",background:T.card,color:T.muted,textAlign:"left",borderBottom:`1px solid ${T.border}`,fontSize:"10px"}}>{h}</th>)}
              </tr></thead>
              <tbody>
                {/* Nuova riga in cima se si sta aggiungendo */}
                {editingIdx===-1&&(
                  <tr style={{background:`${T.gold}08`,borderBottom:`1px solid ${T.border}`}}>
                    {isHK ? <>
                      <td colSpan={2} style={{padding:"4px 6px"}}>
                        <input value={editRow.codeInput||""} onChange={e=>setEditRow((r:any)=>({...r,codeInput:e.target.value}))}
                          placeholder="N HK o IFB No" autoFocus
                          style={{...inputStyle(),fontSize:"10px",padding:"3px 6px",width:"130px",fontFamily:"monospace"}}/>
                      </td>
                      <td style={{padding:"4px 6px",fontSize:"10px",color:T.muted}}>
                        {(()=>{const raw=String(editRow.codeInput||"").trim();const p=products.find((p:any)=>p.nHK===raw||p.code===raw);return p?.description||"—";})()}
                      </td>
                      <td style={{padding:"4px 6px",textAlign:"center"}}>
                        <label style={{cursor:"pointer",display:"flex",alignItems:"center",gap:"6px",justifyContent:"center"}}>
                          <input type="checkbox" checked={!!editRow.hasAlcTax} onChange={e=>setEditRow((r:any)=>({...r,hasAlcTax:e.target.checked}))}/>
                          <span style={{fontSize:"10px",color:editRow.hasAlcTax?T.orange:T.muted}}>{editRow.hasAlcTax?"SÌ":"NO"}</span>
                        </label>
                      </td>
                    </> : <>
                      {/* CAN add: code → anagrafica; LT + Grado; €/LT e Totale auto */}
                      <td style={{padding:"2px 4px",fontSize:"10px",color:T.muted,fontFamily:"monospace"}}>
                        {(()=>{const c=String(editRow.codeInput||"").trim();return getNComit(c)||"—";})()}
                      </td>
                      <td style={{padding:"2px 4px"}}>
                        <input value={editRow.codeInput||""} onChange={e=>setEditRow((r:any)=>({...r,codeInput:e.target.value}))}
                          placeholder="IFB No o N COMIT" autoFocus
                          style={{...inputStyle(),fontSize:"10px",padding:"2px 4px",width:"110px",fontFamily:"monospace"}}/>
                      </td>
                      <td style={{padding:"2px 4px",fontSize:"10px",color:T.muted,maxWidth:"150px",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                        {(()=>{const c=String(editRow.codeInput||"").trim();const p=findProduct(c,products,xrefs);return p?.description||"—";})()}
                      </td>
                      <td style={{padding:"2px 4px"}}>
                        <input value={editRow.ltPerUnit||""} onChange={e=>setEditRow((r:any)=>({...r,ltPerUnit:e.target.value}))}
                          placeholder="LT" style={{...inputStyle(),fontSize:"10px",padding:"2px 4px",width:"60px",fontFamily:"monospace"}}/>
                      </td>
                      <td style={{padding:"2px 4px"}}>
                        <input value={editRow.gradoAlcolico||""} onChange={e=>setEditRow((r:any)=>({...r,gradoAlcolico:e.target.value}))}
                          placeholder="°" style={{...inputStyle(),fontSize:"10px",padding:"2px 4px",width:"60px",fontFamily:"monospace"}}/>
                      </td>
                      <td style={{padding:"2px 4px",textAlign:"right",fontSize:"10px",color:T.dim}}>
                        {(()=>{const g=parseFloat(String(editRow.gradoAlcolico||"0").replace(",","."));const gp=g>0&&g<1?g*100:g;return gp>0?roundN((gp/100)*CAN_ALC_EUR_PER_LT,4).toFixed(4):"—";})()}
                      </td>
                      <td style={{padding:"2px 4px",textAlign:"right",fontSize:"10px",color:T.orange,fontWeight:"bold"}}>
                        {(()=>{const lt=parseFloat(String(editRow.ltPerUnit||"0").replace(",","."));const g=parseFloat(String(editRow.gradoAlcolico||"0").replace(",","."));const gp=g>0&&g<1?g*100:g;const t=lt>0&&gp>0?roundN((gp/100)*CAN_ALC_EUR_PER_LT*lt,2):0;return t>0?t.toFixed(4):"—";})()}
                      </td>
                    </>}
                    <td style={{padding:"2px 4px",display:"flex",gap:"4px"}}>
                      <MiniBtn label="✓" onClick={()=>saveEdit(-1)} color={T.green}/>
                      <MiniBtn label="✕" onClick={()=>{setEditingIdx(null);setEditRow({});}} color={T.red}/>
                    </td>
                  </tr>
                )}
                {displayed.map((b:any,i:number)=>{
                  const realIdx = bevInfo.indexOf(b);
                  const prod = products.find((p:any)=>p.code===b.ifbNo);
                  const nComit = getNComit(b.ifbNo);
                  const isEditing = editingIdx===realIdx;
                  return(
                    <tr key={b.ifbNo} style={{borderBottom:`1px solid ${T.border}`,background:isEditing?`${T.gold}08`:i%2===0?T.bg:T.surface}}>
                      {isEditing ? (
                        <>
                          {isHK ? <>
                            <td style={{padding:"2px 6px",fontFamily:"monospace",fontSize:"10px",color:T.gold}}>{b.nHK||"—"}</td>
                            <td style={{padding:"2px 6px",fontFamily:"monospace",fontSize:"10px",color:T.gold}}>{b.ifbNo}</td>
                            <td style={{padding:"2px 6px",fontSize:"10px",color:T.muted}}>{prod?.description||"—"}</td>
                            <td style={{padding:"2px 6px",textAlign:"center"}}>
                              <label style={{cursor:"pointer",display:"flex",alignItems:"center",gap:"6px",justifyContent:"center"}}>
                                <input type="checkbox" checked={!!editRow.hasAlcTax} onChange={e=>setEditRow((r:any)=>({...r,hasAlcTax:e.target.checked}))}/>
                                <span style={{fontSize:"10px",color:editRow.hasAlcTax?T.orange:T.muted}}>{editRow.hasAlcTax?"SÌ":"NO"}</span>
                              </label>
                            </td>
                          </> : <>
                            {/* CAN edit: N COMIT auto; codice editabile; descrizione auto; LT+Grado input; €/LT+Totale auto */}
                            <td style={{padding:"2px 4px",fontSize:"10px",color:T.muted,fontFamily:"monospace"}}>{nComit||"—"}</td>
                            <td style={{padding:"2px 4px"}}>
                              <input value={editRow.codeInput||editRow.ifbNo||""} onChange={e=>setEditRow((r:any)=>({...r,codeInput:e.target.value}))}
                                style={{...inputStyle(),fontSize:"10px",padding:"2px 4px",width:"100px",fontFamily:"monospace"}}/>
                            </td>
                            <td style={{padding:"2px 4px",fontSize:"10px",color:T.muted,maxWidth:"150px",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                              {(()=>{const c=String(editRow.codeInput||editRow.ifbNo||"").trim();const p=findProduct(c,products,xrefs)||products.find((pp:any)=>pp.code===c);return p?.description||prod?.description||"—";})()}
                            </td>
                            <td style={{padding:"2px 4px"}}>
                              <input value={editRow.ltPerUnit||""} onChange={e=>setEditRow((r:any)=>({...r,ltPerUnit:e.target.value}))}
                                style={{...inputStyle(),fontSize:"10px",padding:"2px 4px",width:"60px",fontFamily:"monospace"}}/>
                            </td>
                            <td style={{padding:"2px 4px"}}>
                              <input value={editRow.gradoAlcolico||""} onChange={e=>setEditRow((r:any)=>({...r,gradoAlcolico:e.target.value}))}
                                style={{...inputStyle(),fontSize:"10px",padding:"2px 4px",width:"60px",fontFamily:"monospace"}}/>
                            </td>
                            <td style={{padding:"2px 4px",textAlign:"right",fontSize:"10px",color:T.dim}}>{CAN_ALC_EUR_PER_LT}</td>
                            <td style={{padding:"2px 4px",textAlign:"right",fontSize:"10px",color:T.orange,fontWeight:"bold"}}>
                              {(()=>{const lt=parseFloat(String(editRow.ltPerUnit||"0").replace(",","."));const g=parseFloat(String(editRow.gradoAlcolico||"0").replace(",","."));const gp=g>0&&g<1?g*100:g;const t=lt>0&&gp>0?roundN((gp/100)*CAN_ALC_EUR_PER_LT*lt,2):0;return t>0?t.toFixed(4):"—";})()}
                            </td>
                          </>}
                          <td style={{padding:"2px 4px",display:"flex",gap:"4px"}}>
                            <MiniBtn label="✓" onClick={()=>saveEdit(realIdx)} color={T.green}/>
                            <MiniBtn label="✕" onClick={()=>{setEditingIdx(null);setEditRow({});}} color={T.muted}/>
                          </td>
                        </>
                      ) : (
                        <>
                          {isHK ? <>
                            <td style={{padding:"3px 6px",fontFamily:"monospace",fontSize:"10px",color:T.gold}}>{b.nHK||"—"}</td>
                            <td style={{padding:"3px 6px",fontFamily:"monospace",fontSize:"10px",color:T.gold}}>{b.ifbNo}</td>
                            <td style={{padding:"3px 6px",fontSize:"10px",maxWidth:"200px",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{prod?.description||<span style={{color:T.orange}}>⚠ non in anagrafica</span>}</td>
                            <td style={{padding:"3px 6px",fontSize:"10px",textAlign:"center",fontWeight:"bold",color:b.hasAlcTax?T.orange:T.muted}}>{b.hasAlcTax?"✓ SÌ":"—"}</td>
                          </> : <>
                            <td style={{padding:"3px 6px",fontFamily:"monospace",fontSize:"10px",color:T.muted}}>{nComit||"—"}</td>
                            <td style={{padding:"3px 6px",fontFamily:"monospace",fontSize:"10px",color:T.gold}}>{b.ifbNo}</td>
                            <td style={{padding:"3px 6px",fontSize:"10px",maxWidth:"200px",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{prod?.description||<span style={{color:T.orange}}>⚠ non in anagrafica</span>}</td>
                            <td style={{padding:"3px 6px",fontFamily:"monospace",fontSize:"10px",textAlign:"right"}}>{b.ltPerUnit||"—"}</td>
                            <td style={{padding:"3px 6px",fontFamily:"monospace",fontSize:"10px",textAlign:"right"}}>{b.gradoAlcolico||"—"}°</td>
                            <td style={{padding:"3px 6px",fontFamily:"monospace",fontSize:"10px",textAlign:"right"}}>{b.eurPerLt>0?b.eurPerLt.toFixed(2):"—"}</td>
                            <td style={{padding:"3px 6px",fontFamily:"monospace",fontSize:"10px",textAlign:"right",color:(prod?.category||"").toUpperCase()==="SPIRITS"?T.orange:T.dim,fontWeight:"bold"}}>
                              {(prod?.category||"").toUpperCase()==="SPIRITS"
                                ? (b.totaleBottiglia>0?b.totaleBottiglia.toFixed(4):"—")
                                : <span style={{fontSize:"9px",fontWeight:"normal"}}>non spirits</span>}
                            </td>
                          </>}
                          <td style={{padding:"3px 6px",display:"flex",gap:"4px"}}>
                            <MiniBtn label="✎" onClick={()=>{setEditingIdx(realIdx);setEditRow({...b});}} color={T.blue}/>
                            <MiniBtn label="✕" onClick={()=>{const n=bevInfo.filter((_:any,j:number)=>j!==realIdx);setBevInfo(n);IDB.set(`ifb_bevinfo_${branch}`,n);}} color={T.red}/>
                          </td>
                        </>
                      )}
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </Section>
      ) : (
        <div style={{color:T.muted,textAlign:"center",padding:"40px",fontSize:"13px"}}>
          Nessun dato beverage. Carica il file o clicca "+ Aggiungi".
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
  const [search, setSearchRaw] = useState(()=>psGet("pg_meatprice_search",""));
  const setSearch=(v:string)=>{setSearchRaw(v);psSet("pg_meatprice_search",v);};

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
                          <td style={{padding:"3px 6px",fontSize:"10px",fontFamily:"monospace",whiteSpace:"nowrap"}}>
                            <span style={{color:T.gold}}>{m.code}</span>
                            {prod && <span style={{marginLeft:"6px",fontSize:"9px",color:T.green}}>✓ {prod.code}</span>}
                          </td>
                          <td style={{padding:"3px 6px",fontSize:"10px",maxWidth:"200px",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{m.description}</td>
                          <td style={{padding:"3px 6px",fontSize:"10px",fontFamily:"monospace",whiteSpace:"nowrap"}}>
                            <span style={{color:T.green,fontWeight:"bold"}}>€ {m.pricePerKg?.toFixed(2)||"—"}</span>
                          </td>
                          <td style={{padding:"3px 6px",fontSize:"10px"}}>
                            {m.foglio && <Chip label={m.foglio} color={T.blue}/>}
                          </td>
                          <td style={{padding:"3px 6px",fontSize:"10px",color:T.muted}}>{m.fonte||"—"}</td>
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

function BcBanner({icon="ℹ", title, children}:any){
  return(
    <div style={{background:`${T.blue}15`,border:`1px solid ${T.blue}44`,borderRadius:"8px",padding:"12px 16px",marginBottom:"16px",fontSize:"12px",color:T.text}}>
      <div style={{fontWeight:"bold",color:T.blue,marginBottom:"6px"}}>{icon} {title}</div>
      <div style={{color:T.muted,lineHeight:"1.6"}}>{children}</div>
    </div>
  );
}

function PageHeader({title,sub,srcKey=null}:any){
  return(
    <div style={{marginBottom:"20px"}}>
      <h2 style={{color:T.gold,margin:"0 0 4px",fontSize:"18px",display:"flex",alignItems:"center",flexWrap:"wrap",gap:"6px"}}>
        {title}
        {srcKey&&<SourceBadge dataKey={srcKey}/>}
      </h2>
      {sub&&<div style={{fontSize:"12px",color:T.muted}}>{sub}</div>}
    </div>
  );
}

// Helper: salva sorgente dati in localStorage
function setDataSource(key:string, src:"bc"|"manual") {
  LS.set(`ifb_dsrc_${key}`, JSON.stringify({src, ts: Date.now()}));
}
function getDataSource(key:string):{src:"bc"|"manual",ts:number}|null {
  try { return JSON.parse(LS.get(`ifb_dsrc_${key}`) || "null"); } catch { return null; }
}

function SourceBadge({dataKey}:{dataKey:string}) {
  const info = getDataSource(dataKey);
  if(!info) return null;
  const date = new Date(info.ts).toLocaleDateString("it-IT",{day:"2-digit",month:"2-digit",year:"2-digit"});
  const isBc = info.src === "bc";
  return(
    <span style={{
      display:"inline-flex",alignItems:"center",gap:"4px",
      fontSize:"10px",fontWeight:600,letterSpacing:"0.04em",
      padding:"2px 8px",borderRadius:"12px",marginLeft:"10px",
      background: isBc ? `${T.blue}22` : `${T.muted}22`,
      color: isBc ? T.blue : T.muted,
      border: `1px solid ${isBc ? T.blue+"55" : T.muted+"55"}`,
      verticalAlign:"middle",
    }}>
      {isBc ? "🔄 BC" : "📁 Manuale"} · {date}
    </span>
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
function SearchBar({value,onChange,placeholder,style:extraStyle}:any){
  return<input value={value}
    onChange={e=>onChange(e.target.value)}
    onInput={e=>onChange((e.target as HTMLInputElement).value)}
    placeholder={placeholder||"Cerca..."}
    autoComplete="off" autoCorrect="off" spellCheck={false}
    style={{...inputStyle(),maxWidth:"320px",marginBottom:"14px",...(extraStyle||{})}}/>;
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
            padding: "3px 6px",
            background: T.card,
            color: T.muted,
            textAlign: "left",
            borderBottom: `1px solid ${T.border}`,
            fontSize: "10px",
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
  return<td style={{padding:"3px 6px",borderBottom:`1px solid ${T.border}`,fontSize:"10px",fontFamily:mono?"monospace":"inherit",verticalAlign:"middle",whiteSpace:"nowrap"}}>{children}</td>;
}
function Chip({label,color}){
  return<span style={{padding:"2px 7px",background:`${color}22`,color,borderRadius:"4px",fontSize:"10px",fontWeight:"bold",letterSpacing:"0.04em"}}>{label}</span>;
}
function MiniBtn({label,onClick,color}){
  return<button onClick={onClick} style={{padding:"3px 8px",background:"none",border:`1px solid ${color||T.border}`,borderRadius:"4px",color:color||T.muted,cursor:"pointer",fontSize:"11px"}}>{label}</button>;
}