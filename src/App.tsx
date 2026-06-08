import { useState, useEffect, useMemo, useCallback, useRef } from 'react';
import * as XLSX from 'xlsx';

// ─────────────────────────────────────────────────────────────────────────────
// DESIGN TOKENS
// ─────────────────────────────────────────────────────────────────────────────
const T = {
  bg: '#0B0F14',
  surface: '#111720',
  card: '#161E28',
  border: 'rgba(255,255,255,0.07)',
  borderHi: 'rgba(255,255,255,0.14)',
  text: '#E2D9CC',
  muted: 'rgba(226,217,204,0.45)',
  dim: 'rgba(226,217,204,0.22)',
  gold: '#C9A84C',
  goldDim: 'rgba(201,168,76,0.18)',
  blue: '#4A8FB5',
  green: '#4BA87A',
  red: '#B5534A',
  orange: '#C47A3B',
  purple: '#7B5AC4',
};

const BRANCH_CFG = {
  HK: {
    label: 'Hong Kong',
    flag: '🇭🇰',
    color: T.gold,
    currency: 'HKD',
    defaultRate: 9.1437,
  },
  CAN: {
    label: 'Canarie',
    flag: '🇮🇨',
    color: T.blue,
    currency: 'EUR',
    defaultRate: 1,
  },
  AUS: {
    label: 'Australia',
    flag: '🇦🇺',
    color: T.orange,
    currency: 'AUD',
    defaultRate: 1.642,
  },
};

const NOW = () => new Date().toISOString().slice(0, 7);
const FMT = (n, d = 4) => (typeof n === 'number' ? n.toFixed(d) : '—');
const PCT = (a, b) => (b ? ((a - b) / b) * 100 : null);

// ─────────────────────────────────────────────────────────────────────────────
// SEED DATA
// ─────────────────────────────────────────────────────────────────────────────
const SEED_PRODUCTS = [
  {
    id: 'P001',
    code: 'BAL-001',
    description: 'Balsamic Vinegar 500ml CASA EMILIA',
    category: 'FOOD',
    uom: 'PCS',
    qtyPerBox: 12,
    boxPerPallet: 70,
    kgPerBox: null,
    temperature: 'DRY',
    active: true,
  },
  {
    id: 'P002',
    code: 'BAL-002',
    description: 'Balsamic Cream 250ml',
    category: 'FOOD',
    uom: 'PCS',
    qtyPerBox: 12,
    boxPerPallet: 130,
    kgPerBox: null,
    temperature: 'DRY',
    active: true,
  },
  {
    id: 'P003',
    code: 'CHE-001',
    description: 'Bavarian Emmental "Cepparo" 3kg',
    category: 'FOOD',
    uom: 'KG',
    qtyPerBox: 10,
    boxPerPallet: 60,
    kgPerBox: 10,
    temperature: 'FRESH',
    active: true,
  },
  {
    id: 'P004',
    code: 'CHE-002',
    description: 'Parmigiano Reggiano 24M 2kg',
    category: 'FOOD',
    uom: 'KG',
    qtyPerBox: 8,
    boxPerPallet: 99,
    kgPerBox: 8,
    temperature: 'FRESH',
    active: true,
  },
  {
    id: 'P005',
    code: 'PAS-001',
    description: 'Tortelli Ricotta & Spinach 3kg',
    category: 'FOOD',
    uom: 'BOX',
    qtyPerBox: 1,
    boxPerPallet: 160,
    kgPerBox: null,
    temperature: 'FROZEN',
    active: true,
  },
  {
    id: 'P006',
    code: 'WIN-001',
    description: "Lambrusco Dell'Emilia IGT 75cl",
    category: 'WINE',
    uom: 'PCS',
    qtyPerBox: 6,
    boxPerPallet: 80,
    kgPerBox: null,
    temperature: 'FRESH',
    active: true,
  },
  {
    id: 'P007',
    code: 'WIN-002',
    description: 'Prosecco DOC Extra Dry 75cl',
    category: 'WINE',
    uom: 'PCS',
    qtyPerBox: 6,
    boxPerPallet: 96,
    kgPerBox: null,
    temperature: 'FRESH',
    active: true,
  },
  {
    id: 'P008',
    code: 'MEA-001',
    description: 'Mortadella BLU 3kg Ø14cm',
    category: 'MEAT',
    uom: 'KG',
    qtyPerBox: 6,
    boxPerPallet: 72,
    kgPerBox: 6,
    temperature: 'FRESH',
    active: true,
  },
  {
    id: 'P009',
    code: 'MEA-002',
    description: 'Spianata Calabra 2kg',
    category: 'MEAT',
    uom: 'KG',
    qtyPerBox: 4,
    boxPerPallet: 150,
    kgPerBox: 4,
    temperature: 'FRESH',
    active: true,
  },
  {
    id: 'P010',
    code: 'FRZ-001',
    description: 'Arancini Mozzarella & Tomato 3kg',
    category: 'FOOD',
    uom: 'BOX',
    qtyPerBox: 1,
    boxPerPallet: 162,
    kgPerBox: null,
    temperature: 'FROZEN',
    active: true,
  },
];

const SEED_LOGISTIC = [
  {
    productId: 'P001',
    branch: 'HK',
    area: 'NORD',
    ubicazione: 'MTO',
    pltPerContainer: 25,
    hasCert: false,
    hasAlcTax: false,
    alcTax: 0,
  },
  {
    productId: 'P002',
    branch: 'HK',
    area: 'NORD',
    ubicazione: 'MTO',
    pltPerContainer: 25,
    hasCert: false,
    hasAlcTax: false,
    alcTax: 0,
  },
  {
    productId: 'P003',
    branch: 'HK',
    area: 'NORD',
    ubicazione: 'MTO',
    pltPerContainer: 23,
    hasCert: true,
    hasAlcTax: false,
    alcTax: 0,
  },
  {
    productId: 'P004',
    branch: 'HK',
    area: 'NORD',
    ubicazione: 'FOR',
    pltPerContainer: 23,
    hasCert: true,
    hasAlcTax: false,
    alcTax: 0,
  },
  {
    productId: 'P005',
    branch: 'HK',
    area: 'NORD',
    ubicazione: 'MTO',
    pltPerContainer: 23,
    hasCert: false,
    hasAlcTax: false,
    alcTax: 0,
  },
  {
    productId: 'P006',
    branch: 'HK',
    area: 'NORD',
    ubicazione: 'MTO',
    pltPerContainer: 23,
    hasCert: false,
    hasAlcTax: true,
    alcTax: 0.45,
  },
  {
    productId: 'P007',
    branch: 'HK',
    area: 'NORD',
    ubicazione: 'MTO',
    pltPerContainer: 23,
    hasCert: false,
    hasAlcTax: true,
    alcTax: 0.38,
  },
  {
    productId: 'P008',
    branch: 'HK',
    area: 'NORD',
    ubicazione: 'MTO',
    pltPerContainer: 23,
    hasCert: true,
    hasAlcTax: false,
    alcTax: 0,
  },
  {
    productId: 'P009',
    branch: 'HK',
    area: 'NORD',
    ubicazione: 'MTO',
    pltPerContainer: 23,
    hasCert: true,
    hasAlcTax: false,
    alcTax: 0,
  },
  {
    productId: 'P010',
    branch: 'HK',
    area: 'NORD',
    ubicazione: 'MTS',
    pltPerContainer: 23,
    hasCert: false,
    hasAlcTax: false,
    alcTax: 0,
  },
];

const SEED_PRICES = [
  { productId: 'P001', branch: 'HK', month: '2026-05', priceEur: 1.4941 },
  { productId: 'P002', branch: 'HK', month: '2026-05', priceEur: 2.12 },
  { productId: 'P003', branch: 'HK', month: '2026-05', priceEur: 5.5556 },
  { productId: 'P004', branch: 'HK', month: '2026-05', priceEur: 19.181 },
  { productId: 'P005', branch: 'HK', month: '2026-05', priceEur: 19.0 },
  { productId: 'P006', branch: 'HK', month: '2026-05', priceEur: 3.2 },
  { productId: 'P007', branch: 'HK', month: '2026-05', priceEur: 2.95 },
  { productId: 'P008', branch: 'HK', month: '2026-05', priceEur: 5.2828 },
  { productId: 'P009', branch: 'HK', month: '2026-05', priceEur: 8.5172 },
  { productId: 'P010', branch: 'HK', month: '2026-05', priceEur: 18.809 },
  { productId: 'P001', branch: 'HK', month: '2026-06', priceEur: 1.551 },
  { productId: 'P002', branch: 'HK', month: '2026-06', priceEur: 2.12 },
  { productId: 'P003', branch: 'HK', month: '2026-06', priceEur: 5.8 },
  { productId: 'P004', branch: 'HK', month: '2026-06', priceEur: 19.181 },
  { productId: 'P005', branch: 'HK', month: '2026-06', priceEur: 19.0 },
  { productId: 'P006', branch: 'HK', month: '2026-06', priceEur: 3.05 },
  { productId: 'P007', branch: 'HK', month: '2026-06', priceEur: 2.95 },
  { productId: 'P008', branch: 'HK', month: '2026-06', priceEur: 5.2828 },
  { productId: 'P009', branch: 'HK', month: '2026-06', priceEur: 8.9 },
  { productId: 'P010', branch: 'HK', month: '2026-06', priceEur: 18.809 },
];

const SEED_FX = [
  { branch: 'HK', month: '2026-05', rate: 9.12 },
  { branch: 'HK', month: '2026-06', rate: 9.1437 },
  { branch: 'CAN', month: '2026-05', rate: 1.0 },
  { branch: 'CAN', month: '2026-06', rate: 1.0 },
  { branch: 'AUS', month: '2026-05', rate: 1.62 },
  { branch: 'AUS', month: '2026-06', rate: 1.642 },
];

// ─────────────────────────────────────────────────────────────────────────────
// BC COLUMN MAPPING CONFIG
// Chiave = campo IFB interno, values = possibili nomi colonna BC (case-insensitive)
// ─────────────────────────────────────────────────────────────────────────────
const BC_FIELD_ALIASES = {
  code: [
    'no.',
    'no',
    'item no.',
    'item no',
    'codice',
    'code',
    'item_no',
    'item no_',
  ],
  description: [
    'description',
    'descrizione',
    'desc',
    'item description',
    'item_description',
  ],
  category: [
    'item category code',
    'item category',
    'category',
    'categoria',
    'cat',
  ],
  uom: [
    'base unit of measure',
    'uom',
    'unit of measure',
    'base uom',
    'unit',
    'unità',
  ],
  qtyPerBox: [
    'units per parcel',
    'qty per box',
    'qty/box',
    'pz per cartone',
    'units_per_parcel',
    'unità per collo',
  ],
  boxPerPallet: [
    'parcels per pallet',
    'box per pallet',
    'cartoni per pallet',
    'parcels_per_pallet',
    'colli per pallet',
  ],
  kgPerBox: [
    'net weight',
    'kg per box',
    'peso netto',
    'net_weight',
    'peso netto cartone',
  ],
  temperature: [
    'item tracking code',
    'temperatura',
    'temperature',
    'tracking code',
    'item_tracking_code',
    'storage',
  ],
  active: ['blocked', 'bloccato', 'active', 'attivo', 'stato'],
};

// Mapping valori BC → valori IFB
const BC_VALUE_MAP = {
  category: {
    food: 'FOOD',
    alimenti: 'FOOD',
    'f&b': 'FOOD',
    beverage: 'WINE',
    wine: 'WINE',
    vino: 'WINE',
    bevande: 'WINE',
    meat: 'MEAT',
    carni: 'MEAT',
    fish: 'MEAT',
    pesce: 'MEAT',
    salumi: 'MEAT',
  },
  uom: {
    pcs: 'PCS',
    pz: 'PCS',
    pezzo: 'PCS',
    pezzi: 'PCS',
    piece: 'PCS',
    box: 'BOX',
    ctn: 'BOX',
    cartone: 'BOX',
    collo: 'BOX',
    kg: 'KG',
    kgs: 'KG',
    kilogram: 'KG',
  },
  temperature: {
    dry: 'DRY',
    secco: 'DRY',
    ambient: 'DRY',
    amb: 'DRY',
    drygoods: 'DRY',
    fresh: 'FRESH',
    fresco: 'FRESH',
    chilled: 'FRESH',
    refrigerated: 'FRESH',
    frozen: 'FROZEN',
    surgelato: 'FROZEN',
    congelato: 'FROZEN',
    freeze: 'FROZEN',
  },
  active: {
    // "Blocked" in BC: true = bloccato = active:false
    true: false,
    yes: false,
    si: false,
    '1': false,
    blocked: false,
    false: true,
    no: true,
    '0': true,
    '': true,
  },
};

function mapBCValue(field, raw) {
  if (!BC_VALUE_MAP[field]) return raw;
  const key = String(raw || '')
    .toLowerCase()
    .trim();
  return BC_VALUE_MAP[field][key] !== undefined
    ? BC_VALUE_MAP[field][key]
    : raw;
}

function detectBCColumn(headers, field) {
  const aliases = BC_FIELD_ALIASES[field] || [];
  for (const h of headers) {
    const hLow = h.toLowerCase().trim();
    if (aliases.some((a) => hLow === a || hLow.includes(a))) return h;
  }
  return null;
}

// ─────────────────────────────────────────────────────────────────────────────
// COST ENGINE (HK model)
// ─────────────────────────────────────────────────────────────────────────────
const C = {
  FOB_HK: {
    DRY: { NORD: 2000, CENTRO: 0, SUD: 1108.55 },
    FRESH: { NORD: 3500, CENTRO: 3500, SUD: 0 },
    FROZEN: { NORD: 4000, CENTRO: 0, SUD: 0 },
  },
  LIC: 863.9828,
  VGM: 100,
  HC: 80,
  PLT: 30,
  MTO: { DRY: 8.16, FRESH: 10.2, FROZEN: 12.24 },
  MTS_D: { DRY: 14.4228, FRESH: 16.4832, FROZEN: 24.7248 },
  MTS_I: { DRY: 2.5755, FRESH: 3.6057, FROZEN: 3.6057 },
  MTS_P: { DRY: 0.303, FRESH: 0.3434, FROZEN: 0.3535 },
};

function calcHK({ priceEur, product, logistic, eurToHkd }) {
  const { uom, qtyPerBox, boxPerPallet, kgPerBox, temperature } = product;
  const { area, ubicazione, pltPerContainer, hasCert, hasAlcTax, alcTax } =
    logistic;
  let unitsPerPlt;
  if (uom === 'BOX') unitsPerPlt = boxPerPallet;
  else if (uom === 'KG') unitsPerPlt = (kgPerBox || qtyPerBox) * boxPerPallet;
  else unitsPerPlt = qtyPerBox * boxPerPallet;
  const pickDiv =
    uom === 'BOX' ? 1 : uom === 'KG' ? kgPerBox || qtyPerBox : qtyPerBox;
  const totalUnits = unitsPerPlt * pltPerContainer;
  if (!totalUnits) return null;
  const fob = (C.FOB_HK[temperature]?.[area] ?? 0) / totalUnits;
  const lic = C.LIC / totalUnits;
  const vgm = C.VGM / totalUnits;
  const hc = hasCert ? C.HC / totalUnits : 0;
  const plt = C.PLT / unitsPerPlt;
  const alc = hasAlcTax ? alcTax || 0 : 0;
  const step1 = priceEur + fob + lic + vgm + hc + plt + alc;
  let wh = 0,
    whDetail = {};
  if (ubicazione === 'MTO') {
    wh = C.MTO[temperature] / unitsPerPlt;
    whDetail = { type: 'MTO', total: wh };
  } else if (ubicazione === 'MTS') {
    const d = C.MTS_D[temperature] / unitsPerPlt;
    const i = C.MTS_I[temperature] / unitsPerPlt;
    const p = C.MTS_P[temperature] / pickDiv;
    wh = d + i + p;
    whDetail = { type: 'MTS', dep: d, inbound: i, picking: p, total: wh };
  } else {
    whDetail = { type: 'FOR', total: 0 };
  }
  const step2Eur = step1 + wh;
  const rate = eurToHkd || 9.1437;
  return {
    priceEur,
    fob,
    lic,
    vgm,
    hc,
    plt,
    alc,
    step1Eur: step1,
    step1Hkd: step1 * rate,
    wh,
    whDetail,
    step2Eur,
    step2Hkd: step2Eur * rate,
    rate,
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// LOCAL STORAGE
// ─────────────────────────────────────────────────────────────────────────────
const LS = {
  get: (k, def) => {
    try {
      const v = localStorage.getItem(k);
      return v ? JSON.parse(v) : def;
    } catch {
      return def;
    }
  },
  set: (k, v) => {
    try {
      localStorage.setItem(k, JSON.stringify(v));
    } catch {}
  },
};

// ─────────────────────────────────────────────────────────────────────────────
// ROOT APP
// ─────────────────────────────────────────────────────────────────────────────
export default function App() {
  const [products, setProducts] = useState(() =>
    LS.get('ifb_products', SEED_PRODUCTS)
  );
  const [logistics, setLogistics] = useState(() =>
    LS.get('ifb_logistics', SEED_LOGISTIC)
  );
  const [prices, setPrices] = useState(() => LS.get('ifb_prices', SEED_PRICES));
  const [fx, setFx] = useState(() => LS.get('ifb_fx', SEED_FX));
  const [sentMails, setSentMails] = useState(() => LS.get('ifb_mails', []));
  const [importLogs, setImportLogs] = useState(() =>
    LS.get('ifb_importlogs', [])
  );

  const [page, setPage] = useState('dashboard');
  const [branch, setBranch] = useState('HK');
  const [month, setMonth] = useState(NOW());
  const [toast, setToast] = useState(null);

  useEffect(() => {
    LS.set('ifb_products', products);
  }, [products]);
  useEffect(() => {
    LS.set('ifb_logistics', logistics);
  }, [logistics]);
  useEffect(() => {
    LS.set('ifb_prices', prices);
  }, [prices]);
  useEffect(() => {
    LS.set('ifb_fx', fx);
  }, [fx]);
  useEffect(() => {
    LS.set('ifb_mails', sentMails);
  }, [sentMails]);
  useEffect(() => {
    LS.set('ifb_importlogs', importLogs);
  }, [importLogs]);

  const showToast = (msg, color = T.green) => {
    setToast({ msg, color });
    setTimeout(() => setToast(null), 3500);
  };

  const costRows = useMemo(() => {
    const fxRate =
      fx.find((f) => f.branch === branch && f.month === month)?.rate ||
      BRANCH_CFG[branch].defaultRate;
    const prevMonth =
      month.slice(0, 4) +
      '-' +
      String(parseInt(month.slice(5)) - 1).padStart(2, '0');
    return products
      .filter((p) => p.active)
      .map((prod) => {
        const log = logistics.find(
          (l) => l.productId === prod.id && l.branch === branch
        );
        const pr = prices.find(
          (p) =>
            p.productId === prod.id && p.branch === branch && p.month === month
        );
        const prPrev = prices.find(
          (p) =>
            p.productId === prod.id &&
            p.branch === branch &&
            p.month === prevMonth
        );
        if (!log || !pr) return { ...prod, cost: null, prevCost: null, log };
        const cost = calcHK({
          priceEur: pr.priceEur,
          product: prod,
          logistic: log,
          eurToHkd: fxRate,
        });
        const prevCost = prPrev
          ? calcHK({
              priceEur: prPrev.priceEur,
              product: prod,
              logistic: log,
              eurToHkd: fxRate,
            })
          : null;
        const delta = prevCost ? PCT(cost.step2Eur, prevCost.step2Eur) : null;
        return {
          ...prod,
          cost,
          prevCost,
          delta,
          priceEur: pr.priceEur,
          isNew: !prPrev,
          flagged: delta !== null && Math.abs(delta) >= 3,
        };
      });
  }, [products, logistics, prices, fx, branch, month]);

  const branchCfg = BRANCH_CFG[branch];

  const NAV = [
    { id: 'dashboard', icon: '⬡', label: 'Dashboard' },
    { id: 'products', icon: '◈', label: 'Anagrafica' },
    { id: 'import', icon: '⇪', label: 'Import BC' },
    { id: 'logistics', icon: '◎', label: 'Logistica' },
    { id: 'prices', icon: '◉', label: 'Listini' },
    { id: 'fx', icon: '◌', label: 'Cambi' },
    { id: 'costs', icon: '◆', label: 'Standard Cost' },
    { id: 'mail', icon: '◻', label: 'Mail Mensile' },
  ];

  const pages = {
    dashboard: (
      <Dashboard
        costRows={costRows}
        branch={branch}
        month={month}
        branchCfg={branchCfg}
        setPage={setPage}
      />
    ),
    products: (
      <Products
        products={products}
        setProducts={setProducts}
        showToast={showToast}
      />
    ),
    import: (
      <ImportBC
        products={products}
        setProducts={setProducts}
        importLogs={importLogs}
        setImportLogs={setImportLogs}
        showToast={showToast}
      />
    ),
    logistics: (
      <Logistics
        logistics={logistics}
        setLogistics={setLogistics}
        products={products}
        branch={branch}
        showToast={showToast}
      />
    ),
    prices: (
      <Prices
        prices={prices}
        setPrices={setPrices}
        products={products}
        branch={branch}
        month={month}
        showToast={showToast}
      />
    ),
    fx: (
      <FxRates
        fx={fx}
        setFx={setFx}
        branch={branch}
        month={month}
        showToast={showToast}
      />
    ),
    costs: (
      <CostTable
        costRows={costRows}
        branch={branch}
        month={month}
        branchCfg={branchCfg}
      />
    ),
    mail: (
      <MailGen
        costRows={costRows}
        branch={branch}
        month={month}
        branchCfg={branchCfg}
        sentMails={sentMails}
        setSentMails={setSentMails}
        showToast={showToast}
      />
    ),
  };

  return (
    <div
      style={{
        display: 'flex',
        minHeight: '100vh',
        background: T.bg,
        fontFamily: "'Palatino Linotype','Book Antiqua',Palatino,serif",
        color: T.text,
      }}
    >
      {/* SIDEBAR */}
      <div
        style={{
          width: '220px',
          flexShrink: 0,
          background: T.surface,
          borderRight: `1px solid ${T.border}`,
          display: 'flex',
          flexDirection: 'column',
          position: 'sticky',
          top: 0,
          height: '100vh',
        }}
      >
        <div
          style={{
            padding: '28px 20px 20px',
            borderBottom: `1px solid ${T.border}`,
          }}
        >
          <div
            style={{
              fontSize: '11px',
              letterSpacing: '3px',
              color: T.gold,
              textTransform: 'uppercase',
              marginBottom: '4px',
            }}
          >
            IFB Platform
          </div>
          <div
            style={{ fontSize: '17px', fontWeight: 'bold', lineHeight: 1.2 }}
          >
            Cost Intelligence
          </div>
        </div>
        <div
          style={{
            padding: '14px 16px',
            borderBottom: `1px solid ${T.border}`,
          }}
        >
          <div
            style={{
              fontSize: '9px',
              letterSpacing: '2px',
              color: T.dim,
              textTransform: 'uppercase',
              marginBottom: '8px',
            }}
          >
            Filiale
          </div>
          <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
            {Object.entries(BRANCH_CFG).map(([key, cfg]) => (
              <button
                key={key}
                onClick={() => setBranch(key)}
                style={{
                  padding: '7px 10px',
                  background: branch === key ? `${cfg.color}20` : 'transparent',
                  border: `1px solid ${
                    branch === key ? cfg.color : 'transparent'
                  }`,
                  borderRadius: '6px',
                  color: branch === key ? cfg.color : T.muted,
                  cursor: 'pointer',
                  fontFamily: 'inherit',
                  fontSize: '12px',
                  textAlign: 'left',
                  display: 'flex',
                  alignItems: 'center',
                  gap: '7px',
                  transition: 'all 0.2s',
                }}
              >
                <span>{cfg.flag}</span>
                {cfg.label}
                {key === 'AUS' && (
                  <span
                    style={{
                      fontSize: '8px',
                      color: T.orange,
                      marginLeft: 'auto',
                      background: `${T.orange}22`,
                      padding: '1px 5px',
                      borderRadius: '3px',
                    }}
                  >
                    SOON
                  </span>
                )}
              </button>
            ))}
          </div>
        </div>
        <div
          style={{
            padding: '14px 16px',
            borderBottom: `1px solid ${T.border}`,
          }}
        >
          <div
            style={{
              fontSize: '9px',
              letterSpacing: '2px',
              color: T.dim,
              textTransform: 'uppercase',
              marginBottom: '6px',
            }}
          >
            Mese di riferimento
          </div>
          <input
            type="month"
            value={month}
            onChange={(e) => setMonth(e.target.value)}
            style={{
              width: '100%',
              padding: '6px 8px',
              background: 'rgba(255,255,255,0.05)',
              border: `1px solid ${T.border}`,
              borderRadius: '6px',
              color: T.text,
              fontFamily: 'inherit',
              fontSize: '12px',
              outline: 'none',
              boxSizing: 'border-box',
            }}
          />
        </div>
        <nav
          style={{
            flex: 1,
            padding: '10px',
            display: 'flex',
            flexDirection: 'column',
            gap: '2px',
          }}
        >
          {NAV.map((n) => (
            <button
              key={n.id}
              onClick={() => setPage(n.id)}
              style={{
                padding: '9px 12px',
                background: page === n.id ? T.goldDim : 'transparent',
                border: `1px solid ${
                  page === n.id ? T.gold + '44' : 'transparent'
                }`,
                borderRadius: '7px',
                color: page === n.id ? T.gold : T.muted,
                cursor: 'pointer',
                fontFamily: 'inherit',
                fontSize: '13px',
                textAlign: 'left',
                display: 'flex',
                alignItems: 'center',
                gap: '10px',
                transition: 'all 0.18s',
              }}
            >
              <span style={{ fontSize: '12px', opacity: 0.8 }}>{n.icon}</span>
              {n.label}
              {n.id === 'import' && (
                <span
                  style={{
                    marginLeft: 'auto',
                    fontSize: '8px',
                    background: `${T.blue}33`,
                    color: T.blue,
                    padding: '1px 5px',
                    borderRadius: '3px',
                  }}
                >
                  BC
                </span>
              )}
            </button>
          ))}
        </nav>
        <div
          style={{
            padding: '14px 16px',
            borderTop: `1px solid ${T.border}`,
            fontSize: '10px',
            color: T.dim,
          }}
        >
          Inalca Food & Beverage
          <br />© 2026 · Cost Platform v2.0
        </div>
      </div>

      {/* MAIN */}
      <div
        style={{
          flex: 1,
          display: 'flex',
          flexDirection: 'column',
          minWidth: 0,
        }}
      >
        <div
          style={{
            padding: '16px 32px',
            borderBottom: `1px solid ${T.border}`,
            background: T.surface,
            display: 'flex',
            alignItems: 'center',
            gap: '12px',
            position: 'sticky',
            top: 0,
            zIndex: 10,
          }}
        >
          <span style={{ fontSize: '18px' }}>{branchCfg.flag}</span>
          <span style={{ fontSize: '14px', fontWeight: 'bold' }}>
            {branchCfg.label}
          </span>
          <span style={{ color: T.dim }}>·</span>
          <span style={{ fontSize: '13px', color: T.muted }}>
            {NAV.find((n) => n.id === page)?.label}
          </span>
          <span style={{ color: T.dim }}>·</span>
          <span style={{ fontSize: '12px', color: T.gold }}>{month}</span>
          <div style={{ marginLeft: 'auto', display: 'flex', gap: '8px' }}>
            <button
              onClick={() => setPage('import')}
              style={{
                padding: '6px 14px',
                background: 'rgba(74,143,181,0.15)',
                border: `1px solid ${T.blue}44`,
                borderRadius: '6px',
                color: T.blue,
                cursor: 'pointer',
                fontFamily: 'inherit',
                fontSize: '11px',
              }}
            >
              ⇪ Import BC
            </button>
            <button
              onClick={() => setPage('costs')}
              style={{
                padding: '6px 14px',
                background: 'rgba(255,255,255,0.05)',
                border: `1px solid ${T.border}`,
                borderRadius: '6px',
                color: T.muted,
                cursor: 'pointer',
                fontFamily: 'inherit',
                fontSize: '11px',
              }}
            >
              📊 Vedi Costi
            </button>
            <button
              onClick={() => setPage('mail')}
              style={{
                padding: '6px 14px',
                background: T.gold,
                border: `1px solid ${T.gold}`,
                borderRadius: '6px',
                color: T.bg,
                cursor: 'pointer',
                fontFamily: 'inherit',
                fontSize: '11px',
                fontWeight: 'bold',
              }}
            >
              ✉ Mail Mensile
            </button>
          </div>
        </div>
        <div style={{ flex: 1, padding: '28px 32px', overflow: 'auto' }}>
          {pages[page]}
        </div>
      </div>

      {toast && (
        <div
          style={{
            position: 'fixed',
            bottom: '28px',
            right: '28px',
            padding: '12px 20px',
            background: toast.color,
            borderRadius: '8px',
            color: '#fff',
            fontSize: '13px',
            fontWeight: 'bold',
            boxShadow: '0 8px 24px rgba(0,0,0,0.4)',
            zIndex: 1000,
          }}
        >
          {toast.msg}
        </div>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// IMPORT BC MODULE
// ─────────────────────────────────────────────────────────────────────────────
function ImportBC({
  products,
  setProducts,
  importLogs,
  setImportLogs,
  showToast,
}) {
  const [step, setStep] = useState('upload'); // upload | map | preview | done
  const [rawRows, setRawRows] = useState([]);
  const [headers, setHeaders] = useState([]);
  const [fileName, setFileName] = useState('');
  const [mapping, setMapping] = useState({});
  const [preview, setPreview] = useState([]);
  const [actions, setActions] = useState({}); // productId/rowIdx → "NEW"|"UPDATE"|"SKIP"
  const dropRef = useRef();

  // ── Step 1: Parse file ──
  const parseFile = (file) => {
    if (!file) return;
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'binary' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
        if (data.length < 2) {
          showToast('File vuoto o formato non valido', T.red);
          return;
        }
        const hdrs = data[0].map((h) => String(h).trim()).filter((h) => h);
        const rows = data.slice(1).filter((r) => r.some((cell) => cell !== ''));
        setHeaders(hdrs);
        setRawRows(rows);
        // Auto-detect mapping
        const autoMap = {};
        Object.keys(BC_FIELD_ALIASES).forEach((field) => {
          const detected = detectBCColumn(hdrs, field);
          if (detected) autoMap[field] = detected;
        });
        setMapping(autoMap);
        setStep('map');
      } catch (err) {
        showToast('Errore lettura file: ' + err.message, T.red);
      }
    };
    reader.readAsBinaryString(file);
  };

  const handleDrop = (e) => {
    e.preventDefault();
    const file = e.dataTransfer.files[0];
    if (file) parseFile(file);
  };

  const handleFileInput = (e) => {
    const file = e.target.files[0];
    if (file) parseFile(file);
  };

  // ── Step 2: Build preview from mapping ──
  const buildPreview = () => {
    const mapped = rawRows
      .map((row, idx) => {
        const get = (field) => {
          const col = mapping[field];
          if (!col) return null;
          const colIdx = headers.indexOf(col);
          return colIdx >= 0 ? row[colIdx] : null;
        };
        const code = String(get('code') || '').trim();
        if (!code) return null;
        const desc = String(get('description') || '').trim();
        const rawCat = get('category');
        const rawUom = get('uom');
        const rawTemp = get('temperature');
        const rawBlocked = get('active');
        const existing = products.find((p) => p.code === code);
        return {
          _idx: idx,
          code,
          description: desc || existing?.description || '',
          category:
            mapBCValue('category', rawCat) || existing?.category || 'FOOD',
          uom: mapBCValue('uom', rawUom) || existing?.uom || 'PCS',
          temperature:
            mapBCValue('temperature', rawTemp) ||
            existing?.temperature ||
            'DRY',
          active:
            rawBlocked !== null
              ? mapBCValue('active', rawBlocked)
              : existing?.active ?? true,
          qtyPerBox: parseFloat(get('qtyPerBox')) || existing?.qtyPerBox || 1,
          boxPerPallet:
            parseFloat(get('boxPerPallet')) || existing?.boxPerPallet || 1,
          kgPerBox: parseFloat(get('kgPerBox')) || existing?.kgPerBox || null,
          _isNew: !existing,
          _existing: existing,
        };
      })
      .filter(Boolean);
    setPreview(mapped);
    const defaultActions = {};
    mapped.forEach((r) => {
      defaultActions[r._idx] = r._isNew ? 'NEW' : 'UPDATE';
    });
    setActions(defaultActions);
    setStep('preview');
  };

  // ── Step 3: Execute import ──
  const executeImport = () => {
    const toProcess = preview.filter((r) => actions[r._idx] !== 'SKIP');
    let newCount = 0,
      updateCount = 0;
    const updated = [...products];
    toProcess.forEach((r) => {
      const action = actions[r._idx];
      if (action === 'NEW') {
        updated.push({
          id: 'P_BC_' + Date.now() + '_' + r._idx,
          code: r.code,
          description: r.description,
          category: r.category,
          uom: r.uom,
          qtyPerBox: r.qtyPerBox,
          boxPerPallet: r.boxPerPallet,
          kgPerBox: r.kgPerBox,
          temperature: r.temperature,
          active: r.active,
        });
        newCount++;
      } else if (action === 'UPDATE') {
        const idx = updated.findIndex((p) => p.code === r.code);
        if (idx >= 0) {
          updated[idx] = {
            ...updated[idx],
            description: r.description,
            category: r.category,
            uom: r.uom,
            qtyPerBox: r.qtyPerBox,
            boxPerPallet: r.boxPerPallet,
            kgPerBox: r.kgPerBox,
            temperature: r.temperature,
            active: r.active,
          };
          updateCount++;
        }
      }
    });
    setProducts(updated);
    setImportLogs((logs) => [
      {
        id: Date.now(),
        fileName,
        importedAt: new Date().toISOString(),
        totalRows: rawRows.length,
        newCount,
        updateCount,
        skippedCount: preview.filter((r) => actions[r._idx] === 'SKIP').length,
      },
      ...logs,
    ]);
    showToast(
      `Import completato: ${newCount} nuovi, ${updateCount} aggiornati ✓`,
      T.gold
    );
    setStep('done');
  };

  const reset = () => {
    setStep('upload');
    setRawRows([]);
    setHeaders([]);
    setFileName('');
    setMapping({});
    setPreview([]);
    setActions({});
  };

  const mappedFields = Object.keys(BC_FIELD_ALIASES);
  const requiredFields = ['code', 'description'];

  return (
    <div>
      <PageHeader
        title="⇪ Import Anagrafica da Business Central"
        sub="Carica un export Excel/CSV da Power BI o BC · mappatura automatica delle colonne"
      />

      {/* Progress indicator */}
      <div
        style={{
          display: 'flex',
          alignItems: 'center',
          gap: '0',
          marginBottom: '28px',
        }}
      >
        {[
          ['upload', '1. Carica file'],
          ['map', '2. Mappa colonne'],
          ['preview', '3. Preview & azioni'],
          ['done', '4. Completato'],
        ].map(([s, l], i, arr) => (
          <div key={s} style={{ display: 'flex', alignItems: 'center' }}>
            <div
              style={{
                display: 'flex',
                flexDirection: 'column',
                alignItems: 'center',
                gap: '4px',
              }}
            >
              <div
                style={{
                  width: '28px',
                  height: '28px',
                  borderRadius: '50%',
                  background:
                    step === s
                      ? T.gold
                      : ['upload', 'map', 'preview', 'done'].indexOf(step) > i
                      ? `${T.gold}44`
                      : 'rgba(255,255,255,0.06)',
                  border: `2px solid ${
                    step === s
                      ? T.gold
                      : ['upload', 'map', 'preview', 'done'].indexOf(step) > i
                      ? `${T.gold}66`
                      : T.border
                  }`,
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  fontSize: '11px',
                  fontWeight: 'bold',
                  color:
                    step === s
                      ? T.bg
                      : ['upload', 'map', 'preview', 'done'].indexOf(step) > i
                      ? T.gold
                      : T.dim,
                }}
              >
                {i + 1}
              </div>
              <span
                style={{
                  fontSize: '10px',
                  color: step === s ? T.gold : T.dim,
                  whiteSpace: 'nowrap',
                }}
              >
                {l}
              </span>
            </div>
            {i < arr.length - 1 && (
              <div
                style={{
                  width: '60px',
                  height: '1px',
                  background: T.border,
                  margin: '0 4px',
                  marginBottom: '18px',
                }}
              />
            )}
          </div>
        ))}
      </div>

      {/* ── STEP 1: UPLOAD ── */}
      {step === 'upload' && (
        <div>
          <Section
            title="Carica file di export Business Central / Power BI"
            mb="20px"
          >
            <div
              ref={dropRef}
              onDragOver={(e) => e.preventDefault()}
              onDrop={handleDrop}
              style={{
                border: `2px dashed ${T.borderHi}`,
                borderRadius: '12px',
                padding: '48px 32px',
                textAlign: 'center',
                background: 'rgba(255,255,255,0.02)',
                cursor: 'pointer',
                transition: 'border-color 0.2s',
              }}
              onClick={() => document.getElementById('bc_file_input').click()}
            >
              <div style={{ fontSize: '40px', marginBottom: '12px' }}>⇪</div>
              <div
                style={{ fontSize: '15px', color: T.text, marginBottom: '6px' }}
              >
                Trascina qui il file Excel o CSV
              </div>
              <div
                style={{
                  fontSize: '12px',
                  color: T.muted,
                  marginBottom: '18px',
                }}
              >
                Supportati: .xlsx, .xls, .csv — Prima riga = intestazioni
              </div>
              <ActionBtn
                label="Sfoglia file…"
                onClick={(e) => {
                  e.stopPropagation();
                  document.getElementById('bc_file_input').click();
                }}
                primary
              />
              <input
                id="bc_file_input"
                type="file"
                accept=".xlsx,.xls,.csv"
                onChange={handleFileInput}
                style={{ display: 'none' }}
              />
            </div>
          </Section>

          {/* Format guide */}
          <Section title="Colonne riconosciute automaticamente (nomi BC standard)">
            <div
              style={{
                display: 'grid',
                gridTemplateColumns: 'repeat(2,1fr)',
                gap: '8px',
              }}
            >
              {[
                ['Codice prodotto', 'No. · Item No. · Codice'],
                ['Descrizione', 'Description · Descrizione'],
                ['Categoria', 'Item Category Code · Categoria'],
                ['Unità di misura', 'Base Unit of Measure · UOM'],
                ['Pz per cartone', 'Units per Parcel · Qty per Box'],
                ['Cartoni/pallet', 'Parcels per Pallet · Box per Pallet'],
                ['Peso netto', 'Net Weight · Kg per Box'],
                ['Temperatura', 'Item Tracking Code · Temperature'],
                ['Bloccato', 'Blocked · Bloccato (invertito → Active)'],
              ].map(([campo, bc]) => (
                <div
                  key={campo}
                  style={{
                    display: 'flex',
                    gap: '10px',
                    padding: '8px 10px',
                    background: 'rgba(255,255,255,0.03)',
                    borderRadius: '6px',
                  }}
                >
                  <span
                    style={{
                      fontSize: '11px',
                      color: T.gold,
                      minWidth: '130px',
                    }}
                  >
                    {campo}
                  </span>
                  <span style={{ fontSize: '11px', color: T.dim }}>{bc}</span>
                </div>
              ))}
            </div>
            <div
              style={{
                marginTop: '14px',
                padding: '10px 14px',
                background: `${T.blue}11`,
                border: `1px solid ${T.blue}33`,
                borderRadius: '8px',
                fontSize: '12px',
                color: T.muted,
              }}
            >
              💡 <strong style={{ color: T.text }}>Consiglio:</strong> Se le
              colonne hanno nomi diversi da quelli standard, potrai mapparle
              manualmente nel passo successivo. Non devi rinominare nulla nel
              file BC.
            </div>
          </Section>
        </div>
      )}

      {/* ── STEP 2: COLUMN MAPPING ── */}
      {step === 'map' && (
        <div>
          <Section
            title={`Mappatura colonne — ${fileName} · ${rawRows.length} righe rilevate`}
            mb="20px"
          >
            <div
              style={{ fontSize: '12px', color: T.muted, marginBottom: '16px' }}
            >
              La mappatura automatica ha rilevato le colonne seguenti. Verifica
              e correggi se necessario.
            </div>
            <div
              style={{
                display: 'grid',
                gridTemplateColumns: 'repeat(2,1fr)',
                gap: '12px',
              }}
            >
              {mappedFields.map((field) => {
                const isRequired = requiredFields.includes(field);
                return (
                  <div key={field}>
                    <label
                      style={{
                        display: 'block',
                        fontSize: '11px',
                        color: isRequired ? T.gold : T.muted,
                        marginBottom: '5px',
                      }}
                    >
                      {field === 'code'
                        ? 'Codice *'
                        : field === 'description'
                        ? 'Descrizione *'
                        : field === 'category'
                        ? 'Categoria'
                        : field === 'uom'
                        ? 'Unità di misura'
                        : field === 'qtyPerBox'
                        ? 'Pz per cartone'
                        : field === 'boxPerPallet'
                        ? 'Cartoni/pallet'
                        : field === 'kgPerBox'
                        ? 'Kg per cartone'
                        : field === 'temperature'
                        ? 'Temperatura'
                        : 'Attivo/Bloccato'}
                    </label>
                    <select
                      value={mapping[field] || ''}
                      onChange={(e) =>
                        setMapping((m) => ({
                          ...m,
                          [field]: e.target.value || null,
                        }))
                      }
                      style={{
                        ...inputStyle(),
                        cursor: 'pointer',
                        borderColor:
                          !mapping[field] && isRequired
                            ? T.red + '88'
                            : T.border,
                      }}
                    >
                      <option value="">— non mappato —</option>
                      {headers.map((h) => (
                        <option key={h} value={h}>
                          {h}
                        </option>
                      ))}
                    </select>
                  </div>
                );
              })}
            </div>
            {/* Sample data */}
            {rawRows.length > 0 && (
              <div style={{ marginTop: '16px' }}>
                <div
                  style={{
                    fontSize: '10px',
                    color: T.dim,
                    letterSpacing: '1px',
                    textTransform: 'uppercase',
                    marginBottom: '8px',
                  }}
                >
                  Anteprima prime 3 righe del file
                </div>
                <div style={{ overflowX: 'auto' }}>
                  <table
                    style={{
                      borderCollapse: 'collapse',
                      fontSize: '11px',
                      whiteSpace: 'nowrap',
                    }}
                  >
                    <thead>
                      <tr>
                        {headers.map((h) => (
                          <th
                            key={h}
                            style={{
                              padding: '4px 10px',
                              color: T.muted,
                              borderBottom: `1px solid ${T.border}`,
                              textAlign: 'left',
                              fontWeight: 'normal',
                            }}
                          >
                            {h}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {rawRows.slice(0, 3).map((row, i) => (
                        <tr
                          key={i}
                          style={{ borderBottom: `1px solid ${T.border}` }}
                        >
                          {headers.map((h, j) => (
                            <td
                              key={j}
                              style={{
                                padding: '4px 10px',
                                color: T.text,
                                fontFamily: 'monospace',
                              }}
                            >
                              {String(row[j] || '')}
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
            <div style={{ display: 'flex', gap: '10px', marginTop: '18px' }}>
              <ActionBtn label="← Ricarica file" onClick={reset} />
              <ActionBtn
                label="Continua → Preview"
                onClick={buildPreview}
                primary
                disabled={requiredFields.some((f) => !mapping[f])}
              />
            </div>
          </Section>
        </div>
      )}

      {/* ── STEP 3: PREVIEW ── */}
      {step === 'preview' && (
        <div>
          {/* Summary pills */}
          <div style={{ display: 'flex', gap: '12px', marginBottom: '20px' }}>
            {[
              [
                preview.filter((r) => r._isNew).length,
                'Nuovi prodotti',
                T.gold,
              ],
              [
                preview.filter((r) => !r._isNew).length,
                'Da aggiornare',
                T.blue,
              ],
              [
                preview.filter((r) => actions[r._idx] === 'SKIP').length,
                'Da saltare',
                T.muted,
              ],
              [preview.length, 'Totale righe', T.text],
            ].map(([n, l, c]) => (
              <div
                key={l}
                style={{
                  padding: '10px 16px',
                  background: T.card,
                  border: `1px solid ${T.border}`,
                  borderRadius: '8px',
                }}
              >
                <div style={{ fontSize: '20px', fontWeight: 'bold', color: c }}>
                  {n}
                </div>
                <div
                  style={{ fontSize: '10px', color: T.dim, marginTop: '2px' }}
                >
                  {l}
                </div>
              </div>
            ))}
          </div>

          <Section
            title={`Preview import · ${preview.length} prodotti · clicca sull'azione per cambiarla`}
          >
            <div style={{ overflowX: 'auto' }}>
              <table
                style={{
                  width: '100%',
                  borderCollapse: 'collapse',
                  fontSize: '12px',
                }}
              >
                <THead
                  cols={[
                    'Codice',
                    'Descrizione',
                    'Cat.',
                    'UOM',
                    'Temp.',
                    'Qty/Box',
                    'Box/Plt',
                    'Stato',
                    'Azione',
                  ]}
                />
                <tbody>
                  {preview.map((r) => {
                    const act = actions[r._idx] || 'SKIP';
                    const actColor =
                      act === 'NEW'
                        ? T.gold
                        : act === 'UPDATE'
                        ? T.blue
                        : T.dim;
                    const isSkipped = act === 'SKIP';
                    return (
                      <tr
                        key={r._idx}
                        style={{
                          borderBottom: `1px solid ${T.border}`,
                          opacity: isSkipped ? 0.4 : 1,
                          background: r._isNew
                            ? `${T.gold}07`
                            : act === 'UPDATE'
                            ? `${T.blue}07`
                            : '',
                        }}
                      >
                        <td
                          style={{
                            padding: '7px 10px',
                            fontFamily: 'monospace',
                            color: T.text,
                            fontSize: '11px',
                          }}
                        >
                          {r.code}
                        </td>
                        <td style={{ padding: '7px 10px', color: T.text }}>
                          {r.description}
                          {r._isNew && (
                            <span
                              style={{
                                marginLeft: '6px',
                                fontSize: '9px',
                                background: `${T.gold}22`,
                                color: T.gold,
                                padding: '1px 5px',
                                borderRadius: '3px',
                              }}
                            >
                              NUOVO
                            </span>
                          )}
                          {!r._isNew && act === 'UPDATE' && (
                            <span
                              style={{
                                marginLeft: '6px',
                                fontSize: '9px',
                                background: `${T.blue}22`,
                                color: T.blue,
                                padding: '1px 5px',
                                borderRadius: '3px',
                              }}
                            >
                              AGG.
                            </span>
                          )}
                        </td>
                        <td style={{ padding: '7px 10px' }}>
                          <Chip label={r.category} />
                        </td>
                        <td style={{ padding: '7px 10px' }}>
                          <Chip label={r.uom} color={T.blue} />
                        </td>
                        <td style={{ padding: '7px 10px' }}>
                          <TempChip t={r.temperature} />
                        </td>
                        <td
                          style={{
                            padding: '7px 10px',
                            fontFamily: 'monospace',
                            color: T.text,
                          }}
                        >
                          {r.qtyPerBox || '—'}
                        </td>
                        <td
                          style={{
                            padding: '7px 10px',
                            fontFamily: 'monospace',
                            color: T.text,
                          }}
                        >
                          {r.boxPerPallet || '—'}
                        </td>
                        <td style={{ padding: '7px 10px' }}>
                          <Chip
                            label={r.active ? 'Attivo' : 'Sospeso'}
                            color={r.active ? T.green : T.red}
                          />
                        </td>
                        <td style={{ padding: '7px 10px' }}>
                          <div style={{ display: 'flex', gap: '4px' }}>
                            {(r._isNew
                              ? ['NEW', 'SKIP']
                              : ['UPDATE', 'SKIP', 'NEW']
                            ).map((a) => (
                              <button
                                key={a}
                                onClick={() =>
                                  setActions((ac) => ({ ...ac, [r._idx]: a }))
                                }
                                style={{
                                  padding: '2px 8px',
                                  fontSize: '10px',
                                  fontWeight: act === a ? 'bold' : 'normal',
                                  cursor: 'pointer',
                                  fontFamily: 'inherit',
                                  background:
                                    act === a
                                      ? `${
                                          a === 'NEW'
                                            ? T.gold
                                            : a === 'UPDATE'
                                            ? T.blue
                                            : T.muted
                                        }22`
                                      : 'transparent',
                                  border: `1px solid ${
                                    act === a
                                      ? a === 'NEW'
                                        ? T.gold
                                        : a === 'UPDATE'
                                        ? T.blue
                                        : T.muted
                                      : T.border
                                  }`,
                                  borderRadius: '4px',
                                  color:
                                    act === a
                                      ? a === 'NEW'
                                        ? T.gold
                                        : a === 'UPDATE'
                                        ? T.blue
                                        : T.muted
                                      : T.dim,
                                }}
                              >
                                {a}
                              </button>
                            ))}
                          </div>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>

            {/* Bulk actions */}
            <div
              style={{
                display: 'flex',
                gap: '8px',
                marginTop: '14px',
                paddingTop: '14px',
                borderTop: `1px solid ${T.border}`,
              }}
            >
              <span
                style={{ fontSize: '11px', color: T.dim, alignSelf: 'center' }}
              >
                Selezione rapida:
              </span>
              <MiniBtn
                label="Tutto UPDATE"
                onClick={() =>
                  setActions((a) => {
                    const n = { ...a };
                    preview
                      .filter((r) => !r._isNew)
                      .forEach((r) => {
                        n[r._idx] = 'UPDATE';
                      });
                    return n;
                  })
                }
              />
              <MiniBtn
                label="Tutto SKIP (esistenti)"
                onClick={() =>
                  setActions((a) => {
                    const n = { ...a };
                    preview
                      .filter((r) => !r._isNew)
                      .forEach((r) => {
                        n[r._idx] = 'SKIP';
                      });
                    return n;
                  })
                }
              />
              <MiniBtn
                label="Salta tutti nuovi"
                onClick={() =>
                  setActions((a) => {
                    const n = { ...a };
                    preview
                      .filter((r) => r._isNew)
                      .forEach((r) => {
                        n[r._idx] = 'SKIP';
                      });
                    return n;
                  })
                }
                color={T.red}
              />
            </div>
          </Section>

          <div style={{ display: 'flex', gap: '10px', marginTop: '18px' }}>
            <ActionBtn
              label="← Torna alla mappatura"
              onClick={() => setStep('map')}
            />
            <ActionBtn
              label={`✓ Esegui import (${
                preview.filter((r) => actions[r._idx] !== 'SKIP').length
              } prodotti)`}
              onClick={executeImport}
              primary
            />
          </div>
        </div>
      )}

      {/* ── STEP 4: DONE ── */}
      {step === 'done' && importLogs.length > 0 && (
        <div>
          <Section title="✓ Import completato" accent={T.green} mb="20px">
            <div
              style={{
                padding: '20px',
                background: `${T.green}11`,
                border: `1px solid ${T.green}33`,
                borderRadius: '8px',
                marginBottom: '16px',
              }}
            >
              <div
                style={{
                  fontSize: '16px',
                  fontWeight: 'bold',
                  color: T.green,
                  marginBottom: '8px',
                }}
              >
                Import eseguito con successo
              </div>
              <div
                style={{ fontSize: '13px', color: T.muted, lineHeight: '1.8' }}
              >
                📄 File:{' '}
                <strong style={{ color: T.text }}>
                  {importLogs[0]?.fileName}
                </strong>
                <br />✦ Nuovi prodotti:{' '}
                <strong style={{ color: T.gold }}>
                  {importLogs[0]?.newCount}
                </strong>
                <br />↻ Aggiornati:{' '}
                <strong style={{ color: T.blue }}>
                  {importLogs[0]?.updateCount}
                </strong>
                <br />⊘ Saltati:{' '}
                <strong style={{ color: T.muted }}>
                  {importLogs[0]?.skippedCount}
                </strong>
              </div>
            </div>
            <div style={{ display: 'flex', gap: '10px' }}>
              <ActionBtn label="⇪ Nuovo import" onClick={reset} primary />
              <ActionBtn label="◈ Vai ad Anagrafica" onClick={() => {}} />
            </div>
          </Section>
        </div>
      )}

      {/* Import history */}
      {importLogs.length > 0 && step !== 'done' && (
        <Section title="Storico import" mt="24px">
          <table style={{ width: '100%', borderCollapse: 'collapse' }}>
            <THead
              cols={[
                'File',
                'Data import',
                'Righe totali',
                'Nuovi',
                'Aggiornati',
                'Saltati',
              ]}
            />
            <tbody>
              {importLogs.map((log) => (
                <tr
                  key={log.id}
                  style={{ borderBottom: `1px solid ${T.border}` }}
                >
                  <TD mono>{log.fileName}</TD>
                  <TD mono>
                    {new Date(log.importedAt).toLocaleDateString('it-IT', {
                      day: '2-digit',
                      month: '2-digit',
                      year: 'numeric',
                      hour: '2-digit',
                      minute: '2-digit',
                    })}
                  </TD>
                  <TD mono>{log.totalRows}</TD>
                  <TD>
                    <Chip label={String(log.newCount)} color={T.gold} />
                  </TD>
                  <TD>
                    <Chip label={String(log.updateCount)} color={T.blue} />
                  </TD>
                  <TD>
                    <span style={{ fontSize: '12px', color: T.dim }}>
                      {log.skippedCount}
                    </span>
                  </TD>
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
// DASHBOARD
// ─────────────────────────────────────────────────────────────────────────────
function Dashboard({ costRows, branch, month, branchCfg, setPage }) {
  const flagged = costRows.filter((r) => r.flagged);
  const newItems = costRows.filter((r) => r.isNew && r.cost);
  const missing = costRows.filter((r) => !r.cost);

  return (
    <div>
      <PageHeader
        title={`Dashboard — ${branchCfg.flag} ${branchCfg.label}`}
        sub={`Riepilogo Standard Cost · ${month}`}
      />
      <div
        style={{
          display: 'grid',
          gridTemplateColumns: 'repeat(4,1fr)',
          gap: '14px',
          marginBottom: '28px',
        }}
      >
        <KPI
          label="Prodotti attivi"
          value={costRows.length}
          color={T.blue}
          icon="◈"
        />
        <KPI
          label="Variazioni ≥ ±3%"
          value={flagged.length}
          color={flagged.length ? T.red : T.green}
          icon="◉"
        />
        <KPI
          label="Nuovi prodotti"
          value={newItems.length}
          color={T.gold}
          icon="+"
        />
        <KPI
          label="Senza prezzo/log"
          value={missing.length}
          color={missing.length ? T.orange : T.green}
          icon="◌"
        />
      </div>

      {flagged.length > 0 && (
        <Section
          title="⚠️ Variazioni significative (≥ ±3%)"
          accent={T.red}
          mb="20px"
        >
          <table style={{ width: '100%', borderCollapse: 'collapse' }}>
            <THead
              cols={[
                'Codice',
                'Descrizione',
                'Prev. SC2 €',
                'Nuovo SC2 €',
                'Δ %',
                'Trend',
              ]}
            />
            <tbody>
              {flagged.map((r) => (
                <tr
                  key={r.id}
                  style={{ borderBottom: `1px solid ${T.border}` }}
                >
                  <TD mono>{r.code}</TD>
                  <TD>{r.description}</TD>
                  <TD mono>{FMT(r.prevCost?.step2Eur)}</TD>
                  <TD mono bold>
                    {FMT(r.cost?.step2Eur)}
                  </TD>
                  <TD>
                    <DeltaBadge delta={r.delta} />
                  </TD>
                  <TD>{r.delta > 0 ? '📈' : '📉'}</TD>
                </tr>
              ))}
            </tbody>
          </table>
        </Section>
      )}

      {newItems.length > 0 && (
        <Section title="✦ Nuovi prodotti questo mese" accent={T.gold} mb="20px">
          <table style={{ width: '100%', borderCollapse: 'collapse' }}>
            <THead
              cols={[
                'Codice',
                'Descrizione',
                'Cat.',
                'Temp.',
                'SC2 €',
                `SC2 ${branchCfg.currency}`,
              ]}
            />
            <tbody>
              {newItems.map((r) => (
                <tr
                  key={r.id}
                  style={{ borderBottom: `1px solid ${T.border}` }}
                >
                  <TD mono>{r.code}</TD>
                  <TD>{r.description}</TD>
                  <TD>
                    <Chip label={r.category} />
                  </TD>
                  <TD>
                    <TempChip t={r.temperature} />
                  </TD>
                  <TD mono bold>
                    {FMT(r.cost?.step2Eur)}
                  </TD>
                  <TD mono>{FMT(r.cost?.step2Hkd, 2)}</TD>
                </tr>
              ))}
            </tbody>
          </table>
        </Section>
      )}

      <div style={{ display: 'flex', gap: '12px', marginTop: '8px' }}>
        <ActionBtn
          label="📊 Vedi tutti i costi"
          onClick={() => setPage('costs')}
          primary
        />
        <ActionBtn
          label="✉ Genera mail mensile"
          onClick={() => setPage('mail')}
        />
        <ActionBtn label="⇪ Import da BC" onClick={() => setPage('import')} />
        <ActionBtn
          label="💶 Aggiorna listini"
          onClick={() => setPage('prices')}
        />
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// PRODUCTS (Anagrafica)
// ─────────────────────────────────────────────────────────────────────────────
function Products({ products, setProducts, showToast }) {
  const EMPTY_P = {
    id: '',
    code: '',
    description: '',
    category: 'FOOD',
    uom: 'PCS',
    qtyPerBox: '',
    boxPerPallet: '',
    kgPerBox: '',
    temperature: 'DRY',
    active: true,
  };
  const [form, setForm] = useState(EMPTY_P);
  const [editId, setEditId] = useState(null);
  const [search, setSearch] = useState('');
  const setF = (k, v) => setForm((f) => ({ ...f, [k]: v }));

  const handleSave = () => {
    if (!form.code || !form.description) {
      showToast('Codice e descrizione obbligatori', T.red);
      return;
    }
    if (editId) {
      setProducts((ps) =>
        ps.map((p) => (p.id === editId ? { ...form, id: editId } : p))
      );
      showToast('Prodotto aggiornato ✓');
    } else {
      setProducts((ps) => [...ps, { ...form, id: 'P' + Date.now() }]);
      showToast('Prodotto aggiunto ✓');
    }
    setForm(EMPTY_P);
    setEditId(null);
  };

  const filtered = products.filter(
    (p) =>
      p.description.toLowerCase().includes(search.toLowerCase()) ||
      p.code.toLowerCase().includes(search.toLowerCase())
  );

  return (
    <div>
      <PageHeader
        title="◈ Anagrafica Prodotti"
        sub="Master data condiviso tra tutte le filiali"
      />
      <Section
        title={editId ? 'Modifica prodotto' : 'Nuovo prodotto'}
        mb="24px"
      >
        <div
          style={{
            display: 'grid',
            gridTemplateColumns: 'repeat(3,1fr)',
            gap: '12px',
          }}
        >
          <FormField
            label="Codice"
            value={form.code}
            onChange={(v) => setF('code', v)}
            placeholder="es. BAL-001"
          />
          <FormField
            label="Descrizione"
            value={form.description}
            onChange={(v) => setF('description', v)}
            placeholder="Nome prodotto"
            span={2}
          />
          <SelectField
            label="Categoria"
            value={form.category}
            onChange={(v) => setF('category', v)}
            opts={[
              ['FOOD', 'Food'],
              ['WINE', 'Beverage'],
              ['MEAT', 'Meat/Fish'],
            ]}
          />
          <SelectField
            label="UOM"
            value={form.uom}
            onChange={(v) => setF('uom', v)}
            opts={[
              ['PCS', 'PCS — Pezzo'],
              ['BOX', 'BOX — Cartone'],
              ['KG', 'KG — Chilogrammo'],
            ]}
          />
          <SelectField
            label="Temperatura"
            value={form.temperature}
            onChange={(v) => setF('temperature', v)}
            opts={[
              ['DRY', 'Dry'],
              ['FRESH', 'Fresh'],
              ['FROZEN', 'Frozen'],
            ]}
          />
          <FormField
            label={form.uom === 'KG' ? 'Kg per cartone' : 'Pz per cartone'}
            value={form.qtyPerBox}
            onChange={(v) => setF('qtyPerBox', v)}
            type="number"
          />
          <FormField
            label="Cartoni per pallet"
            value={form.boxPerPallet}
            onChange={(v) => setF('boxPerPallet', v)}
            type="number"
          />
          {form.uom === 'KG' && (
            <FormField
              label="Kg netti per cartone"
              value={form.kgPerBox}
              onChange={(v) => setF('kgPerBox', v)}
              type="number"
            />
          )}
        </div>
        <div style={{ display: 'flex', gap: '10px', marginTop: '14px' }}>
          <ActionBtn
            label={editId ? 'Salva modifiche' : 'Aggiungi prodotto'}
            onClick={handleSave}
            primary
          />
          {editId && (
            <ActionBtn
              label="Annulla"
              onClick={() => {
                setForm(EMPTY_P);
                setEditId(null);
              }}
            />
          )}
        </div>
      </Section>
      <div style={{ marginBottom: '12px' }}>
        <input
          placeholder="🔍  Cerca per codice o descrizione…"
          value={search}
          onChange={(e) => setSearch(e.target.value)}
          style={inputStyle()}
        />
      </div>
      <Section title={`${filtered.length} prodotti`}>
        <table style={{ width: '100%', borderCollapse: 'collapse' }}>
          <THead
            cols={[
              'Codice',
              'Descrizione',
              'Cat.',
              'UOM',
              'Temp.',
              'Qty/Box',
              'Box/Plt',
              'Stato',
              'Azioni',
            ]}
          />
          <tbody>
            {filtered.map((p) => (
              <tr
                key={p.id}
                style={{
                  borderBottom: `1px solid ${T.border}`,
                  opacity: p.active ? 1 : 0.45,
                }}
              >
                <TD mono>{p.code}</TD>
                <TD>{p.description}</TD>
                <TD>
                  <Chip label={p.category} />
                </TD>
                <TD>
                  <Chip label={p.uom} color={T.blue} />
                </TD>
                <TD>
                  <TempChip t={p.temperature} />
                </TD>
                <TD mono>{p.qtyPerBox}</TD>
                <TD mono>{p.boxPerPallet}</TD>
                <TD>
                  <Chip
                    label={p.active ? 'Attivo' : 'Sospeso'}
                    color={p.active ? T.green : T.red}
                  />
                </TD>
                <TD>
                  <div style={{ display: 'flex', gap: '6px' }}>
                    <MiniBtn
                      label="✎"
                      onClick={() => {
                        setForm(p);
                        setEditId(p.id);
                      }}
                    />
                    <MiniBtn
                      label={p.active ? '⊘' : '✓'}
                      onClick={() =>
                        setProducts((ps) =>
                          ps.map((pp) =>
                            pp.id === p.id ? { ...pp, active: !pp.active } : pp
                          )
                        )
                      }
                      color={p.active ? T.red : T.green}
                    />
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
function Logistics({ logistics, setLogistics, products, branch, showToast }) {
  const [editId, setEditId] = useState(null);
  const [form, setForm] = useState({});
  const setF = (k, v) => setForm((f) => ({ ...f, [k]: v }));
  const branchLog = logistics.filter((l) => l.branch === branch);
  const coveredIds = branchLog.map((l) => l.productId);
  const uncovered = products.filter(
    (p) => p.active && !coveredIds.includes(p.id)
  );

  const startEdit = (log) => {
    setEditId(log.productId);
    setForm({ ...log });
  };
  const startNew = (prod) => {
    setEditId('NEW_' + prod.id);
    setForm({
      productId: prod.id,
      branch,
      area: 'NORD',
      ubicazione: 'MTO',
      pltPerContainer: 23,
      hasCert: false,
      hasAlcTax: false,
      alcTax: 0,
    });
  };
  const handleSave = () => {
    if (editId?.startsWith('NEW_')) setLogistics((ls) => [...ls, { ...form }]);
    else
      setLogistics((ls) =>
        ls.map((l) =>
          l.productId === editId && l.branch === branch ? form : l
        )
      );
    showToast('Parametri logistici salvati ✓');
    setEditId(null);
    setForm({});
  };

  return (
    <div>
      <PageHeader
        title="◎ Parametri Logistici"
        sub={`Configurazione FOR/MTO/MTS — ${branch}`}
      />
      {editId && (
        <Section title="Modifica parametri" mb="22px">
          <div
            style={{
              display: 'grid',
              gridTemplateColumns: 'repeat(3,1fr)',
              gap: '12px',
            }}
          >
            <SelectField
              label="Ubicazione"
              value={form.ubicazione}
              onChange={(v) => setF('ubicazione', v)}
              opts={[
                ['MTO', 'MTO — Cross-docking'],
                ['FOR', 'FOR — Franco fornitore'],
                ['MTS', 'MTS — Magazzino Unifreddo'],
              ]}
            />
            <SelectField
              label="Area fornitore"
              value={form.area}
              onChange={(v) => setF('area', v)}
              opts={[
                ['NORD', 'Nord Italia'],
                ['CENTRO', 'Centro Italia'],
                ['SUD', 'Sud Italia'],
              ]}
            />
            <FormField
              label="Pallet per container"
              value={form.pltPerContainer}
              onChange={(v) => setF('pltPerContainer', parseFloat(v) || 23)}
              type="number"
            />
            <div
              style={{
                display: 'flex',
                alignItems: 'center',
                gap: '10px',
                paddingTop: '18px',
              }}
            >
              <CheckBox
                label="Health Certificate"
                checked={form.hasCert}
                onChange={(v) => setF('hasCert', v)}
              />
            </div>
            <div
              style={{
                display: 'flex',
                alignItems: 'center',
                gap: '10px',
                paddingTop: '18px',
              }}
            >
              <CheckBox
                label="Tassa alcolici"
                checked={form.hasAlcTax}
                onChange={(v) => setF('hasAlcTax', v)}
              />
            </div>
            {form.hasAlcTax && (
              <FormField
                label="Importo tassa alcolici (€/u)"
                value={form.alcTax}
                onChange={(v) => setF('alcTax', parseFloat(v) || 0)}
                type="number"
              />
            )}
          </div>
          <div style={{ display: 'flex', gap: '10px', marginTop: '14px' }}>
            <ActionBtn label="Salva" onClick={handleSave} primary />
            <ActionBtn
              label="Annulla"
              onClick={() => {
                setEditId(null);
                setForm({});
              }}
            />
          </div>
        </Section>
      )}
      {uncovered.length > 0 && (
        <Section
          title={`${uncovered.length} prodotti senza configurazione logistica`}
          accent={T.orange}
          mb="20px"
        >
          <table style={{ width: '100%', borderCollapse: 'collapse' }}>
            <THead cols={['Codice', 'Descrizione', 'Temp.', 'Azione']} />
            <tbody>
              {uncovered.map((p) => (
                <tr
                  key={p.id}
                  style={{ borderBottom: `1px solid ${T.border}` }}
                >
                  <TD mono>{p.code}</TD>
                  <TD>{p.description}</TD>
                  <TD>
                    <TempChip t={p.temperature} />
                  </TD>
                  <TD>
                    <MiniBtn
                      label="+ Configura"
                      onClick={() => startNew(p)}
                      color={T.gold}
                    />
                  </TD>
                </tr>
              ))}
            </tbody>
          </table>
        </Section>
      )}
      <Section title="Configurazioni attive">
        <table style={{ width: '100%', borderCollapse: 'collapse' }}>
          <THead
            cols={[
              'Codice',
              'Descrizione',
              'Ubicazione',
              'Area',
              'Plt/Cont',
              'HC',
              'Alc.Tax',
              'Modifica',
            ]}
          />
          <tbody>
            {branchLog.map((l) => {
              const p = products.find((pp) => pp.id === l.productId);
              if (!p) return null;
              return (
                <tr
                  key={l.productId}
                  style={{ borderBottom: `1px solid ${T.border}` }}
                >
                  <TD mono>{p.code}</TD>
                  <TD>{p.description}</TD>
                  <TD>
                    <UbicChip u={l.ubicazione} />
                  </TD>
                  <TD>
                    <Chip label={l.area} color={T.muted} />
                  </TD>
                  <TD mono>{l.pltPerContainer}</TD>
                  <TD>{l.hasCert ? '✓' : '—'}</TD>
                  <TD>{l.hasAlcTax ? `${l.alcTax}€` : '—'}</TD>
                  <TD>
                    <MiniBtn label="✎" onClick={() => startEdit(l)} />
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

// ─────────────────────────────────────────────────────────────────────────────
// PRICES
// ─────────────────────────────────────────────────────────────────────────────
function Prices({ prices, setPrices, products, branch, month, showToast }) {
  const [editingId, setEditingId] = useState(null);
  const [tempVal, setTempVal] = useState('');
  const monthPrices = prices.filter(
    (p) => p.branch === branch && p.month === month
  );
  const prevMonth =
    month.slice(0, 4) +
    '-' +
    String(parseInt(month.slice(5)) - 1).padStart(2, '0');
  const prevPrices = prices.filter(
    (p) => p.branch === branch && p.month === prevMonth
  );

  const handleSave = (productId) => {
    const v = parseFloat(tempVal);
    if (!v || v <= 0) {
      showToast('Prezzo non valido', T.red);
      return;
    }
    const exists = prices.find(
      (p) =>
        p.productId === productId && p.branch === branch && p.month === month
    );
    if (exists)
      setPrices((ps) =>
        ps.map((p) =>
          p.productId === productId && p.branch === branch && p.month === month
            ? { ...p, priceEur: v }
            : p
        )
      );
    else setPrices((ps) => [...ps, { productId, branch, month, priceEur: v }]);
    showToast('Prezzo salvato ✓');
    setEditingId(null);
    setTempVal('');
  };

  return (
    <div>
      <PageHeader
        title="◉ Listini — Prezzi di Acquisto"
        sub={`Prezzi in EUR per ${branch} · ${month}`}
      />
      <Section title="Inserimento prezzi mensili">
        <div style={{ fontSize: '12px', color: T.muted, marginBottom: '14px' }}>
          Clicca su un prezzo per modificarlo.
        </div>
        <table style={{ width: '100%', borderCollapse: 'collapse' }}>
          <THead
            cols={[
              'Codice',
              'Descrizione',
              'UOM',
              'Mese prec. (€)',
              'Mese corrente (€)',
              'Δ %',
            ]}
          />
          <tbody>
            {products
              .filter((p) => p.active)
              .map((prod) => {
                const cur = monthPrices.find((p) => p.productId === prod.id);
                const prev = prevPrices.find((p) => p.productId === prod.id);
                const delta =
                  cur && prev ? PCT(cur.priceEur, prev.priceEur) : null;
                const isEditing = editingId === prod.id;
                return (
                  <tr
                    key={prod.id}
                    style={{
                      borderBottom: `1px solid ${T.border}`,
                      background: !cur ? 'rgba(196,122,59,0.07)' : '',
                    }}
                  >
                    <TD mono>{prod.code}</TD>
                    <TD>{prod.description}</TD>
                    <TD>
                      <Chip label={prod.uom} color={T.blue} />
                    </TD>
                    <TD mono>
                      {prev ? (
                        `€ ${prev.priceEur.toFixed(4)}`
                      ) : (
                        <span style={{ color: T.dim }}>—</span>
                      )}
                    </TD>
                    <TD>
                      {isEditing ? (
                        <div
                          style={{
                            display: 'flex',
                            gap: '6px',
                            alignItems: 'center',
                          }}
                        >
                          <input
                            autoFocus
                            type="number"
                            value={tempVal}
                            onChange={(e) => setTempVal(e.target.value)}
                            style={{
                              ...inputStyle(),
                              width: '110px',
                              padding: '4px 8px',
                              fontSize: '13px',
                            }}
                          />
                          <MiniBtn
                            label="✓"
                            onClick={() => handleSave(prod.id)}
                            color={T.green}
                          />
                          <MiniBtn
                            label="✕"
                            onClick={() => {
                              setEditingId(null);
                              setTempVal('');
                            }}
                            color={T.red}
                          />
                        </div>
                      ) : (
                        <span
                          onClick={() => {
                            setEditingId(prod.id);
                            setTempVal(cur?.priceEur || '');
                          }}
                          style={{
                            cursor: 'pointer',
                            color: cur ? T.text : T.orange,
                            textDecoration: 'underline dotted',
                            fontFamily: 'monospace',
                          }}
                        >
                          {cur ? `€ ${cur.priceEur.toFixed(4)}` : '— inserisci'}
                        </span>
                      )}
                    </TD>
                    <TD>
                      {delta !== null ? (
                        <DeltaBadge delta={delta} small />
                      ) : (
                        <span style={{ color: T.dim }}>—</span>
                      )}
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

// ─────────────────────────────────────────────────────────────────────────────
// FX RATES
// ─────────────────────────────────────────────────────────────────────────────
function FxRates({ fx, setFx, branch, month, showToast }) {
  const [val, setVal] = useState('');
  const cur = fx.find((f) => f.branch === branch && f.month === month);
  const cfg = BRANCH_CFG[branch];
  const handleSave = () => {
    const v = parseFloat(val);
    if (!v || v <= 0) {
      showToast('Tasso non valido', T.red);
      return;
    }
    if (cur)
      setFx((fs) =>
        fs.map((f) =>
          f.branch === branch && f.month === month ? { ...f, rate: v } : f
        )
      );
    else setFx((fs) => [...fs, { branch, month, rate: v }]);
    showToast('Tasso aggiornato ✓');
    setVal('');
  };
  const history = fx
    .filter((f) => f.branch === branch)
    .sort((a, b) => b.month.localeCompare(a.month));
  return (
    <div>
      <PageHeader
        title="◌ Tassi di Cambio"
        sub={`EUR → ${cfg.currency} · ${branch}`}
      />
      <Section title={`Tasso corrente — ${month}`} mb="22px">
        <div style={{ display: 'flex', alignItems: 'flex-end', gap: '16px' }}>
          <div>
            <div
              style={{ fontSize: '11px', color: T.muted, marginBottom: '6px' }}
            >
              Tasso attuale EUR/{cfg.currency}
            </div>
            <div
              style={{ fontSize: '32px', fontWeight: 'bold', color: cfg.color }}
            >
              {cur ? (
                cur.rate.toFixed(4)
              ) : (
                <span style={{ color: T.dim }}>—</span>
              )}
            </div>
            <div style={{ fontSize: '11px', color: T.dim, marginTop: '3px' }}>
              Default modello: {cfg.defaultRate.toFixed(4)}
            </div>
          </div>
          <div
            style={{
              display: 'flex',
              gap: '8px',
              alignItems: 'center',
              paddingBottom: '4px',
            }}
          >
            <input
              type="number"
              placeholder={`es. ${cfg.defaultRate}`}
              value={val}
              onChange={(e) => setVal(e.target.value)}
              style={{ ...inputStyle(), width: '140px' }}
            />
            <ActionBtn label="Aggiorna" onClick={handleSave} primary />
          </div>
        </div>
      </Section>
      <Section title={`Storico EUR/${cfg.currency}`}>
        <table style={{ width: '100%', borderCollapse: 'collapse' }}>
          <THead
            cols={['Mese', 'Tasso EUR/' + cfg.currency, 'Δ vs mese prec.']}
          />
          <tbody>
            {history.map((f, i) => {
              const prev = history[i + 1];
              const delta = prev ? PCT(f.rate, prev.rate) : null;
              return (
                <tr
                  key={f.month}
                  style={{ borderBottom: `1px solid ${T.border}` }}
                >
                  <TD mono>{f.month}</TD>
                  <TD mono bold>
                    {f.rate.toFixed(4)}
                  </TD>
                  <TD>
                    {delta !== null ? (
                      <DeltaBadge delta={delta} small />
                    ) : (
                      <span style={{ color: T.dim }}>—</span>
                    )}
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

// ─────────────────────────────────────────────────────────────────────────────
// COST TABLE
// ─────────────────────────────────────────────────────────────────────────────
function CostTable({ costRows, branch, month, branchCfg }) {
  const [filter, setFilter] = useState('ALL');
  const [expand, setExpand] = useState(null);
  const filtered = costRows.filter((r) => {
    if (filter === 'FLAGGED') return r.flagged;
    if (filter === 'NEW') return r.isNew && r.cost;
    if (filter === 'MISSING') return !r.cost;
    return true;
  });
  return (
    <div>
      <PageHeader
        title="◆ Standard Cost"
        sub={`Calcolo completo · ${branch} · ${month}`}
      />
      <div style={{ display: 'flex', gap: '8px', marginBottom: '20px' }}>
        {[
          ['ALL', 'Tutti'],
          ['FLAGGED', 'Variazioni ≥ ±3%'],
          ['NEW', 'Nuovi'],
          ['MISSING', 'Mancanti'],
        ].map(([k, l]) => (
          <button
            key={k}
            onClick={() => setFilter(k)}
            style={{
              padding: '6px 14px',
              background: filter === k ? T.goldDim : 'rgba(255,255,255,0.04)',
              border: `1px solid ${filter === k ? T.gold : T.border}`,
              borderRadius: '6px',
              color: filter === k ? T.gold : T.muted,
              cursor: 'pointer',
              fontFamily: 'inherit',
              fontSize: '12px',
            }}
          >
            {l}
          </button>
        ))}
      </div>
      <Section title={`${filtered.length} prodotti`}>
        <table style={{ width: '100%', borderCollapse: 'collapse' }}>
          <THead
            cols={[
              'Codice',
              'Descrizione',
              'Cat.',
              'Temp.',
              'Ubic.',
              'Prezzo €',
              'SC2 €',
              `SC2 ${branchCfg.currency}`,
              'Δ %',
              'Det.',
            ]}
          />
          <tbody>
            {filtered.map((r) => (
              <>
                <tr
                  key={r.id}
                  style={{
                    borderBottom: `1px solid ${T.border}`,
                    background: r.flagged
                      ? 'rgba(181,83,74,0.05)'
                      : r.isNew
                      ? 'rgba(201,168,76,0.05)'
                      : '',
                  }}
                >
                  <TD mono>{r.code}</TD>
                  <TD>{r.description}</TD>
                  <TD>
                    <Chip label={r.category} />
                  </TD>
                  <TD>
                    <TempChip t={r.temperature} />
                  </TD>
                  <TD>
                    <UbicChip u={r.log?.ubicazione} />
                  </TD>
                  <TD mono>
                    {r.cost ? (
                      FMT(r.priceEur)
                    ) : (
                      <span style={{ color: T.dim }}>—</span>
                    )}
                  </TD>
                  <TD mono bold>
                    {r.cost ? (
                      FMT(r.cost.step2Eur)
                    ) : (
                      <span style={{ color: T.dim }}>—</span>
                    )}
                  </TD>
                  <TD mono>
                    {r.cost ? (
                      FMT(r.cost.step2Hkd, 2)
                    ) : (
                      <span style={{ color: T.dim }}>—</span>
                    )}
                  </TD>
                  <TD>
                    {r.delta !== null ? (
                      <DeltaBadge delta={r.delta} />
                    ) : r.isNew ? (
                      <Chip label="NUOVO" color={T.gold} />
                    ) : (
                      <span style={{ color: T.dim }}>—</span>
                    )}
                  </TD>
                  <TD>
                    {r.cost && (
                      <MiniBtn
                        label={expand === r.id ? '▲' : '▼'}
                        onClick={() => setExpand(expand === r.id ? null : r.id)}
                      />
                    )}
                  </TD>
                </tr>
                {expand === r.id && r.cost && (
                  <tr
                    key={r.id + '_exp'}
                    style={{ background: 'rgba(255,255,255,0.02)' }}
                  >
                    <td colSpan={10} style={{ padding: '14px 20px' }}>
                      <CostBreakdown
                        cost={r.cost}
                        currency={branchCfg.currency}
                      />
                    </td>
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

function CostBreakdown({ cost }) {
  const rows = [
    ['Prezzo acquisto', cost.priceEur],
    ['FOB + Nolo', cost.fob],
    ['Local Import Charges', cost.lic],
    ['VGM', cost.vgm],
    cost.hc > 0 && ['Health Certificate', cost.hc],
    ['Costo pallet', cost.plt],
    cost.alc > 0 && ['Tassa alcolici', cost.alc],
    ['→ Step 1', cost.step1Eur, true],
    cost.whDetail.type === 'MTO' && ['MTO (cross-docking)', cost.wh],
    cost.whDetail.type === 'MTS' && ['MTS — Deposito', cost.whDetail.dep],
    cost.whDetail.type === 'MTS' && ['MTS — Inbound', cost.whDetail.inbound],
    cost.whDetail.type === 'MTS' && ['MTS — Picking', cost.whDetail.picking],
    cost.whDetail.type === 'FOR' && ['FOR — nessun handling', 0],
    ['→ Step 2 (SC finale)', cost.step2Eur, true],
  ].filter(Boolean);
  return (
    <div
      style={{
        display: 'grid',
        gridTemplateColumns: '1fr 1fr',
        gap: '6px 20px',
      }}
    >
      {rows.map(([label, eur, bold], i) => (
        <div
          key={i}
          style={{
            display: 'flex',
            justifyContent: 'space-between',
            fontSize: '12px',
            fontWeight: bold ? 'bold' : 'normal',
            color: bold ? T.gold : T.muted,
            borderTop: bold ? `1px solid ${T.border}` : 'none',
            paddingTop: bold ? '4px' : 0,
          }}
        >
          <span>{label}</span>
          <span style={{ fontFamily: 'monospace' }}>€ {eur.toFixed(4)}</span>
        </div>
      ))}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// MAIL GENERATOR
// ─────────────────────────────────────────────────────────────────────────────
function MailGen({
  costRows,
  branch,
  month,
  branchCfg,
  sentMails,
  setSentMails,
  showToast,
}) {
  const [sent, setSent] = useState(false);
  const [viewing, setViewing] = useState(null);
  const flagged = costRows.filter((r) => r.flagged && r.cost);
  const ups = flagged.filter((r) => r.delta > 0);
  const downs = flagged.filter((r) => r.delta < 0);
  const newItems = costRows.filter((r) => r.isNew && r.cost);
  const monthLabel = new Date(month + '-01').toLocaleDateString('it-IT', {
    month: 'long',
    year: 'numeric',
  });

  const handleSend = () => {
    setSentMails((ms) => [
      {
        id: Date.now(),
        branch,
        month,
        sentAt: new Date().toISOString(),
        newCount: newItems.length,
        changedCount: flagged.length,
      },
      ...ms,
    ]);
    setSent(true);
    showToast('Mail registrata nello storico ✓', T.gold);
  };

  if (viewing !== null) {
    const m = sentMails[viewing];
    return (
      <div>
        <PageHeader
          title="✉ Mail inviata — storico"
          sub={`${m.branch} · ${m.month}`}
        />
        <ActionBtn label="← Torna" onClick={() => setViewing(null)} />
        <div
          style={{
            marginTop: '16px',
            padding: '16px',
            background: T.card,
            borderRadius: '8px',
            fontSize: '12px',
            color: T.muted,
          }}
        >
          Inviata il {new Date(m.sentAt).toLocaleString('it-IT')} · {m.newCount}{' '}
          nuovi · {m.changedCount} variazioni
        </div>
      </div>
    );
  }

  return (
    <div>
      <PageHeader
        title="✉ Mail Mensile"
        sub={`Standard Cost Update · ${branchCfg.flag} ${branchCfg.label} · ${monthLabel}`}
      />
      <Section title="Anteprima comunicazione" mb="24px">
        <div
          style={{
            background: '#fff',
            borderRadius: '10px',
            padding: '40px 48px',
            color: '#1A1A1A',
            fontFamily: "'Helvetica Neue',Arial,sans-serif",
            maxWidth: '700px',
            boxShadow: '0 4px 32px rgba(0,0,0,0.4)',
          }}
        >
          <div
            style={{
              borderBottom: '3px solid #1A2B3C',
              paddingBottom: '20px',
              marginBottom: '28px',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'space-between',
            }}
          >
            <div>
              <div
                style={{
                  fontSize: '11px',
                  letterSpacing: '3px',
                  color: '#C9A84C',
                  textTransform: 'uppercase',
                  marginBottom: '4px',
                }}
              >
                Inalca Food & Beverage
              </div>
              <div
                style={{
                  fontSize: '22px',
                  fontWeight: 'bold',
                  color: '#1A2B3C',
                }}
              >
                Standard Cost Update
              </div>
              <div
                style={{ fontSize: '13px', color: '#666', marginTop: '3px' }}
              >
                {branchCfg.flag} {branchCfg.label} · {monthLabel}
              </div>
            </div>
            <div style={{ fontSize: '36px' }}>{branchCfg.flag}</div>
          </div>
          <p
            style={{
              fontSize: '13px',
              color: '#333',
              lineHeight: '1.7',
              marginBottom: '24px',
            }}
          >
            Si comunica l'aggiornamento mensile degli{' '}
            <strong>Standard Cost</strong> per la filiale di{' '}
            <strong>{branchCfg.label}</strong> con decorrenza{' '}
            <strong>{monthLabel}</strong>.
          </p>
          {newItems.length > 0 && (
            <div style={{ marginBottom: '28px' }}>
              <div
                style={{
                  fontSize: '12px',
                  fontWeight: 'bold',
                  color: '#C9A84C',
                  textTransform: 'uppercase',
                  letterSpacing: '2px',
                  marginBottom: '12px',
                }}
              >
                ✦ Nuovi Prodotti ({newItems.length})
              </div>
              <MailTable
                rows={newItems}
                currency={branchCfg.currency}
                showPrev={false}
              />
            </div>
          )}
          {ups.length > 0 && (
            <div style={{ marginBottom: '28px' }}>
              <div
                style={{
                  fontSize: '12px',
                  fontWeight: 'bold',
                  color: '#B5534A',
                  textTransform: 'uppercase',
                  letterSpacing: '2px',
                  marginBottom: '12px',
                }}
              >
                ↑ Aumenti ≥ +3% ({ups.length})
              </div>
              <MailTable
                rows={ups}
                currency={branchCfg.currency}
                showPrev={true}
              />
            </div>
          )}
          {downs.length > 0 && (
            <div style={{ marginBottom: '28px' }}>
              <div
                style={{
                  fontSize: '12px',
                  fontWeight: 'bold',
                  color: '#2D7A50',
                  textTransform: 'uppercase',
                  letterSpacing: '2px',
                  marginBottom: '12px',
                }}
              >
                ↓ Riduzioni ≥ -3% ({downs.length})
              </div>
              <MailTable
                rows={downs}
                currency={branchCfg.currency}
                showPrev={true}
              />
            </div>
          )}
          {newItems.length === 0 && flagged.length === 0 && (
            <div
              style={{
                padding: '20px',
                background: '#F5F5F0',
                borderRadius: '8px',
                textAlign: 'center',
                color: '#666',
                fontSize: '13px',
              }}
            >
              Nessuna variazione significativa questo mese.
            </div>
          )}
          <div
            style={{
              borderTop: '1px solid #E8E8E0',
              paddingTop: '18px',
              marginTop: '28px',
              fontSize: '11px',
              color: '#999',
            }}
          >
            Generato da{' '}
            <strong style={{ color: '#1A2B3C' }}>
              IFB Cost Intelligence Platform
            </strong>{' '}
            · {new Date().toLocaleDateString('it-IT')}
          </div>
        </div>
      </Section>
      <div style={{ display: 'flex', gap: '12px' }}>
        <ActionBtn
          label={sent ? '✓ Registrata' : '✉ Registra invio e salva in storico'}
          onClick={handleSend}
          primary
          disabled={sent}
        />
      </div>
      {sentMails.length > 0 && (
        <Section title="Storico mail inviate" mt="28px">
          <table style={{ width: '100%', borderCollapse: 'collapse' }}>
            <THead
              cols={[
                'Filiale',
                'Mese',
                'Data invio',
                'Nuovi',
                'Variazioni',
                'Azione',
              ]}
            />
            <tbody>
              {sentMails.map((m, i) => (
                <tr
                  key={m.id}
                  style={{ borderBottom: `1px solid ${T.border}` }}
                >
                  <TD>
                    {BRANCH_CFG[m.branch]?.flag} {m.branch}
                  </TD>
                  <TD mono>{m.month}</TD>
                  <TD mono>{new Date(m.sentAt).toLocaleDateString('it-IT')}</TD>
                  <TD mono>{m.newCount}</TD>
                  <TD mono>{m.changedCount}</TD>
                  <TD>
                    <MiniBtn label="Visualizza" onClick={() => setViewing(i)} />
                  </TD>
                </tr>
              ))}
            </tbody>
          </table>
        </Section>
      )}
    </div>
  );
}

function MailTable({ rows, currency, showPrev }) {
  const cols = showPrev
    ? ['Codice', 'Descrizione', 'SC2 prec. (€)', `SC2 nuovo (€)`, 'Δ %']
    : ['Codice', 'Descrizione', 'Temp.', 'SC2 (€)', `SC2 (${currency})`];
  return (
    <table
      style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px' }}
    >
      <thead>
        <tr style={{ background: '#F5F5F0' }}>
          {cols.map((h) => (
            <th
              key={h}
              style={{
                padding: '8px 10px',
                textAlign: 'left',
                color: '#444',
                fontWeight: 'bold',
                fontSize: '11px',
              }}
            >
              {h}
            </th>
          ))}
        </tr>
      </thead>
      <tbody>
        {rows.map((r, i) => (
          <tr
            key={r.id}
            style={{
              background: i % 2 === 0 ? '#fff' : '#FAFAF8',
              borderBottom: '1px solid #E8E8E0',
            }}
          >
            <td
              style={{
                padding: '8px 10px',
                fontFamily: 'monospace',
                fontSize: '11px',
              }}
            >
              {r.code}
            </td>
            <td style={{ padding: '8px 10px' }}>{r.description}</td>
            {showPrev ? (
              <>
                <td style={{ padding: '8px 10px', fontFamily: 'monospace' }}>
                  € {FMT(r.prevCost?.step2Eur)}
                </td>
                <td
                  style={{
                    padding: '8px 10px',
                    fontFamily: 'monospace',
                    fontWeight: 'bold',
                  }}
                >
                  € {FMT(r.cost?.step2Eur)}
                </td>
                <td style={{ padding: '8px 10px' }}>
                  <span
                    style={{
                      padding: '2px 8px',
                      borderRadius: '4px',
                      fontSize: '11px',
                      fontWeight: 'bold',
                      background: r.delta > 0 ? '#FDECEA' : '#E8F5EE',
                      color: r.delta > 0 ? '#B5534A' : '#2D7A50',
                    }}
                  >
                    {r.delta > 0 ? '+' : ''}
                    {r.delta?.toFixed(1)}%
                  </span>
                </td>
              </>
            ) : (
              <>
                <td style={{ padding: '8px 10px' }}>{r.temperature}</td>
                <td style={{ padding: '8px 10px', fontFamily: 'monospace' }}>
                  € {FMT(r.cost?.step2Eur)}
                </td>
                <td style={{ padding: '8px 10px', fontFamily: 'monospace' }}>
                  {currency} {FMT(r.cost?.step2Hkd, 2)}
                </td>
              </>
            )}
          </tr>
        ))}
      </tbody>
    </table>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// SHARED UI COMPONENTS
// ─────────────────────────────────────────────────────────────────────────────
function PageHeader({ title, sub }) {
  return (
    <div style={{ marginBottom: '24px' }}>
      <h1
        style={{
          fontSize: '22px',
          fontWeight: 'bold',
          margin: 0,
          marginBottom: '4px',
        }}
      >
        {title}
      </h1>
      <div style={{ fontSize: '12px', color: T.muted }}>{sub}</div>
    </div>
  );
}
function Section({ title, accent, children, mb, mt }) {
  return (
    <div style={{ marginBottom: mb || '0', marginTop: mt || '0' }}>
      {title && (
        <div
          style={{
            fontSize: '10px',
            letterSpacing: '2px',
            textTransform: 'uppercase',
            color: accent || T.gold,
            marginBottom: '10px',
            display: 'flex',
            alignItems: 'center',
            gap: '6px',
          }}
        >
          <div
            style={{
              height: '1px',
              width: '16px',
              background: accent || T.gold,
              opacity: 0.5,
            }}
          />
          {title}
        </div>
      )}
      <div
        style={{
          background: T.card,
          border: `1px solid ${T.border}`,
          borderRadius: '10px',
          padding: '18px',
        }}
      >
        {children}
      </div>
    </div>
  );
}
function KPI({ label, value, color, icon }) {
  return (
    <div
      style={{
        background: T.card,
        border: `1px solid ${T.border}`,
        borderRadius: '10px',
        padding: '20px',
      }}
    >
      <div style={{ fontSize: '24px', marginBottom: '4px' }}>{icon}</div>
      <div style={{ fontSize: '28px', fontWeight: 'bold', color }}>{value}</div>
      <div style={{ fontSize: '11px', color: T.muted, marginTop: '2px' }}>
        {label}
      </div>
    </div>
  );
}
function THead({ cols }) {
  return (
    <thead>
      <tr>
        {cols.map((c) => (
          <th
            key={c}
            style={{
              padding: '8px 10px',
              textAlign: 'left',
              fontSize: '10px',
              letterSpacing: '1px',
              textTransform: 'uppercase',
              color: T.dim,
              borderBottom: `1px solid ${T.border}`,
              fontWeight: 'normal',
            }}
          >
            {c}
          </th>
        ))}
      </tr>
    </thead>
  );
}
function TD({ children, mono, bold }) {
  return (
    <td
      style={{
        padding: '9px 10px',
        fontSize: '13px',
        fontFamily: mono ? 'monospace' : 'inherit',
        fontWeight: bold ? 'bold' : 'normal',
        color: T.text,
      }}
    >
      {children}
    </td>
  );
}
function Chip({ label, color }) {
  const c =
    color ||
    (label === 'FOOD'
      ? T.gold
      : label === 'WINE'
      ? T.purple
      : label === 'MEAT'
      ? T.red
      : T.muted);
  return (
    <span
      style={{
        padding: '2px 7px',
        borderRadius: '4px',
        fontSize: '10px',
        background: `${c}22`,
        color: c,
        border: `1px solid ${c}33`,
        letterSpacing: '0.5px',
      }}
    >
      {label}
    </span>
  );
}
function TempChip({ t }) {
  const c = t === 'DRY' ? T.gold : t === 'FRESH' ? T.blue : T.purple;
  const e = t === 'DRY' ? '🌾' : t === 'FRESH' ? '❄️' : '🧊';
  return <Chip label={`${e} ${t}`} color={c} />;
}
function UbicChip({ u }) {
  if (!u) return <span style={{ color: T.dim }}>—</span>;
  const c = u === 'MTO' ? T.blue : u === 'FOR' ? T.green : T.orange;
  return <Chip label={u} color={c} />;
}
function DeltaBadge({ delta, small }) {
  if (delta === null || delta === undefined) return null;
  const up = delta > 0;
  return (
    <span
      style={{
        padding: small ? '1px 6px' : '3px 9px',
        borderRadius: '5px',
        fontSize: small ? '10px' : '12px',
        fontWeight: 'bold',
        background: up ? 'rgba(181,83,74,0.18)' : 'rgba(75,168,122,0.18)',
        color: up ? T.red : T.green,
      }}
    >
      {up ? '+' : ''}
      {delta.toFixed(1)}%
    </span>
  );
}
function ActionBtn({ label, onClick, primary, disabled }) {
  return (
    <button
      onClick={onClick}
      disabled={disabled}
      style={{
        padding: '9px 18px',
        background: disabled
          ? 'rgba(255,255,255,0.04)'
          : primary
          ? T.gold
          : 'rgba(255,255,255,0.06)',
        border: `1px solid ${
          disabled ? T.border : primary ? T.gold : T.borderHi
        }`,
        borderRadius: '7px',
        color: disabled ? T.dim : primary ? T.bg : T.text,
        cursor: disabled ? 'not-allowed' : 'pointer',
        fontFamily: 'inherit',
        fontSize: '13px',
        fontWeight: primary ? 'bold' : 'normal',
        transition: 'all 0.18s',
        opacity: disabled ? 0.6 : 1,
      }}
    >
      {label}
    </button>
  );
}
function MiniBtn({ label, onClick, color }) {
  return (
    <button
      onClick={onClick}
      style={{
        padding: '3px 9px',
        background: 'rgba(255,255,255,0.04)',
        border: `1px solid ${color || T.border}`,
        borderRadius: '5px',
        color: color || T.muted,
        cursor: 'pointer',
        fontFamily: 'inherit',
        fontSize: '11px',
        whiteSpace: 'nowrap',
      }}
    >
      {label}
    </button>
  );
}
function FormField({ label, value, onChange, placeholder, type, span }) {
  return (
    <div style={{ gridColumn: span ? `span ${span}` : undefined }}>
      <label
        style={{
          display: 'block',
          fontSize: '11px',
          color: T.muted,
          marginBottom: '5px',
        }}
      >
        {label}
      </label>
      <input
        type={type || 'text'}
        placeholder={placeholder}
        value={value || ''}
        onChange={(e) => onChange(e.target.value)}
        style={inputStyle()}
      />
    </div>
  );
}
function SelectField({ label, value, onChange, opts }) {
  return (
    <div>
      <label
        style={{
          display: 'block',
          fontSize: '11px',
          color: T.muted,
          marginBottom: '5px',
        }}
      >
        {label}
      </label>
      <select
        value={value}
        onChange={(e) => onChange(e.target.value)}
        style={{ ...inputStyle(), cursor: 'pointer' }}
      >
        {opts.map(([v, l]) => (
          <option key={v} value={v}>
            {l}
          </option>
        ))}
      </select>
    </div>
  );
}
function CheckBox({ label, checked, onChange }) {
  return (
    <div
      onClick={() => onChange(!checked)}
      style={{
        display: 'flex',
        alignItems: 'center',
        gap: '8px',
        cursor: 'pointer',
        userSelect: 'none',
      }}
    >
      <div
        style={{
          width: '16px',
          height: '16px',
          borderRadius: '4px',
          background: checked ? T.gold : 'rgba(255,255,255,0.08)',
          border: `1px solid ${checked ? T.gold : T.border}`,
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          fontSize: '10px',
          color: T.bg,
          flexShrink: 0,
        }}
      >
        {checked ? '✓' : ''}
      </div>
      <span style={{ fontSize: '12px', color: checked ? T.text : T.muted }}>
        {label}
      </span>
    </div>
  );
}
function inputStyle() {
  return {
    width: '100%',
    padding: '8px 11px',
    background: 'rgba(255,255,255,0.05)',
    border: `1px solid ${T.border}`,
    borderRadius: '7px',
    color: T.text,
    fontFamily: 'inherit',
    fontSize: '13px',
    outline: 'none',
    boxSizing: 'border-box',
  };
}
