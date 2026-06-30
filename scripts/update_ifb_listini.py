"""
Aggiorna docs/data/ifb_listini.json con i listini prezzi da BC Italia.
Logica derivata dal confronto con Power BI:
  - FCA/DAP Price = purchase_price / 0.99  (markup 1% su costo acquisto)
  - Carriage      = DAP_sell - FCA_sell
  - Discount      = da listino vendita cliente (solo record non scaduti)
  - Discounted    = Price * (1 - Discount/100)
  - DAP Final     = DAP Discounted
"""
import os, json, requests
from datetime import date, timedelta
from pathlib import Path
from collections import defaultdict

TENANT_ID     = "2acd007b-8d3f-4be0-9681-cf248264a0e2"
CLIENT_ID     = "925de6e4-e71f-4c24-9e0a-f3ae544ae644"
CLIENT_SECRET = os.environ.get("BC_CLIENT_SECRET", "")
BC_ENV        = "Production"
BC_COMPANY    = "Inalca%20Food%20%26%20Beverage%20s.r.l."
BC_ENV_HK     = "Production_HK"
BC_COMPANY_HK = "BRIGHT%20VIEW%20TRADING%20HK%20LIMITED"
OUT_PATH      = Path(__file__).parent.parent / "docs" / "data" / "ifb_listini.json"

TODAY  = date.today().isoformat()
MARKUP = 1 / 0.98  # markup IFB: PowerBI usa dir_unit_cost_conv * Fatt_Conv / 0.98

# Mappa temperatura BC -> chiave interna
TEMP_NORM = {
    "congelato": "FROZEN", "frozen": "FROZEN",
    "fresco": "FRESH", "fresh": "FRESH", "refrigerato": "FRESH", "refrigerated": "FRESH",
    "secco": "DRY", "dry": "DRY", "ambient": "DRY", "ambiente": "DRY",
}

CUSTOMERS = {
    "HK":  "40000854",
    "CAN": "40000175",
    "MAC": "40001358",
}


def get_token():
    r = requests.post(
        f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token",
        data={"grant_type": "client_credentials", "client_id": CLIENT_ID,
              "client_secret": CLIENT_SECRET,
              "scope": "https://api.businesscentral.dynamics.com/.default"})
    r.raise_for_status()
    return r.json()["access_token"]


def bc_fetch_all(token, entity, filt=None):
    base_url = (f"https://api.businesscentral.dynamics.com/v2.0/{TENANT_ID}"
                f"/{BC_ENV}/ODataV4/Company('{BC_COMPANY}')/{entity}")
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    base_params = {"$top": 5000}
    if filt:
        base_params["$filter"] = filt
    results = []
    next_url = None
    skip = 0
    while True:
        if next_url:
            r = requests.get(next_url, headers=headers)
        else:
            r = requests.get(base_url, headers=headers, params={**base_params, "$skip": skip})
        r.raise_for_status()
        data = r.json()
        batch = data.get("value", [])
        results.extend(batch)
        next_url = data.get("@odata.nextLink")
        if next_url:
            skip = 0
        elif len(batch) == 5000:
            # BC non ha restituito nextLink ma potrebbe esserci un batch successivo
            skip += 5000
        else:
            break
    return results


def bc_fetch_hk(token, entity, filt=None):
    """Fetch da BC Brightview HK (Production_HK)."""
    base = (f"https://api.businesscentral.dynamics.com/v2.0/{TENANT_ID}"
            f"/{BC_ENV_HK}/ODataV4/Company('{BC_COMPANY_HK}')/")
    url = base + entity
    params = {"$top": 5000}
    if filt:
        params["$filter"] = filt
    results = []
    while url:
        r = requests.get(url,
                         headers={"Authorization": f"Bearer {token}",
                                  "Accept": "application/json"},
                         params=params if url == base + entity else {})
        r.raise_for_status()
        data = r.json()
        results.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return results


def fetch_price_lines(token, filt):
    return bc_fetch_all(token, "IFB_Price_List_Line", filt=filt)


def classify_ship(code):
    c = (code or "").strip().upper()
    if c in {"MTS", "EXW"}:
        return "MTS"
    if c in {"DAP", "CIF", "DDU", "DDP"}:
        return "DAP"
    return "FCA"


def is_expired(enddate_str):
    ed = (enddate_str or "").split("T")[0]
    return ed not in ("", "0001-01-01") and ed < TODAY

def is_active_date(enddate_str):
    """
    Logica identica a PowerBI:
    endingdate = null  OR  endingdate > oggi  OR  endingdate < 2010-01-01
    (endingdate < 2010 cattura il valore BC "vuoto" = 0001-01-01)
    """
    ed = (enddate_str or "").split("T")[0]
    if not ed or ed < "2010-01-01":   # null o 0001-01-01 -> nessuna scadenza
        return True
    if ed > TODAY:                     # scadenza futura -> ancora valido
        return True
    return False                       # scaduto



def build_item_card_data(token):
    """Fetch dati articolo: fornitore, temperatura, pz/box, box/pallet."""
    print("  Fetch Item Card data (fornitore, temp, uom)...")
    sel = "No,AltICMIFB_Item,AltICMQuantity_x_Packaging,AltICMPackaging_x_Pallet,AltICMProduct_Type,AltICMVendor_Name"
    rows = bc_fetch_all(token, f"Item_Card_Excel?$select={sel}")
    result = {}
    for r in rows:
        ifb_code = str(r.get("AltICMIFB_Item") or "").strip()
        hk_code  = str(r.get("No") or "").strip()
        key = ifb_code or hk_code
        if not key:
            continue
        temp_raw = str(r.get("AltICMProduct_Type") or "").strip().lower()
        result[key] = {
            "vendorName":   str(r.get("AltICMVendor_Name") or "").strip(),
            "temperature":  TEMP_NORM.get(temp_raw, temp_raw.upper()),
            "qtyPerBox":    float(r.get("AltICMQuantity_x_Packaging") or 0),
            "boxPerPallet": float(r.get("AltICMPackaging_x_Pallet") or 0),
        }
    print(f"    {len(result)} articoli con dati card")
    return result


def build_transport_costs(token):
    """Fetch Tabella_Costi_di_Trasporto_Excel -> {(vendor_name, temp): pallet1_cost}"""
    print("  Fetch Tabella Costi Trasporto...")
    rows = bc_fetch_all(token, "Tabella_Costi_di_Trasporto_Excel")
    costs = {}
    for r in rows:
        vendor   = str(r.get("Vendor_Name") or "").strip()
        temp_raw = str(r.get("Temperature") or "").strip().lower()
        temp     = TEMP_NORM.get(temp_raw, temp_raw.upper())
        pallet1  = float(r.get("Pallet1") or 0)
        if vendor and pallet1 > 0:
            costs[(vendor, temp)] = pallet1
    print(f"    {len(costs)} combinazioni (fornitore, temperatura) con Pallet1 > 0")
    return costs


def build_item_prices(rows, pricetype_filter=None):
    """
    Price list lines (vendita o acquisto) per un cliente / fornitore.
    PowerBI usa totlinediscountperc (non linediscount).
    Discount = spurc + ssale: somma sconto acquisto + sconto vendita per cliente HK.
    Restituisce: (item_codes, sale_slots)
      - sale_slots: assetno -> {FCA/MTS/DAP: {unitprice, discount_purch, discount_sale, ...}}
    """
    item_codes = set()
    sale_slots = defaultdict(lambda: {"FCA": {}, "MTS": {}, "DAP": {}})

    for r in rows:
        code = str(r.get("assetno") or "").strip()
        if not code:
            continue
        if pricetype_filter and str(r.get("pricetype") or "").strip().lower() != pricetype_filter:
            continue
        item_codes.add(code)
        ed       = str(r.get("endingdate")   or "")
        sd       = str(r.get("startingdate") or "")
        ship     = classify_ship(r.get("shipmentmethodcode"))
        slot     = sale_slots[code][ship]
        up       = float(r.get("unitprice")          or 0)
        dc       = float(r.get("directunitcost")     or 0)
        # PowerBI usa totlinediscountperc (fallback linediscount)
        disc     = float(r.get("totlinediscountperc") or r.get("linediscount") or 0)
        ptype    = str(r.get("pricetype") or "").strip().lower()
        active   = is_active_date(ed) and sd <= TODAY

        # Prezzo assoluto (amounttype in {Price, Price & Discount}): usa il più recente valido
        atype = str(r.get("amounttype") or "").strip().lower()
        if up > 0 and active and atype in ("price", "price & discount"):
            if slot.get("_sd_price", "") <= sd:
                slot.update({
                    "unitprice":   up,
                    "directcost":  dc,
                    "description": str(r.get("description") or "").strip(),
                    "uom":         str(r.get("unitofmeasurecode") or "").strip(),
                    "startdate":   sd,
                    "enddate":     ed,
                    "amounttype":  str(r.get("amounttype") or ""),
                    "_sd_price":   sd,
                })
        # Sconto %: solo righe non scadute, amounttype in {Discount, Price & Discount}
        if disc > 0 and active and atype in ("discount", "price & discount"):
            if slot.get("_sd_disc", "") <= sd:
                disc_key = "discount_purch" if ptype == "purchase" else "discount_sale"
                slot.update({disc_key: disc, "_sd_disc": sd})
        if not slot.get("description"):
            slot["description"] = str(r.get("description") or "").strip()
        if not slot.get("uom"):
            slot["uom"] = str(r.get("unitofmeasurecode") or "").strip()

    return item_codes, sale_slots


def build_uom_conversions(token):
    """
    Fetch tabella unità di misura articoli (IFB BC).
    Restituisce: { item_no: { uom_code: qty_per_base_uom } }
    La base UoM ha qty=1. Es: Z3774 -> {BOX: 6, KG: 1, PCS: 1, PLT: 432}
    """
    print("  Fetch UoM conversioni articoli...")
    rows = bc_fetch_all(token, "IFB_Item_Unit_of_Measure")
    print(f"    {len(rows)} righe UoM")
    if rows:
        # Debug: mostra campi disponibili della prima riga
        sample = {k: v for k, v in rows[0].items() if not k.startswith("@")}
        print(f"    Campi disponibili: {list(sample.keys())}")
        print(f"    Esempio prima riga: {sample}")
    conv = {}
    for r in rows:
        # Prova tutti i possibili nomi campo per item code
        item = str(r.get("itemno") or r.get("Item_No") or r.get("item_no") or
                   r.get("No_") or r.get("assetno") or "").strip()
        # Prova tutti i possibili nomi campo per uom code
        code = str(r.get("code") or r.get("Code") or r.get("uomcode") or
                   r.get("UoM") or "").strip().upper()
        # Prova tutti i possibili nomi campo per qty
        qty_raw = (r.get("qtyperunitofmeasure") or r.get("QtyPerUnitOfMeasure") or
                   r.get("qty_per_unit_of_measure") or r.get("Qty") or 1)
        qty = float(qty_raw)
        if item and code:
            conv.setdefault(item, {})[code] = qty
    # Debug: verifica conversioni per articoli chiave
    for test_code in ["HA7021-IB", "Z3774", "BD0501", "CF0051-IFA"]:
        if test_code in conv:
            print(f"    {test_code}: {conv[test_code]}")
        else:
            print(f"    {test_code}: NON TROVATO in UoM table")
    print(f"    Totale articoli con conversioni UoM: {len(conv)}")
    return conv


def build_purchase_prices(token, uom_conv=None):
    """
    Listini acquisto fornitore (pricetype=Purchase, Active).
    Restituisce: (purch_dict, all_codes)
      - purch_dict: assetno -> {FCA/DAP/MTS: {price}, uom, desc}
        prezzo: preferisce record aperti (no end date, start<=oggi), poi non-scaduti, poi il più recente
      - all_codes: TUTTI i codici nel listino acquisto (anche solo con record scaduti)
    """
    # Filtro identico a PowerBI: status=Active + shipmentmethod DAP/FCA
    # NON filtra per pricetype in OData (PowerBI non lo fa) -> filtriamo in Python
    print("  Fetch listini acquisto (status=Active, DAP/FCA, pricetype=Purchase)...")
    rows = fetch_price_lines(token, "status eq 'Active' and (shipmentmethodcode eq 'DAP' or shipmentmethodcode eq 'FCA')")
    rows = [r for r in rows if str(r.get("pricetype") or "").strip().lower() == "purchase"]
    print(f"    {len(rows)} righe purchase dopo filtro pricetype")
    print(f"    {len(rows)} righe totali acquisto")
    result    = defaultdict(lambda: {"FCA": {}, "DAP": {}, "MTS": {}, "uom": "", "desc": ""})
    all_codes = set()
    for r in rows:
        code = str(r.get("assetno") or "").strip()
        if not code:
            continue
        all_codes.add(code)
        sd_r     = str(r.get("startingdate") or "")
        ed       = str(r.get("endingdate") or "")
        # PowerBI usa directunitcost (dir_unit_cost_conv = directunitcost / fatt_conv)
        price    = float(r.get("directunitcost") or 0)
        # Anche sconto acquisto (spurc in PowerBI = totlinediscountperc da Purchase)
        disc_purch = float(r.get("totlinediscountperc") or r.get("linediscount") or 0)
        puom = str(r.get("unitofmeasurecode") or "").strip().upper()
        if price and puom and puom not in ("PCS", "", " ") and uom_conv:
            qty = (uom_conv.get(code) or {}).get(puom)
            if qty and qty > 1:
                price = price / qty
        ship     = classify_ship(r.get("shipmentmethodcode"))
        slot     = result[code][ship]
        expired  = not is_active_date(ed)
        is_open  = is_active_date(ed) and (not sd_r or sd_r <= TODAY)
        slot_open    = slot.get("_open", False)
        slot_expired = slot.get("_expired", True)
        # Priorità: open (no end, start<=oggi) > futuro (no end, start>oggi) > non-scaduto-con-end > scaduto
        # Un record "open" non viene mai sostituito da uno futuro o non-open
        def better(io=is_open, exp=expired, so=slot_open, se=slot_expired, sd=sd_r, sl=slot):
            if io and not so:  return True   # nuovo open, slot non-open -> sostituisci
            if not io and so:  return False  # nuovo non-open, slot open -> tieni
            if not exp and se: return True   # nuovo non-scaduto, slot scaduto -> sostituisci
            if exp and not se: return False  # nuovo scaduto, slot non-scaduto -> tieni
            return sl.get("_sd", "") <= sd   # stessa categoria: più recente per startdate
        if better():
            slot.update({"price": price, "disc_purch": disc_purch,
                         "_sd": sd_r, "_open": is_open, "_expired": expired})
        elif disc_purch > 0 and not slot.get("disc_purch"):
            slot["disc_purch"] = disc_purch
        if not result[code]["uom"]:
            result[code]["uom"]   = str(r.get("unitofmeasurecode") or "").strip()
            result[code]["desc"]  = str(r.get("description") or "").strip()
            result[code]["puom"]  = puom  # UoM del prezzo di acquisto (BOX/PCS/KG)
    print(f"    {len(all_codes)} codici unici ({len(result)} con almeno un record processato)")
    return dict(result), all_codes


def _iss_row_to_price(r):
    fca   = float(r.get("standardcostbranch") or 0)
    dap_v1 = float(r.get("pricedaptotal") or 0)
    dap_v2 = float(r.get("standardcost")  or 0)
    dap    = dap_v1 if dap_v1 > 0 else dap_v2
    if dap == 0 and fca > 0:
        dap = round(fca + float(r.get("carriagecost") or 0), 6)
    return {
        "fca":           fca,
        "dap":           dap,
        "carriage":      float(r.get("carriagecost") or 0),
        "executiondate": str(r.get("executiondate")  or ""),
    }


def build_iss_prices(token, target_codes):
    """
    Fetch IFB_Item_Statistics_SC da BC Italia filtrando per un singolo giorno con $skip
    (ISS ha >5000 righe/giorno, $skip permette paginazione completa).
    Campi ISS:
      standardcostbranch = FCA purchase price (EUR)
      pricedaptotal      = DAP (FCA + carriagecost), dal 2026
      standardcost       = DAP alternativo (fino al 2025)
    Restituisce: {itemno: {fca, dap, carriage, executiondate}}
    """
    if not target_codes:
        return {}
    target_codes = set(target_codes)
    base = (f"https://api.businesscentral.dynamics.com/v2.0/{TENANT_ID}"
            f"/{BC_ENV}/ODataV4/Company('{BC_COMPANY}')/IFB_Item_Statistics_SC")
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    # Prova dal giorno piu' recente; con $skip raccoglie tutti i batch finche' trovati tutti i codici
    for days_back in range(0, 7):
        d = (date.today() - timedelta(days=days_back)).isoformat()
        result = {}
        skip = 0
        while True:
            r = requests.get(base, headers=headers,
                             params={"$filter": f"executiondate eq {d}",
                                     "$top": 5000, "$skip": skip})
            r.raise_for_status()
            rows = r.json().get("value", [])
            for row in rows:
                code = str(row.get("itemno") or "").strip()
                if code in target_codes and code not in result:
                    entry = _iss_row_to_price(row)
                    if entry["fca"] > 0:
                        result[code] = entry
            if len(rows) < 5000:
                break
            if len(result) == len(target_codes):
                break
            skip += 5000
        if result:
            batches = skip // 5000 + 1
            print(f"  ISS data={d}: {len(result)}/{len(target_codes)} trovati in {batches} batch")
            break
    found   = [c for c in target_codes if c in result]
    missing = [c for c in target_codes if c not in result]
    if missing:
        print(f"    Non trovati in ISS: {missing[:10]}{' ...' if len(missing)>10 else ''}")
    for chk in ("LSM30",):
        if chk in result:
            print(f"    {chk} ISS: {result[chk]}")
    return result


def compute_row(branch, code, sale_slots, purch, item_card=None, transport_costs=None, iss_carriage=None):
    """
    Logica identica a PowerBI MILLE SAPORI (HK) — applicata a tutti i branch:
      FCA cost  = directunitcost_purch (convertito in base UoM) * MARKUP (1/0.98)
      DAP cost  = idem da listino DAP; fallback FCA + carriagecost ISS
      Discount  = spurc + ssale (sconto acquisto + sconto vendita cliente)
      Discounted = Price * (1 - Discount/100)
    """
    fca_sale = sale_slots.get("FCA", {})
    mts_sale = sale_slots.get("MTS", {})
    dap_sale = sale_slots.get("DAP", {})
    pur      = purch.get(code, {})

    purch_fca = pur.get("FCA", {}).get("price") or 0.0
    purch_dap = pur.get("DAP", {}).get("price") or 0.0
    purch_mts = pur.get("MTS", {}).get("price") or 0.0

    def to_sell(p):
        return round(p * MARKUP, 6) if p else 0.0

    # Prezzo: acquisto × MARKUP (PowerBI: dir_unit_cost_conv * Fatt_Conv / 0.98)
    # Il listino vendita a volte ha un prezzo assoluto — in quel caso lo usa PowerBI?
    # NO: le formule PowerBI usano SEMPRE directunitcost dal listino acquisto.
    # Il listino vendita serve solo per gli sconti (discount_sale).
    fca_price = to_sell(purch_fca)
    dap_price = to_sell(purch_dap)
    mts_price = to_sell(purch_mts)
    carriage  = round(dap_price - fca_price, 4) if dap_price and fca_price else 0.0

    # Discount = spurc (purch) + ssale (sale) — PowerBI: spurc + ssale
    purch_fca_disc = pur.get("FCA", {}).get("disc_purch") or 0.0
    purch_dap_disc = pur.get("DAP", {}).get("disc_purch") or 0.0
    sale_fca_disc  = fca_sale.get("discount_sale") or fca_sale.get("discount") or 0.0
    sale_dap_disc  = dap_sale.get("discount_sale") or dap_sale.get("discount") or 0.0
    sale_mts_disc  = mts_sale.get("discount_sale") or mts_sale.get("discount") or 0.0

    fca_disc = purch_fca_disc + sale_fca_disc
    mts_disc = sale_mts_disc
    # DAP discount: if spurc_dap + ssale_dap > 0, usa quelli; altrimenti FCA discount
    raw_dap_disc = purch_dap_disc + sale_dap_disc
    dap_disc = raw_dap_disc if raw_dap_disc > 0 else fca_disc

    def apply(price, disc):
        return round(price * (1 - disc / 100), 6) if price else 0.0

    # Se DAP=0 e FCA>0: usa ISS carriagecost (campo BC compilato manualmente) poi tabella trasporti
    if dap_price == 0 and fca_price > 0:
        iss_cr = float((iss_carriage or {}).get(code, 0))
        if iss_cr > 0:
            carriage  = round(iss_cr, 6)
            dap_price = round(fca_price + carriage, 6)
        elif item_card and transport_costs:
            ic = item_card.get(code, {})
            vendor   = ic.get("vendorName", "")
            temp     = ic.get("temperature", "")
            pallet1  = transport_costs.get((vendor, temp), 0)
            qty_per_box    = ic.get("qtyPerBox", 0)
            box_per_pallet = ic.get("boxPerPallet", 0)
            pcs_per_pallet = qty_per_box * box_per_pallet
            if pallet1 > 0 and pcs_per_pallet > 0:
                carriage  = round(pallet1 / pcs_per_pallet, 6)
                dap_price = round(fca_price + carriage, 6)

    fca_discounted = apply(fca_price, fca_disc)
    mts_discounted = apply(mts_price, mts_disc)
    dap_discounted = apply(dap_price, dap_disc)

    desc = (fca_sale.get("description") or mts_sale.get("description")
            or dap_sale.get("description") or pur.get("desc") or "")
    uom  = (fca_sale.get("uom") or mts_sale.get("uom")
            or dap_sale.get("uom") or pur.get("uom") or "")
    sd   = (fca_sale.get("startdate") or mts_sale.get("startdate")
            or dap_sale.get("startdate") or "")
    ed   = (fca_sale.get("enddate") or mts_sale.get("enddate")
            or dap_sale.get("enddate") or "")

    puom = pur.get("puom", "") if pur else ""
    return {
        "b":  branch,
        "n":  code,
        "d":  desc[:60] if desc else "",
        "pu": puom,  # UoM acquisto (BOX/PCS/KG) — serve per conversione in app
        "fp": round(fca_price, 6),
        "fd": round(fca_disc, 4),
        "fc": round(fca_discounted, 6),
        "mp": round(mts_price, 6),
        "dp": round(dap_price, 6),
        "dd": round(dap_disc, 4),
        "dc": round(dap_discounted, 6),
        "cr": round(carriage, 6),
    }


if __name__ == "__main__":
    if not CLIENT_SECRET:
        raise RuntimeError("BC_CLIENT_SECRET non impostato")

    print("Ottengo token BC...")
    token = get_token()

    uom_conv = {}
    try:
        uom_conv = build_uom_conversions(token)
    except Exception as e:
        print(f"  Warning: UoM conversioni non disponibili ({e}) — prezzi senza conversione")

    item_card = {}
    try:
        item_card = build_item_card_data(token)
    except Exception as e:
        print(f"  Warning: Item Card data non disponibile ({e})")

    transport_costs = {}
    try:
        transport_costs = build_transport_costs(token)
    except Exception as e:
        print(f"  Warning: Costi trasporto non disponibili ({e})")

    purch, all_purchase_codes = build_purchase_prices(token, uom_conv)
    print(f"  Articoli con prezzo acquisto valido: {sum(1 for v in purch.values() if v.get('FCA',{}).get('price') or v.get('DAP',{}).get('price'))}")
    print(f"  Articoli totali nel listino acquisto BC Italia: {len(all_purchase_codes)}")

    # Fetch ISS per TUTTI gli articoli: carriagecost per-articolo da BC + articoli HK non nel listino
    iss_carriage = {}
    try:
        hk_data_path = Path(__file__).parent.parent / "docs" / "data" / "hk_anagrafica.json"
        hk_items = json.loads(hk_data_path.read_text(encoding="utf-8"))
        hk_codes = {str(item.get("code") or "").strip() for item in hk_items if item.get("code")}
        all_target = all_purchase_codes | hk_codes
        print(f"  Fetch ISS per {len(all_target)} articoli (carriagecost per-articolo + HK mancanti)...")
        iss_prices = build_iss_prices(token, all_target)
        iss_carriage = {code: iss["carriage"] for code, iss in iss_prices.items() if iss.get("carriage", 0) > 0}
        print(f"  ISS: {len(iss_carriage)} articoli con carriagecost > 0")
        # Supplementa purch con articoli HK non nel listino acquisto BC Italia
        missing_hk = hk_codes - all_purchase_codes
        print(f"  {len(hk_codes)} codici IFB in hk_anagrafica, {len(missing_hk)} non nel listino acquisto BC Italia")
        for code in missing_hk:
            if code in iss_prices:
                iss = iss_prices[code]
                fca = iss.get("fca") or 0
                dap = iss.get("dap") or 0
                entry = {"FCA": {}, "DAP": {}, "MTS": {}, "uom": "", "desc": "", "puom": "PCS"}
                if fca > 0:
                    entry["FCA"] = {"price": fca, "_open": True, "_expired": False}
                if dap > 0:
                    entry["DAP"] = {"price": dap, "_open": True, "_expired": False}
                purch[code] = entry
        all_purchase_codes = all_purchase_codes | set(iss_prices.keys())
        print(f"  Totale articoli dopo merge ISS: {len(all_purchase_codes)}")
    except Exception as e:
        import traceback; traceback.print_exc()
        print(f"  Warning: fetch ISS fallito ({e})")

    all_rows = []

    for branch, cust_no in CUSTOMERS.items():
        print(f"\nFetch listino vendita {branch} (customer {cust_no})...")
        rows = fetch_price_lines(token, f"assigntono eq '{cust_no}' and status eq 'Active'")
        print(f"  {len(rows)} righe trovate")

        _item_codes, active_discounts = build_item_prices(rows, pricetype_filter="sale")
        with_price = sum(1 for s in active_discounts.values() for sh in s.values() if sh.get("unitprice",0) > 0)
        with_disc  = sum(1 for s in active_discounts.values() for sh in s.values() if sh.get("discount",0) > 0)
        print(f"  {with_price} slot con prezzo assoluto, {with_disc} con sconto attivo")

        all_codes = all_purchase_codes
        print(f"  {len(all_codes)} articoli totali nel listino acquisto -> usati come base listino")

        for code in all_codes:
            slots = active_discounts.get(code, {"FCA": {}, "MTS": {}, "DAP": {}})
            row = compute_row(branch, code, slots, purch, item_card, transport_costs, iss_carriage)
            all_rows.append(row)

    print(f"\nTotale {len(all_rows)} righe listino (HK+CAN+MAC)")

    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)

    # Scrivi un file per branch (più leggero da caricare sul client)
    for br in CUSTOMERS:
        br_rows = [r for r in all_rows if r["b"] == br]
        br_path = OUT_PATH.parent / f"ifb_listini_{br}.json"
        br_path.write_text(json.dumps(br_rows, ensure_ascii=False, separators=(",",":")), encoding="utf-8")
        with_price = sum(1 for r in br_rows if r["fp"] > 0 or r["dp"] > 0)
        print(f"  {br}: {len(br_rows)} righe ({with_price} con prezzo) -> {br_path.name}")

    # Mantieni anche il file unico per compatibilità
    OUT_PATH.write_text(json.dumps(all_rows, ensure_ascii=False, separators=(",",":")), encoding="utf-8")
    print(f"Scritto {len(all_rows)} righe in {OUT_PATH.name}")
