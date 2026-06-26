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
from datetime import date
from pathlib import Path
from collections import defaultdict

TENANT_ID     = "2acd007b-8d3f-4be0-9681-cf248264a0e2"
CLIENT_ID     = "925de6e4-e71f-4c24-9e0a-f3ae544ae644"
CLIENT_SECRET = os.environ.get("BC_CLIENT_SECRET", "")
BC_ENV        = "Production"
BC_COMPANY    = "Inalca%20Food%20%26%20Beverage%20s.r.l."
OUT_PATH      = Path(__file__).parent.parent / "docs" / "data" / "ifb_listini.json"

TODAY  = date.today().isoformat()
MARKUP = 100 / 99  # markup IFB su costo acquisto (~1%)

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
    base = (f"https://api.businesscentral.dynamics.com/v2.0/{TENANT_ID}"
            f"/{BC_ENV}/ODataV4/Company('{BC_COMPANY}')/")
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

def is_valid_purchase(startdate_str, enddate_str):
    """Righe acquisto valide: senza data fine E con data inizio <= oggi."""
    ed = (enddate_str or "").split("T")[0]
    sd = (startdate_str or "").split("T")[0]
    if ed not in ("", "0001-01-01"):
        return False  # ha data fine → salta
    if sd and sd > TODAY:
        return False  # non ancora attivo → salta
    return True



def build_item_prices(rows):
    """
    Price list lines di vendita per un cliente.
    Restituisce: (item_codes, active_discounts)
      - item_codes: set di assetno presenti nel listino (anche scaduti)
      - active_discounts: assetno -> {FCA/MTS/DAP: {...}} solo record NON scaduti
    """
    item_codes = set()
    active = defaultdict(lambda: {"FCA": {}, "MTS": {}, "DAP": {}})

    for r in rows:
        code = str(r.get("assetno") or "").strip()
        if not code:
            continue
        item_codes.add(code)
        ed = str(r.get("endingdate") or "")
        if is_expired(ed):
            continue
        ship = classify_ship(r.get("shipmentmethodcode"))
        slot = active[code][ship]
        up   = float(r.get("unitprice")      or 0)
        dc   = float(r.get("directunitcost") or 0)
        disc = float(r.get("linediscount")   or 0)
        sd   = str(r.get("startingdate") or "")
        if slot.get("_sd", "") <= sd:
            slot.update({
                "unitprice":   up,
                "directcost":  dc,
                "discount":    disc,
                "description": str(r.get("description") or "").strip(),
                "uom":         str(r.get("unitofmeasurecode") or "").strip(),
                "startdate":   sd,
                "enddate":     ed,
                "amounttype":  str(r.get("amounttype") or ""),
                "_sd":         sd,
            })

    return item_codes, active


def build_purchase_prices(token):
    """
    Listini acquisto fornitore (pricetype=Purchase, Active).
    Restituisce: (purch_dict, all_codes)
      - purch_dict: assetno -> {FCA/DAP/MTS: {price}, uom, desc}
        prezzo: preferisce record aperti (no end date, start<=oggi), poi non-scaduti, poi il più recente
      - all_codes: TUTTI i codici nel listino acquisto (anche solo con record scaduti)
    """
    print("  Fetch listini acquisto (pricetype=Purchase)...")
    rows = fetch_price_lines(token, "pricetype eq 'Purchase' and status eq 'Active'")
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
        dc       = float(r.get("directunitcost") or 0)
        up       = float(r.get("unitprice")      or 0)
        price    = dc or up
        ship     = classify_ship(r.get("shipmentmethodcode"))
        slot     = result[code][ship]
        expired  = is_expired(ed)
        is_open  = not expired and ed in ("", "0001-01-01") and (not sd_r or sd_r <= TODAY)
        slot_open    = slot.get("_open", False)
        slot_expired = slot.get("_expired", True)
        # Priorità: open > non-scaduto > scaduto; a parità, startingdate più recente
        def better(io=is_open, exp=expired, so=slot_open, se=slot_expired, sd=sd_r, sl=slot):
            if io and not so:           return True
            if not exp and se:          return True
            if exp and not se:          return False
            return sl.get("_sd", "") <= sd
        if better():
            slot.update({"price": price, "_sd": sd_r, "_open": is_open, "_expired": expired})
        if not result[code]["uom"]:
            result[code]["uom"]  = str(r.get("unitofmeasurecode") or "").strip()
            result[code]["desc"] = str(r.get("description") or "").strip()
    print(f"    {len(all_codes)} codici unici ({len(result)} con almeno un record processato)")
    return dict(result), all_codes


def compute_row(branch, code, sale_slots, purch):
    """
    FCA/DAP Price = purchase * MARKUP (100/99)
    Carriage      = DAP_sell - FCA_sell
    Discount      = da listino vendita (solo non scaduti)
    Discounted    = Price * (1 - Discount/100)
    """
    fca_sale = sale_slots.get("FCA", {})
    mts_sale = sale_slots.get("MTS", {})
    dap_sale = sale_slots.get("DAP", {})
    pur      = purch.get(code, {})

    purch_fca = pur.get("FCA", {}).get("price") or 0.0
    purch_dap = pur.get("DAP", {}).get("price") or 0.0
    purch_mts = pur.get("MTS", {}).get("price") or 0.0

    def to_sell(p):
        return round(p * MARKUP, 2) if p else 0.0

    # Se il listino vendita ha un prezzo assoluto (amounttype=Price, unitprice>0), usalo direttamente.
    # Altrimenti calcola da acquisto × markup.
    sale_fca_abs = float(fca_sale.get("unitprice") or 0)
    sale_dap_abs = float(dap_sale.get("unitprice") or 0)
    sale_mts_abs = float(mts_sale.get("unitprice") or 0)

    fca_price = sale_fca_abs if sale_fca_abs > 0 else to_sell(purch_fca)
    dap_price = sale_dap_abs if sale_dap_abs > 0 else to_sell(purch_dap)
    mts_price = sale_mts_abs if sale_mts_abs > 0 else to_sell(purch_mts)
    carriage  = round(dap_price - fca_price, 4) if dap_price and fca_price else 0.0

    fca_disc = fca_sale.get("discount") or 0.0
    mts_disc = mts_sale.get("discount") or 0.0
    dap_disc = dap_sale.get("discount") or fca_disc

    def apply(price, disc):
        return round(price * (1 - disc / 100), 6) if price else 0.0

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

    return {
        "Branch":         branch,
        "No_":            code,
        "Description":    desc,
        "UoM":            uom,
        "MTS_Price":      round(mts_price, 6),
        "MTS_Discount":   round(mts_disc, 4),
        "MTS_Discounted": round(mts_discounted, 6),
        "FCA_Price":      round(fca_price, 6),
        "FCA_Discount":   round(fca_disc, 4),
        "FCA_Discounted": round(fca_discounted, 6),
        "Carriage":       round(carriage, 6),
        "DAP_Price":      round(dap_price, 6),
        "DAP_Discount":   round(dap_disc, 4),
        "DAP_Discounted": round(dap_discounted, 6),
        "DAP_Final":      round(dap_discounted, 6),
        "StartDate":      sd,
        "EndDate":        ed,
    }


if __name__ == "__main__":
    if not CLIENT_SECRET:
        raise RuntimeError("BC_CLIENT_SECRET non impostato")

    print("Ottengo token BC...")
    token = get_token()

    purch, all_purchase_codes = build_purchase_prices(token)
    print(f"  Articoli con prezzo acquisto valido: {sum(1 for v in purch.values() if v.get('FCA',{}).get('price') or v.get('DAP',{}).get('price'))}")
    print(f"  Articoli totali nel listino acquisto: {len(all_purchase_codes)}")

    all_rows = []

    for branch, cust_no in CUSTOMERS.items():
        print(f"\nFetch listino vendita {branch} (customer {cust_no})...")
        rows = fetch_price_lines(token, f"assigntono eq '{cust_no}' and status eq 'Active'")
        print(f"  {len(rows)} righe trovate")

        _item_codes, active_discounts = build_item_prices(rows)
        print(f"  {len(active_discounts)} articoli con sconto attivo")

        # Tutti gli articoli nel listino acquisto (anche quelli con solo record scaduti)
        all_codes = all_purchase_codes
        print(f"  {len(all_codes)} articoli totali nel listino acquisto → usati come base listino")

        for code in all_codes:
            slots = active_discounts.get(code, {"FCA": {}, "MTS": {}, "DAP": {}})
            row = compute_row(branch, code, slots, purch)
            all_rows.append(row)

    print(f"\nTotale {len(all_rows)} righe listino (HK+CAN+MAC)")

    hk_ex  = next((r for r in all_rows if r["Branch"] == "HK"  and r["FCA_Price"] > 0), None)
    can_ex = next((r for r in all_rows if r["Branch"] == "CAN" and r["FCA_Price"] > 0), None)
    if hk_ex:  print(f"  Esempio HK:  {hk_ex}")
    if can_ex: print(f"  Esempio CAN: {can_ex}")

    with_price = sum(1 for r in all_rows if r["FCA_Price"] > 0 or r["DAP_Price"] > 0)
    disc_only  = sum(1 for r in all_rows if r["FCA_Price"] == 0 and r["FCA_Discount"] > 0)
    print(f"  Con prezzo assoluto: {with_price}")
    print(f"  Solo sconto (prezzo base non in BC): {disc_only}")

    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    OUT_PATH.write_text(json.dumps(all_rows, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\nScritto {len(all_rows)} righe in {OUT_PATH}")
