"""
Aggiorna docs/data/mac_anagrafica.json con l'anagrafica articoli MACAO da BC.

Logica:
  1. Prende TUTTI i prodotti da BrightView (qualunque vendor) con il loro Standard_Cost (HKD)
  2. Prende i prodotti HOFF Macau con vendor = BrightView
  3. Fa il join per codice (stesso codice nei due sistemi)
  4. standardCostHkd viene da BrightView Standard_Cost

Company BV:   BRIGHT VIEW TRADING HK LIMITED (Production_HK)
Company HOFF: HOFF Macau (Production_HK)
"""
import os, json, requests
from pathlib import Path

TENANT_ID     = "2acd007b-8d3f-4be0-9681-cf248264a0e2"
CLIENT_ID     = "925de6e4-e71f-4c24-9e0a-f3ae544ae644"
CLIENT_SECRET = os.environ.get("BC_CLIENT_SECRET", "")
BC_ENV        = "Production_HK"
COMPANY_BV    = "BRIGHT%20VIEW%20TRADING%20HK%20LIMITED"
COMPANY_HOFF  = "HOFF%20Macau"
BASE_BV       = f"https://api.businesscentral.dynamics.com/v2.0/{TENANT_ID}/{BC_ENV}/ODataV4/Company('{COMPANY_BV}')"
BASE_HOFF     = f"https://api.businesscentral.dynamics.com/v2.0/{TENANT_ID}/{BC_ENV}/ODataV4/Company('{COMPANY_HOFF}')"
OUT_PATH      = Path(__file__).parent.parent / "docs" / "data" / "mac_anagrafica.json"

UOM_MAP  = {"pcs": "PCS", "pz": "PCS", "piece": "PCS", "box": "BOX", "ctn": "BOX",
            "cartone": "BOX", "kg": "KG", "kgs": "KG", "gr": "GR", "gram": "GR"}
TEMP_MAP = {"dry": "DRY", "fresh": "FRESH", "frozen": "FROZEN",
            "ambient": "DRY", "refrigerated": "FRESH"}

BV_VENDOR_KEYWORDS = ["bright view", "brightview"]


def norm(v, mapping):
    s = str(v or "").strip().lower()
    return mapping.get(s, str(v or "").upper() if v else "")


def get_token():
    r = requests.post(
        f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token",
        data={"grant_type": "client_credentials", "client_id": CLIENT_ID,
              "client_secret": CLIENT_SECRET,
              "scope": "https://api.businesscentral.dynamics.com/.default"},
    )
    r.raise_for_status()
    return r.json()["access_token"]


def bc_get(token, base, endpoint):
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    results, url = [], f"{base}/{endpoint}"
    while url:
        r = requests.get(url, headers=headers)
        if not r.ok:
            print(f"  ERRORE {r.status_code}: {r.text[:300]}")
            r.raise_for_status()
        data = r.json()
        results.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return results


def fetch_bv_items(token):
    """Tutti i prodotti BrightView (qualunque vendor) con Standard_Cost."""
    for entity in ["Item_Card_Excel", "Item"]:
        try:
            print(f"  BV: provo entity '{entity}'...")
            rows = bc_get(token, BASE_BV, f"{entity}?$top=5000")
            if rows:
                print(f"  BV: {len(rows)} items trovati")
                return {str(r.get("No") or r.get("No_") or "").strip(): r for r in rows if r.get("No") or r.get("No_")}
        except Exception as e:
            print(f"  BV: '{entity}' non disponibile: {e}")
    return {}


def is_bv_vendor(vendor_name: str) -> bool:
    vl = (vendor_name or "").strip().lower()
    return any(k in vl for k in BV_VENDOR_KEYWORDS)


def fetch_hoff_items(token):
    """Prodotti HOFF Macau con vendor = BrightView."""
    for entity in ["Item_Card_Excel", "Item"]:
        try:
            print(f"  HOFF: provo entity '{entity}'...")
            rows = bc_get(token, BASE_HOFF, f"{entity}?$top=5000")
            if rows:
                bv_rows = [r for r in rows if is_bv_vendor(
                    r.get("AltICMVendor_Name") or r.get("AltMACVendor_Name") or r.get("Vendor_Name") or ""
                )]
                print(f"  HOFF: {len(rows)} totali, {len(bv_rows)} con vendor=BrightView")
                if bv_rows:
                    return bv_rows, entity
                print(f"  HOFF: nessun vendor=BrightView trovato, uso tutti i {len(rows)} items")
                return rows, entity
        except Exception as e:
            print(f"  HOFF: '{entity}' non disponibile: {e}")
    return [], None


def f_field(item, keys):
    for k in (keys if isinstance(keys, list) else [keys]):
        v = item.get(k)
        if v is not None:
            return v
    return None


def parse_hoff_item(hoff, bv_item, entity):
    macao_no = str(f_field(hoff, ["No", "No_"]) or "").strip()
    # Il codice HOFF Macau è uguale al codice BrightView (stessa codifica)
    bv_code  = macao_no

    description  = str(f_field(hoff, ["Description"]) or "").strip()
    product_type = norm(f_field(hoff, ["AltICMProduct_Type","AltMACProduct_Type","Product_Type","Item_Tracking_Code"]), TEMP_MAP) or "DRY"
    vendor_name  = str(f_field(hoff, ["AltICMVendor_Name","AltMACVendor_Name","Vendor_Name"]) or "BRIGHT VIEW TRADING HK LTD").strip()
    section_desc = str(f_field(hoff, ["AltICMSection_Description","AltMACSection_Description","Section_Description","Gen_Prod_Posting_Group"]) or "").strip()

    sales_uom_mac = norm(f_field(hoff, ["Sales_Unit_of_Measure","AltMACSales_UOM","AltICMSales_UOM"]), UOM_MAP) or "KG"
    sales_uom_bv  = norm(f_field(hoff, ["AltMACBV_Sales_UOM","AltICMBV_Sales_UOM","Base_Unit_of_Measure"]), UOM_MAP) or sales_uom_mac

    pcs_gross   = float(f_field(hoff, ["AltICMPcs_Gross_Weight","AltMACPcs_Gross_Weight","Gross_Weight","Pcs_Gross_Weight"]) or 0)
    qty_per_box = float(f_field(hoff, ["AltICMQuantity_x_Packaging","AltMACQuantity_x_Packaging","Quantity_x_Packaging","Units_per_Parcel"]) or 0)

    # Standard Cost HKD: fonte primaria = BrightView Standard_Cost
    std_cost_hkd = 0.0
    bv_sc_source = "hoff"
    if bv_item:
        bv_sc = float(bv_item.get("Standard_Cost") or 0)
        if bv_sc == 0:
            bv_sc = float(bv_item.get("Unit_Cost") or 0)
        if bv_sc > 0:
            std_cost_hkd = bv_sc
            bv_sc_source = "bv"
        # BV Sales UOM: preferisci quello dell'anagrafica BV
        bv_uom = norm(bv_item.get("Sales_Unit_of_Measure") or bv_item.get("Base_Unit_of_Measure") or "", UOM_MAP)
        if bv_uom:
            sales_uom_bv = bv_uom
    if std_cost_hkd == 0:
        std_cost_hkd = float(f_field(hoff, ["Standard_Cost","AltICMStandard_Cost","AltMACStandard_Cost"]) or 0)

    is_hoff  = bool(f_field(hoff, ["AltICMHOFF","AltMACHOFF","HOFF"]))
    blocked  = hoff.get("Blocked") is True

    return {
        "id":              macao_no,
        "nHK":             macao_no,
        "code":            bv_code,
        "description":     description,
        "category":        section_desc,
        "producttype":     product_type,
        "temperature":     product_type,
        "vendorName":      vendor_name,
        "uom":             sales_uom_mac,
        "hkUom":           sales_uom_bv,
        "pcsgrossweight":  pcs_gross,
        "qtyPerBox":       qty_per_box,
        "standardCostHkd": round(std_cost_hkd, 5),
        "isHoff":          is_hoff,
        "active":          not blocked,
        "_bvScSource":     bv_sc_source,
        "_entity":         entity,
    }


if __name__ == "__main__":
    if not CLIENT_SECRET:
        raise RuntimeError("BC_CLIENT_SECRET non impostato")

    print("Ottengo token BC...")
    token = get_token()

    print("Leggo ALL items da BrightView (qualunque vendor)...")
    bv_index = fetch_bv_items(token)
    print(f"  {len(bv_index)} items BV indicizzati per codice")

    print("Leggo items HOFF Macau con vendor=BrightView...")
    hoff_items, entity = fetch_hoff_items(token)

    if not hoff_items:
        print("ATTENZIONE: nessun item trovato in HOFF Macau.")
        exit(1)

    products = []
    matched = 0
    for hoff in hoff_items:
        no = str(f_field(hoff, ["No", "No_"]) or "").strip()
        if not no:
            continue
        bv_item = bv_index.get(no)
        if bv_item:
            matched += 1
        products.append(parse_hoff_item(hoff, bv_item, entity))

    print(f"  {len(products)} prodotti totali, {matched} matchati con BV Standard_Cost")
    if products:
        sc_found = sum(1 for p in products if p["standardCostHkd"] > 0)
        print(f"  {sc_found} con standardCostHkd > 0")
        print(f"  Esempio: {json.dumps(products[0], ensure_ascii=False, indent=2)}")

    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    OUT_PATH.write_text(json.dumps(products, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Scritto {len(products)} prodotti in {OUT_PATH}")
