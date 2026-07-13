"""
Aggiorna docs/data/mac_anagrafica.json con l'anagrafica articoli MACAO da BC.
Company: HOFF Macau (Production_HK environment).
Usato da GitHub Actions. CLIENT_SECRET letto da variabile d'ambiente BC_CLIENT_SECRET.
"""
import os, json, requests
from pathlib import Path

TENANT_ID     = "2acd007b-8d3f-4be0-9681-cf248264a0e2"
CLIENT_ID     = "925de6e4-e71f-4c24-9e0a-f3ae544ae644"
CLIENT_SECRET = os.environ.get("BC_CLIENT_SECRET", "")
BC_ENV        = "Production_HK"
COMPANY       = "HOFF%20Macau"
BASE          = f"https://api.businesscentral.dynamics.com/v2.0/{TENANT_ID}/{BC_ENV}/ODataV4/Company('{COMPANY}')"
OUT_PATH      = Path(__file__).parent.parent / "docs" / "data" / "mac_anagrafica.json"

UOM_MAP  = {"pcs": "PCS", "pz": "PCS", "piece": "PCS", "box": "BOX", "ctn": "BOX",
            "cartone": "BOX", "kg": "KG", "kgs": "KG", "gr": "GR", "gram": "GR"}
TEMP_MAP = {"dry": "DRY", "fresh": "FRESH", "frozen": "FROZEN",
            "ambient": "DRY", "refrigerated": "FRESH"}

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


def bc_get(token, endpoint):
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    results, url = [], f"{BASE}/{endpoint}"
    while url:
        r = requests.get(url, headers=headers)
        if not r.ok:
            print(f"  ERRORE {r.status_code}: {r.text[:300]}")
            r.raise_for_status()
        data = r.json()
        results.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return results


def get_anagrafica(token):
    # Prova prima con Item_Card_Excel (stesso pattern di HK), poi fallback su Item
    for entity in ["Item_Card_Excel", "Item"]:
        try:
            print(f"  Provo entity: {entity}...")
            rows = bc_get(token, f"{entity}?$top=5000")
            if rows:
                print(f"  Trovati {len(rows)} items con '{entity}'")
                return rows, entity
        except Exception as e:
            print(f"  '{entity}' non disponibile: {e}")
    return [], None


def parse_item(item, entity):
    """Mappa i campi BC → struttura app. I nomi dei campi possono variare per HOFF Macau."""
    def f(keys):
        for k in keys if isinstance(keys, list) else [keys]:
            v = item.get(k)
            if v is not None:
                return v
        return None

    # Codici
    macao_no = str(f(["No", "No_"]) or "").strip()
    bv_no    = str(f(["AltICMIFB_Item", "AltMACBV_No", "BV_No", "Cross_Reference_No"]) or "").strip()
    ifb_item = str(f(["AltICMIFB_Item", "AltMACIFB_Item", "IFB_Item"]) or bv_no).strip()

    # Descrizione e categoria
    description     = str(f(["Description"]) or "").strip()
    section_desc    = str(f(["AltICMSection_Description","AltMACSection_Description","Section_Description","Gen_Prod_Posting_Group"]) or "").strip()
    product_type    = norm(f(["AltICMProduct_Type","AltMACProduct_Type","Product_Type","Item_Tracking_Code"]), TEMP_MAP) or "DRY"
    vendor_name     = str(f(["AltICMVendor_Name","AltMACVendor_Name","Vendor_Name"]) or "BRIGHT VIEW TRADING HK LTD").strip()

    # UOM
    sales_uom_mac   = norm(f(["Sales_Unit_of_Measure","AltMACSales_UOM","AltICMSales_UOM"]), UOM_MAP) or "KG"
    sales_uom_bv    = norm(f(["AltMACBV_Sales_UOM","AltICMBV_Sales_UOM","Base_Unit_of_Measure"]), UOM_MAP) or sales_uom_mac

    # Pesi e quantità (chiave per formula 5-step)
    pcs_gross   = float(f(["AltICMPcs_Gross_Weight","AltMACPcs_Gross_Weight","Gross_Weight","Pcs_Gross_Weight"]) or 0)
    pcs_net     = float(f(["AltICMPcs_Net_Weight","AltMACPcs_Net_Weight","Net_Weight","Pcs_Net_Weight"]) or 0)
    qty_per_box = float(f(["AltICMQuantity_x_Packaging","AltMACQuantity_x_Packaging","Quantity_x_Packaging","Units_per_Parcel"]) or 0)
    pcs_x_kg    = float(f(["AltICMPcs_x_Kg","AltMACPcs_x_Kg","Pcs_x_Kg"]) or 0)

    # SC BV (HKD) — valore standard cost da sistema BV, base del calcolo MAC
    std_cost_hkd = float(f(["Standard_Cost","AltICMStandard_Cost","AltMACStandard_Cost"]) or 0)

    # Flag HOFF
    is_hoff = bool(f(["AltICMHOFF","AltMACHOFF","HOFF"]))

    # Stato
    blocked = item.get("Blocked") is True

    return {
        "id":              macao_no or ifb_item,
        "nHK":             macao_no,
        "code":            ifb_item or bv_no,
        "description":     description,
        "category":        section_desc,
        "producttype":     product_type,
        "temperature":     product_type,
        "vendorName":      vendor_name,
        "uom":             sales_uom_mac,
        "hkUom":           sales_uom_bv,
        "pcsgrossweight":  pcs_gross,
        "pcsnetweight":    pcs_net,
        "qtyPerBox":       qty_per_box,
        "pcsxkg":          pcs_x_kg,
        "standardCostHkd": std_cost_hkd,
        "isHoff":          is_hoff,
        "active":          not blocked,
        "_entity":         entity,  # debug: quale entity ha risposto
    }


if __name__ == "__main__":
    if not CLIENT_SECRET:
        raise RuntimeError("BC_CLIENT_SECRET non impostato")

    print("Ottengo token BC...")
    token = get_token()

    print(f"Leggo anagrafica HOFF Macau da BC ({BC_ENV})...")
    items, entity = get_anagrafica(token)

    if not items:
        print("ATTENZIONE: nessun item trovato. Verifica il nome dell'entity in BC HOFF Macau.")
        print("Suggerimento: esegui scripts/discover_entities.py con COMPANY='HOFF Macau' per trovare il nome corretto.")
        exit(1)

    products = [parse_item(i, entity) for i in items if (i.get("No") or i.get("No_"))]
    print(f"  {len(products)} prodotti parsati")

    # Debug: mostra primo prodotto per verifica campi
    if products:
        print(f"  Esempio: {json.dumps(products[0], ensure_ascii=False, indent=2)}")

    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    OUT_PATH.write_text(json.dumps(products, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Scritto {len(products)} prodotti in {OUT_PATH}")
