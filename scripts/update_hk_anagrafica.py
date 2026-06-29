"""
Aggiorna docs/data/hk_anagrafica.json con l'anagrafica articoli HK da Business Central.
Usato da GitHub Actions. CLIENT_SECRET letto da variabile d'ambiente BC_CLIENT_SECRET.

NOTA: verifica i nomi dei campi BC commentati con "# BC field?" se qualcosa non corrisponde.
"""
import os, json, requests
from pathlib import Path

TENANT_ID     = "2acd007b-8d3f-4be0-9681-cf248264a0e2"
CLIENT_ID     = "925de6e4-e71f-4c24-9e0a-f3ae544ae644"
CLIENT_SECRET = os.environ.get("BC_CLIENT_SECRET", "")
BC_ENV        = "Production_HK"
COMPANY       = "BRIGHT%20VIEW%20TRADING%20HK%20LIMITED"
BASE          = f"https://api.businesscentral.dynamics.com/v2.0/{TENANT_ID}/{BC_ENV}/ODataV4/Company('{COMPANY}')"
OUT_PATH      = Path(__file__).parent.parent / "docs" / "data" / "hk_anagrafica.json"

# Mapping BC category values → app values
CAT_MAP = {
    "beef": "BEEF", "beef meat": "BEEF MEAT", "wine": "WINE", "pasta": "PASTA",
    "rice and cereals": "RICE AND CEREALS", "bakery": "BAKERY", "desserts": "DESSERTS AND PREPARATIONS",
    "desserts and preparations": "DESSERTS AND PREPARATIONS", "meat": "MEAT", "food": "FOOD",
    "seafood": "SEAFOOD", "dairy": "DAIRY", "oil": "OIL AND FATS", "condiments": "CONDIMENTS",
    "beverages": "BEVERAGES", "charcuterie": "CHARCUTERIE",
}
UOM_MAP  = {"pcs": "PCS", "pz": "PCS", "piece": "PCS", "box": "BOX", "ctn": "BOX",
            "cartone": "BOX", "kg": "KG", "kgs": "KG"}
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
        r.raise_for_status()
        data = r.json()
        results.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return results


def get_anagrafica(token):
    vendor = "INALCA FOOD %26 BEVERAGE SRL"
    fields = ",".join([
        "No", "AltICMIFB_Item", "Description",
        "AltICMSection_Description",   # category
        "Sales_Unit_of_Measure",       # uom
        "AltICMQuantity_x_Packaging",  # qtyPerBox
        "AltICMPackaging_x_Pallet",    # boxPerPallet
        "AltICMNet_Weight",            # kgPerBox
        "AltICMKg_x_PLT",             # kgxplt
        "AltICMProduct_Type",          # temperature
        "Transportation",              # modalità trasporto (CH AIR / DRY AIR / FR AIR / SEA ...)
        "Standard_Cost",               # standard cost HKD su BC Brightview
        "Blocked",
        "AltICMVendor_Name", "AltICMVendor_Name_2",
        "AltICMHOFF",
    ])
    endpoint = (
        f"Item_Card_Excel"
        f"?$filter=Gen_Prod_Posting_Group eq 'DS' and AltICMVendor_Name eq '{vendor}'"
        f"&$select={fields}&$top=5000"
    )
    return bc_get(token, endpoint)


if __name__ == "__main__":
    if not CLIENT_SECRET:
        raise RuntimeError("BC_CLIENT_SECRET non impostato")

    print("Ottengo token BC...")
    token = get_token()

    print("Leggo anagrafica da BC...")
    items = get_anagrafica(token)
    print(f"  {len(items)} items trovati")
    products = []
    for item in items:
        code = str(item.get("AltICMIFB_Item") or "").strip()
        nHK  = str(item.get("No") or "").strip()
        if not code and not nHK:
            continue
        blocked = item.get("Blocked") is True
        products.append({
            "id":          code or nHK,
            "code":        code,
            "nHK":         nHK,
            "description": str(item.get("Description") or "").strip(),
            "category":    norm(item.get("AltICMSection_Description"), CAT_MAP),
            "uom":         norm(item.get("Sales_Unit_of_Measure"), UOM_MAP),
            "qtyPerBox":   float(item.get("AltICMQuantity_x_Packaging") or 0),
            "boxPerPallet":float(item.get("AltICMPackaging_x_Pallet") or 0),
            "kgPerBox":    float(item.get("AltICMNet_Weight") or 0),
            "kgxplt":      float(item.get("AltICMKg_x_PLT") or 0),
            "temperature":      norm(item.get("AltICMProduct_Type"), TEMP_MAP),
            "bcTransportation": str(item.get("Transportation") or "").strip().upper(),
            "standardCostHkd": float(item.get("Standard_Cost") or 0),
            "isHoff":           bool(item.get("AltICMHOFF")),
            "active":      not blocked,
            "vendorName":  str(item.get("AltICMVendor_Name") or "").strip(),
            "vendorName2": str(item.get("AltICMVendor_Name_2") or "").strip(),
            "aiem":        0,
        })

    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    OUT_PATH.write_text(json.dumps(products, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Scritto {len(products)} prodotti in {OUT_PATH}")
