"""
Aggiorna docs/data/hk_air.json — tutti gli articoli AIR (Transportation contiene 'air')
non bloccati da BC BrightView HK, senza filtro vendor.
Usato da GitHub Actions. CLIENT_SECRET letto da variabile d'ambiente BC_CLIENT_SECRET.
"""
import os, json, requests
from pathlib import Path

TENANT_ID     = "2acd007b-8d3f-4be0-9681-cf248264a0e2"
CLIENT_ID     = "925de6e4-e71f-4c24-9e0a-f3ae544ae644"
CLIENT_SECRET = os.environ.get("BC_CLIENT_SECRET", "")
BC_ENV        = "Production_HK"
COMPANY       = "BRIGHT%20VIEW%20TRADING%20HK%20LIMITED"
BASE          = f"https://api.businesscentral.dynamics.com/v2.0/{TENANT_ID}/{BC_ENV}/ODataV4/Company('{COMPANY}')"
OUT_PATH      = Path(__file__).parent.parent / "docs" / "data" / "hk_air.json"


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


if __name__ == "__main__":
    if not CLIENT_SECRET:
        raise RuntimeError("BC_CLIENT_SECRET non impostato")

    print("Ottengo token BC...")
    token = get_token()

    # Niente $select: Transportation non viene restituito con $select attivo (nota da update_hk_anagrafica.py)
    print("Leggo tutti gli articoli DS da BC HK (senza filtro vendor)...")
    items = bc_get(token, "Item_Card_Excel?$filter=Gen_Prod_Posting_Group eq 'DS'&$top=10000")
    print(f"  {len(items)} items totali")

    air_items = []
    for item in items:
        if item.get("Blocked") is True:
            continue
        transport = str(item.get("Transportation") or "").strip().upper()
        if "air" not in transport.lower():
            continue
        code = str(item.get("AltICMIFB_Item") or "").strip()
        nHK  = str(item.get("No") or "").strip()
        air_items.append({
            "code":             code,
            "nHK":              nHK,
            "description":      str(item.get("Description") or "").strip(),
            "bcTransportation": transport,
            "vendorName":       str(item.get("AltICMVendor_Name") or "").strip(),
        })

    print(f"  {len(air_items)} articoli AIR attivi (tutti i vendor)")
    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    OUT_PATH.write_text(json.dumps(air_items, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Scritto {len(air_items)} articoli in {OUT_PATH}")
