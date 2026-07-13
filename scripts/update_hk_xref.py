"""
Aggiorna docs/data/hk_xref.json con le XRef HK→IFB da Business Central.
Usato da GitHub Actions. CLIENT_SECRET letto da variabile d'ambiente BC_CLIENT_SECRET.
"""
import os, json, requests
from pathlib import Path

TENANT_ID    = "2acd007b-8d3f-4be0-9681-cf248264a0e2"
CLIENT_ID    = "925de6e4-e71f-4c24-9e0a-f3ae544ae644"
CLIENT_SECRET = os.environ.get("BC_CLIENT_SECRET", "")
BC_ENV       = "Production_HK"
BC_COMPANY   = "BRIGHT%20VIEW%20TRADING%20HK%20LIMITED"
OUT_PATH     = Path(__file__).parent.parent / "docs" / "data" / "hk_xref.json"


def get_token():
    r = requests.post(
        f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token",
        data={
            "grant_type": "client_credentials",
            "client_id": CLIENT_ID,
            "client_secret": CLIENT_SECRET,
            "scope": "https://api.businesscentral.dynamics.com/.default",
        },
    )
    r.raise_for_status()
    return r.json()["access_token"]


def get_xref_data(token):
    url = (
        f"https://api.businesscentral.dynamics.com/v2.0/{TENANT_ID}/{BC_ENV}"
        f"/ODataV4/Company('{BC_COMPANY}')/Item_Card_Excel"
        f"?$top=5000"
    )
    results = []
    while url:
        r = requests.get(url, headers={"Authorization": f"Bearer {token}", "Accept": "application/json"})
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

    print("Leggo XRef da BC...")
    records = get_xref_data(token)
    print(f"  {len(records)} record trovati")

    # Converti nel formato atteso dall'app: [{nHK, ifbNo}]
    xrefs = [
        {"nHK": item["No"], "ifbNo": item["AltICMIFB_Item"]}
        for item in records
        if item.get("AltICMIFB_Item")  # escludi righe senza IFB code
    ]
    print(f"  {len(xrefs)} XRef valide (con IFB code)")

    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    OUT_PATH.write_text(json.dumps(xrefs, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Scritto: {OUT_PATH}")
