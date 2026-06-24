"""
Aggiorna docs/data/ifb_fatture.json con le ultime fatture di vendita da BC Italia.
Sostituisce il flusso Power Automate → Excel → import manuale.
"""
import os, json, requests
from pathlib import Path
from datetime import datetime, timedelta

TENANT_ID     = "2acd007b-8d3f-4be0-9681-cf248264a0e2"
CLIENT_ID     = "925de6e4-e71f-4c24-9e0a-f3ae544ae644"
CLIENT_SECRET = os.environ.get("BC_CLIENT_SECRET", "")
BC_ENV        = "Production"
BC_COMPANY    = "Inalca%20Food%20%26%20Beverage%20s.r.l."
OUT_PATH      = Path(__file__).parent.parent / "docs" / "data" / "ifb_fatture.json"

# Ultimi 12 mesi
DATE_FROM = (datetime.today() - timedelta(days=365)).strftime("%Y-%m-%d")


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


def bc_get(token, endpoint):
    base = f"https://api.businesscentral.dynamics.com/v2.0/{TENANT_ID}/{BC_ENV}/ODataV4/Company('{BC_COMPANY}')/"
    url = base + endpoint
    results = []
    while url:
        r = requests.get(url, headers={"Authorization": f"Bearer {token}", "Accept": "application/json"})
        r.raise_for_status()
        data = r.json()
        results.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return results


def get_fatture(token):
    # Nessun filtro: scopriamo prima i campi disponibili
    endpoint = "IFB_Sales_Invoice_Line?$top=10"
    return bc_get(token, endpoint)


if __name__ == "__main__":
    if not CLIENT_SECRET:
        raise RuntimeError("BC_CLIENT_SECRET non impostato")

    print("Ottengo token BC...")
    token = get_token()

    print(f"Leggo fatture da BC Italia (da {DATE_FROM})...")
    items = get_fatture(token)
    print(f"  {len(items)} righe trovate")

    if items:
        print("  Campi disponibili:")
        for k in sorted(k for k in items[0].keys() if not k.startswith("@")):
            print(f"    {k} = {repr(items[0][k])[:80]}")

    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    OUT_PATH.write_text(json.dumps(items, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Scritto {len(items)} righe in {OUT_PATH}")
