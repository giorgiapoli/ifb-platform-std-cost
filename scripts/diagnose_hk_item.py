"""
Stampa tutti i campi BC Brightview per un item specifico (default: CPS03 / CF5004-FG).
Usare per trovare il campo corretto che contiene il valore di transportation/AIR.

Uso:
  BC_CLIENT_SECRET=xxx python scripts/diagnose_hk_item.py
  BC_CLIENT_SECRET=xxx python scripts/diagnose_hk_item.py CF5004-FG
"""
import os, sys, json, requests

TENANT_ID     = "2acd007b-8d3f-4be0-9681-cf248264a0e2"
CLIENT_ID     = "925de6e4-e71f-4c24-9e0a-f3ae544ae644"
CLIENT_SECRET = os.environ.get("BC_CLIENT_SECRET", "")
BC_ENV        = "Production_HK"
COMPANY       = "BRIGHT%20VIEW%20TRADING%20HK%20LIMITED"
BASE          = f"https://api.businesscentral.dynamics.com/v2.0/{TENANT_ID}/{BC_ENV}/ODataV4/Company('{COMPANY}')"

NHK = sys.argv[1] if len(sys.argv) > 1 else "CPS03"

def get_token():
    r = requests.post(
        f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token",
        data={"grant_type": "client_credentials", "client_id": CLIENT_ID,
              "client_secret": CLIENT_SECRET,
              "scope": "https://api.businesscentral.dynamics.com/.default"},
    )
    r.raise_for_status()
    return r.json()["access_token"]

if __name__ == "__main__":
    if not CLIENT_SECRET:
        raise RuntimeError("BC_CLIENT_SECRET non impostato")

    token = get_token()
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}

    # Prova 1: Item_Card_Excel (endpoint attuale)
    print(f"\n=== Item_Card_Excel per No='{NHK}' ===")
    url = f"{BASE}/Item_Card_Excel?$filter=No eq '{NHK}'&$top=1"
    r = requests.get(url, headers=headers)
    if r.ok:
        items = r.json().get("value", [])
        if items:
            item = items[0]
            # Mostra tutti i campi non vuoti
            print("Campi non vuoti:")
            for k, v in sorted(item.items()):
                if v not in (None, "", 0, False):
                    print(f"  {k}: {v!r}")
            print("\nCampi con 'air' o 'transport' nel nome o valore:")
            for k, v in sorted(item.items()):
                kl = k.lower(); vl = str(v).lower()
                if "air" in kl or "transport" in kl or "air" in vl or "freight" in vl:
                    print(f"  {k}: {v!r}")
        else:
            print("  Nessun risultato")
    else:
        print(f"  Errore: {r.status_code} {r.text[:200]}")

    # Prova 2: Item endpoint standard
    print(f"\n=== Item endpoint standard per No='{NHK}' ===")
    url2 = f"{BASE}/Item?$filter=No eq '{NHK}'&$top=1"
    r2 = requests.get(url2, headers=headers)
    if r2.ok:
        items2 = r2.json().get("value", [])
        if items2:
            item2 = items2[0]
            print("Campi con 'air' o 'transport' nel nome o valore:")
            for k, v in sorted(item2.items()):
                kl = k.lower(); vl = str(v).lower()
                if "air" in kl or "transport" in kl or "air" in vl or "freight" in vl:
                    print(f"  {k}: {v!r}")
            if not any("air" in k.lower() or "transport" in k.lower() or "air" in str(v).lower() for k, v in item2.items()):
                print("  (nessun campo trovato — tutti i campi non vuoti:)")
                for k, v in sorted(item2.items()):
                    if v not in (None, "", 0, False):
                        print(f"    {k}: {v!r}")
        else:
            print("  Nessun risultato")
    else:
        print(f"  Errore: {r2.status_code} {r2.text[:200]}")
