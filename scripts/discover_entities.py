"""
Lista tutte le entità OData BC e filtra per parole chiave logistiche.
"""
import os, requests

TENANT_ID     = "2acd007b-8d3f-4be0-9681-cf248264a0e2"
CLIENT_ID     = "925de6e4-e71f-4c24-9e0a-f3ae544ae644"
CLIENT_SECRET = os.environ.get("BC_CLIENT_SECRET", "")
BC_ENV        = "Production"
BC_COMPANY    = "Inalca%20Food%20%26%20Beverage%20s.r.l."

def get_token():
    r = requests.post(
        f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token",
        data={"grant_type":"client_credentials","client_id":CLIENT_ID,
              "client_secret":CLIENT_SECRET,"scope":"https://api.businesscentral.dynamics.com/.default"})
    r.raise_for_status()
    return r.json()["access_token"]

if __name__ == "__main__":
    token = get_token()
    base         = f"https://api.businesscentral.dynamics.com/v2.0/{TENANT_ID}/{BC_ENV}/ODataV4/Company('{BC_COMPANY}')/"
    service_root = f"https://api.businesscentral.dynamics.com/v2.0/{TENANT_ID}/{BC_ENV}/ODataV4/"

    # Il service document root (senza Company) lista tutti gli EntitySet
    r = requests.get(service_root, headers={"Authorization": f"Bearer {token}", "Accept": "application/json"})
    print(f"Status: {r.status_code}")
    print(f"Response (primi 2000 chars): {r.text[:2000]}")
    try:
        data = r.json()
    except Exception as e:
        print(f"Errore parsing JSON: {e}")
        data = {}
    entities = [e["name"] for e in data.get("value", []) if "name" in e]
    print(f"Totale entità: {len(entities)}\n")

    keywords = ["carriage","freight","logistic","cost","transport","nolo","spese","shipping","collo","dap","fca"]
    print("=== Entità con keyword logistiche/prezzi ===")
    for e in sorted(entities):
        el = e.lower()
        if any(k in el for k in keywords):
            print(f"  {e}")

    print("\n=== Tutte le entità IFB_ ===")
    for e in sorted(entities):
        if e.upper().startswith("IFB_"):
            print(f"  {e}")

    print("\n=== Tutte le entità (ordinate) ===")
    for e in sorted(entities):
        print(f"  {e}")
