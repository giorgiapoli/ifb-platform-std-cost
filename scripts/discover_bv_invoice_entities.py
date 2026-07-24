"""
Scopre il nome corretto dell'entity Sales Invoice Line in BC BrightView (Production_HK).
Esegui localmente con: BC_CLIENT_SECRET=xxx python scripts/discover_bv_invoice_entities.py
"""
import os, requests

TENANT_ID     = "2acd007b-8d3f-4be0-9681-cf248264a0e2"
CLIENT_ID     = "925de6e4-e71f-4c24-9e0a-f3ae544ae644"
CLIENT_SECRET = os.environ.get("BC_CLIENT_SECRET", "")
BC_ENV        = "Production_HK"
COMPANY       = "BRIGHT%20VIEW%20TRADING%20HK%20LIMITED"
BASE          = f"https://api.businesscentral.dynamics.com/v2.0/{TENANT_ID}/{BC_ENV}/ODataV4/Company('{COMPANY}')"

def get_token():
    r = requests.post(
        f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token",
        data={"grant_type": "client_credentials", "client_id": CLIENT_ID,
              "client_secret": CLIENT_SECRET,
              "scope": "https://api.businesscentral.dynamics.com/.default"})
    r.raise_for_status()
    return r.json()["access_token"]

if __name__ == "__main__":
    if not CLIENT_SECRET:
        raise RuntimeError("Imposta BC_CLIENT_SECRET")

    token = get_token()
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}

    # Lista tutte le entity del service root
    svc = f"https://api.businesscentral.dynamics.com/v2.0/{TENANT_ID}/{BC_ENV}/ODataV4/"
    r = requests.get(svc, headers=headers)
    data = r.json()
    entities = [e["name"] for e in data.get("value", []) if "name" in e]
    print(f"Totale entity in {BC_ENV}: {len(entities)}\n")

    # Filtra entity con parole chiave invoice/sales/fattura
    keywords = ["invoice", "sales", "fattura", "posted", "ledger", "entry"]
    print("=== Entity con keyword invoice/sales ===")
    for e in sorted(entities):
        if any(k in e.lower() for k in keywords):
            print(f"  {e}")

    # Entità promettenti: mostra TUTTI i campi
    print("\n=== Dettaglio campi entità promettenti ===")
    detail_candidates = [
        "IFB_Item_Ledger_Entry",
        "Posted_Sales_Invoice_ExcelSalesInvLines",
        "Item_Ledger_Entries_Excel",
        "IFB_Customer_Ledger_Entries",
        "IFB_Value_Entry",
        "IFB_Sales_Invoice_Header",
    ]
    for entity in detail_candidates:
        url = f"{BASE}/{entity}?$top=2"
        r2 = requests.get(url, headers=headers)
        if r2.ok:
            rows = r2.json().get("value", [])
            print(f"\n  ✓ {entity} ({len(rows)} righe)")
            if rows:
                clean = {k: v for k, v in rows[0].items() if not k.startswith("@")}
                print(f"     Campi: {list(clean.keys())}")
                print(f"     Esempio: {clean}")
        else:
            print(f"\n  ✗ {entity} — {r2.status_code}")

    # Test vecchi candidati
    print("\n=== Test entity candidate (vecchi) ===")
    candidates = [
        "IFB_Invoice_Line", "IFB_Invoice_Lines", "IFB_SalesInvoiceLine",
        "IFB_Sales_Invoice_Line", "IFB_Posted_Invoice_Line",
        "Sales_Invoice_Line", "BV_Sales_Invoice_Line",
        "Posted_Sales_Invoice_Line", "Posted_Sales_Invoice_Lines",
        "SalesInvoiceLine", "salesInvoiceLines",
        "Item_Ledger_Entry", "Item_Ledger_Entries_Excel",
        "Value_Entry", "Cust__Ledger_Entry",
    ]
    for entity in candidates:
        url = f"{BASE}/{entity}?$top=1"
        r2 = requests.get(url, headers=headers)
        if r2.ok:
            rows = r2.json().get("value", [])
            fields = list(rows[0].keys())[:8] if rows else []
            print(f"  ✓ {entity} — {len(rows)} righe — campi: {fields}")
        else:
            print(f"  ✗ {entity} — {r2.status_code}")
