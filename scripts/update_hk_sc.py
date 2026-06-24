"""
Aggiorna docs/data/hk_sc.json con i costi attuali HK da Business Central.
Usato da GitHub Actions. CLIENT_SECRET letto da variabile d'ambiente BC_CLIENT_SECRET.
"""
import os, json, requests, pandas as pd
from datetime import datetime, timedelta, timezone
from pathlib import Path

TENANT_ID     = "2acd007b-8d3f-4be0-9681-cf248264a0e2"
CLIENT_ID     = "925de6e4-e71f-4c24-9e0a-f3ae544ae644"
CLIENT_SECRET = os.environ.get("BC_CLIENT_SECRET", "")
BC_ENV        = "Production_HK"
COMPANY       = "BRIGHT%20VIEW%20TRADING%20HK%20LIMITED"
BASE          = f"https://api.businesscentral.dynamics.com/v2.0/{TENANT_ID}/{BC_ENV}/ODataV4/Company('{COMPANY}')"
OUT_PATH      = Path(__file__).parent.parent / "docs" / "data" / "hk_sc.json"


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
    results = []
    url = f"{BASE}/{endpoint}"
    while url:
        r = requests.get(url, headers=headers)
        r.raise_for_status()
        data = r.json()
        results.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return results


def get_items(token):
    vendor = "INALCA FOOD %26 BEVERAGE SRL"
    endpoint = (
        f"Item_Card_Excel"
        f"?$filter=Gen_Prod_Posting_Group eq 'DS' and AltICMVendor_Name eq '{vendor}'"
        f"&$select=No,Description,Unit_Cost,Standard_Cost,Inventory"
        f"&$top=5000"
    )
    return bc_get(token, endpoint)


def get_last_purchase_dates(token):
    endpoint = (
        "Item_Ledger_Entries_Excel"
        "?$filter=Entry_Type eq 'Purchase'"
        "&$select=Item_No,Posting_Date"
        "&$top=50000"
    )
    entries = bc_get(token, endpoint)
    if not entries:
        return {}
    df = pd.DataFrame(entries)
    df["Posting_Date"] = pd.to_datetime(df["Posting_Date"])
    return df.groupby("Item_No")["Posting_Date"].max().dt.strftime("%Y-%m-%d").to_dict()


def get_sales_last_3_months(token):
    date_from = (datetime.now(timezone.utc) - timedelta(days=90)).strftime("%Y-%m-%d")
    endpoint = (
        f"Item_Ledger_Entries_Excel"
        f"?$filter=Entry_Type eq 'Sale' and Posting_Date ge {date_from}"
        f"&$select=Item_No,Quantity"
        f"&$top=50000"
    )
    entries = bc_get(token, endpoint)
    if not entries:
        return {}
    df = pd.DataFrame(entries)
    return df.groupby("Item_No")["Quantity"].sum().to_dict()


if __name__ == "__main__":
    if not CLIENT_SECRET:
        raise RuntimeError("BC_CLIENT_SECRET non impostato")

    print("Ottengo token BC...")
    token = get_token()

    print("Leggo items da BC...")
    items = get_items(token)
    print(f"  {len(items)} items trovati")

    print("Leggo Last Purchase Date...")
    last_purchase = get_last_purchase_dates(token)

    print("Leggo Sales ultimi 3 mesi...")
    sales_3m = get_sales_last_3_months(token)

    # Converti nel formato atteso dall'app
    rows = []
    for item in items:
        no = item["No"]
        rows.append({
            "code":            no,
            "description":     item.get("Description") or "",
            "lastSC":          round(float(item.get("Standard_Cost") or 0), 5),
            "fifoUnit":        round(float(item.get("Unit_Cost") or 0), 5),
            "salesLast3m":     round(float(sales_3m.get(no) or 0), 5),
            "lastPurchaseDate": last_purchase.get(no) or "",
            "stockQty":        round(float(item.get("Inventory") or 0), 3),
        })

    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    OUT_PATH.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Scritto {len(rows)} righe in {OUT_PATH}")
