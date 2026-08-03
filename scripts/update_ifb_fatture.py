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

# Clienti per branch (solo IFB Italia → filiali)
# MAC escluso: le fatture MAC vengono da BC BrightView (update_mac_fatture.py)
CUSTOMERS = {
    "HK":  "40000854",  # BRIGHT VIEW TRADING HK LIMITED
    "CAN": "40000175",  # COMERCIAL ITALIANA DE ALIMENTACION S.
    "AUS": "40000312",  # FRESCO GOURMET PTY LTD
}


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


def get_fatture_for_customer(token, customer_no, branch):
    fields = ",".join([
        "no", "description", "postingdate",
        "quantity", "unitprice", "amount", "locationcode",
        "documentno", "billtocustomerno", "selltocustomerno",
    ])
    def make_endpoint(field, value):
        return (
            f"IFB_Sales_Invoice_Line"
            f"?$filter=postingdate ge {DATE_FROM}"
            f" and type eq 'Item'"
            f" and {field} eq '{value}'"
            f"&$select={fields}&$top=50000"
        )
    print(f"  Fetching {branch} (customer {customer_no})...")
    rows_bill = bc_get(token, make_endpoint("billtocustomerno", customer_no))
    rows_sell = bc_get(token, make_endpoint("selltocustomerno", customer_no))
    # Dedup per (documentno, no) — priorità a billtocustomerno
    seen = {(r.get("documentno",""), r.get("no","")) for r in rows_bill}
    rows = rows_bill + [r for r in rows_sell if (r.get("documentno",""), r.get("no","")) not in seen]
    print(f"    bill-to: {len(rows_bill)}, sell-to extra: {len(rows)-len(rows_bill)}, totale: {len(rows)}")
    print(f"    {len(rows)} righe trovate")
    return rows, branch


if __name__ == "__main__":
    if not CLIENT_SECRET:
        raise RuntimeError("BC_CLIENT_SECRET non impostato")

    print("Ottengo token BC...")
    token = get_token()

    print(f"Leggo fatture da BC Italia (da {DATE_FROM})...")
    rows = []
    for branch, customer_no in CUSTOMERS.items():
        items, br = get_fatture_for_customer(token, customer_no, branch)
        for item in items:
            no = str(item.get("no") or "").strip()
            if not no:
                continue
            rows.append({
                "No_":               no,
                "Description":       str(item.get("description") or "").strip(),
                "Vendor Name":       "",
                "Last Posting Date": str(item.get("postingdate") or ""),
                "Quantity":          float(item.get("quantity") or 0),
                "Price":             float(item.get("amount") or 0) / float(item.get("quantity") or 1) if float(item.get("quantity") or 0) != 0 else 0,
                "Location Code":     str(item.get("locationcode") or "").strip(),
                "Section Description": "",
                "Branch":            br,
            })

    print(f"Totale {len(rows)} righe valide")
    if rows:
        print(f"  Esempio HK:  {next((r for r in rows if r['Branch']=='HK'),  None)}")
        print(f"  Esempio CAN: {next((r for r in rows if r['Branch']=='CAN'), None)}")
        print(f"  Esempio AUS: {next((r for r in rows if r['Branch']=='AUS'), None)}")

    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    OUT_PATH.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Scritto {len(rows)} righe in {OUT_PATH}")
