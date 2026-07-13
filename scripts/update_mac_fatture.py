"""
Aggiorna docs/data/mac_fatture.json con le fatture BrightView → MACAO da BC.
Company: BRIGHT VIEW TRADING HK LIMITED (Production_HK).
Clienti Macao: 40001358 (THE HOUSE OF FINE FOODS M...) e 40001359 (THE HOUSE OF FINE FOODS LTD).
Usato da GitHub Actions. CLIENT_SECRET letto da variabile d'ambiente BC_CLIENT_SECRET.
"""
import os, json, requests
from pathlib import Path
from datetime import datetime, timedelta

TENANT_ID     = "2acd007b-8d3f-4be0-9681-cf248264a0e2"
CLIENT_ID     = "925de6e4-e71f-4c24-9e0a-f3ae544ae644"
CLIENT_SECRET = os.environ.get("BC_CLIENT_SECRET", "")
BC_ENV        = "Production_HK"
COMPANY       = "BRIGHT%20VIEW%20TRADING%20HK%20LIMITED"
BASE          = f"https://api.businesscentral.dynamics.com/v2.0/{TENANT_ID}/{BC_ENV}/ODataV4/Company('{COMPANY}')"
OUT_PATH      = Path(__file__).parent.parent / "docs" / "data" / "mac_fatture.json"

MAC_CUSTOMERS = ["40001358", "40001359"]
DATE_FROM     = (datetime.today() - timedelta(days=365)).strftime("%Y-%m-%d")


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


def get_fatture(token, customer_no):
    """Prova vari nomi entity per le righe fattura BV."""
    # Candidati in ordine di probabilità (BV potrebbe avere nomi diversi da IFB)
    for entity in ["IFB_Sales_Invoice_Line", "BV_Sales_Invoice_Line", "Sales_Invoice_Line"]:
        try:
            print(f"  Provo entity '{entity}' per customer {customer_no}...")
            rows = bc_get(token, (
                f"{entity}"
                f"?$filter=postingdate ge {DATE_FROM}"
                f" and type eq 'Item'"
                f" and billtocustomerno eq '{customer_no}'"
                f"&$top=50000"
            ))
            print(f"    Trovate {len(rows)} righe bill-to")
            # Prova anche sell-to
            rows2 = bc_get(token, (
                f"{entity}"
                f"?$filter=postingdate ge {DATE_FROM}"
                f" and type eq 'Item'"
                f" and selltocustomerno eq '{customer_no}'"
                f"&$top=50000"
            ))
            # Dedup
            seen = {(r.get("documentno",""), r.get("no","")) for r in rows}
            extra = [r for r in rows2 if (r.get("documentno",""), r.get("no","")) not in seen]
            print(f"    sell-to extra: {len(extra)}, totale: {len(rows)+len(extra)}")
            return rows + extra, entity
        except Exception as e:
            print(f"    '{entity}' non disponibile: {e}")
    return [], None


if __name__ == "__main__":
    if not CLIENT_SECRET:
        raise RuntimeError("BC_CLIENT_SECRET non impostato")

    print("Ottengo token BC...")
    token = get_token()

    print(f"Leggo fatture BV→MAC da BC ({BC_ENV}, da {DATE_FROM})...")
    all_rows = []
    seen_keys = set()

    for customer_no in MAC_CUSTOMERS:
        rows, entity = get_fatture(token, customer_no)
        if not rows:
            print(f"  ATTENZIONE: nessuna fattura trovata per customer {customer_no}.")
            print("  Suggerimento: esegui discover_entities.py (COMPANY=BRIGHT VIEW TRADING HK LIMITED) per trovare il nome entity corretto.")
            continue

        for item in rows:
            no  = str(item.get("no") or "").strip()
            doc = str(item.get("documentno") or "").strip()
            key = (doc, no)
            if not no or key in seen_keys:
                continue
            seen_keys.add(key)
            qty = float(item.get("quantity") or 0)
            amt = float(item.get("amount") or 0)
            all_rows.append({
                "No_":               no,
                "Description":       str(item.get("description") or "").strip(),
                "Vendor Name":       "BRIGHT VIEW TRADING HK LTD",
                "Last Posting Date": str(item.get("postingdate") or ""),
                "Quantity":          qty,
                "Price":             amt / qty if qty != 0 else 0,
                "Location Code":     str(item.get("locationcode") or "").strip(),
                "Section Description": str(item.get("sectiondescription") or "").strip(),
                "Document No":       doc,
                "Customer No":       customer_no,
                "Branch":            "MAC",
            })

    print(f"Totale {len(all_rows)} righe fattura MAC")
    if all_rows:
        print(f"  Esempio: {json.dumps(all_rows[0], ensure_ascii=False)}")

    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    OUT_PATH.write_text(json.dumps(all_rows, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Scritto {len(all_rows)} righe in {OUT_PATH}")
