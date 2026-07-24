"""
Aggiorna docs/data/mac_fatture.json con le fatture BrightView → MACAO da BC.
Company: BRIGHT VIEW TRADING HK LIMITED (Production_HK).
Clienti Macao: CUST-00715, CUST-02012.

Flusso:
1. IFB_Sales_Invoice_Header  → filtra per customer MAC → ottieni Document_No + postingdate
2. Posted_Sales_Invoice_ExcelSalesInvLines → filtra per Document_No → righe articolo
   Campo AltICMOld_Item_No = codice IFB (ex-IFB_Invoice_Line.no)
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

MAC_CUSTOMERS = ["CUST-00715", "CUST-02012"]
DATE_FROM     = (datetime.today() - timedelta(days=365)).strftime("%Y-%m-%d")
BATCH_SIZE    = 30  # max Document_No per chiamata OData


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


def get_mac_invoice_headers(token, customer_no):
    """Ritorna dict {doc_no: postingdate} per il customer MAC dato."""
    rows = bc_get(token, (
        f"IFB_Sales_Invoice_Header"
        f"?$filter=postingdate ge {DATE_FROM}"
        f" and (billtocustomerno eq '{customer_no}' or selltocustomerno eq '{customer_no}')"
        f"&$select=no,postingdate,billtocustomerno,selltocustomerno"
        f"&$top=5000"
    ))
    result = {r["no"]: r["postingdate"] for r in rows if r.get("no")}
    print(f"    {len(result)} fatture header trovate per {customer_no}")
    return result


def get_lines_for_docs(token, doc_nos_with_dates):
    """Fetch righe articolo per i document_no dati, in batch."""
    doc_list = list(doc_nos_with_dates.items())
    all_lines = []
    for i in range(0, len(doc_list), BATCH_SIZE):
        batch = doc_list[i:i+BATCH_SIZE]
        doc_filter = " or ".join(f"Document_No eq '{d}'" for d, _ in batch)
        rows = bc_get(token, (
            f"Posted_Sales_Invoice_ExcelSalesInvLines"
            f"?$filter=({doc_filter}) and Type eq 'Item'"
            f"&$top=10000"
        ))
        # Aggiunge postingdate dall'header
        for r in rows:
            r["_postingdate"] = doc_nos_with_dates.get(r.get("Document_No", ""), "")
        all_lines.extend(rows)
    return all_lines


if __name__ == "__main__":
    if not CLIENT_SECRET:
        raise RuntimeError("BC_CLIENT_SECRET non impostato")

    print("Ottengo token BC...")
    token = get_token()

    print(f"Leggo fatture BV→MAC da BC ({BC_ENV}, da {DATE_FROM})...")
    all_rows = []
    seen_keys = set()

    for customer_no in MAC_CUSTOMERS:
        print(f"\n  Customer: {customer_no}")
        doc_map = get_mac_invoice_headers(token, customer_no)
        if not doc_map:
            print(f"  ATTENZIONE: nessuna fattura trovata per {customer_no}.")
            continue

        lines = get_lines_for_docs(token, doc_map)
        print(f"    {len(lines)} righe articolo totali")

        for item in lines:
            bv_code  = str(item.get("No") or "").strip()
            ifb_code = str(item.get("AltICMOld_Item_No") or bv_code).strip()
            doc      = str(item.get("Document_No") or "").strip()
            key      = (doc, bv_code)
            if not bv_code or key in seen_keys:
                continue
            seen_keys.add(key)
            qty      = float(item.get("Quantity") or 0)
            price    = float(item.get("Unit_Price") or 0)
            all_rows.append({
                "No_":               ifb_code,         # codice IFB (AltICMOld_Item_No)
                "BV_No":             bv_code,           # codice BrightView
                "Description":       str(item.get("Description") or "").strip(),
                "Vendor Name":       "BRIGHT VIEW TRADING HK LTD",
                "Last Posting Date": item.get("_postingdate", ""),
                "Quantity":          qty,
                "Price":             price,
                "Location Code":     str(item.get("Location_Code") or "").strip(),
                "Section Description": "",
                "Document No":       doc,
                "Customer No":       customer_no,
                "Branch":            "MAC",
            })

    print(f"\nTotale {len(all_rows)} righe fattura MAC")
    if all_rows:
        print(f"  Esempio: {json.dumps(all_rows[0], ensure_ascii=False)}")

    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    OUT_PATH.write_text(json.dumps(all_rows, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Scritto {len(all_rows)} righe in {OUT_PATH}")
