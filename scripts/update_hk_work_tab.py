"""
Legge 08_Work_Tab.xlsx (HK) e aggiorna docs/data/hk_work_tab.json.
Eseguire localmente dopo ogni aggiornamento del file Excel.

Output: { vendor_name: plt_cost_eur } per tutti i fornitori HK con flag X
(= fornitori che gestiscono il trasporto con costo/pallet definito).
La logica si applica a TUTTI gli articoli di quel fornitore, inclusi i futuri.
"""
import sys, json, openpyxl
from pathlib import Path

DEFAULT_PATH = (
    Path.home()
    / "OneDrive - Inalca spa"
    / "Data Analysis - 068 STANDARD COST HONG KONG"
    / "STANDARD_COST_AUTOMATION"
    / "INPUT"
    / "08_Work_Tab.xlsx"
)
OUT_PATH = Path(__file__).parent.parent / "docs" / "data" / "hk_work_tab.json"


def extract(xlsx_path: Path) -> dict:
    wb = openpyxl.load_workbook(str(xlsx_path), read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    vendor_costs = {}
    conflicts = []
    for r in rows[2:]:
        nIFB   = str(r[1] or "").strip()
        vendor = str(r[3] or "").strip()
        fca_x  = str(r[10] or "").strip()
        plt_c  = r[11]
        if fca_x != "X" or not vendor or not nIFB or nIFB == "#N/A":
            continue
        try:
            plt_cost = float(plt_c)
            if plt_cost <= 0:
                continue
            if vendor in vendor_costs and vendor_costs[vendor] != plt_cost:
                conflicts.append(f"{vendor}: {vendor_costs[vendor]} vs {plt_cost}")
            elif vendor not in vendor_costs:
                vendor_costs[vendor] = plt_cost
        except (TypeError, ValueError):
            pass
    if conflicts:
        print(f"  ATTENZIONE: costi pallet diversi per stesso fornitore:")
        for c in conflicts:
            print(f"    {c}")
    print(f"  {len(vendor_costs)} fornitori con costo pallet definito")
    return vendor_costs


if __name__ == "__main__":
    xlsx_path = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_PATH
    if not xlsx_path.exists():
        print(f"File non trovato: {xlsx_path}")
        sys.exit(1)
    print(f"Leggo {xlsx_path.name}...")
    data = extract(xlsx_path)
    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    OUT_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Scritto {len(data)} fornitori in {OUT_PATH}")
