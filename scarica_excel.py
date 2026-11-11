"""
Pulsante «Excel» sotto ogni tabella: scarica la tabella che si vede (stesse righe, stessi filtri) in un file Excel, con i
colori della tabella, l'intestazione fissa e i filtri sulle colonne. Il file si crea solo quando si clicca (non rallenta la pagina).
"""
import io
import re

import pandas as pd
import streamlit as st

_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
_contatore = [0]


def reset():
    """Da chiamare a ogni esecuzione della pagina: i pulsanti ricevono un nome progressivo (serve a Streamlit per distinguerli)."""
    _contatore[0] = 0


def _come_dataframe(dati):
    """(DataFrame, Styler o None) da qualunque cosa passata a una tabella."""
    if hasattr(dati, "to_excel") and hasattr(dati, "data") and not isinstance(dati, pd.DataFrame):      # Styler
        return dati.data, dati
    if isinstance(dati, pd.DataFrame):
        return dati, None
    if isinstance(dati, pd.Series):
        return dati.to_frame(), None
    return pd.DataFrame(dati), None


def xlsx_bytes(dati, indice=None):
    """Excel di una tabella (DataFrame o Styler). indice: True/False; se None si include solo se non e' un numero progressivo."""
    df, sty = _come_dataframe(dati)
    if indice is None:
        indice = not isinstance(df.index, pd.RangeIndex) and not df.index.equals(pd.RangeIndex(len(df)))
    buf = io.BytesIO()
    scritto = False
    if sty is not None:
        try:
            with pd.ExcelWriter(buf, engine="openpyxl") as xw:
                sty.to_excel(xw, sheet_name="Tabella", index=indice)
            scritto = True
        except Exception:
            buf = io.BytesIO()
    if not scritto:
        with pd.ExcelWriter(buf, engine="openpyxl") as xw:
            df.to_excel(xw, sheet_name="Tabella", index=indice)
    # rifinitura: intestazione in grassetto e fissa, filtri, larghezze
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment, Font
    from openpyxl.utils import get_column_letter
    buf.seek(0)
    wb = load_workbook(buf)
    ws = wb.active
    for c in ws[1]:
        c.font = Font(bold=True)
        c.alignment = Alignment(wrap_text=True, vertical="bottom")
    ws.freeze_panes = "A2"
    if ws.max_row > 1 and ws.max_column >= 1:
        ws.auto_filter.ref = ws.dimensions
    for j in range(1, ws.max_column + 1):
        lung = max((len(str(c.value)) for c in ws[get_column_letter(j)][:200] if c.value is not None), default=8)
        ws.column_dimensions[get_column_letter(j)].width = min(max(lung + 2, 9), 60)
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def _nome_file(dati):
    df, _ = _come_dataframe(dati)
    cols = [str(c) for c in list(df.columns)[:2]] or ["tabella"]
    return "tabella_" + re.sub(r"[^\w\-]+", "_", "-".join(cols)).strip("_")[:40] + ".xlsx"


def bottone(dati, contenitore=None, indice=None, nome=None):
    """Mette sotto la tabella il pulsante «Excel». contenitore = st o una colonna (dove si e' messa la tabella)."""
    try:
        if dati is None or len(_come_dataframe(dati)[0]) == 0:
            return
    except Exception:
        return
    _contatore[0] += 1
    dg = contenitore if contenitore is not None else st
    dg.download_button("Excel", data=lambda: xlsx_bytes(dati, indice), file_name=nome or _nome_file(dati), mime=_MIME,
                       key=f"xl_{_contatore[0]}", type="tertiary", icon=":material/download:",
                       help="Scarica questa tabella in Excel")
