"""
Interfaccia web locale: "Variazioni Standard Cost"

Avvio (dalla cartella STANDARD_COST_AUTOMATION):
    python -m streamlit run app.py

Usa il motore di spiega_variazioni.py. Legge i modelli in SOLA LETTURA: non modifica nessun file.
"""
import glob
import hashlib
import io
import os
import sys
import pickle
import re
import tempfile
from datetime import datetime, timedelta

import openpyxl

import altair as alt
import pandas as pd
import streamlit as st

import cronologia
import spiega_variazioni as sv

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
CACHE_DIR = os.path.join(HERE, "cache_std")
DATA_ROOT = os.path.dirname(ROOT)   # cartella OneDrive che contiene le cartelle delle filiali

# tabelle piu' compatte: righe basse (default ~35 px)
# la funzione originale si salva UNA volta sola: lo script gira a ogni aggiornamento della pagina e, senza questa guardia,
# i "ganci" si sovrapporrebbero (un pulsante Excel in piu' a ogni aggiornamento)
if not hasattr(st, "_dataframe_originale"):
    st._dataframe_originale = st.dataframe
_df_originale = st._dataframe_originale


import scarica_excel
scarica_excel.reset()                      # i pulsanti «Excel» sotto le tabelle prendono un nome progressivo a ogni esecuzione


def _df_compatto(data=None, *args, **kw):
    kw.setdefault("row_height", 28)
    r = _df_originale(data, *args, **kw)
    scarica_excel.bottone(data, indice=False if kw.get("hide_index") else None)      # ogni tabella si puo' scaricare in Excel
    return r


st.dataframe = _df_compatto

# stessa cosa per le tabelle messe dentro una colonna (colonna.dataframe(...))
from streamlit.delta_generator import DeltaGenerator as _DG

if not hasattr(_DG, "_dataframe_originale"):
    _DG._dataframe_originale = _DG.dataframe
_dg_dataframe_orig = _DG._dataframe_originale


def _dg_dataframe(self, data=None, *args, **kw):
    r = _dg_dataframe_orig(self, data, *args, **kw)
    scarica_excel.bottone(data, contenitore=self, indice=False if kw.get("hide_index") else None)
    return r


_DG.dataframe = _dg_dataframe

# numeri all'italiana (1.234,56 e 1,18) in tutte le tabelle formattate: virgola decimale, punto per le migliaia
from pandas.io.formats.style import Styler as _Styler

if not hasattr(_Styler, "_format_originale"):
    _Styler._format_originale = _Styler.format
_format_originale = _Styler._format_originale


def _format_it(self, formatter=None, subset=None, na_rep=None, precision=None, decimal=",", thousands=".",
               escape=None, hyperlinks=None):
    return _format_originale(self, formatter, subset, na_rep, precision, decimal, thousands, escape, hyperlinks)


_Styler.format = _format_it

# ---- palette ---------------------------------------------------------------------------------
INK, MUTED, LINE, BG = "#14213D", "#667085", "#E4E7EC", "#F6F7F9"
ACCENT = "#0F766E"
UP, DOWN = "#B42318", "#027A48"        # costo che sale = rosso, che scende = verde
UP_BG, DOWN_BG, WARN_BG, NEU_BG = "#FEF3F2", "#ECFDF3", "#FFFAEB", "#F2F4F7"
GROUP_COLORS = {"Listino": "#0F766E", "Trasporto": "#2563EB", "Assicurazione": "#7C3AED",
                "Pallet": "#D97706", "AIEM": "#DB2777", "MTS/MTO": "#475467", "Altro": "#98A2B3",
                "Unità di misura": "#B54708", "Maggiorazione +2%": "#0E7490",
                "FOB": "#2563EB", "Oneri import": "#7C3AED", "VGM": "#0891B2", "Cert. sanitario": "#65A30D",
                "Imposta alcolici": "#DB2777"}

st.set_page_config(page_title="Standard Cost", page_icon=os.path.join(HERE, "assets", "standard_cost.png"), layout="wide")

st.markdown(f"""
<style>
.block-container {{ padding-top: 1.6rem; padding-left: 2.2rem; padding-right: 2.2rem; max-width: 100%; }}
html, body, [class*="css"] {{ color: {INK}; }}
/* Tipografia: font di sistema Windows (nessun download), cifre allineate, rendering nitido */
.stApp :where(p, span, label, li, h1, h2, h3, h4, h5, input, textarea, button, div, td, th):not([data-testid="stIconMaterial"]):not([class*="material"]):not([data-testid="stAlertDynamicIcon"] *):not([data-testid*="Icon"] *):not([class*="icon" i]) {{
    font-family: "Segoe UI Variable Text", "Segoe UI", "Inter", "Helvetica Neue", Arial, sans-serif;
}}
.stApp {{ -webkit-font-smoothing: antialiased; text-rendering: optimizeLegibility;
         font-feature-settings: "tnum" 1, "lnum" 1; font-variant-numeric: tabular-nums lining-nums; }}
.stApp h1, .stApp h2, .stApp h3, .stApp h4 {{ font-family: "Segoe UI Variable Display", "Segoe UI Semibold", "Segoe UI", Arial, sans-serif; letter-spacing: -0.01em; }}
.stApp h4 {{ font-weight: 600; }}
h1.hero {{ font-size: 1.75rem; font-weight: 600; letter-spacing: -0.015em; margin: 0 0 .2rem 0; }}
p.sub {{ color: {MUTED}; margin: 0 0 1rem 0; font-size: .98rem; }}
.pill {{ display:inline-block; padding: 2px 10px; border-radius: 999px; font-size: .76rem; font-weight: 500;
        background:{NEU_BG}; color:{INK}; border:1px solid {LINE}; margin-right:6px; }}
.kpi {{ background:#fff; border:1px solid {LINE}; border-radius:14px; padding:14px 16px; }}
.kpi .l {{ color:{MUTED}; font-size:.72rem; text-transform:uppercase; letter-spacing:.08em; font-weight:600; }}
.kpi .v {{ font-size:1.7rem; font-weight:600; line-height:1.15; margin-top:4px; letter-spacing:-0.01em; }}
.kpi .s {{ color:{MUTED}; font-size:.8rem; }}
.card {{ background:#fff; border:1px solid {LINE}; border-radius:16px; padding:16px 18px; height:100%; }}
.card .isl {{ font-weight:600; font-size:1rem; letter-spacing:-0.005em; }}
.card .code {{ color:{MUTED}; font-size:.78rem; }}
.card .std {{ color:{MUTED}; font-size:.92rem; margin-top:6px; }}
.card .std b {{ color:{INK}; }}
.delta {{ font-size:1.7rem; font-weight:600; letter-spacing:-0.015em; margin:2px 0 0 0; }}
.delta {{ white-space:nowrap; }}
.delta small {{ display:block; font-size:.9rem; font-weight:600; margin:0 0 4px 0; }}
.up {{ color:{UP}; }} .down {{ color:{DOWN}; }} .flat {{ color:{MUTED}; }}
.chip {{ display:block; border-radius:10px; padding:6px 10px; margin-top:6px; font-size:.86rem; }}
.chip b {{ font-weight:600; }}
.chip span {{ color:{MUTED}; }}
.chip.up {{ background:{UP_BG}; }} .chip.down {{ background:{DOWN_BG}; }}
.status {{ border-radius:12px; padding:10px 12px; margin-top:10px; font-size:.9rem; background:{WARN_BG};
          border:1px solid #FEDF89; }}
.status.info {{ background:{NEU_BG}; border-color:{LINE}; }}
div[data-testid="stTabs"] button {{ font-weight:600; }}
section[data-testid="stSidebar"] {{ background:{BG}; }}
div[data-testid="stTable"] td, div[data-testid="stTable"] th {{ padding: 3px 8px !important; font-size: .86rem; }}
/* ---- adattamento a tutti gli schermi (portatile, monitor grandi) ---- */
/* le colonne non si stringono oltre un minimo: se manca spazio vanno a capo invece di schiacciarsi */
div[data-testid="stHorizontalBlock"]:has(> div[data-testid="stColumn"]:nth-child(4)) {{ flex-wrap: wrap; row-gap: .6rem; }}
div[data-testid="stHorizontalBlock"]:has(> div[data-testid="stColumn"]:nth-child(4)) > div[data-testid="stColumn"] {{ min-width: 135px; }}
.kpi, .card {{ min-width: 0; overflow-wrap: anywhere; }}
.tab-h, .tab-det, .tab-c {{ max-width: 100%; }}
h1.hero {{ font-size: clamp(1.35rem, 0.9rem + 1vw, 1.75rem); }}
.kpi .v {{ font-size: clamp(1.15rem, 0.7rem + 1vw, 1.7rem); }}
.delta {{ font-size: clamp(1.25rem, 0.8rem + 1vw, 1.7rem); }}
div[data-testid="stButton"] button, div[data-testid="stDownloadButton"] button {{ white-space: nowrap; }}
@media (max-width: 1600px) {{
  .block-container {{ padding-left: 1.4rem; padding-right: 1.4rem; }}
  .tab-h table, .tab-det table {{ font-size: .78rem; }}
}}
@media (max-width: 1300px) {{
  .block-container {{ padding-left: 1rem; padding-right: 1rem; }}
  .tab-h th, .tab-h td, .tab-det th, .tab-det td {{ padding-left: 5px !important; padding-right: 5px !important; }}
  div[data-testid="stTabs"] button {{ padding-left: .6rem; padding-right: .6rem; }}
}}
@media (max-width: 900px) {{
  section[data-testid="stSidebar"] {{ min-width: 0; }}
  .pill {{ font-size: .7rem; }}
}}
</style>
""", unsafe_allow_html=True)


import tabelle_ridimensionabili
tabelle_ridimensionabili.inietta()         # tabelle ridimensionabili: le misure scelte si ricordano a ogni aggiornamento
import assistente
assistente.inietta()                       # il fumetto dell'assistente compare vicino al mouse

# ---- dati ------------------------------------------------------------------------------------
def _motore():
    """Versione del motore (data di spiega_variazioni.py): se cambia, i risultati in memoria si rifanno da zero.
    Senza questo l'app in esecuzione continuerebbe a mostrare i risultati calcolati con la logica precedente."""
    return os.path.getmtime(os.path.join(HERE, "spiega_variazioni.py"))


def carica(path, mtime, filiale):
    return _carica(path, mtime, filiale, _motore())


@st.cache_resource(show_spinner=False)
def _carica(path, mtime, filiale, motore):
    """Legge il modello. La lettura da Excel costa ~4 s a file: il risultato resta anche su disco, cosi' al
    riavvio dell'app (o con il file invariato) si ricarica in meno di un secondo."""
    sv.set_branch(filiale)
    key = hashlib.md5(f"load|{filiale}|{path}|{mtime}|{motore}".encode("utf-8")).hexdigest()
    f = os.path.join(CACHE_DIR, "modello_" + key + ".pkl")
    if os.path.exists(f):
        try:
            with open(f, "rb") as fh:
                return pickle.load(fh)
        except Exception:
            pass                                    # file di cache rovinato: lo rifaccio
    out = sv.load(path)
    try:
        os.makedirs(CACHE_DIR, exist_ok=True)
        tmp = f + ".tmp"
        with open(tmp, "wb") as fh:
            pickle.dump(out, fh, protocol=pickle.HIGHEST_PROTOCOL)
        os.replace(tmp, f)
        vecchi = sorted(glob.glob(os.path.join(CACHE_DIR, "modello_*.pkl")), key=os.path.getmtime)[:-12]   # tengo gli ultimi 12
        for v in vecchi:
            os.remove(v)
    except Exception:
        pass
    return out


def trova_modelli(filiale):
    """Modello del mese in corso + modelli dei mesi chiusi, per la filiale scelta."""
    br = sv.BRANCHES[filiale]
    base = os.path.join(DATA_ROOT, br["folder"])
    out = {}
    radice = os.path.join(base, br["storico"].split("*")[0])      # es. .../IFB/STORICO MESI/
    storico = glob.glob(os.path.join(base, br["storico"]), recursive=True)
    for p in sorted(storico):
        if os.path.basename(p).startswith("~$"):
            continue
        rel = os.path.relpath(os.path.dirname(p), radice)
        mese = rel.replace("\\", "/").split("/")[0].strip()          # solo la cartella del mese
        label = "Storico · " + mese
        if label in out:                                                # due modelli nello stesso mese
            label += " · " + os.path.basename(p)
        out[label] = p
    cur = os.path.join(base, br["current"])
    if os.path.exists(cur):
        out["Mese in corso · INPUT"] = cur
    return out


def salva_upload(f):
    d = tempfile.mkdtemp(prefix="stdc_")
    p = os.path.join(d, f.name)
    with open(p, "wb") as fh:
        fh.write(f.getbuffer())
    return p


def n2(x, signed=False):
    return sv.fmt(x, signed) + " " + VAL


# ---- colorazione condizionale delle tabelle -----------------------------------------------------
def _mix(hexc, a):
    """Colore hexc miscelato con il bianco: a=0 bianco, a=1 colore pieno (niente rgba: compatibilita')."""
    r, g, b = (int(hexc[i:i + 2], 16) for i in (1, 3, 5))
    m = lambda c: round(255 + (c - 255) * a)
    return f"#{m(r):02X}{m(g):02X}{m(b):02X}"


def _cella_delta(v, vmax, soglia):
    """Rosso = costo in aumento, verde = in calo; l'intensita' cresce con l'importo. Zero = grigio chiaro."""
    if v is None or pd.isna(v) or abs(v) <= soglia:
        return "color: #98A2B3"
    a = 0.14 + 0.62 * min(1.0, abs(v) / vmax) if vmax else 0.3
    return f"background-color: {_mix(UP if v > 0 else DOWN, a)}; color: {INK}"


def _cella_causa(v):
    c = GROUP_COLORS.get(v)
    return f"background-color: {_mix(c, 0.18)}; color: {INK}; font-weight: 600" if c else ""


def stile_variazioni(df):
    """Styler per la tabella delle variazioni."""
    eur = [c for c in ["Δ"] + sv.GROUPS if c in df.columns]
    vals = pd.concat([df[c].abs() for c in eur]) if len(df) and eur else pd.Series(dtype=float)
    vmax = float(vals.quantile(0.95)) if len(vals) else 0.0       # il 95o percentile evita che un outlier "spenga" il resto
    pmax = float(df["Δ %"].abs().quantile(0.95)) if len(df) else 0.0
    sty = df.style.format({"STD vecchio": "{:.2f} " + VAL, "STD nuovo": "{:.2f} " + VAL, "Δ %": "{:+.1%}",
                           **{c: (lambda v, c=c: "–" if abs(v) <= sv.TOL else (f"{v:+.3f} {VAL}" if c == "Δ" else f"{v:+.3f}")) for c in eur}},
                          na_rep="")
    sty = sty.map(lambda v: _cella_delta(v, vmax, sv.TOL), subset=eur)
    sty = sty.map(lambda v: _cella_delta(v, pmax, 0.0005), subset=["Δ %"])
    sty = sty.map(_cella_causa, subset=["Causa principale"])
    sty = sty.map(lambda v: f"color: {MUTED}", subset=["STD vecchio"])
    sty = sty.map(lambda v: "font-weight: 600", subset=["STD nuovo"])
    # cambio di unita' di misura: le variazioni non sono reali, celle in grigio (niente rosso/verde)
    spente = eur + ["Δ %"]
    sty = sty.apply(lambda riga: ["background-color: #F2F4F7; color: #667085" if (
        riga["Causa principale"] == "Unità di misura" and c in spente) else "" for c in riga.index], axis=1)
    return sty


COLORI_STATO = {"USCITO": "#B54708", "RIGA": "#B54708", "ORA": "#2563EB", "NUOVO": "#027A48", "NESSUNO": "#667085", "Dati": "#B42318"}


def _cella_stato(v):
    for k, c in COLORI_STATO.items():
        if str(v).startswith(k):
            return f"background-color: {_mix(c, 0.16)}; color: {INK}; font-weight: 600"
    return ""


def analisi_completa(p_old, m_old, p_new, m_new, isole, soglia, filiale):
    return _analisi_completa(p_old, m_old, p_new, m_new, isole, soglia, filiale, _motore())


@st.cache_data(show_spinner=False)
def _analisi_completa(p_old, m_old, p_new, m_new, isole, soglia, filiale, motore):
    sv.set_branch(filiale)
    old, new = carica(p_old, m_old, filiale), carica(p_new, m_new, filiale)
    rows_out, stati, notes, _ = sv.analizza(old, new, None, list(isole), soglia)
    return rows_out, stati, notes




@st.cache_data(show_spinner=False)
def valori_std(path, mtime, filiale):
    """STD Step 2 di tutti gli articoli di un modello, {(foglio, codice): {destinazione: valore}}.
    Salvato su disco: un modello gia' letto (es. un mese chiuso) non va riletto da Excel."""
    key = hashlib.md5(f"{filiale}|{path}|{mtime}|{_motore()}".encode("utf-8")).hexdigest()
    f = os.path.join(CACHE_DIR, key + ".pkl")
    if os.path.exists(f):
        with open(f, "rb") as fh:
            return pickle.load(fh)
    sv.set_branch(filiale)
    out = {}
    for sn, sh in carica(path, mtime, filiale).items():
        for code, row in sh.items.items():
            out[(sn, code)] = {d: (sv.std_finale(sh, row, d) or None) for d in sv.ISLANDS}
    os.makedirs(CACHE_DIR, exist_ok=True)
    with open(f, "wb") as fh:
        pickle.dump(out, fh)
    return out


def serie_storica(versioni, filiale, sn, code, dest):
    """STD Step 2 di un articolo in tutte le versioni disponibili. versioni = ((etichetta, path, mtime), ...)"""
    recs = []
    for label, path, mt in versioni:
        vals = valori_std(path, mt, filiale).get((sn, code))
        if not vals:
            continue
        for d in dest:
            recs.append({"Versione": label, "Destinazione": d, "Standard cost (€)": vals.get(d)})
    return recs


def indice_articoli(new, old):
    """(foglio, No_, IFB, descrizione) di tutti gli articoli, per la ricerca."""
    out = {}
    for src in (old, new):  # il nuovo sovrascrive il vecchio
        for sn, sh in src.items():
            for code, row in sh.items.items():
                ifb = str(sh.val(row, "ifbitem") or "")
                out[(sn, code)] = (sn, code, "" if ifb.startswith("#") else ifb, str(sh.val(row, "description") or ""))
    return list(out.values())


@st.cache_data(show_spinner=False)
def leggi_acquisti(path, mtime):
    """Righe di 02_Ultime_Fatture (aggiornato ogni giorno): codice IFB, data dell'ultima registrazione, quantita'."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb["Export"] if "Export" in wb.sheetnames else wb.worksheets[0]
    it = ws.iter_rows(values_only=True)
    hdr = [str(c).strip().lower() if c is not None else "" for c in next(it)]
    ix = lambda *nomi: next((hdr.index(n) for n in nomi if n in hdr), None)
    ic, idt, iq = ix("no_", "item no", "no", "item"), ix("last posting date", "posting date"), ix("quantity", "quantita")
    recs = []
    for r in it:
        if ic is None or idt is None or len(r) <= max(ic, idt) or r[ic] in (None, ""):
            continue
        d = pd.to_datetime(r[idt], errors="coerce")
        if pd.isna(d):
            continue
        q = r[iq] if iq is not None and len(r) > iq and isinstance(r[iq], (int, float)) else None
        recs.append((str(r[ic]).strip(), d, q))
    wb.close()
    return pd.DataFrame(recs, columns=["Codice", "Data", "Quantita"])


def causa_filtro(main, tag):
    """Etichetta della causa per il filtro: 'Trasporto · pallettizzazione', senza i valori dell'articolo."""
    t = (tag or "").strip()
    if not main:
        return ""
    if main == "Unità di misura":
        return "Cambio unità di misura"
    chiavi =[("fornitore", "cambio fornitore"), ("cambio resa", "cambio resa"), ("modalita", "cambio modalità/area"),
              ("cambio destinazioni", "cambio destinazioni"), ("pallettizzazione", "pallettizzazione/dati logistici"),
              ("aliquota", "aliquota AIEM"), ("grado alcolico", "grado alcolico"),
              ("costo bancale", "costo bancale cambiato"), ("tariffe unifreddo", "tariffe Unifreddo cambiate"),
              ("tariffe trasporto cambiate", "tariffe trasporto cambiate"), ("tariffe cambiate", "tariffe cambiate"),
              ("tariffe trasporto", "tariffe trasporto"), ("tariffa pallet", "tariffa pallet"),
              ("tariffe logistica", "tariffe logistica MTS/MTO")]
    for k, label in chiavi:
        if t.lower().startswith(k):
            return f"{main} · {label}"
    return main        # causa generica (es. Listino: prezzo di acquisto da listino)


def df_variazioni(rows_out):
    recs = []
    for sn, code, desc, forn, isl, e in rows_out:
        m = re.search(r"\(IFB (.+?)\)$", desc)
        main0 = e["parts"][0][0] if e["parts"] else ""
        main = "Unità di misura" if e.get("um") else main0     # cambio di unita': non e' un aumento di prezzo
        recs.append({
            "Articolo": code, "IFB": m.group(1) if m else code, "Descrizione": re.sub(r"\s*\(IFB .+?\)$", "", desc),
            "Fornitore": forn, "Destinazione": isl, "STD vecchio": e["old"], "STD nuovo": e["new"],
            "Δ": e["delta"], "Δ %": (e["pct"] or 0) / 100, "Causa principale": main,
            "Dettaglio causa": e["tag"][main0] if main0 else "",
            "Causa (filtro)": causa_filtro(main, e["tag"][main0] if main0 else ""),
            "Cambi a monte": "; ".join(e.get("a_monte", [])),
            **{g: e["g"][g] for g in sv.GROUPS}, "Spiegazione": sv.sentence(code, desc, isl, e),
        })
    return pd.DataFrame(recs)


def df_stati(stati, isole):
    recs = []
    for sn, code, desc, forn, testo, std_old in stati:
        tipo = testo.split(" (")[0].split(" [")[0]
        recs.append({"Tipo": tipo, "Articolo": code, "Descrizione": desc, "Fornitore": forn, "Stato": testo,
                     **{f"STD prec. {i}": (std_old or {}).get(i) for i in isole}})
    return pd.DataFrame(recs)


def excel_bytes(p_old, m_old, p_new, m_new, isole, soglia, filiale):
    return _excel_bytes(p_old, m_old, p_new, m_new, isole, soglia, filiale, _motore())


@st.cache_data(show_spinner=False)
def _excel_bytes(p_old, m_old, p_new, m_new, isole, soglia, filiale, motore):
    rows_out, stati, notes = analisi_completa(p_old, m_old, p_new, m_new, isole, soglia, filiale)
    p = os.path.join(tempfile.mkdtemp(prefix="stdc_"), "report.xlsx")
    sv.write_report(p, rows_out, notes, os.path.basename(p_old), os.path.basename(p_new), stati, list(isole))
    with open(p, "rb") as fh:
        return fh.read()


# ---- sidebar ---------------------------------------------------------------------------------
with st.sidebar:
    st.markdown("### Standard Cost")
    st.caption("Confronta due versioni del modello e mostra **cosa e perche'** e' cambiato.")
    sv.applica_filiali_utente()          # la lista delle filiali puo' essere cambiata (aggiunta o tolta)
    _etichette = {sv.BRANCHES[k]["label"]: k for k in sv.BRANCHES}
    import percorsi
    _f_ult = os.path.join(percorsi.DATI, "filiale_ultima.json")
    if st.session_state.get("filiale_sel") not in _etichette:
        # la filiale scelta resta finche' non la cambi tu (anche dopo un ricaricamento della pagina o un riavvio)
        try:
            import json as _js
            _ult = _js.load(open(_f_ult, encoding="utf-8")).get("label")
        except Exception:
            _ult = None
        st.session_state["filiale_sel"] = _ult if _ult in _etichette else list(_etichette)[0]
    if len(_etichette) > 3:
        filiale = _etichette[st.selectbox("Filiale", list(_etichette), key="filiale_sel")]
    else:
        filiale = _etichette[st.radio("Filiale", list(_etichette), horizontal=True, key="filiale_sel")]
    try:
        import json as _js
        if st.session_state.get("_filiale_salvata") != st.session_state["filiale_sel"]:
            _js.dump({"label": st.session_state["filiale_sel"]}, open(_f_ult, "w", encoding="utf-8"))
            st.session_state["_filiale_salvata"] = st.session_state["filiale_sel"]
    except Exception:
        pass
    sv.set_branch(filiale)
    import storico_std as _sto
    import margini_fifo as _mfx
    VAL = "HKD" if filiale == "HONG KONG" else "€"          # valuta dello standard cost della filiale
    _sto.imposta_filiale(filiale, VAL)            # ogni filiale ha il suo storico a sistema e i suoi mesi di margini; Hong Kong e' in HKD
    _mfx.imposta_filiale(filiale)
    if sys.platform == "emscripten":          # versione web: non ci sono cartelle sul PC, i file si caricano sempre
        sorgente = "Carica i file"
    else:
        sorgente = st.radio("Sorgente dei file", ["Dalle cartelle", "Carica i file"], horizontal=True)
    p_old = p_new = None
    modelli = {}
    if sorgente == "Dalle cartelle":
        modelli = trova_modelli(filiale)
        if len(modelli) < 2:
            st.warning(f"Non trovo almeno due modelli di {sv.BRANCH['label']} nelle cartelle (INPUT e storico).")
        else:
            nomi = list(modelli)
            n_old = st.selectbox("Versione precedente", nomi, index=max(len(nomi) - 2, 0), key=f"old_{filiale}")
            n_new = st.selectbox("Versione nuova", nomi, index=len(nomi) - 1, key=f"new_{filiale}")
            p_old, p_new = modelli[n_old], modelli[n_new]
    else:
        f1 = st.file_uploader("Versione precedente (.xlsm)", type=["xlsm", "xlsx"], key=f"f1_{filiale}")
        f2 = st.file_uploader("Versione nuova (.xlsm)", type=["xlsm", "xlsx"], key=f"f2_{filiale}")
        if f1 and f2:
            p_old, p_new = salva_upload(f1), salva_upload(f2)
    st.divider()
    isole = list(sv.ISLANDS)       # sempre tutte le destinazioni della filiale
    soglia = sv.TOL                # variazioni sotto 0,005 € considerate nulle
    st.caption("I file vengono letti in sola lettura. Niente esce dal tuo PC.")
    st.toggle("Suggerimenti dell'assistente", value=True, key="assistente_on",
              help="Un fumetto sotto il campo che spiega a cosa serve. Si chiude con la X.")

# ---- testata ---------------------------------------------------------------------------------
_c_tit, _c_btn = st.columns([6, 1.1], vertical_alignment="center")
_c_tit.markdown(f'<h1 class="hero">Standard Cost <span class="pill" style="vertical-align:middle">{sv.BRANCH["label"]}</span></h1>',
                unsafe_allow_html=True)
_in_file = bool(st.session_state.get("_vista_file"))
if True:                          # pulsante in alto a destra: apre (e chiude) la pagina con tutti i caricamenti e il pannello delle filiali
    import assistente
    _c_btn.button("Torna all'analisi" if _in_file else "File da caricare",
                  icon=":material/arrow_back:" if _in_file else ":material/upload_file:", type="primary", width="stretch",
                  key="vai_file", on_click=lambda: st.session_state.update(_vista_file=not st.session_state.get("_vista_file", False)))
    with _c_btn:
        assistente.bolla("vai_file", "Tutti i file in un posto solo",
                         "Qui carichi i report (storico dello standard, margini del mese, Cost Analysis, non IFB) e vedi da quale "
                         "report Power BI si esporta ognuno." if not _in_file else "Torni alle schede di analisi: i file caricati restano.",
                         larghezza=360, destra=True)
if not (p_old and p_new and isole):
    st.markdown('<p class="sub">Scegli (o carica) le due versioni del modello nella barra a sinistra per iniziare.</p>',
                unsafe_allow_html=True)
    st.stop()
if p_old == p_new:
    st.warning("Hai scelto lo stesso file due volte: scegli due versioni diverse.")
    st.stop()

def _etichetta_versione(p):
    """Nome mostrato nel confronto: il modello del mese in corso (cartella INPUT) si chiama OGGI."""
    if sorgente != "Dalle cartelle":
        return os.path.basename(p)
    cartella = os.path.basename(os.path.dirname(p)).strip()
    return "OGGI" if cartella.upper() == "INPUT" else cartella


_MESI_NOMI = ["Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", "Luglio", "Agosto", "Settembre", "Ottobre",
              "Novembre", "Dicembre"]


def _nome_mese(p):
    """Solo il nome del mese del modello: 'Ottobre' dalla cartella '7-OTTOBRE'; il modello in corso (INPUT) e' il mese
    della sua data di salvataggio."""
    if sorgente != "Dalle cartelle":
        return os.path.basename(p)
    cartella = os.path.basename(os.path.dirname(p)).strip()
    if cartella.upper() == "INPUT":
        return _MESI_NOMI[datetime.fromtimestamp(os.path.getmtime(p)).month - 1]
    nome = re.sub(r"^\s*\d+\s*[-–.]?\s*", "", cartella).strip()
    return nome.capitalize() if nome else cartella


st.markdown(
    f'<p class="sub"><span class="pill" title="{os.path.basename(p_old)}">{_nome_mese(p_old)}</span>→'
    f'<span class="pill" style="margin-left:6px" title="Modello salvato il '
    f'{datetime.fromtimestamp(os.path.getmtime(p_new)):%d/%m/%Y alle %H:%M}">{_nome_mese(p_new)}</span></p>',
    unsafe_allow_html=True)

with st.spinner("Leggo i due modelli e calcolo le variazioni… (la prima volta richiede circa un minuto)"):
    old = carica(p_old, os.path.getmtime(p_old), filiale)
    new = carica(p_new, os.path.getmtime(p_new), filiale)
    rows_out, stati, notes = analisi_completa(p_old, os.path.getmtime(p_old), p_new, os.path.getmtime(p_new),
                                              tuple(isole), soglia, filiale)
dfv, dfs = df_variazioni(rows_out), df_stati(stati, isole)

# ---- KPI -------------------------------------------------------------------------------------
n_art = dfv["Articolo"].nunique() if len(dfv) else 0
_reali = dfv[dfv["Causa principale"] != "Unità di misura"] if len(dfv) else dfv   # i cambi di unita' non sono aumenti/cali
n_um = int((dfv["Causa principale"] == "Unità di misura").sum()) if len(dfv) else 0
n_up = int((_reali["Δ"] > 0).sum()) if len(dfv) else 0
n_dn = int((_reali["Δ"] < 0).sum()) if len(dfv) else 0
n_out = int(dfs["Tipo"].str.startswith(("USCITO", "RIGA SPARITA")).sum()) if len(dfs) else 0


def _banner_fonte(tipo, testo):
    """Fascia colorata che dice DA DOVE vengono le informazioni che seguono: dai modelli Excel o dal sistema (macchina)."""
    if tipo == "modelli":
        bg, bordo, col, nome = "#EFF8FF", "#B2DDFF", "#175CD3", "SIMULATO · dal tuo modello Excel"
    else:
        bg, bordo, col, nome = "#F4F3FF", "#D9D6FE", "#5925DC", "REALE · in macchina"
    st.markdown(f'<div style="background:{bg};border:1px solid {bordo};border-left:6px solid {col};border-radius:8px;'
                f'padding:8px 14px;margin:14px 0 10px 0"><span style="color:{col};font-weight:700;letter-spacing:.03em">{nome}</span>'
                f'<span style="color:#344054"> · {testo}</span></div>', unsafe_allow_html=True)


def _nome_mese_versione(label):
    """'7-OTTOBRE' -> 'Ottobre'; 'OGGI' -> il mese del modello in corso."""
    t = str(label).strip()
    if t.upper() == "OGGI":
        return "Mese in corso"
    n = re.sub(r"^\s*\d+\s*[-–.]?\s*", "", t).strip()
    return n.capitalize() if n else t


def _storia_modello(serie, ordine, sn, so, sn_, code, isole):
    """Storia dell'articolo nei modelli Excel: da quanto partiva, quando e' cambiato, se e' bloccato. [(tipo, data, testo)]"""
    righe = []
    if sn == "STORICO":
        return [("info", None, "Prodotto non IFB: non è nel modello standard cost.")]
    e4 = lambda v: f"{v:.4f}".replace(".", ",") + " " + VAL
    if serie is not None and not serie.empty:
        piv = serie.pivot_table(index="Versione", columns="Destinazione", values="Standard cost (€)").reindex(
            [o for o in ordine if o in set(serie["Versione"])])
        primo = piv.index[0]
        righe.append(("info", None, f"{_nome_mese_versione(primo)}: " + " · ".join(
            f"{d} <b>{e4(piv.iloc[0][d])}</b>" for d in piv.columns if pd.notna(piv.iloc[0][d]))))
        voci = []
        for i in range(1, len(piv)):
            camb = []
            forte = False
            for d in piv.columns:
                a, b = piv.iloc[i - 1][d], piv.iloc[i][d]
                if pd.notna(a) and pd.notna(b) and a and abs(b / a - 1) >= 0.15:
                    forte = True                      # variazione oltre il 15%: da evidenziare
                if pd.notna(a) and pd.notna(b) and abs(b - a) > 0.005:
                    camb.append(f"{d}: da {e4(a)} a <b>{e4(b)}</b> ({f"{(b / a - 1) * 100:+.1f}".replace(".", ",")}%)" if a else f"{d}: <b>{e4(b)}</b>")
            if camb:
                voci.append(("alert" if forte else ("up" if piv.iloc[i].sum() > piv.iloc[i - 1].sum() else "down"), None, f"{_nome_mese_versione(piv.index[i])}: " + " · ".join(camb)))
        if not voci:
            righe.append(("ok", None, "Standard <b>sempre uguale</b> in tutti i modelli."))
        else:
            if len(voci) > 5:
                righe.append(("info", None, f"… e altri {len(voci) - 5} cambi precedenti"))
            righe += voci[-5:]
    else:
        righe.append(("stop", None, "Nessuno standard calcolato nei modelli mensili."))
    # situazione nel confronto scelto (bloccato, uscito dal calcolo, nuovo...)
    stati = []
    try:
        ro, rn = so.items.get(code) if so else None, sn_.items.get(code) if sn_ else None
        for isl in isole:
            if so is not None and sn_ is not None and ro is not None and rn is not None and sv.explain(so, ro, sn_, rn, isl) is None:
                s_ = sv.stato(so, ro, sn_, rn, code, isl)
                if s_ and s_[0] not in stati:
                    stati.append(s_[0])
    except Exception:
        pass
    for t in stati:
        righe.append(("stop", None, f"Nel confronto {_nome_mese(p_old)} → {_nome_mese(p_new)}: {t}"))
    return righe


def kpi(col, label, value, sub, color=INK):
    col.markdown(f'<div class="kpi"><div class="l">{label}</div><div class="v" style="color:{color}">{value}</div>'
                 f'<div class="s">{sub}</div></div>', unsafe_allow_html=True)



_nomi_tab = ["Storico", "Tutte le variazioni", "Caricamenti"]
if True:
    _nomi_tab.append("Margini vs FIFO")      # controllo margini: standard cost contro FIFO (per ogni filiale)
if _in_file:          # pagina dei caricamenti: al posto delle schede, si torna indietro con il pulsante in alto a destra
    import file_caricati
    file_caricati.render(filiale, os.path.join(DATA_ROOT, sv.BRANCH["folder"], "STANDARD_COST_AUTOMATION", "INPUT"),
                         modelli, p_old, p_new, sorgente, sv=sv, data_root=DATA_ROOT)
    st.stop()
_tabs = st.tabs(_nomi_tab)
tab1, tab2, tab5 = _tabs[:3]
with tab2:        # Logistica, Anagrafica e Listini sono sottopagine di «Tutte le variazioni»
    tabV, tabL, tabA, tabP = st.tabs(["Variazioni", "Logistica", "Anagrafica", "Listini"])

# ---- tab 5: caricamenti (output della macro) ---------------------------------------------------
MESI = {"gen": 1, "feb": 2, "mar": 3, "apr": 4, "mag": 5, "may": 5, "giu": 6, "jun": 6, "lug": 7, "jul": 7,
        "ago": 8, "aug": 8, "set": 9, "sep": 9, "ott": 10, "oct": 10, "nov": 11, "dic": 12, "dec": 12}


def _mesi_in(testo):
    """Mesi (1-12) citati in un testo, nell'ordine: 'SET-OTT 2026' -> [9, 10]; '6- SETTEMBRE' -> [9]."""
    out = []
    for t in re.findall(r"[a-z]+", testo.lower()):
        m = MESI.get(t[:3])
        if m and m not in out:
            out.append(m)
    return out


def motiva_caricamento(periodo, df):
    """Perche' e' cambiato lo standard di ogni prodotto di un caricamento (output macro).
    Confronta il modello del primo mese del periodo con quello del secondo (es. SET-OTT: settembre vs ottobre).
    Ritorna una lista di testi (uno per riga di df) oppure None se non trovo i due modelli."""
    mesi = _mesi_in(periodo)
    if len(mesi) < 2 or sorgente != "Dalle cartelle":
        return None
    per_mese = {}
    for nome, pth in modelli.items():
        if nome.startswith("Storico"):
            ms = _mesi_in(nome.split("·", 1)[1])
            if ms:
                per_mese[ms[0]] = pth
    if mesi[0] not in per_mese or mesi[1] not in per_mese:
        return None
    sv.set_branch(filiale)
    old = carica(per_mese[mesi[0]], os.path.getmtime(per_mese[mesi[0]]), filiale)
    new = carica(per_mese[mesi[1]], os.path.getmtime(per_mese[mesi[1]]), filiale)
    isl = list(sv.ISLANDS)[0]
    old_col = next((c for c in df.columns if c.lower().startswith("old_sc")), None)
    out = []
    for _, riga in df.iterrows():
        code = str(riga["Code"]).strip()
        ifb = str(riga["IFB"]).strip() if "IFB" in df.columns else ""
        trovato = None
        for sn, sh in new.items():
            if code in sh.items:
                trovato = (sn, code)
                break
        if trovato is None and ifb:
            for sn, sh in new.items():
                if sh.ifb.get(ifb):
                    trovato = (sn, sh.ifb[ifb][0])
                    break
        if trovato is None:
            out.append("Articolo non presente nel modello del mese")
            continue
        sn, k = trovato
        if sn not in old or k not in old[sn].items:
            out.append("Nuovo articolo nel modello (non c'era nel mese precedente)")
            continue
        so, sn_ = old[sn], new[sn]
        e = sv.explain(so, so.items[k], sn_, sn_.items[k], isl)
        if e is None:
            st_ = sv.stato(so, so.items[k], sn_, sn_.items[k], k, isl)
            out.append(st_[0].split(" (")[0] if st_ else "Non calcolabile nel modello")
            continue
        if e.get("um"):
            u = e["um"]
            out.append(f"Cambio unità di misura {u['da']} -> {u['a']}"
                       + (f": a parità di unità {sv.fmt(e['old'], False)} -> {sv.fmt(u['eq'], False)} {VAL}"
                          if u["eq"] is not None else ""))
            continue
        if abs(e["delta"]) <= sv.TOL:
            testo = "Invariato tra i due modelli"
        else:   # testo breve: l'etichetta della causa si scrive solo quando dice qualcosa in piu' (es. pallettizzazione)
            ovvie = {"prezzo di acquisto da listino", "tariffe trasporto", "su nuovo prezzo/trasporto", "tariffa",
                     "effetto sulla nuova base (listino/trasporto)", "tariffa pallet", "tariffe logistica MTS/MTO",
                     "tariffe/costi di filiale", "residuo/arrotondamenti", "effetto sul nuovo prezzo/aliquota"}
            testo = "; ".join(f"{k} {sv.fmt(v)}" + ("" if e["tag"][k] in ovvie else
                                                    f" ({e['tag'][k].replace(' (COSTS LOG)', '')})")
                              for k, v in e["parts"][:3])
        if old_col is not None:
            try:
                v = float(riga[old_col])
                if abs(v - e["old"]) > 0.02:
                    testo += f" · a sistema {sv.fmt(v, False)} {VAL}, nel modello prec. {sv.fmt(e['old'], False)} {VAL}"
            except (TypeError, ValueError):
                pass
        out.append(testo)
    return out


with tab5:
    import caricamenti
    caricamenti.render(filiale, os.path.join(DATA_ROOT, sv.BRANCH["folder"]), filiale, motiva=motiva_caricamento)

# ---- tab 7: margini STD contro FIFO (solo Canarie) --------------------------------------------
if True:
    with _tabs[3]:
        import margini_ui
        margini_ui.render(p_new, filiale)

# ---- tab 1: singolo articolo -----------------------------------------------------------------
_trovati_storico = []         # tutti gli articoli trovati dalla ricerca (per il confronto in fondo alla scheda)
_item_storico = None          # articolo scelto in alto: serve per mostrare il suo storico giornaliero in fondo alla scheda
_modello_cmp = None           # standard dell'articolo nei modelli mensili (data, destinazione, valore): finisce nel grafico unico
with tab1:
    idx = indice_articoli(new, old)
    if True:          # anche gli articoli non IFB: ci sono nell'export dello standard giornaliero, non nel modello
        import storico_std
        _gia = {r[1] for r in idx}
        idx = list(idx) + [("STORICO", c, i, d) for c, i, d in storico_std.elenco_articoli() if c not in _gia]
    lab = lambda r: f"{r[1]}" + (f"  ·  IFB {r[2]}" if r[2] and r[2] != r[1] else "") + f"  ·  {r[3][:60]}"
    _mappa = {}
    for r in idx:
        _l = lab(r)
        if _l in _mappa:               # stesso articolo in piu' fogli: si distingue con il nome del foglio
            _l += "  ·  " + r[0].replace("ITEM STANDARD COST", "Standard Cost").strip("_ ")
        _mappa[_l] = r
    # barra di ricerca con suggerimenti: scrivendo "z11" l'elenco si restringe agli articoli con quel codice o in quella descrizione;
    # si sceglie dall'elenco oppure si preme Invio per cercare il testo scritto
    if st.session_state.get("_fil_ricerca") != filiale:      # cambiando filiale la ricerca riparte: l'articolo scelto era di un'altra
        st.session_state.pop("cerca_art", None)
        st.session_state["_fil_ricerca"] = filiale
    q = st.selectbox("Cerca un articolo", list(_mappa), index=None, accept_new_options=True, key="cerca_art",
                     placeholder="Codice filiale (No_), codice IFB o parte della descrizione…",
                     label_visibility="collapsed") or ""
    import assistente
    if q.strip():
        assistente.bolla("cerca_art", "Vuoi guardare un altro articolo?",
                         "Scrivi per vedere i suggerimenti (codice o descrizione) e scegline uno, oppure premi Invio per cercare "
                         "il testo: se trovo più articoli, sotto vedi il confronto di tutti.")
    else:
        assistente.bolla("cerca_art", "Cerca un articolo, ti mostro cosa è cambiato",
                         "Scrivi il <b>codice</b> (filiale o IFB) o una parte della <b>descrizione</b>: l'elenco si restringe mentre "
                         "scrivi. Ottieni la variazione dello standard divisa per destinazione, da dove viene (<b>modelli</b> o "
                         "<b>macchina</b>) e lo storico giorno per giorno.",
                         ("9624", "FL186", "lievito", "san pellegrino"))
    ql = q.strip().lower()
    if q in _mappa:
        match = [_mappa[q]]
    else:
        match = [r for r in idx if ql and (ql == r[1].lower() or ql == r[2].lower() or ql in r[3].lower())][:60]
    _trovati_storico = [r[1] for r in match]
    if not ql:
        pass          # il suggerimento compare passando sopra alla barra di ricerca (fumetto dell'assistente)
    elif not match:
        st.warning("Nessun articolo trovato con questo testo.")
    else:
        if len(match) > 1:
            st.caption(f"{len(match)} articoli trovati: scegline uno per la spiegazione; in fondo trovi il confronto di tutti.")
        sel = st.selectbox("Articolo", match, format_func=lab, label_visibility="collapsed") if len(match) > 1 else match[0]
        sn, code, ifb, desc = sel
        _item_storico = code
        so, sn_ = old.get(sn), new.get(sn)
        st.markdown(f"#### {desc}")
        _slot_storia = st.container()          # la storia dell'articolo (reale e simulato) si riempie piu' sotto, quando i dati sono pronti
        st.caption(f"Codice filiale **{code}**" + (f" · codice IFB **{ifb}**" if ifb else "") + ("" if sn == "STORICO" else f" · foglio {sn.replace('ITEM STANDARD COST', 'Standard Cost').strip('_ ')}"))
        if sn != "STORICO":
            _banner_fonte("modelli", f"variazione calcolata confrontando i due modelli: {_nome_mese(p_old)} → {_nome_mese(p_new)}")
        if sn == "STORICO":
            st.markdown('<div class="status info">Questo articolo <b>non è nel modello standard cost</b> (prodotto non IFB): la '
                        'scomposizione della variazione per componenti non è disponibile. Qui sotto trovi lo storico giornaliero '
                        'del suo standard.</div>', unsafe_allow_html=True)
        elif code not in (sn_.items if sn_ else {}):
            st.markdown('<div class="status">Questa riga <b>non e\' piu\' nel modello nuovo</b> (delistato o rimosso): '
                        'a sistema resta lo standard cost precedente, non verra\' aggiornato.</div>', unsafe_allow_html=True)
        elif code not in so.items:
            st.markdown('<div class="status info">Articolo <b>nuovo</b>: non c\'era nella versione precedente, '
                        'quindi non c\'e\' una variazione da spiegare.</div>', unsafe_allow_html=True)
        else:
            ro, rn = so.items[code], sn_.items[code]
            in_del = code in sn_.delist
            cols = st.columns(max(len(isole), 3))
            for col, isl in zip(cols, isole):
                e = sv.explain(so, ro, sn_, rn, isl)
                head = f'<div class="isl">{sv.ISLANDS[isl]["nome"]}</div><div class="code">{isl}</div>'
                if e is None:
                    st_ = sv.stato(so, ro, sn_, rn, code, isl)
                    if st_ is None:
                        body = '<div class="status info">Destinazione non servita (standard cost = 0).</div>'
                    else:
                        extra = f'<div class="std">STD a sistema: <b>{n2(st_[1])}</b></div>' if st_[1] is not None else ""
                        body = f'{extra}<div class="status">{st_[0]}</div>'
                    col.markdown(f'<div class="card">{head}{body}</div>', unsafe_allow_html=True)
                    continue
                d = e["delta"]
                cls = "up" if d > sv.TOL else "down" if d < -sv.TOL else "flat"
                arrow = "▲" if cls == "up" else "▼" if cls == "down" else "■"
                pct = f'<small>{sv.fmt(e["pct"])}%</small>' if e["pct"] is not None and cls != "flat" else ""
                chips = "".join(
                    f'<div class="chip {"up" if v > 0 else "down"}"><b>{k} {n2(v, True)}</b><br><span>{e["tag"][k]}</span></div>'
                    for k, v in e["parts"])
                if cls == "flat":
                    chips = (f'<div class="status info">Invariato <b>nel modello Excel</b> tra {_nome_mese(p_old)} e {_nome_mese(p_new)}: '
                             f'{n2(e["new"])} in entrambi i modelli.</div>')
                col.markdown(
                    f'<div class="card">{head}<div class="std"><b>{n2(e["old"])}</b> → <b>{n2(e["new"])}</b></div>'
                    f'<div class="delta {cls}">{arrow} {n2(d, True)}{pct}</div>{chips}</div>', unsafe_allow_html=True)
            if in_del:
                st.markdown('<div class="status">Attenzione: questo articolo e\' presente nel foglio <b>DELISTATI</b> '
                            'ma e\' ancora calcolato nel modello.</div>', unsafe_allow_html=True)

            # cambi a monte (anagrafica, resa, modalita', destinazioni, tariffe): le cause dietro le componenti
            a_monte = []
            for i in isole:
                e_ = sv.explain(so, ro, sn_, rn, i)
                for x in (e_["a_monte"] if e_ else []):
                    if x not in a_monte:
                        a_monte.append(x)
            if a_monte:
                st.markdown('<div class="status"><b>Cambi a monte</b> (anagrafica, resa, modalita\' di trasporto, destinazioni, '
                            'tariffe): ' + " · ".join(a_monte) + "</div>", unsafe_allow_html=True)
            # frasi pronte
            frasi ="\n".join(sv.sentence(code, desc, i, sv.explain(so, ro, sn_, rn, i), (_nome_mese(p_old), _nome_mese(p_new)))
                              for i in isole if sv.explain(so, ro, sn_, rn, i))
            if frasi:
                with st.expander("Riga pronta da copiare (mail, nota, risposta al collega)"):
                    st.code(frasi, language=None)

        # ---- andamento nel tempo (tutte le versioni disponibili) ----
        versioni = tuple(cronologia.versioni_per_pagine(sorgente, modelli, p_old, p_new))
        prog = st.progress(0.0, text="Carico le versioni storiche…")
        for k, (label, pth, mt) in enumerate(versioni):
            prog.progress(k / len(versioni), text=f"Leggo {label} ({k + 1} di {len(versioni)}) — i mesi chiusi si leggono una volta sola")
            valori_std(pth, mt, filiale)
        prog.empty()
        serie = pd.DataFrame(serie_storica(versioni, filiale, sn, code, tuple(isole)))
        serie = serie.dropna(subset=["Standard cost (€)"]) if not serie.empty else serie
        ordine = [v[0] for v in versioni]
        _ha_macchina = False
        _righe_reale = _righe_sim = None
        if True:
            import storico_std
            _righe_reale = storico_std.storia_reale([code, ifb])
            try:          # se l'articolo e' in un caricamento (output macro) la richiesta e' stata inviata: va scritto nella storia
                import caricamenti as _car_m
                _car = _car_m.righe_articolo(filiale, os.path.join(DATA_ROOT, sv.BRANCH["folder"]), [code, ifb])
            except Exception:
                _car = []
            if _car:
                _st = _righe_reale or []
                _testa = [x for x in _st if x[1] is None and x[0] == "info"]
                _datate = sorted([x for x in _st if x[1] is not None and x[0] != "ok"] + _car, key=lambda x: pd.Timestamp(x[1]))
                _coda = [x for x in _st if x[0] == "ok" or (x[1] is None and x[0] != "info")]
                _righe_reale = _testa + _datate + _coda
            _righe_sim = _storia_modello(serie, ordine, sn, so, sn_, code, isole)
            with _slot_storia:
                _c1, _c2 = st.columns(2)
                _c1.markdown(storico_std.scheda_storia(
                    "reale", "In macchina", "lo standard che c'è davvero a sistema, giorno per giorno",
                    _righe_reale or [("info", None, "Questo articolo non è nell'export dello storico caricato"
                                                 + (": " + storico_std.nota_mancante() if storico_std.nota_mancante() else "")
                                                 + ". Rifai l'export includendolo (anche a pezzi) e caricalo in «File da caricare».")]),
                    unsafe_allow_html=True)
                _c2.markdown(storico_std.scheda_storia(
                    "simulato", "Nel modello", "lo standard calcolato dal tuo modello Excel, mese per mese", _righe_sim),
                    unsafe_allow_html=True)
        if True:
            import storico_std
            _ha_macchina = any(storico_std.a_sistema([code, ifb], d) is not None for d in tuple(isole))
            if _ha_macchina and not serie.empty:
                _mt = {v[0]: v[2] for v in versioni}
                _modello_cmp = pd.DataFrame({
                    "data": [pd.Timestamp(datetime.fromtimestamp(_mt[r["Versione"]]).date()) for _, r in serie.iterrows()],
                    "modello": [_nome_mese_versione(r["Versione"]) for _, r in serie.iterrows()],
                    "dest": serie["Destinazione"].values, "valore": serie["Standard cost (€)"].values})

        if _ha_macchina:
            pass          # un solo grafico: standard a sistema con sopra i punti dei modelli (piu' sotto, nella parte «in macchina»)
        elif serie.empty or serie["Versione"].nunique() < 2:
            pass          # meno di due versioni con uno standard cost: la sezione non si mostra
        else:
            st.divider()
            st.markdown("#### Andamento nei modelli mensili")
            marc = cronologia.marche_caricamento(filiale, os.path.join(DATA_ROOT, sv.BRANCH["folder"]), code, ifb, ordine)
            st.altair_chart(cronologia.grafico_andamento(serie, ordine, "Standard cost (€)", "Destinazione", marc), width="stretch")
            if not marc.empty:
                st.caption("▼ = standard caricato dalla macro perche' il prodotto e' stato acquistato: passa il mouse sulla "
                           "freccia per vedere caricamento, data del file, data di acquisto e quantita'.")
            with st.expander("Valori"):
                tab = serie.pivot_table(index="Versione", columns="Destinazione", values="Standard cost (€)").reindex(
                    [o for o in ordine if o in set(serie["Versione"])])
                st.dataframe(tab, width="stretch", column_config={c: st.column_config.NumberColumn(format="%.2f " + VAL) for c in tab.columns})

# ---- sotto la spiegazione: storico giornaliero dello standard (solo Canarie) -------------------
if True:
    with tab1:
        import storico_std
        _banner_fonte("sistema", "lo standard cost registrato davvero a sistema, giorno per giorno (export del report)")
        storico_std.render(filiale, _item_storico, _trovati_storico, _modello_cmp)

# ---- tab 2: tutte le variazioni ---------------------------------------------------------------
with tabV:
    if dfv.empty:
        st.info("Nessuna variazione sopra la soglia.")
    else:
        c1, c3, c4 = st.columns([4, 2, 3])
        import assistente
        causa = c1.multiselect("Causa principale", sorted(c for c in dfv["Causa (filtro)"].unique() if c), key=f"causa_{filiale}")
        verso = c3.radio("Direzione", ["Tutti", "Aumenti", "Cali"], horizontal=True, key=f"dir_{filiale}")
        testo = c4.text_input("Cerca", placeholder="codice, IFB, descrizione, fornitore", key=f"cercav_{filiale}")
        with c1:
            assistente.bolla(f"causa_{filiale}", "Perché è cambiato lo standard?",
                             "Scegli il motivo principale della variazione (prezzo di acquisto, logistica, AIEM, unità di misura…) "
                             "per vedere solo gli articoli cambiati per quella ragione. Puoi sceglierne più di uno.", larghezza=420)
        with c3:
            assistente.bolla(f"dir_{filiale}", "Sale o scende?",
                             "<b>Aumenti</b> = il costo è salito, <b>Cali</b> = è sceso. I cambi di unità di misura non sono aumenti "
                             "veri e non contano.", larghezza=360)
        with c4:
            assistente.bolla(f"cercav_{filiale}", "Cerca nella tabella",
                             "Scrivi un codice, un codice IFB, una parte della descrizione o il fornitore.", larghezza=360, destra=True)
        # filtro "acquistati nel periodo": usa 02_Ultime_Fatture (aggiornato ogni giorno), come la macro (ultimi 30 giorni)
        p_fatt = os.path.join(DATA_ROOT, sv.BRANCH["folder"], "STANDARD_COST_AUTOMATION", "INPUT", "02_Ultime_Fatture.xlsx")
        f1c, f2c, f3c = st.columns([3, 2, 4])
        if os.path.exists(p_fatt):
            solo_acq = f1c.checkbox("Solo prodotti acquistati nel periodo", value=True, key=f"acq_{filiale}",   # sempre attivo di partenza
                                    help="Usa il file 02_Ultime_Fatture in cartella INPUT (si aggiorna ogni giorno).")
            with f1c:
                assistente.bolla(f"acq_{filiale}", "Solo quello che hai comprato",
                                 "Mostra i prodotti acquistati davvero nel periodo (dal file Fatture, aggiornato ogni giorno): "
                                 "sono quelli su cui la macro aggiorna lo standard a sistema.", larghezza=400)
            _acq0 = leggi_acquisti(p_fatt, os.path.getmtime(p_fatt))
            _dmin, _dmax = _acq0["Data"].min().date(), _acq0["Data"].max().date()
            periodo_acq = f1c.date_input("Acquistato dal / al", value=(max(_dmin, _dmax - timedelta(days=30)), _dmax),
                                         min_value=_dmin, max_value=_dmax, format="DD/MM/YYYY",
                                         key=f"dd_{filiale}", disabled=not solo_acq)
            _ritardo = (datetime.now().date() - _dmax).days
            f1c.caption(f"Ultima registrazione nelle fatture: **{_dmax:%d/%m/%Y}** · file aggiornato il "
                        f"{datetime.fromtimestamp(os.path.getmtime(p_fatt)):%d/%m/%Y alle %H:%M}"
                        + (f" · le fatture degli ultimi {_ritardo} giorni non sono ancora registrate a sistema, per questo non compaiono."
                           if _ritardo >= 2 else ""))
            if not solo_acq:          # senza il filtro sugli acquisti: dico esplicitamente a quale periodo si riferisce la variazione
                f1c.markdown(
                    f'<div style="width:max-content;max-width:calc(100vw - 420px);background:#EFF8FF;border:1px solid #B2DDFF;border-radius:8px;'
                    f'padding:8px 14px;font-size:.88rem;color:#344054"><b>Periodo di riferimento della variazione</b><br>'
                    f'tutti i prodotti del modello, indipendentemente dagli acquisti: da <b>{_nome_mese(p_old)}</b> '
                    f'(modello salvato il {datetime.fromtimestamp(os.path.getmtime(p_old)):%d/%m/%Y}) a <b>{_nome_mese(p_new)}</b> '
                    f'(modello salvato il {datetime.fromtimestamp(os.path.getmtime(p_new)):%d/%m/%Y})</div>', unsafe_allow_html=True)
        else:
            solo_acq, periodo_acq = False, None
            f1c.caption("File 02_Ultime_Fatture non trovato: filtro acquisti non disponibile.")
        f = dfv
        if solo_acq:
            acq = leggi_acquisti(p_fatt, os.path.getmtime(p_fatt))
            if isinstance(periodo_acq, (tuple, list)) and len(periodo_acq) == 2:
                limite, fine = pd.Timestamp(periodo_acq[0]), pd.Timestamp(periodo_acq[1])
            else:                                  # selezione a meta' (una sola data): da quella data in poi
                limite = pd.Timestamp(periodo_acq[0] if isinstance(periodo_acq, (tuple, list)) else periodo_acq)
                fine = acq["Data"].max()
            fin = acq[(acq["Data"] >= limite) & (acq["Data"] <= fine)].groupby("Codice").agg(
                **{"Ultimo acquisto": ("Data", "max"), "Qtà acquistata": ("Quantita", "sum")})
            f = f.join(fin, on="IFB")
            mancanti = f["Ultimo acquisto"].isna()               # se la fattura usa il codice della filiale
            if mancanti.any():
                f2 = f.loc[mancanti, ["Articolo"]].join(fin, on="Articolo")
                f.loc[mancanti, "Ultimo acquisto"] = f2["Ultimo acquisto"]
                f.loc[mancanti, "Qtà acquistata"] = f2["Qtà acquistata"]
            f = f[f["Ultimo acquisto"].notna()].copy()
            f["Ultimo acquisto"] = pd.to_datetime(f["Ultimo acquisto"]).dt.strftime("%d/%m/%Y")
        if causa:
            f = f[f["Causa (filtro)"].isin(causa)]
        if verso == "Aumenti":
            f = f[f["Δ"] > 0]
        elif verso == "Cali":
            f = f[f["Δ"] < 0]
        if testo:
            t = testo.lower()
            f = f[f[["Articolo", "IFB", "Descrizione", "Fornitore"]].astype(str).apply(
                lambda s: s.str.lower().str.contains(t, regex=False)).any(axis=1)]
        import filtri
        show = f.drop(columns=["Spiegazione", "Causa (filtro)"])
        show = filtri.filtra(show, f"var_{filiale}")            # filtro per colonna: pulsante «Filtra» sopra la tabella
        vuote = [g for g in sv.GROUPS if len(show) and show[g].abs().max() <= sv.TOL]
        show = show.drop(columns=vuote)       # nascondo le cause che non incidono in nessuna riga visibile
        st.caption(f"{len(f)} righe su {len(dfv)}"
                   + (f" · cause senza impatto non mostrate: {', '.join(vuote)}" if vuote else "")
                   + " · per piu' spazio chiudi la barra laterale con «")
        num = lambda fmt_, w=70: st.column_config.NumberColumn(format=fmt_, width=w)
        txt = lambda w: st.column_config.TextColumn(width=w)
        col = lambda w: st.column_config.Column(width=w)
        _tab_var = st.dataframe(
            stile_variazioni(show.reset_index(drop=True)), hide_index=True, width="stretch",
            height=min(760, 40 + 29 * max(len(show), 3)), on_select="rerun", selection_mode="single-row", key=f"var_tab_{filiale}",
            column_config={
                "Articolo": col(72), "IFB": col(64), "Descrizione": col(170), "Fornitore": col(90),
                "Destinazione": col(74), "Causa principale": col(104), "Dettaglio causa": col(210),
                "Cambi a monte": col(230), "Ultimo acquisto": col(86), "Qtà acquistata": col(80),
                "STD vecchio": st.column_config.Column(width=66, help="Dal modello Excel del mese precedente. Clicca una riga per vedere da dove vengono i valori."),
                "STD nuovo": st.column_config.Column(width=66, help="Dal modello Excel del mese nuovo. Clicca una riga per vedere da dove vengono i valori."),
                "Δ": col(66), "Δ %": col(66),
                **{g: col(70) for g in sv.GROUPS},
            })
        st.caption("Rosso = costo in aumento · verde = costo in calo · piu' e' scuro, piu' pesa la variazione · "
                   "– = nessun impatto. Clicca una riga per vedere da dove vengono i valori (modelli o macchina).")
        _rs = _tab_var.selection.rows if _tab_var is not None else []
        if _rs:
            _r = show.reset_index(drop=True).iloc[_rs[0]]
            st.markdown(f"##### Da dove vengono i valori · {_r['Articolo']} · {_r['Descrizione']} · {_r['Destinazione']}")
            _chip = lambda tipo: ('<span style="background:#EFF8FF;color:#175CD3;border:1px solid #B2DDFF;border-radius:999px;padding:1px 10px;'
                                  'font-size:.78rem;font-weight:700">MODELLO</span>' if tipo == "m" else
                                  '<span style="background:#F4F3FF;color:#5925DC;border:1px solid #D9D6FE;border-radius:999px;padding:1px 10px;'
                                  'font-size:.78rem;font-weight:700">MACCHINA</span>')
            _righe = [(_chip("m"), f"<b>STD vecchio {n2(_r['STD vecchio'])}</b> · dal modello Excel di <b>{_nome_mese(p_old)}</b>"),
                      (_chip("m"), f"<b>STD nuovo {n2(_r['STD nuovo'])}</b> · dal modello Excel di <b>{_nome_mese(p_new)}</b>")]
            if True:
                import storico_std
                _s = storico_std.a_sistema([_r["Articolo"], _r["IFB"]], _r["Destinazione"])
                if _s is None:
                    _righe.append((_chip("s"), "Nessun dato a sistema per questo articolo: carica lo storico in «File da caricare» "
                                               "(o l'articolo manca nell'export)."))
                else:
                    _v = _s["valore"]
                    if abs(_v - _r["STD nuovo"]) <= 0.005:
                        _esito = "uguale al modello nuovo: a sistema è già aggiornato"
                    elif abs(_v - _r["STD vecchio"]) <= 0.005:
                        _esito = "uguale al modello vecchio: a sistema <b>non risulta ancora aggiornato</b>"
                    else:
                        _esito = "diverso da entrambi i modelli"
                    _cambio = f" · ultimo cambio a sistema il {_s['ultimo_cambio']:%d/%m/%Y}" if _s["ultimo_cambio"] is not None else ""
                    _righe.append((_chip("s"), f"<b>Standard a sistema {n2(_v)}</b> · export del {_s['data']:%d/%m/%Y}{_cambio} · {_esito}"))
            for _c, _t in _righe:
                st.markdown(f'<div style="display:flex;gap:10px;align-items:center;margin:4px 0">{_c}<span>{_t}</span></div>',
                            unsafe_allow_html=True)
        st.download_button("Scarica il report Excel completo",
                           excel_bytes(p_old, os.path.getmtime(p_old), p_new, os.path.getmtime(p_new), tuple(isole), soglia, filiale),
                           file_name="Spiegazione_variazioni.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ---- pagine storiche: logistica, anagrafica, listini ------------------------------------------
_VERS = cronologia.versioni_per_pagine(sorgente, modelli, p_old, p_new)

with tabL:
    import costi_log
    costi_log.render(filiale, p_old, p_new, os.path.getmtime(p_old), os.path.getmtime(p_new))
    st.divider()
    cronologia.pagina_logistica(filiale, _VERS, carica)

with tabA:
    cronologia.pagina_anagrafica(filiale, _VERS, carica)

with tabP:
    cronologia.pagina_listini(filiale, _VERS, carica)
    st.divider()
    st.markdown("##### Maggiori variazioni nel confronto selezionato")
    if dfv.empty:
        st.info("Nessuna variazione sopra la soglia.")
    else:
        x, y = st.columns(2)
        base = dfv[dfv["Causa principale"] != "Unità di misura"].sort_values("Δ")   # senza i cambi di unita'
        cfg = {"Δ": st.column_config.NumberColumn(format="%+.3f " + VAL), "Δ %": st.column_config.NumberColumn(format="percent")}
        cols_show = ["Articolo", "Descrizione", "Destinazione", "Δ", "Δ %", "Causa principale"]
        x.markdown("**Maggiori aumenti**")
        x.dataframe(stile_variazioni(base.tail(10).iloc[::-1][cols_show].reset_index(drop=True)) if False else
                    base.tail(10).iloc[::-1][cols_show].reset_index(drop=True).style
                    .map(lambda v: _cella_delta(v, float(base["Δ"].abs().quantile(0.95)), sv.TOL), subset=["Δ"])
                    .map(_cella_causa, subset=["Causa principale"])
                    .format({"Δ": "{:+.3f} " + VAL, "Δ %": "{:+.1%}"}),
                    hide_index=True, width="stretch")
        y.markdown("**Maggiori cali**")
        y.dataframe(base.head(10)[cols_show].reset_index(drop=True).style
                    .map(lambda v: _cella_delta(v, float(base["Δ"].abs().quantile(0.95)), sv.TOL), subset=["Δ"])
                    .map(_cella_causa, subset=["Causa principale"])
                    .format({"Δ": "{:+.3f} €", "Δ %": "{:+.1%}"}),
                    hide_index=True, width="stretch")
        um_df = dfv[dfv["Causa principale"] == "Unità di misura"]
        if len(um_df):
            st.markdown("**Cambi di unità di misura** (non sono aumenti di prezzo: lo standard è riferito a un'altra unità)")
            st.dataframe(um_df[["Articolo", "Descrizione", "Destinazione", "STD vecchio", "STD nuovo", "Dettaglio causa"]]
                         .reset_index(drop=True).style.format({"STD vecchio": "{:.2f} " + VAL, "STD nuovo": "{:.2f} " + VAL}),
                         hide_index=True, width="stretch")
