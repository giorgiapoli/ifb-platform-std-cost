"""
Scheda "Caricamenti": legge gli output della macro (SC_Analisi_*.xlsx), cioe' i prodotti acquistati nel mese
per cui e' stato caricato/aggiornato lo standard cost, e li rende consultabili nel tempo.
Sola lettura: non modifica nessun file.
"""
import glob
import os
import re
from datetime import datetime

import openpyxl
import pandas as pd
import streamlit as st

INK, MUTED, UP, DOWN = "#14213D", "#667085", "#B42318", "#027A48"

# dove cercare gli output della macro, per filiale (relativi alla cartella della filiale)
GLOBS = {
    "CANARIE": ["STANDARD_COST_AUTOMATION/OUTPUT/SC_Analisi_*.xlsx"],
    "HONG KONG": ["STANDARD_COST_AUTOMATION/OUTPUT/SC_Analisi_*.xlsx", "IFB/**/OUTPUT/SC_Analisi_*.xlsx"],
}
# nomi di colonna diversi nei vari mesi -> nome unico
ALIAS = {
    "n_comit": "Code", "code": "Code", "item no_": "Code", "descrizione": "Description",
    "nuovo": "NEW", "azione": "ACTION", "last_date": "Last Date Invoiced", "ultima data": "Last Date Invoiced",
}
MESI = {"gen": 1, "feb": 2, "mar": 3, "apr": 4, "mag": 5, "may": 5, "giu": 6, "jun": 6, "lug": 7, "jul": 7,
        "ago": 8, "aug": 8, "set": 9, "sep": 9, "ott": 10, "oct": 10, "nov": 11, "dic": 12, "dec": 12}
NOMI_MESI = ["", "gennaio", "febbraio", "marzo", "aprile", "maggio", "giugno", "luglio", "agosto", "settembre",
             "ottobre", "novembre", "dicembre"]
STATI = {"TO UPDATE": "#B54708", "DA AGGIORNARE": "#B54708", "NEW": "#027A48", "NUOVO": "#027A48"}


def _mix(hexc, a):
    r, g, b = (int(hexc[i:i + 2], 16) for i in (1, 3, 5))
    m = lambda c: round(255 + (c - 255) * a)
    return f"#{m(r):02X}{m(g):02X}{m(b):02X}"


def _cella_delta(v, vmax):
    if v is None or pd.isna(v) or abs(v) < 1e-9:
        return "color: #98A2B3"
    a = 0.14 + 0.62 * min(1.0, abs(v) / vmax) if vmax else 0.3
    return f"background-color: {_mix(UP if v > 0 else DOWN, a)}; color: {INK}"


def _cella_azione(v):
    for k, c in STATI.items():
        if str(v).strip().upper().startswith(k):
            return f"background-color: {_mix(c, 0.16)}; color: {INK}; font-weight: 600"
    return ""


def trova_file(filiale, base):
    """Output della macro trovati: [(etichetta, path, mtime)], dal piu' recente. Duplicati (stesso nome) uniti."""
    visti = {}
    for pat in GLOBS.get(filiale, []):
        for p in glob.glob(os.path.join(base, pat), recursive=True):
            nome = os.path.basename(p)
            if nome.startswith("~$"):
                continue
            if nome not in visti or os.path.getmtime(p) > os.path.getmtime(visti[nome]):
                visti[nome] = p
    out = []
    for nome, p in visti.items():
        mt = os.path.getmtime(p)
        periodo = re.sub(r"^SC_Analisi_", "", os.path.splitext(nome)[0]).replace("_", " ").strip()
        out.append((f"{periodo}  ·  file del {datetime.fromtimestamp(mt).strftime('%d/%m/%Y')}", p, mt))
    return sorted(out, key=lambda x: -x[2])


@st.cache_data(show_spinner=False)
def leggi(path, mtime, foglio=None):
    """Legge un foglio dell'output macro (default: il primo). Ritorna (nomi_fogli, foglio_letto, DataFrame)."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    nomi = list(wb.sheetnames)
    ws = wb[foglio] if foglio in nomi else wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    titolo = ws.title
    wb.close()
    hi = next((i for i, r in enumerate(rows[:10]) if sum(c is not None for c in r) >= 5), None)
    if hi is None:
        return nomi, titolo, pd.DataFrame()
    hdr, visti = [], {}
    for j, c in enumerate(rows[hi]):
        n = str(c).strip() if c is not None else ""
        n = ALIAS.get(n.lower(), n) or f"col{j + 1}"
        visti[n] = visti.get(n, 0) + 1
        hdr.append(n if visti[n] == 1 else f"{n} ({visti[n]})")
    data = [(list(r) + [None] * len(hdr))[:len(hdr)] for r in rows[hi + 1:] if r and r[0] not in (None, "")]
    df = pd.DataFrame(data, columns=hdr).dropna(axis=1, how="all")
    for c in df.columns[:3]:       # codice, IFB, descrizione: sempre testo (nei file sono a volte numeri, a volte testo)
        df[c] = df[c].map(lambda v: "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v))
    return nomi, titolo, df


def _tipo_colonne(df):
    delta = [c for c in df.columns if c.lower().startswith("d_") or "delta" in c.lower() or c.strip().endswith("%")]
    date = [c for c in df.columns if "date" in c.lower() or c.lower().startswith("data")]
    azione = [c for c in df.columns if c.lower() in ("action", "azione")]
    num = [c for c in df.columns if c not in delta + date + azione and (
        re.match(r"(old|new)[_ ]", c.lower()) or "standard cost" in c.lower()
        or c.lower() in ("quantity", "qty", "quantita'", "quantita"))]
    return delta, date, azione, num


def stile(df):
    df = df.copy().reset_index(drop=True)
    delta, date, azione, num = _tipo_colonne(df)
    for c in delta + num:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    for c in date:
        df[c] = pd.to_datetime(df[c], errors="coerce")
    for c in df.columns:           # colonne di testo: celle vuote = stringa vuota (la tabella altrimenti scrive "None")
        if c not in delta + num + date:
            df[c] = df[c].map(lambda v: "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v))
    fm = {}
    for c in delta:
        fm[c] = lambda v: "" if pd.isna(v) else f"{v:+.1%}"
    for c in num:
        fm[c] = lambda v: "" if pd.isna(v) else (f"{v:,.2f}" if abs(v) < 1e5 else f"{v:,.0f}")
    for c in date:
        fm[c] = lambda v: "" if pd.isna(v) else v.strftime("%d/%m/%Y")
    for c in df.columns:           # tutte le altre colonne: testo, celle vuote davvero vuote (non "None")
        if c not in fm:
            fm[c] = lambda v: "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v)
    sty = df.style.format(fm, na_rep="")
    for c in delta:
        vmax = float(df[c].abs().quantile(0.95)) if df[c].notna().any() else 0.0
        sty = sty.map(lambda v, vm=vmax: _cella_delta(v, vm), subset=[c])
    for c in azione:
        sty = sty.map(_cella_azione, subset=[c])
    return sty


CSS_TABELLA = """<style>
[data-testid="stTable"] table { font-size: .78rem; }
[data-testid="stTable"] th, [data-testid="stTable"] td { padding: .25rem .4rem !important; vertical-align: top; }
[data-testid="stTable"] th { position: sticky; top: 0; background: #F6F7F9; z-index: 1; white-space: normal;
                              word-break: break-word; max-width: 78px; line-height: 1.15; }
</style>"""


def _mostra(df, altezza):
    """Tabella con testo che va a capo (st.dataframe taglierebbe la motivazione): si scorre solo in verticale."""
    sty = stile(df).hide(axis="index")
    if True:                                           # tabella HTML ridimensionabile (larghezza colonne e area si ricordano)
        import margini_ui
        margini_ui._tabella_html(sty, a_sinistra=tuple(c for c in df.columns if df[c].dtype == object),
                                 altezza=altezza or 680, a_capo=tuple(c for c in df.columns if c == "Motivazione" or c.lower() in ("description", "descrizione")))
        return
    # regole per posizione di colonna (st.table ignora white-space/min-width impostati cella per cella)
    regole = ['[data-testid="stTable"] table { width: max-content; min-width: 100%; }',
              '[data-testid="stTable"] td { white-space: nowrap; }']
    for i, c in enumerate(df.columns, 1):
        if c == "Motivazione":
            regole.append(f'[data-testid="stTable"] td:nth-child({i}) {{ white-space: normal; min-width: 300px; max-width: 420px; }}')
        elif c.lower() in ("description", "descrizione"):
            regole.append(f'[data-testid="stTable"] td:nth-child({i}) {{ white-space: normal; min-width: 130px; max-width: 160px; }}')
    st.markdown(CSS_TABELLA + "<style>" + " ".join(regole) + "</style>", unsafe_allow_html=True)
    with st.container(height=altezza):
        st.table(sty)


def _con_motivazione(df, periodo, motiva):
    """Aggiunge la colonna Motivazione (calcolata da `motiva`) se si riesce ad abbinare i modelli del periodo."""
    if motiva is None or df.empty or "Code" not in df.columns:
        return df
    mot = motiva(periodo, df)
    if mot is None:
        return df
    df = df.copy()
    df["Motivazione"] = mot
    return df


def _storico_articolo(files, motiva, key):
    """Ricerca di un articolo in tutti i caricamenti (output macro): quando e' stato caricato."""
    st.markdown("#### Storico di un articolo")
    q = st.text_input("Cerca in tutti i caricamenti", placeholder="codice, IFB o descrizione",
                      key=f"carq_{key}", label_visibility="collapsed")
    if q.strip():
        ql = q.strip().lower()
        frames = []
        for lab, pth, mtime in files:
            _, _, d = leggi(pth, mtime)
            if d.empty:
                continue
            ident = d.iloc[:, :3].astype(str).apply(lambda s: s.str.lower().str.contains(ql, regex=False)).any(axis=1)
            if ident.any():
                part = _con_motivazione(d[ident].copy(), lab.split("  ·  ")[0], motiva)
                part.insert(0, "Caricamento", lab.split("  ·  ")[0])
                part.insert(1, "Data file", datetime.fromtimestamp(mtime))
                frames.append(part)
        if not frames:
            st.info("Questo articolo non compare in nessun caricamento.")
        else:
            tot = pd.concat(frames, ignore_index=True).dropna(axis=1, how="all")
            st.caption(f"{len(tot)} comparse in {len(frames)} caricamenti")
            _mostra(tot, 520 if len(tot) > 10 else None)


def render(filiale, base, key, motiva=None):
    st.markdown("#### Caricamenti mensili")
    st.caption("Output della macro: prodotti acquistati nel mese per cui lo standard cost e' stato caricato o "
               "aggiornato. La data indicata e' quella del file generato dalla macro.")
    files = trova_file(filiale, base)
    if not files:
        st.info("Non trovo output della macro (SC_Analisi_*.xlsx) per questa filiale.")
        return
    mappa = {lab: (p, mt) for lab, p, mt in files}
    c1, c2 = st.columns([3, 2])
    scelta = c1.selectbox("Caricamento", list(mappa), key=f"car_{key}")
    p, mt = mappa[scelta]
    nomi, _, _ = leggi(p, mt)
    foglio = c2.selectbox("Foglio", nomi, key=f"carf_{key}_{scelta}") if len(nomi) > 1 else nomi[0]
    _, _, df = leggi(p, mt, foglio)
    df = _con_motivazione(df, scelta.split("  ·  ")[0], motiva)
    if df.empty:
        st.info("Questo foglio non contiene un elenco leggibile.")
    else:
        delta, date, azione, _ = _tipo_colonne(df)
        # titolo preciso: il mese del caricamento (secondo mese del periodo, es. SET-OTT -> ottobre) e l'anno
        periodo = scelta.split("  ·  ")[0]
        mesi = [MESI[t[:3]] for t in re.findall(r"[a-z]+", periodo.lower()) if t[:3] in MESI]
        anno = (re.search(r"(20\d\d)", periodo) or re.search(r"(20\d\d)", datetime.fromtimestamp(mt).isoformat())).group(1)
        dett = [f"periodo {periodo}", f"file generato il {datetime.fromtimestamp(mt):%d/%m/%Y}"]
        if date:
            dd0 = pd.to_datetime(df[date[0]], errors="coerce").dropna()
            if len(dd0):
                dett.insert(1, f"acquisti dal {dd0.min():%d/%m/%Y} al {dd0.max():%d/%m/%Y}")
        if mesi:
            st.markdown(f'<div class="status info" style="font-size:1rem;margin-bottom:18px"><b>Caricamento di {NOMI_MESI[mesi[-1]]} {anno}</b>'
                        f' · {" · ".join(dett)}</div>', unsafe_allow_html=True)
        k = st.columns(4)

        def kpi(col, label, value, sub=""):
            col.markdown(f'<div class="kpi"><div class="l">{label}</div><div class="v">{value}</div>'
                         f'<div class="s">{sub}</div></div>', unsafe_allow_html=True)

        kpi(k[0], "Prodotti", len(df), "righe nel foglio")
        if azione:
            conta = df[azione[0]].astype(str).str.strip().value_counts()
            top = conta.index[0]
            kpi(k[1], "Azione piu' frequente", int(conta.iloc[0]), top)
        if delta:
            d = pd.to_numeric(df[delta[0]], errors="coerce")
            kpi(k[2], "Variazione media", "" if d.notna().sum() == 0 else f"{d.mean():+.1%}", delta[0])
        if date:
            dd = pd.to_datetime(df[date[0]], errors="coerce").dropna()
            if len(dd):
                kpi(k[3], "Acquisti dal / al", f"{dd.min():%d/%m} - {dd.max():%d/%m}", date[0])
        st.write("")
        _storico_articolo(files, motiva, key)          # ricerca di un articolo: sopra la tabella
        st.markdown("##### Elenco del caricamento")
        _mostra(df, 680 if len(df) > 14 else None)
        if delta:
            st.caption("Rosso = costo in aumento · verde = costo in calo · piu' e' scuro, piu' pesa la variazione.")
        if "Motivazione" in df.columns:
            st.caption("Motivazione: confronto tra il modello del mese precedente e quello del mese del caricamento "
                       "(prima destinazione). Se lo standard a sistema ('Old' del file) non coincide con quello del "
                       "modello precedente, e' indicato: vuol dire che a sistema non era stato aggiornato.")
        elif motiva is not None and "Code" in df.columns:
            st.caption("Motivazione non disponibile: non trovo nello storico i due modelli di questo periodo.")
