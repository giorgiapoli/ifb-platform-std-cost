"""Pagina 'Margini vs FIFO' dell'app (solo Canarie / COMIT). La logica sta in margini_fifo.py."""
import hashlib
import os
import re

import altair as alt
import numpy as np
import pandas as pd
import streamlit as st

import margini_fifo as mf
import filtri
import memoria_file
import scarica_excel

INK, MUTED, LINE = "#14213D", "#667085", "#E4E7EC"
SEMAFORO = {"verde": ("#027A48", "#ECFDF3", "#ABEFC6"), "giallo": ("#B54708", "#FFFAEB", "#FEDF89"),
            "rosso": ("#B42318", "#FEF3F2", "#FECDCA")}


def _motore():
    p = os.path.join(os.path.dirname(os.path.abspath(__file__)), "margini_fifo.py")
    return os.path.getmtime(p), os.path.getmtime(p.replace("margini_fifo", "spiega_variazioni"))


# ---- lettura con cache (i file caricati arrivano come byte) ---------------------------------------
@st.cache_data(show_spinner=False)
def _margini(data, motore):
    return mf.leggi_margini(data)


@st.cache_data(show_spinner=False)
def _entrate(data, motore):
    return mf.leggi_entrate(data)


@st.cache_data(show_spinner=False)
def _items_modello_hk(path, mtime):
    """Codici degli articoli presenti nei fogli ITEM STDC del modello di Hong Kong (quelli che gestiamo noi)."""
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    out = set()
    for n in wb.sheetnames:
        if n.upper().startswith("ITEM STDC"):
            for r in wb[n].iter_rows(min_row=6, max_col=1, values_only=True):
                if r[0] not in (None, ""):
                    out.add(str(r[0]).strip())
    wb.close()
    return out


@st.cache_data(show_spinner=False)
def _listino(path, mtime):
    return mf.leggi_listino(path)


@st.cache_data(show_spinner=False)
def _entrate_hk(data, motore):
    return mf.leggi_entrate_hk(data)


from contextlib import contextmanager


@contextmanager
def _cx(tab):
    """Contenitore di una scheda; se la scheda non esiste (filiale senza quei controlli) il contenuto si scarta."""
    if tab is not None:
        with tab:
            yield
    else:
        ph = st.empty()
        with ph.container():
            yield
        ph.empty()


@st.cache_data(show_spinner=False)
def _non_ifb(data, motore):
    return mf.leggi_non_ifb(data)


@st.cache_data(show_spinner=False)
def _anagrafica(path, mtime, motore):
    return mf.leggi_anagrafica(path)


@st.cache_data(show_spinner=False)
def _std_ifb(path, mtime, motore):
    return mf.leggi_std_ifb(path)


def _kpi(col, label, value, sub="", color=INK):
    col.markdown(f'<div class="kpi"><div class="l">{label}</div><div class="v" style="color:{color}">{value}</div>'
                 f'<div class="s">{sub}</div></div>', unsafe_allow_html=True)


def _it(txt):
    """Formato italiano per un numero gia' scritto all'americana (1,234.5 -> 1.234,5)."""
    return txt.replace(",", "§").replace(".", ",").replace("§", ".")


def _pct(v, signed=True, dec=1):
    return "" if v is None or pd.isna(v) else _it(f"{v:+.{dec}%}" if signed else f"{v:.{dec}%}")


def _eur(v, signed=False):
    if v is None or pd.isna(v):
        return ""
    return _it(f"{v:+,.0f} €" if signed else f"{v:,.0f} €")


def _banner(codice, titolo, testo):
    c, bg, bd = SEMAFORO[codice]
    st.markdown(f'<div style="background:{bg};border:1px solid {bd};border-radius:12px;padding:12px 16px;margin:6px 0 12px 0">'
                f'<div style="color:{c};font-weight:600;font-size:1.05rem">{titolo}</div>'
                f'<div style="color:{INK};font-size:.92rem;margin-top:2px">{testo}</div></div>', unsafe_allow_html=True)


def _evidenzia(v, soglia=0.0):
    """Rosso se il FIFO costa piu' dello STD (Δ margine positivo), verde se costa meno."""
    if v is None or pd.isna(v) or abs(v) <= soglia:
        return ""
    return "color:#B42318;font-weight:600" if v > 0 else "color:#027A48;font-weight:600"


def _xlsx_outlier(tb):
    """Excel con la stessa tabella e gli stessi colori della pagina (STD verde, FIFO rosso, righe a standard zero in giallo)."""
    import io
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as xw:
        wb = xw.book
        ws = wb.add_worksheet("Outlier")
        xw.sheets["Outlier"] = ws
        num = {"Q.tà venduta": '#,##0', "Prezzo vendita per UM": '#,##0.000 "€"', "STD per UM": '#,##0.000 "€"',
               "FIFO per UM": '#,##0.000 "€"', "Vendite": '#,##0 "€"', "Δ margine": '+#,##0 "€";-#,##0 "€"',
               "Δ % vendite": '+0.0%;-0.0%', "Δ prezzo acquisto": '+0.000 "€";-0.000 "€"', "Δ logistica": '+0.000 "€";-0.000 "€"',
               "Δ AIEM": '+0.000 "€";-0.000 "€"', "Non spiegato": '+0.000 "€";-0.000 "€"'}
        larg = {"Descrizione": 42, "Cause probabili": 80, "Articolo": 10, "Origine": 9, "UM": 6}
        testa = wb.add_format({"bold": True, "bg_color": "#F9FAFB", "font_color": "#667085", "text_wrap": True,
                               "align": "center", "valign": "bottom", "border": 1, "border_color": "#E4E7EC"})
        cache = {}

        def fmt(c, bg=None, col=None, bold=False):
            k = (c, bg, col, bold)
            if k not in cache:
                d = {"valign": "top", "border": 1, "border_color": "#F2F4F7"}
                if c in num:
                    d["num_format"] = num[c]
                    d["align"] = "right"
                else:
                    d["align"] = "left"
                if c in ("Descrizione", "Cause probabili"):
                    d["text_wrap"] = True
                if bg:
                    d["bg_color"] = bg
                if col:
                    d["font_color"] = col
                if bold:
                    d["bold"] = True
                cache[k] = wb.add_format(d)
            return cache[k]

        for j, c in enumerate(tb.columns):
            ws.write(0, j, c, testa)
            ws.set_column(j, j, larg.get(c, 14))
        ws.set_row(0, 32)
        ws.freeze_panes(1, 2)
        for i, (_, r) in enumerate(tb.iterrows(), start=1):
            zero = "standard cost a zero" in str(r.get("Cause probabili", ""))
            for j, c in enumerate(tb.columns):
                v = r[c]
                bg = col = None
                bold = False
                if c == "STD per UM":
                    bg, col = "#D1FADF", "#05603A"
                elif c == "FIFO per UM":
                    bg, col, bold = "#FEE4E2", "#B42318", True
                elif c in ("Δ margine", "Δ % vendite") and pd.notna(v) and v != 0:
                    col, bold = ("#B42318" if v > 0 else "#027A48"), True
                if zero:
                    bg = "#FEF9C3"
                if v is None or (not isinstance(v, str) and pd.isna(v)):
                    ws.write_blank(i, j, None, fmt(c, bg, col, bold))
                elif isinstance(v, str):
                    ws.write_string(i, j, v, fmt(c, bg, col, bold))
                else:
                    ws.write_number(i, j, float(v), fmt(c, bg, col, bold))
        ws.autofilter(0, 0, len(tb), len(tb.columns) - 1)
    return buf.getvalue()


LUNGHE = ("Cause probabili", "Segnale", "Anomalia", "Motivo", "Esito", "Motivazione")
VERDE_C = "background-color:#D1FADF;color:#05603A"
ROSSO_C = "background-color:#FEE4E2;color:#B42318;font-weight:600"
AMBRA_C = "background-color:#FEF0C7;color:#93370D"


@st.cache_data(show_spinner=False)
def _ultima_fattura(path, mtime, codice):
    """Data dell'ultima fattura di un codice IFB in 02_Ultime_Fatture (ultima registrazione con quantita' positiva, altrimenti qualsiasi)."""
    d = pd.read_excel(path, usecols=[0, 3, 4])
    d.columns = ["cod", "data", "qty"]
    d = d[d["cod"].astype(str).str.strip() == codice]
    d["data"] = pd.to_datetime(d["data"], errors="coerce")
    d = d[d["data"].notna()]
    if d.empty:
        return None
    pos = d[pd.to_numeric(d["qty"], errors="coerce").fillna(0) > 0]
    return (pos if len(pos) else d)["data"].max()


@st.dialog("Storico dell'articolo", width="large")
def _dialog_storico(codice, descr, ifb, filiale, trasp=""):
    """Si apre cliccando il codice di un articolo: la sua storia a sistema (cambi, caricamenti, da quanto non si aggiorna)."""
    import caricamenti
    import spiega_variazioni as sv
    import storico_std
    st.markdown(f"**{codice} · {descr}**" + (f" · IFB {ifb}" if ifb else "") + (f" · Transportation <b>{trasp}</b>" if trasp and trasp != "—" else ""),
                unsafe_allow_html=True)
    if ifb:          # solo articoli IFB: data dell'ultima fattura (da «Ultime fatture»)
        try:
            _pf = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), sv.BRANCHES[filiale]["folder"],
                               "STANDARD_COST_AUTOMATION", "INPUT", "02_Ultime_Fatture.xlsx")
            _uf = _ultima_fattura(_pf, os.path.getmtime(_pf), ifb) if os.path.exists(_pf) else None
            st.markdown("Ultima fattura (da Ultime fatture): " + (f"**{_uf:%d/%m/%Y}**" if _uf is not None else "nessuna nel file"))
        except Exception:
            pass
    righe = storico_std.storia_reale([codice, ifb])
    base = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), sv.BRANCHES[filiale]["folder"])
    try:
        car = caricamenti.righe_articolo(filiale, base, [codice, ifb])
    except Exception:
        car = []
    righe = storico_std.unisci_caricamenti(righe, car)
    if not righe:
        st.info("Questo articolo non è nell'export dello storico caricato" +
                ((": " + storico_std.nota_mancante()) if storico_std.nota_mancante() else "") + ".")
        return
    st.markdown(storico_std.scheda_storia("reale", "In macchina", "lo standard che c'è davvero a sistema, giorno per giorno", righe),
                unsafe_allow_html=True)


def _tabella_html(sty, a_sinistra=(), altezza=520, a_capo=(), excel=True, nome=None, apri=None):
    """Tabella HTML: le intestazioni vanno a capo quando la finestra si restringe (st.dataframe le taglia).
    a_sinistra = nomi delle colonne di testo da allineare a sinistra; a_capo = colonne di testo lungo che vanno a capo;
    scorre in orizzontale e in verticale."""
    cols = list(sty.data.columns)
    cid = "t" + hashlib.md5("|".join(map(str, cols)).encode("utf-8")).hexdigest()[:8]      # le regole valgono solo per questa tabella
    reg = "".join(f".tab-h.{cid} td.col{i},.tab-h.{cid} th.col{i}{{text-align:left}}" for i, c in enumerate(cols)
                  if c in a_sinistra)
    reg += "".join(f".tab-h.{cid} td.col{i}{{white-space:normal;min-width:{300 if c in LUNGHE else 170}px;"
                   f"max-width:{520 if c in LUNGHE else 260}px;line-height:1.3;overflow-wrap:anywhere}}" for i, c in enumerate(cols) if c in a_capo)
    st.markdown(
        "<style>.tab-h{overflow:auto;resize:both;max-height:78vh;height:" + str(min(altezza, 130 + 29 * len(sty.data))) + "px;min-height:90px;"
        "border:1px solid #E4E7EC;border-radius:10px}"
        ".tab-h table{border-collapse:collapse;font-size:.82rem;width:100%}"
        ".tab-h th{position:sticky;top:0;z-index:2;white-space:normal;text-align:center;vertical-align:bottom;background:#F9FAFB;color:#667085;"
        "font-weight:500;padding:6px 8px;border-bottom:1px solid #E4E7EC;min-width:76px;line-height:1.2}"
        ".tab-h td{white-space:nowrap;padding:5px 8px;border-bottom:1px solid #F2F4F7;text-align:right}" + reg
        + (f".tab-h.{cid} td.col0{{cursor:pointer;color:#0F766E;text-decoration:underline dotted}}"
           f".st-key-{apri}{{position:absolute;left:-9999px;height:0;overflow:hidden}}" if apri else "")
        + f"</style><div class='tab-h {cid}'" + (f" data-apri='{apri}'" if apri else "") + ">" + sty.hide(axis="index").to_html() + "</div>",
        unsafe_allow_html=True)
    if apri:          # campo nascosto: lo script della pagina ci scrive il codice cliccato e questo apre la finestra dello storico
        _v = st.text_input("Apri storico articolo", key=apri, label_visibility="collapsed")
        if _v and st.session_state.get(apri + "_ult") != _v:
            st.session_state[apri + "_ult"] = _v
            st.session_state["_apri_cod"] = _v.split("|")[0]
    if excel:
        scarica_excel.bottone(sty, nome=nome, indice=False)


def _motivo(r):
    """Perche' una spedizione e' tra le anomalie, in parole semplici."""
    m = []
    if r["esito"].startswith("Volume del collo"):
        m.append("trasporto diviso con il volume sbagliato")
    elif r["esito"].startswith("Ripartizione non coerente"):
        m.append("trasporto diviso con un criterio non chiaro")
    elif r["sbilancio_eur"] >= 1:
        m.append(f"piccolo scostamento di volume ({_eur(r['sbilancio_eur'])} spostati)")
    if r["imposte_indebite_eur"] >= 1:
        n = int(r["imposte_su_senza_aiem"])
        m.append(f"imposte per {_eur(r['imposte_indebite_eur'])} su {n} {'articolo' if n == 1 else 'articoli'} senza AIEM")
    if r["esito"].startswith("Ripartito a peso"):
        m.append("ripartito a peso: spedizione aerea, coerente" + (f" · IBRIDO: {int(r['n_ibridi'])} articoli etichettati via mare "
                                                                  "viaggiano in aereo" if r.get("n_ibridi", 0) else ""))
    if not m and r["esito"].startswith("Non verificabile"):
        m.append("volume non verificabile (meno di 3 righe con le misure del box): nessuna anomalia rilevata")
    return " · ".join(m)


def _date_ingresso(date):
    """Data (o intervallo) di registrazione dell'entrata: '15/10/2026' oppure '15/10/2026 – 16/10/2026'."""
    if date is None or len(date) == 0:
        return ""
    a, b = date.min(), date.max()
    return f"{a:%d/%m/%Y}" if a == b else f"{a:%d/%m/%Y} – {b:%d/%m/%Y}"


def _barre_container(titolo, totale, g, c_ok, c_reg, pal_ok, pal_reg, txt_ok, txt_reg, col_ok, col_reg):
    """HTML di un disegno: due container affiancati divisi in blocchi (uno per articolo). g ha item, desc e, per ogni
    grandezza, la quota (c_ok / c_reg) e il valore (v_ok / v_reg). Stesso colore/posizione = stesso articolo."""
    def barra(col, palette, vcol):
        blocchi = []
        for i, r in g.iterrows():
            q = r[col]
            testo = (f'<b>{r["item"]}</b><br>{q:.0%}' if q >= 0.09 else (f'{q:.0%}' if q >= 0.045 else ""))
            tip = f'{r["item"]} · {str(r["desc"]).strip()} · {q:.1%} · {r[vcol]:,.0f} €'.replace(",", ".")
            blocchi.append(f'<div title="{tip}" style="flex:{max(q, 0.004):.4f} 1 0;background:{palette[i % len(palette)]};'
                           'color:#fff;font-size:.7rem;line-height:1.15;text-align:center;display:flex;align-items:center;'
                           'justify-content:center;border-right:2px solid #fff;min-width:2px;overflow:hidden">'
                           f'<span>{testo}</span></div>')
        return ('<div style="display:flex;align-items:stretch">'
                '<div style="flex:1;display:flex;height:60px;border:3px solid #475467;border-right:0;border-radius:8px 0 0 8px;'
                'overflow:hidden;background:#fff">' + "".join(blocchi) + '</div>'
                '<div style="width:12px;border:3px solid #475467;border-radius:0 8px 8px 0;background:'
                'repeating-linear-gradient(90deg,#98A2B3 0 3px,#D0D5DD 3px 6px)"></div></div>')

    leg = "".join(
        f'<span style="display:inline-block;margin:2px 9px 2px 0;font-size:.78rem"><span style="display:inline-block;width:9px;height:9px;'
        f'background:{pal_ok[i % len(pal_ok)]};margin-right:3px"></span><span style="display:inline-block;width:9px;height:9px;'
        f'background:{pal_reg[i % len(pal_reg)]};margin-right:4px"></span><b>{r["item"]}</b> {str(r["desc"]).strip()[:22]}: '
        f'{r[c_ok]:.0%} → <span style="color:{col_reg};font-weight:600">{r[c_reg]:.0%}</span></span>'
        for i, r in g.iterrows())
    return ('<div style="background:#fff;border:1px solid #E4E7EC;border-radius:12px;padding:12px 14px;margin:8px 0">'
            f'<div style="font-weight:600;margin-bottom:4px">{titolo} ' + f'({totale:,.0f} €)'.replace(",", ".") + ' dentro la spedizione</div>'
            f'<div style="color:{col_ok};font-weight:600;font-size:.86rem;margin:6px 0 3px 0">{txt_ok}</div>'
            + barra(c_ok, pal_ok, "v_ok") +
            f'<div style="color:{col_reg};font-weight:600;font-size:.86rem;margin:10px 0 3px 0">{txt_reg}</div>'
            + barra(c_reg, pal_reg, "v_reg") +
            '<div style="margin-top:6px">' + leg + '</div></div>')


def _disegno_container(d, totale_trasporto, con_aiem=True):
    """Due disegni affiancati: a sinistra la ripartizione del trasporto (verde = come dovrebbe essere, per volume; rosso = come
    e' stato diviso), a destra quella dell'AIEM (giallo = atteso, arancione = registrato). Stesso articolo = stessa posizione."""
    g = d.groupby("item").agg(desc=("desc", "first"), vol=("v_ok", "sum"), reg=("reg", "sum"), att=("atteso", "sum"),
                              ai_reg=("imp_reg", "sum"), ai_att=("aiem_att", "sum")).reset_index()
    g = g.sort_values("vol", ascending=False).reset_index(drop=True)
    g["q_ok"] = g["vol"] / g["vol"].sum()
    g["q_reg"] = g["reg"] / g["reg"].sum()
    g["v_ok"], g["v_reg"] = g["att"], g["reg"]
    verdi = ["#027A48", "#12B76A", "#05603A", "#32D583"]
    rossi = ["#B42318", "#F04438", "#7A271A", "#F97066"]
    gialli = ["#CA8A04", "#EAB308", "#854D0E", "#FACC15"]
    arancio = ["#C2410C", "#F97316", "#7C2D12", "#FB923C"]
    sx, dx = st.columns(2) if con_aiem else (st.container(), None)
    sx.markdown(_barre_container("Il trasporto", totale_trasporto, g, "q_ok", "q_reg", verdi, rossi,
                                 "Come dovrebbe essere diviso (in base al volume dei box)", "Come è stato diviso davvero",
                                 "#027A48", "#B42318"), unsafe_allow_html=True)
    if dx is None:          # Hong Kong: niente AIEM
        return
    tot_att, tot_reg = float(g["ai_att"].sum()), float(g["ai_reg"].sum())
    if tot_att <= 0 and tot_reg <= 0:
        dx.markdown('<div style="background:#fff;border:1px solid #E4E7EC;border-radius:12px;padding:12px 14px;margin:8px 0">'
                    '<div style="font-weight:600;margin-bottom:4px">L\'AIEM dentro la spedizione</div>'
                    '<div style="color:#667085;font-size:.88rem">Nessuna imposta AIEM in questa spedizione.</div></div>',
                    unsafe_allow_html=True)
        return
    h = g.copy()
    h["q_ok"] = h["ai_att"] / tot_att if tot_att > 0 else 0.0
    h["q_reg"] = h["ai_reg"] / tot_reg if tot_reg > 0 else 0.0
    h["v_ok"], h["v_reg"] = h["ai_att"], h["ai_reg"]
    dx.markdown(_barre_container("L'AIEM", tot_reg, h, "q_ok", "q_reg", gialli, arancio,
                                 "Come dovrebbe essere ripartita (aliquota × merce + trasporto atteso)", "Come è stata registrata davvero",
                                 "#A16207", "#C2410C"), unsafe_allow_html=True)


_NC = {"c": "COMIT", "m": " Canarie"}          # nomi che cambiano da filiale a filiale: COMIT e' il cliente di Canarie


def render(p_modello, filiale="CANARIE"):
    """Pagina Margini. Per Hong Kong gli importi (vendite, margini, costi, FOB...) sono in HKD: mentre si disegna la pagina il
    simbolo € diventa HKD in testi, riquadri e tabelle (il prezzo di listino resta in euro, scritto «EUR»)."""
    if filiale == "CANARIE":
        return _render(p_modello, filiale)
    from pandas.io.formats.style import Styler as _St
    from streamlit.delta_generator import DeltaGenerator as _DGm
    orig_md, orig_cap, orig_fmt, orig_st_md, orig_st_cap = _DGm.markdown, _DGm.caption, _St.format, st.markdown, st.caption
    sost = lambda t: t.replace("€", "HKD") if isinstance(t, str) else t

    def _f(f):
        if isinstance(f, str):
            return sost(f)
        if callable(f):
            return lambda v, _f=f: sost(_f(v))
        if isinstance(f, dict):
            return {k: _f(v) for k, v in f.items()}
        return f

    def md(self, body, *a, **k):
        return orig_md(self, sost(body), *a, **k)

    def cap(self, body, *a, **k):
        return orig_cap(self, sost(body), *a, **k)

    def fmt(self, formatter=None, *a, **k):
        return orig_fmt(self, _f(formatter), *a, **k)

    _DGm.markdown, _DGm.caption, _St.format = md, cap, fmt
    st.markdown, st.caption = (lambda body, *a, **k: md(st._main, body, *a, **k)), (lambda body, *a, **k: cap(st._main, body, *a, **k))
    try:
        return _render(p_modello, filiale)
    finally:
        _DGm.markdown, _DGm.caption, _St.format = orig_md, orig_cap, orig_fmt
        st.markdown, st.caption = orig_st_md, orig_st_cap


def _render(p_modello, filiale="CANARIE"):
    mf.imposta_filiale(filiale)
    _NC["c"], _NC["m"] = ("COMIT", " Canarie") if filiale == "CANARIE" else ("il cliente", "")
    sfx = "" if filiale == "CANARIE" else "_" + filiale.replace(" ", "_")          # chiavi separate per filiale (mese scelto, file caricati)
    k_mese = "mf_mese" + sfx
    st.markdown("#### Margini " + ("COMIT" if filiale == "CANARIE" else "") + ": standard cost contro FIFO")
    st.caption("Carica ogni mese i report: i file vengono letti in sola lettura e non escono dal PC. "
               "Differenza margine = margine STD − margine FIFO: negativa = standard prudente, "
               "positiva = i costi reali superano lo standard. Fascia buona: tra −3% e −0,5% delle vendite.")
    # ---- mesi salvati: un pulsante per mese ---------------------------------------------------
    salvati = mf.mesi_salvati()
    sel = st.session_state.get(k_mese)
    if sel not in salvati:
        sel = None
    if salvati:
        st.markdown("**Mesi in memoria**")
        cols = st.columns([3] * len(salvati) + [3, max(18 - 3 * len(salvati), 3)])
        for col, k in zip(cols, salvati):
            if col.button(mf.etichetta_mese(k), key=f"mf_btn_{k}", type="primary" if k == sel else "secondary",
                          width="stretch"):
                st.session_state[k_mese] = k
                st.rerun()
        if cols[len(salvati)].button("＋ Nuovo mese", key="mf_nuovo", type="primary" if sel is None else "secondary",
                                     width="stretch"):
            st.session_state[k_mese] = None
            st.rerun()
    b_marg = b_cost = b_non = None
    giorni_rec = 30          # entrate recenti a costi accessori zero = fatture probabilmente non ancora registrate
    if sel:
        b_marg, b_cost, b_non = mf.leggi_file_mese(sel)
        # se in «File da caricare» c'e' un report diverso da quello salvato per questo mese, si puo' sostituire il mese
        _n = [st.session_state.get("dato_" + k + sfx) or memoria_file.carica(k + sfx) for k in ("mf_marg", "mf_cost", "mf_non")]
        # avviso solo se il report in «File da caricare» e' stato caricato DOPO l'ultimo salvataggio del mese (e quindi e' piu' recente)
        _t_up = os.path.getmtime(os.path.join(memoria_file.DIR, "mf_marg" + sfx + ".bin"))             if os.path.exists(os.path.join(memoria_file.DIR, "mf_marg" + sfx + ".bin")) else 0
        _f_mese = os.path.join(mf._dir_mesi(), sel, mf._FILE_MESE["marg"])
        _t_mese = os.path.getmtime(_f_mese) if os.path.exists(_f_mese) else 0
        if _n[0] and _n[0].getvalue() != b_marg and _t_up > _t_mese:
            _w1, _w2 = st.columns([5, 2], vertical_alignment="center")
            _w1.warning(f"L'ultimo report margini caricato («{getattr(_n[0], 'name', '')}») è diverso da quello salvato per "
                        f"{mf.etichetta_mese(sel)}: qui stai guardando quello salvato.")
            if _w2.button("Sostituisci il mese con l'ultimo caricamento", key="mf_sost_" + sel, type="primary"):
                mf.salva_file_mese(sel, _n[0].getvalue(), _n[1].getvalue() if _n[1] else b_cost,
                                   _n[2].getvalue() if _n[2] else b_non)
                st.rerun()
    else:
        f_marg, f_cost, f_non = (st.session_state.get("dato_" + k + sfx) or memoria_file.carica(k + sfx) for k in ("mf_marg", "mf_cost", "mf_non"))   # caricati in «File da caricare»
        b_marg = f_marg.getvalue() if f_marg else None
        b_cost = f_cost.getvalue() if f_cost else None
        b_non = f_non.getvalue() if f_non else None
    if not b_marg:
        st.info("Carica i tre file del mese nella pagina «File da caricare» (pulsante in alto a destra) (il report dei margini è obbligatorio; con il Cost "
                "Analysis e il report non IFB l'app spiega anche perché gli articoli si allontanano dallo standard).")
        return
    mt = _motore()
    try:
        art, tot = _margini(b_marg, mt)
    except Exception as ex:
        st.error(f"Non riesco a leggere il report dei margini: {ex}")
        return
    ent_hk = None
    try:
        if b_cost and filiale != "CANARIE":
            ent_hk = _entrate_hk(b_cost, mt)
        ent = _entrate(b_cost, mt) if (b_cost and filiale == "CANARIE") else (mf.adatta_entrate_hk(ent_hk) if ent_hk is not None else None)
    except Exception as ex:
        st.error(f"Non riesco a leggere il Cost Analysis: {ex}")
        ent = None
    non = _non_ifb(b_non, mt) if b_non else {}
    if sel is None:        # file appena caricati: li si puo' tenere in memoria con il nome del mese
        mese_nuovo_def = f"{ent['data'].max():%Y-%m}" if ent is not None and len(ent) else ""
        s1, s2, s3 = st.columns([1, 1, 3])
        chiave_nuova = s1.text_input("Mese da salvare", value=mese_nuovo_def, placeholder="es. 2026-10",
                                     help="Anno-mese: il pulsante prendera' il nome del mese (es. 2026-10 = Ottobre 2026).")
        s2.write("")
        if s2.button("Tieni in memoria", disabled=not chiave_nuova.strip(), width="stretch"):
            k = mf.salva_file_mese(chiave_nuova, b_marg, b_cost, b_non)
            st.session_state[k_mese] = k
            st.rerun()
        s3.caption("Salva i file caricati: dalla prossima volta basta cliccare il pulsante del mese.")
    ifb = {}
    if filiale == "CANARIE" and p_modello and os.path.exists(p_modello):
        with st.spinner("Leggo gli standard cost IFB dal modello…"):
            ifb = _std_ifb(p_modello, os.path.getmtime(p_modello), mt)
    a = mf.perimetri(art, ifb, non)
    a["trasp"] = ""
    if filiale != "CANARIE":
        # Hong Kong: IFB = articoli con codice IFB nel report; Transportation (SEA / AIR / ALTRO) dall'anagrafica del modello
        _ifb_ok = a["ifb"].notna() & a["ifb"].astype(str).str.strip().ne("") if "ifb" in a.columns else pd.Series(False, index=a.index)
        a["perimetro"] = np.where(_ifb_ok, "IFB", "Non IFB")
        _cl = lambda t: "SEA" if "SEA" in t else ("AIR" if "AIR" in t else "ALTRO")
        _ana_t = _anagrafica(p_modello, os.path.getmtime(p_modello), mt) if (p_modello and os.path.exists(p_modello)) else {}
        a["trasp"] = [_cl(_ana_t.get(i, {}).get("_trasp") or "") for i in a["item"]]
        a["trasp_txt"] = [(_ana_t.get(i, {}).get("_trasp") or "—") for i in a["item"]]          # transportation come scritta in anagrafica
    perim = mf.tabella_perimetri(a)
    # lo storico si aggiorna da solo a ogni caricamento (mese in memoria o file appena caricati con un mese valido)
    chiave_auto = sel if sel is not None else chiave_nuova.strip()
    if re.fullmatch(r"\d{4}-\d{2}", chiave_auto):
        mf.salva_mese(chiave_auto, tot, perim)

    # dati di imballo di tutti gli articoli (modello per gli IFB, anagrafica COMIT per gli altri e per i mancanti)
    # l'anagrafica del mese e' gia' nel modello (foglio «Anagrafica Articoli»): si legge da li', senza altri file
    p_ana = p_modello if (p_modello and os.path.exists(p_modello)) else None
    ana = _anagrafica(p_ana, os.path.getmtime(p_ana), mt) if p_ana else {}
    dati_art = mf.dati_articoli(ifb, non, ana)

    # ---- analisi sulle entrate -------------------------------------------------------------
    w = med = aiem = cont = det = trasp = None
    provv = pd.DataFrame()
    if sel is not None:
        mese_chiave = sel
    else:
        mese_chiave = chiave_nuova
    if ent is not None and len(ent):
        w, inizio, fine = mf.finestra_mese(ent, mese_chiave)
        if w.empty:
            st.warning(f"Nel Cost Analysis non ci sono entrate di {mf.etichetta_mese(mese_chiave)}: i controlli sulle entrate non sono "
                       "disponibili per questo mese.")
            ent = None
    if ent is not None and len(ent):
        w, provv, recenti_dal = mf.escludi_provvisorie(w, ent["data"].max(), giorni_rec)
        med = mf.medie_entrate(w)
        aiem = mf.controllo_aiem(w, ifb, non, recenti_dal)
        cont, det = mf.controllo_container(w, dati_art)
        if filiale != "CANARIE" and cont is not None and len(cont):
            # Hong Kong: le spedizioni via aerea ripartiscono a PESO, non a volume. Se la ripartizione registrata torna con il
            # peso la spedizione e' coerente (non e' un'anomalia); gli articoli «SEA» dentro sono IBRIDI (viaggiano in aereo).
            pk = mf.classifica_peso_hk(w, ana)
            cont = cont.merge(pk, on="ext", how="left")
            aer = cont["a_peso"].fillna(False).astype(bool) & (cont["esito"] != "Ripartito a volume correttamente")
            cont.loc[aer, "esito"] = "Ripartito a peso (spedizione aerea)"
            cont.loc[aer, ["impatto_eur", "sbilancio_eur"]] = 0.0
            cont["n_ibridi"] = [len(x) if isinstance(x, list) and a_ else 0 for x, a_ in zip(cont["ibridi"], aer)]
        trasp = mf.controllo_trasporto(med, ifb, non)

    if ent is not None and len(ent) and len(provv):
        with st.expander(f"{len(provv)} entrate recenti con costi accessori a zero: fatture probabilmente non ancora registrate "
                         f"(escluse dai controlli)"):
            st.dataframe(provv.sort_values("data", ascending=False)[["data", "item", "desc", "ubic", "ext", "qty", "prezzo"]].rename(
                columns={"data": "Data", "item": "Articolo", "desc": "Descrizione", "ubic": "Ubicazione",
                         "ext": "Spedizione (fattura)", "qty": "Q.tà", "prezzo": "Prezzo"}).style.format(
                {"Data": "{:%d/%m/%Y}", "Q.tà": "{:,.0f}", "Prezzo": "{:.3f} €"}), hide_index=True, width="stretch")
    t_hk = None
    if filiale == "CANARIE":
        t_esito, t_cont, t_art, t_aiem, t_tr, t_sto = st.tabs(
            ["Esito e outlier", "Spedizioni", "Un articolo", "AIEM", "Trasporto", "Storico mesi"])
    else:          # Hong Kong: niente AIEM e niente confronto con la logistica del modello (solo ripartizione delle spedizioni)
        t_esito, t_cont, t_art, t_hk, t_sto = st.tabs(["Esito e outlier", "Spedizioni", "Un articolo", "Costi d'entrata", "Storico mesi"])
        t_aiem = t_tr = None

    # ================== esito e outlier =====================================================
    with t_esito:
        scelta = None
        if True:
            scelta = st.segmented_control("Mostra", ["Tutti", "IFB", "Non IFB"], default="Tutti", key="mf_perim",
                                          label_visibility="collapsed",
                                          help="IFB = standard dal modello" + _NC["m"] + ". Non IFB = prodotti che " + _NC["c"] + " compra direttamente.")
            scelta = scelta or "Tutti"
        scelta = scelta or "Tutti"
        a_v = a if scelta == "Tutti" else a[a["perimetro"] == scelta]
        _ts = "Tutti"
        if filiale != "CANARIE":          # filtro Transportation (Hong Kong): SEA / AIR / ALTRO
            _ts = st.segmented_control("Transportation", ["Tutti", "SEA", "AIR", "ALTRO"], default="Tutti", key="mf_trasp_hk",
                                       label_visibility="collapsed",
                                       help="SEA = via mare, AIR = via aerea, ALTRO = le altre transportation. Dall'anagrafica del modello.") or "Tutti"
            if _ts != "Tutti":
                a_v = a_v[a_v["trasp"] == _ts]
        tot_v = tot if (scelta == "Tutti" and _ts == "Tutti") else mf.totali(a_v)
        cod, tit, txt = mf.esito(tot_v["diff_pct"])
        _banner(cod, f"{tit}: scostamento {_pct(tot_v['diff_pct'], True, 2)} delle vendite", txt)
        if tot_v.get("filtri"):
            st.warning(f"Il report dei margini è filtrato ({tot_v['filtri']}): questo scostamento riguarda solo quegli articoli e può differire "
                       "dal totale del mese. Per il dato complessivo ricarica l'export senza filtro sul trasporto.", icon="⚠️")
        k1, k2, k3, k4 = st.columns(4)
        _kpi(k1, "Vendite del mese", _eur(tot_v["vendite"]), f"{len(a_v)} articoli")
        _kpi(k2, "Margine a standard", _pct(tot_v["p_std"], False), _eur(tot_v["m_std"]))
        _kpi(k3, "Margine a FIFO", _pct(tot_v["p_fifo"], False), _eur(tot_v["m_fifo"]))
        _kpi(k4, "Scostamento STD − FIFO", _pct(tot_v["diff_pct"], True, 2), _eur(tot_v["diff"], True), SEMAFORO[cod][0])
        if scelta == "Tutti" and (a["perimetro"] == "Fuori modello").any():
            st.caption(f"«Tutti» comprende anche {int((a['perimetro'] == 'Fuori modello').sum())} articoli senza standard "
                       "né nel modello IFB né nel report non IFB.")
        st.write("")

        st.markdown("#### Outlier")
        o1, o2, o3, o6, o5, o4 = st.columns([1, 1, 1.3, 1, 1.6, 2])
        import assistente
        min_pct = o1.number_input("Scostamento minimo (% vendite)", 0.0, 100.0, 10.0, 1.0, key="mf_min_pct") / 100
        min_eur = o2.number_input("Scostamento minimo (€)", 0.0, 100000.0, 50.0, 10.0, key="mf_min_eur")
        verso = o3.selectbox("Verso", ["Tutti", "FIFO peggiore dello STD", "STD troppo alto"], key="mf_verso")
        cerca = o4.text_input("Cerca", placeholder="codice o parte della descrizione", key="mf_cerca")
        with o1:
            assistente.bolla("mf_min_pct", "Quanto pesa sulle vendite?",
                             "Tieni solo gli articoli in cui la differenza tra margine a standard e a FIFO supera questa percentuale "
                             "delle loro vendite.", larghezza=340)
        with o2:
            assistente.bolla("mf_min_eur", "E in euro?",
                             "Toglie le differenze piccole: resta chi perde almeno questi euro di margine.", larghezza=320)
        with o3:
            assistente.bolla("mf_verso", "In che direzione?",
                             "<b>FIFO peggiore dello STD</b> = i costi reali superano lo standard, il margine è a rischio. "
                             "<b>STD troppo alto</b> = lo standard è più prudente del necessario.", larghezza=380)
        entrate = o5.selectbox("Entrate nel periodo", ["Tutti", "Nascondi quelli senza entrate", "Solo quelli senza entrate"], key="mf_entrate",
                               help="«Senza entrate» = nessuna entrata nel periodo analizzato: il FIFO arriva da giacenza più vecchia (danno del passato che si trascina).")
        origine = o6.selectbox("Origine", ["Tutti", "IFB", "Non IFB"], key="mf_origine",
                               help="IFB = standard dal modello" + _NC["m"] + ". Non IFB = prodotti che " + _NC["c"] + " compra direttamente.")
        with o6:
          if True:
            assistente.bolla("mf_origine", "IFB o non IFB?",
                             "<b>IFB</b> = standard dal modello" + _NC["m"] + ". <b>Non IFB</b> = prodotti che " + _NC["c"] + " compra direttamente.",
                             larghezza=340, destra=True)
        with o5:
            assistente.bolla("mf_entrate", "Danni del passato?",
                             "«Senza entrate» = nessun acquisto nel periodo: il FIFO arriva da giacenza più vecchia, è un errore di "
                             "ingresso che si trascina. Nascondili per vedere solo i problemi del mese.", larghezza=400, destra=True)
        with o4:
            assistente.bolla("mf_cerca", "Cerca un articolo", "Scrivi il codice o una parte della descrizione.",
                             larghezza=320, destra=True)
        out = mf.outlier(a_v, min_pct, min_eur)
        if origine != "Tutti":
            out = out[out["perimetro"] == origine]
        if verso == "FIFO peggiore dello STD":
            out = out[out["diff"] > 0]
        elif verso == "STD troppo alto":
            out = out[out["diff"] < 0]
        if cerca:
            q = cerca.strip().lower()
            out = out[out["item"].str.lower().str.contains(q, regex=False) | out["desc"].astype(str).str.lower().str.contains(q, regex=False)]
        dg = None
        if med is not None:
            if filiale == "CANARIE":
                dg = mf.diagnosi(out, med, aiem, trasp, det, ifb, non)
            else:
                dg = mf.diagnosi_hk(out, ent, mese_chiave, _items_modello_hk(p_modello, os.path.getmtime(p_modello)) if p_modello else set())
            out = out.join(dg, on="item")
            _senza = out["cause"].astype(str).str.contains("nessuna entrata nel periodo", regex=False)
            if entrate.startswith("Nascondi"):
                out = out[~_senza]
            elif entrate.startswith("Solo"):
                out = out[_senza]
        peggio = out[out["diff"] > 0]
        st.markdown(f"**{len(out)} articoli outlier** · {len(peggio)} con FIFO peggiore dello standard "
                    f"({_eur(peggio['diff'].sum(), True)} di margine) · {len(out) - len(peggio)} con standard troppo alto "
                    f"({_eur(out[out['diff'] < 0]['diff'].sum(), True)}).")
        if out.empty:
            st.success("Nessun articolo sopra le soglie scelte.")
        else:
            cols = ["item", "desc", "perimetro"] + (["trasp_txt"] if filiale != "CANARIE" else []) + ["qty", "uom", "prezzo", "c_std", "c_fifo", "vendite", "diff", "diff_pct"]
            nomi = {"item": "Articolo", "desc": "Descrizione", "perimetro": "Origine", "trasp_txt": "Transportation", "qty": "Q.tà venduta", "uom": "UM",
                    "prezzo": "Prezzo vendita per UM", "c_std": "STD per UM", "c_fifo": "FIFO per UM", "vendite": "Vendite",
                    "diff": "Δ margine", "diff_pct": "Δ % vendite"}
            if dg is not None:
                cols += ["cause", "d_prezzo", "d_logistica"] + (["d_aiem"] if filiale == "CANARIE" else []) + ["d_altro"]
                nomi.update({"d_prezzo": "Δ prezzo acquisto", "d_logistica": "Δ logistica" if filiale == "CANARIE" else "Δ FOB",
                             "d_aiem": "Δ AIEM", "d_altro": "Non spiegato", "cause": "Cause probabili"})
            tb = out[cols].rename(columns=nomi).reset_index(drop=True)
            tb = filtri.filtra(tb, "outlier")
            fm = {"Q.tà venduta": "{:,.0f}", "Prezzo vendita per UM": "{:.3f} €", "Vendite": "{:,.0f} €", "STD per UM": "{:.3f} €",
                  "FIFO per UM": "{:.3f} €", "Δ margine": "{:+,.0f} €", "Δ % vendite": "{:+.1%}",
                  "Δ prezzo acquisto": "{:+.3f} €", "Δ logistica": "{:+.3f} €", "Δ AIEM": "{:+.3f} €", "Non spiegato": "{:+.3f} €"}
            fmt = lambda f: (lambda v: "" if v is None or pd.isna(v) else f.format(v))
            sty = tb.style.format({k: fmt(v) for k, v in fm.items() if k in tb.columns})
            sty = sty.map(_evidenzia, subset=[c for c in ("Δ margine", "Δ % vendite") if c in tb.columns])
            sty = sty.map(lambda v: VERDE_C, subset=["STD per UM"]).map(lambda v: ROSSO_C, subset=["FIFO per UM"])   # verde = standard, rosso = costo reale
            if "Cause probabili" in tb.columns:        # standard a zero: tutta la riga in giallo tenue, per farla risaltare
                sty = sty.apply(lambda r: ["background-color:#FEF9C3"] * len(r)
                                if "standard cost a zero" in str(r["Cause probabili"]) else [""] * len(r), axis=1)
            _tabella_html(sty, a_sinistra=("Articolo", "Descrizione", "Origine", "Transportation", "UM", "Cause probabili"), altezza=640, apri="mf_apri_art",
                          a_capo=("Descrizione", "Cause probabili"))
            _cod = st.session_state.pop("_apri_cod", None)          # codice cliccato nella tabella: si apre lo storico dell'articolo
            if _cod:
                _r = a[a["item"] == _cod]
                _dialog_storico(_cod, str(_r["desc"].iloc[0]) if len(_r) and "desc" in _r.columns else "",
                                str(_r["ifb"].iloc[0]) if len(_r) and "ifb" in _r.columns and pd.notna(_r["ifb"].iloc[0]) else "", filiale,
                                str(_r["trasp_txt"].iloc[0]) if len(_r) and "trasp_txt" in _r.columns else "")
            if dg is not None:
                st.caption("Prezzo vendita, STD e FIFO sono per unità di misura (UM), come nel report: per un pezzo se la UM è PCS, per un kg se è KG. "
                           "STD per UM = lo standard che c'era a sistema nel mese su quelle vendite; FIFO per UM = il costo reale. "
                           "Le colonne Δ (€ per UM) dicono "
                           "di quanto il costo reale delle entrate recenti supera lo standard: prezzo di acquisto, logistica "
                           "(trasporto + despachos), AIEM. «Non spiegato» = quanto le entrate non spiegano (giacenza più vecchia, "
                           "maggiorazione concordata, unità di misura).")
        if med is not None:
            st.download_button("Scarica il rapporto in Excel", mf.excel_report(
                tot, perim, out.drop(columns=[c for c in out.columns if c.startswith("_")]),
                mf.riepilogo_aiem(aiem, w), cont, det, trasp),
                file_name="controllo_margini_fifo.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    needs = "Carica anche il file Cost Analysis (entrate di magazzino) per attivare questo controllo."

    # ================== AIEM ==============================================================
    with _cx(t_aiem):
        if aiem is None:
            st.info(needs)
        else:
            st.markdown("#### AIEM registrato contro anagrafica")
            st.caption(f"Entrate dal {inizio:%d/%m/%Y} al {fine:%d/%m/%Y}. L'AIEM atteso viene dall'anagrafica articolo "
                       "(modello IFB / report non IFB): una voce IMPUESTOS su un articolo senza AIEM, o l'assenza su un articolo "
                       "che ce l'ha, è una registrazione ambigua da segnalare a " + _NC["c"] + ".")
            if aiem.empty:
                st.success("Nessuna entrata con AIEM incoerente nel periodo.")
            else:
                ri = mf.riepilogo_aiem(aiem, w)
                a1, a2, a3 = st.columns(3)
                x = ri[ri["tipo"].str.startswith("AIEM registrato, ma")]
                y = ri[ri["tipo"].str.startswith("AIEM non registrato")]
                z = ri[ri["tipo"].str.contains("molto diverso")]
                _kpi(a1, "Imposte su articoli senza AIEM", f"{len(x)} articoli", f"{_eur(x['extra_eur'].sum(), True)} caricati in più", "#B42318")
                _kpi(a2, "AIEM non registrato", f"{len(y)} articoli", f"{_eur(y['extra_eur'].sum(), True)} rispetto all'atteso", "#B54708")
                _kpi(a3, "AIEM molto diverso dall'atteso", f"{len(z)} articoli", "più del doppio o meno della metà")
                st.write("")
                tipo = st.selectbox("Tipo di anomalia", ["Tutte"] + sorted(ri["tipo"].unique()))
                vis = ri if tipo == "Tutte" else ri[ri["tipo"] == tipo]
                tb = vis.rename(columns={"item": "Articolo", "desc": "Descrizione", "tipo": "Anomalia", "entrate_anomale": "Entrate anomale",
                                         "entrate_totali": "Entrate nel periodo", "extra_eur": "Imposte vs atteso",
                                         "ultima": "Ultima entrata", "aliquota": "AIEM % anagrafica",
                                         "spedizioni": "External Document No_ (entrate anomale)"})
                ta = tb[["Articolo", "Descrizione", "Anomalia", "External Document No_ (entrate anomale)", "AIEM % anagrafica",
                         "Entrate anomale", "Entrate nel periodo", "Imposte vs atteso", "Ultima entrata"]]
                ta = filtri.filtra(ta, "aiem")
                sty_a = ta.style.format({"Imposte vs atteso": "{:+,.0f} €", "AIEM % anagrafica": "{:.0f}%", "Ultima entrata": "{:%d/%m/%Y}"})
                # verde = aliquota prevista in anagrafica; rosso = imposte registrate diverse dall'atteso
                sty_a = sty_a.map(lambda v: VERDE_C, subset=["AIEM % anagrafica"])
                sty_a = sty_a.map(lambda v: ROSSO_C if pd.notna(v) and abs(v) >= 0.5 else "", subset=["Imposte vs atteso"])
                _tabella_html(sty_a, a_sinistra=("Articolo", "Descrizione", "Anomalia", "External Document No_ (entrate anomale)"),
                              altezza=560, a_capo=("Descrizione", "Anomalia", "External Document No_ (entrate anomale)"))
                with st.expander("Dettaglio per entrata"):
                    d = aiem if tipo == "Tutte" else aiem[aiem["tipo"] == tipo]
                    st.dataframe(d.sort_values("data", ascending=False).rename(columns={
                        "item": "Articolo", "desc": "Descrizione", "data": "Data", "ubic": "Ubicazione", "ext": "Spedizione (fattura)",
                        "doc": "Doc.", "qty": "Q.tà", "uom": "UM", "aliquota": "AIEM %", "imp_reg": "Imposte reg.",
                        "imp_att": "Imposte attese", "extra_eur": "Δ", "tipo": "Anomalia"}).style.format(
                        {"Data": "{:%d/%m/%Y}", "Imposte reg.": "{:.3f} €", "Imposte attese": "{:.3f} €", "Δ": "{:+,.0f} €",
                         "Q.tà": "{:,.0f}"}), hide_index=True, width="stretch")

    # ================== container =========================================================
    with _cx(t_cont):
        if cont is None:
            st.info(needs)
        else:
            st.markdown("#### Ripartizione dei costi sulle spedizioni")
            st.caption(("Ogni spedizione è un documento di acquisto (Document No_): tutte le righe che lo condividono viaggiano insieme. "
                        if filiale != "CANARIE" else
                        "Il container si isola con il «External Document No_» (il numero fattura IFB: tutte le righe che lo condividono "
                        "viaggiano insieme). ") + " Si confronta il trasporto registrato su ogni riga con la quota di volume attesa "
                       "(box × volume di un box, con box = pezzi ÷ pezzi per box). Verificabile per gli articoli con le misure "
                       "del box nel modello (IFB) o nell'anagrafica " + _NC["c"] + " (tutti gli altri).")
            # i dati di imballo del modello devono coincidere con l'anagrafica: lo verifico sempre
            dif_ana = pd.DataFrame(columns=["item", "campo", "modello", "anagrafica"])
            if p_ana and ifb:
                dif_ana, n_conf, n_ass = mf.confronta_anagrafica(ifb, ana)
                nome_ana = os.path.basename(p_ana)
                if not dif_ana.empty:          # nessun messaggio se tutto coincide: si segnala solo la discrepanza
                    st.warning(f"Dati di imballo diversi tra modello e anagrafica «{nome_ana}»: {dif_ana['item'].nunique()} articoli "
                               f"su {n_conf} confrontati. Il volume delle spedizioni con questi articoli può essere sbagliato.")
                    with st.expander("Articoli con dati diversi"):
                        st.dataframe(dif_ana.rename(columns={"item": "Articolo", "campo": "Dato", "modello": "Modello",
                                                              "anagrafica": "Anagrafica"}), hide_index=True, width="stretch")
            if cont.empty:
                st.info("Nessuna spedizione nel periodo.")
            else:
                q_art = st.text_input("Cerca un articolo", placeholder="codice o parte della descrizione: vedi le spedizioni in cui "
                                      "compare e il trasporto che ha preso", key="mf_sped_q")
                q = q_art.strip().lower()
                righe_q = None
                if q:
                    righe_q = w[w["item"].str.lower().str.contains(q, regex=False)
                                | w["desc"].astype(str).str.lower().str.contains(q, regex=False)]
                    cont = cont[cont["ext"].isin(righe_q["ext"])]
                    if cont.empty:
                        st.info("Nessuna spedizione del mese contiene questo articolo.")
                vc = cont["esito"].value_counts()
                n_bug = int(vc.get("Volume del collo x pezzi (manca la divisione per i pezzi per collo)", 0))
                n_ok = int(vc.get("Ripartito a volume correttamente", 0))
                n_nc = int(vc.get("Ripartizione non coerente con il volume", 0))
                hk = filiale != "CANARIE"          # Hong Kong: niente AIEM ne' despachos
                c1, c2, c3, c4 = st.columns(4)
                _kpi(c1, "Spedizioni nel periodo", len(cont), "nel mese" if hk else f"{int(cont['stessa_chiave'].sum())} con imposte ripartite come il trasporto")
                _kpi(c2, "A volume corretto", n_ok, "ripartizione attesa", "#027A48")
                _kpi(c3, "Errore «collo × pezzi»", n_bug, "manca la divisione per i pezzi per collo", "#B42318" if n_bug else INK)
                _kpi(c4, "Non coerenti", n_nc, "né volume corretto né errore noto", "#B54708" if n_nc else INK)
                st.write("")
                f = st.selectbox("Mostra", ["Tutte le spedizioni", "Spedizioni con anomalie"], key="mf_mostra_sped")
                import assistente
                assistente.bolla("mf_mostra_sped", "Quali spedizioni?",
                                 ("<b>Con anomalie</b> = solo quelle in cui il FOB è stato registrato diverso dall'atteso " if filiale != "CANARIE" else
                                  "<b>Con anomalie</b> = solo quelle in cui trasporto o AIEM sono stati registrati diversi dall'atteso ") +
                                 "per almeno 1 €. Le altre sono corrette o non verificabili.", larghezza=420)
                vis = cont if f == "Tutte le spedizioni" else cont[cont["impatto_eur"] >= 1]
                vis = vis.sort_values("impatto_eur", ascending=False)      # dalla problematica maggiore
                vis = vis.assign(motivo=[_motivo(r) for _, r in vis.iterrows()])
                colonne = ["Spedizione (fattura)", "Ultima entrata", "Ubicazione", "Righe"]
                if righe_q is not None and len(righe_q):
                    agg = righe_q.assign(_t=righe_q["trasp"] * righe_q["qty"]).groupby("ext").agg(
                        q_art=("qty", "sum"), t_art=("_t", "sum"), um_art=("uom", "first")).reset_index()
                    agg["t_unit"] = agg["t_art"] / agg["q_art"]
                    vis = vis.merge(agg, on="ext", how="left")
                    colonne += ["Q.tà articolo", "UM", "Trasporto sull'articolo", "Trasporto per UM"]
                colonne += ["Valore merce", "Trasporto totale", "Trasporto su merce", "Despachos totali", "Imposte totali", "Esito", "Motivo",
                            "Impatto stimato", "Costo spostato tra righe", "Imposte su articoli senza AIEM",
                            "Imposte ripartite come il trasporto"]
                if hk:
                    colonne = [c for c in colonne if c not in ("Despachos totali", "Imposte totali", "Imposte su articoli senza AIEM",
                                                               "Imposte ripartite come il trasporto")]
                tb = vis.rename(columns={"ext": "Spedizione (fattura)", "data": "Ultima entrata", "ubic": "Ubicazione",
                                         "righe": "Righe", "trasporto_eur": "Trasporto totale", "esito": "Esito",
                                         "despachos_eur": "Despachos totali", "imposte_eur": "Imposte totali",
                                         "merce_eur": "Valore merce", "trasp_pct": "Trasporto su merce",
                                         "stessa_chiave": "Imposte ripartite come il trasporto",
                                         "imposte_su_senza_aiem": "Righe con imposte su articoli senza AIEM",
                                         "impatto_eur": "Impatto stimato", "imposte_indebite_eur": "Imposte su articoli senza AIEM",
                                         "sbilancio_eur": "Costo spostato tra righe", "motivo": "Motivo", "q_art": "Q.tà articolo", "um_art": "UM",
                                         "t_art": "Trasporto sull'articolo", "t_unit": "Trasporto per UM"})
                with st.expander("Elenco delle spedizioni"):
                    tb = tb.assign(**{"Imposte ripartite come il trasporto": tb["Imposte ripartite come il trasporto"].map(
                        lambda v: "Sì" if v else "")})
                    tb = tb.loc[filtri.filtra(tb[colonne], "sped").index]
                    sty_el = tb[colonne].style.format(
                        {k: v for k, v in {"Ultima entrata": "{:%d/%m/%Y}", "Valore merce": "{:,.0f} €", "Trasporto totale": "{:,.0f} €",
                         "Trasporto su merce": "{:.1%}", "Despachos totali": "{:,.0f} €", "Imposte totali": "{:,.0f} €",
                         "Costo spostato tra righe": "{:,.0f} €", "Impatto stimato": "{:,.0f} €",
                         "Imposte su articoli senza AIEM": "{:,.0f} €", "Q.tà articolo": "{:,.0f}",
                         "Trasporto sull'articolo": "{:,.0f} €", "Trasporto per UM": "{:.3f} €"}.items() if k in colonne}, na_rep="")
                    _tabella_html(sty_el, a_sinistra=("Spedizione (fattura)", "Ubicazione", "Esito", "Motivo", "UM"),
                                  a_capo=("Spedizione (fattura)", "Esito", "Motivo"))
                if len(vis):
                    scelta = st.selectbox("Dettaglio di una spedizione", list(vis["ext"]), key="mf_sped_sel")
                    import assistente
                    assistente.bolla("mf_sped_sel", "Guarda dentro la spedizione",
                                     "Sotto trovi il disegno del container (come doveva essere diviso e come è stato diviso), "
                                     "la tabella articolo per articolo e il testo pronto per la segnalazione a " + _NC["c"] + ".", larghezza=440)
                    tt = cont[cont["ext"] == scelta].iloc[0]
                    st.markdown((f"**Spedizione {scelta}**: FOB totale **{_eur(tt['trasporto_eur'])}** · altri oneri "
                                 f"{_eur(tt['despachos_eur'])} · valore merce " if hk else
                                 f"**Spedizione {scelta}**: trasporto totale **{_eur(tt['trasporto_eur'])}** · despachos "
                                 f"{_eur(tt['despachos_eur'])} · imposte {_eur(tt['imposte_eur'])} · valore merce ")
                                + f"{_eur(tt['merce_eur'])} · {int(tt['righe'])} righe · ingresso in magazzino "
                                + _date_ingresso(w[w["ext"] == scelta]["data"]))
                    if hk:          # il tasso di cambio non e' nei report: lo scrive l'utente e resta salvato per questa spedizione
                        _t0 = mf.leggi_tassi().get(str(scelta), 0.0)
                        _lst = _listino(p_modello, os.path.getmtime(p_modello)) if (p_modello and os.path.exists(p_modello)) else {}
                        _ws = w[w["ext"] == scelta]
                        _rap = [r_.prezzo / _lst[r_.item][0] for r_ in _ws.itertuples()
                                if r_.item in _lst and _lst[r_.item][0] > 0 and _lst[r_.item][1] in ("", str(r_.uom).upper())]
                        _imp = float(np.median(_rap)) if _rap else 0.0
                        _tc, _ = st.columns([1, 3])
                        _tv = _tc.number_input("Tasso di cambio usato alla registrazione dell'entrata"
                                               + (f" (dalle righe: {_imp:.4f})" if _imp else ""), min_value=0.0,
                                               value=float(_t0 or round(_imp, 4)),
                                               step=0.0001, format="%.4f", key=f"mf_tasso_{scelta}",
                                               help="Il cambio è quello del giorno in cui l'entrata è registrata (via mare circa un mese dopo "
                                                    "la partenza, via aerea pochi giorni). Lascia 0 se non lo conosci.")
                        if abs(_tv - (_t0 or round(_imp, 4))) > 1e-9:
                            mf.salva_tasso(scelta, _tv)
                    d_sped = det[det["ext"] == scelta].copy() if len(det) else pd.DataFrame()
                    # righe della spedizione che non entrano nel controllo del volume: lo dico, con il motivo
                    w_sp = w[w["ext"] == scelta].drop_duplicates("item")
                    gia = set(d_sped["item"]) if len(d_sped) else set()
                    fuori = []
                    for r in w_sp.itertuples():
                        if r.item in gia:
                            continue
                        motivo = ("mancano le misure del box (né nel modello né nell'anagrafica)" if r.item in dati_art
                                  else "articolo senza dati di imballo")
                        fuori.append(f"{r.item} {str(r.desc).strip()[:30]} ({motivo})")
                    if fuori and len(d_sped):
                        st.caption("Righe non verificabili sul volume, quindi fuori dal disegno e dalla tabella: " + "; ".join(fuori) + ".")
                    if len(d_sped):
                        with st.expander("Disegno del container: come dovrebbe essere diviso il FOB e come è stato diviso" if hk else
                                         "Disegno del container: come dovrebbero essere divisi trasporto e AIEM e come sono stati divisi"):
                            _disegno_container(d_sped, tt["trasporto_eur"], con_aiem=not hk)
                    if len(d_sped):
                        d = d_sped.copy()
                        d["pz_scatola"] = d["qty"] / d["scatole"]
                        d["vol_scatola"] = d["v_ok"] / d["scatole"]          # m³
                        d["quota_vol"] = d["v_ok"] / d["v_ok"].sum()
                        d["quota_reg"] = d["reg"] / d["reg"].sum()
                        d["tr_reg_u"] = d["reg"] / d["qty"]
                        d["tr_att_u"] = d["atteso"] / d["qty"]
                        d["ai_reg_u"] = d["imp_reg"] / d["qty"]
                        d["ai_att_u"] = d["aiem_att"] / d["qty"]
                        d["de_reg_u"] = d["desp_reg"] / d["qty"]
                        d["de_att_u"] = d["desp_att"] / d["qty"]
                        d["tot_r_u"] = d["tot_reale"] / d["qty"]
                        d["tot_a_u"] = d["tot_atteso"] / d["qty"]
                        d["aliq"] = d["aliquota"].where(d["aliquota"] >= 0)
                        d = d.rename(columns={"item": "Articolo", "desc": "Descrizione", "uom": "UM", "qty": "Q.tà",
                                              "pz_scatola": "UM x box", "scatole": "Box", "vol_scatola": "Volume box (m³)",
                                              "v_ok": "Volume totale (m³)", "quota_vol": "Quota di volume", "data": "Data ingresso",
                                              "quota_reg": "Quota registrata", "reg": "Trasporto reale", "atteso": "Trasporto atteso",
                                              "sbilancio": "Δ trasporto", "aliq": "AIEM %", "imp_reg": "AIEM reale",
                                              "aiem_att": "AIEM atteso", "aiem_diff": "Δ AIEM",
                                              "tot_reale": "Totale reale", "tot_atteso": "Totale atteso",
                                              "tr_reg_u": "Trasporto reale per UM", "tr_att_u": "Trasporto atteso per UM",
                                              "ai_reg_u": "AIEM reale per UM", "ai_att_u": "AIEM atteso per UM", "de_reg_u": "Despachos reale per UM", "de_att_u": "Despachos atteso per UM", "tot_r_u": "Totale reale per UM", "tot_a_u": "Totale atteso per UM",
                                              "merce": "Purchase price", "desp_reg": "Despachos reale", "desp_att": "Despachos atteso"})
                        d = d.sort_values("Δ trasporto", key=abs, ascending=False)
                        if hk:          # peso e quota di peso: le spedizioni aeree si ripartiscono a kg
                            _ws2 = w[w["ext"] == scelta]
                            _pk = pd.Series({r_.item: mf.peso_riga(r_.item, r_.qty, r_.uom, ana) for r_ in _ws2.itertuples()})
                            _pk = _ws2.assign(_p=[mf.peso_riga(r_.item, r_.qty, r_.uom, ana) for r_ in _ws2.itertuples()]).groupby("item")["_p"].sum()
                            d["Peso (kg)"] = d["Articolo"].map(_pk)
                            d["Quota di peso"] = d["Peso (kg)"] / d["Peso (kg)"].sum() if d["Peso (kg)"].sum() > 0 else np.nan
                        if hk:          # oneri di Hong Kong: certificato sanitario, altri oneri, dogana, pallet (totali della riga)
                            ws_ = w[w["ext"] == scelta]
                            ch_ = pd.DataFrame({"Cert. sanitario": ws_["c_salute"] * ws_["qty"], "Other Charge": ws_["c_altro"] * ws_["qty"],
                                                "Other Custom Charge": ws_["c_dogana"] * ws_["qty"], "Pallet": ws_["c_pallet"] * ws_["qty"],
                                                "item": ws_["item"]}).groupby("item").sum()
                            d = d.join(ch_, on="Articolo")
                        cols = ["Articolo", "Descrizione", "UM", "Q.tà", "UM x box", "Box", "Volume box (m³)", "Volume totale (m³)",
                                "Quota di volume", "Quota registrata", "Trasporto reale", "Trasporto atteso", "Δ trasporto",
                                "Trasporto reale per UM", "Trasporto atteso per UM",
                                "AIEM %", "AIEM reale", "AIEM atteso", "Δ AIEM", "AIEM reale per UM", "AIEM atteso per UM",
                                "Purchase price", "Despachos reale", "Despachos atteso", "Despachos reale per UM", "Despachos atteso per UM", "Totale reale per UM", "Totale atteso per UM"]
                        if hk:
                            cols = ["Articolo", "Descrizione", "UM", "Q.tà", "UM x box", "Box", "Volume box (m³)", "Volume totale (m³)",
                                    "Quota di volume", "Peso (kg)", "Quota di peso", "Quota registrata", "Trasporto reale", "Trasporto atteso", "Δ trasporto",
                                    "Trasporto reale per UM", "Trasporto atteso per UM", "Cert. sanitario", "Other Charge",
                                    "Other Custom Charge", "Pallet", "Purchase price", "Totale reale per UM", "Totale atteso per UM"]
                        tab = d[cols].copy()
                        tot_riga = {c: None for c in cols}
                        tot_riga.update({"Articolo": "TOTALE", "Box": d["Box"].sum(), "Volume totale (m³)": d["Volume totale (m³)"].sum(),
                                         "Quota di volume": d["Quota di volume"].sum(), "Quota registrata": d["Quota registrata"].sum(),
                                         "Trasporto reale": d["Trasporto reale"].sum(), "Trasporto atteso": d["Trasporto atteso"].sum(),
                                         "AIEM reale": d["AIEM reale"].sum(),
                                         "AIEM atteso": d["AIEM atteso"].sum(), 
                                         "Purchase price": d["Purchase price"].sum(), "Despachos reale": d["Despachos reale"].sum(),
                                         "Despachos atteso": d["Despachos atteso"].sum()})
                        if hk:
                            for c_ in ("Cert. sanitario", "Other Charge", "Other Custom Charge", "Pallet"):
                                tot_riga[c_] = d[c_].sum()
                        tot_riga = {k: v for k, v in tot_riga.items() if k in cols}
                        tab = pd.concat([tab, pd.DataFrame([tot_riga])], ignore_index=True)
                        for c in tab.columns:                        # niente "None" nella riga dei totali
                            if c == "Data ingresso":
                                tab[c] = tab[c].map(lambda v: "" if pd.isna(v) else f"{v:%d/%m/%Y}")
                            elif c in ("Articolo", "Descrizione", "UM"):
                                tab[c] = tab[c].fillna("")
                            else:
                                tab[c] = pd.to_numeric(tab[c], errors="coerce")
                        sty = tab.style.format(
                            {k: v for k, v in {"Peso (kg)": "{:,.1f}", "Quota di peso": "{:.1%}", "Cert. sanitario": "{:,.0f} €", "Other Charge": "{:,.0f} €", "Other Custom Charge": "{:,.0f} €",
                                               "Pallet": "{:,.0f} €", "Q.tà": "{:,.0f}", "UM x box": "{:,.0f}", "Box": "{:,.0f}", "Volume box (m³)": "{:.4f}",
                             "Volume totale (m³)": "{:.2f}", "Quota di volume": "{:.1%}", "Quota registrata": "{:.1%}",
                             "Trasporto reale": "{:,.0f} €", "Trasporto atteso": "{:,.0f} €", "Δ trasporto": "{:+,.0f} €",
                             "AIEM %": "{:.0f}%", "AIEM reale": "{:,.0f} €", "AIEM atteso": "{:,.0f} €", "Δ AIEM": "{:+,.0f} €",
                             "Trasporto reale per UM": "{:.3f} €", "Trasporto atteso per UM": "{:.3f} €",
                             "Totale reale per UM": "{:.3f} €", "Totale atteso per UM": "{:.3f} €", "AIEM reale per UM": "{:.3f} €", "AIEM atteso per UM": "{:.3f} €",
                             "Purchase price": "{:,.0f} €", "Despachos reale": "{:,.2f} €", "Despachos atteso": "{:,.2f} €", "Despachos reale per UM": "{:.4f} €", "Despachos atteso per UM": "{:.4f} €"}.items() if k in tab.columns}, na_rep="")
                        sty = sty.map(_evidenzia, subset=pd.IndexSlice[tab.index[:-1], [c_ for c_ in ("Δ trasporto", "Δ AIEM") if c_ in tab.columns]])
                        # verde = come dovrebbe essere; rosso = com'e' stato registrato, se si scosta dall'atteso
                        verdi_c = ["Quota di volume", "Trasporto atteso", "AIEM atteso", "Totale atteso", "Trasporto atteso per UM", "AIEM atteso per UM", "Despachos atteso", "Despachos atteso per UM"]
                        rossi_t = ["Quota registrata", "Trasporto reale", "Trasporto reale per UM"]
                        rossi_a = ["AIEM reale", "AIEM reale per UM"]
                        rossi_d = ["Despachos reale", "Despachos reale per UM"]
                        soglia_t = 0.5        # rosso per qualsiasi differenza dall'atteso (sotto 0,5 € sono arrotondamenti)
                        soglia_a = 0.5
                        ultimo = tab.index[-1]

                        def colora(r):
                            if r.name == ultimo:
                                return [""] * len(r)
                            err_t = abs(r["Δ trasporto"]) >= soglia_t
                            err_tot = abs(r["Totale reale per UM"] - r["Totale atteso per UM"]) * r["Q.tà"] >= soglia_t
                            err_a = "Δ AIEM" in r.index and pd.notna(r["Δ AIEM"]) and abs(r["Δ AIEM"]) >= soglia_a
                            err_d = "Despachos reale" in r.index and abs(r["Despachos reale"] - r["Despachos atteso"]) >= 0.05      # importi piccoli: la soglia di 0,5 € nasconderebbe le differenze
                            rosso = "background-color:#FEE4E2;color:#B42318;font-weight:600"
                            ok = "background-color:#F6FEF9;color:#05603A"
                            out = []
                            for c in r.index:
                                if c in ("Totale atteso", "Totale atteso per UM"):                      # totali: testo nero e grassetto, sfondo verde/rosso
                                    out.append("background-color:#D1FADF;color:#000000;font-weight:700")
                                elif c in ("Totale reale", "Totale reale per UM"):
                                    out.append(("background-color:#FEE4E2" if err_tot else "background-color:#F6FEF9")
                                               + ";color:#000000;font-weight:700")
                                elif c in verdi_c:
                                    out.append("background-color:#D1FADF;color:#05603A")
                                elif c in rossi_t:
                                    out.append(rosso if err_t else ok)
                                elif c == "Totale reale":
                                    out.append(rosso if err_tot else ok)
                                elif c in rossi_d:
                                    out.append(rosso if err_d else ok)
                                elif c in rossi_a:
                                    out.append(rosso if err_a else ok)
                                else:
                                    out.append("")
                            return out

                        sty = sty.apply(colora, axis=1)
                        sty = sty.apply(lambda r: ["font-weight:700;background-color:#F2F4F7" if r.name == ultimo else ""
                                                   for _ in r], axis=1)
                        if q:
                            trovata = (tab["Articolo"].astype(str).str.lower().str.contains(q, regex=False)
                                       | tab["Descrizione"].astype(str).str.lower().str.contains(q, regex=False))
                            gialle = ["Articolo", "Descrizione", "UM", "Q.tà", "UM x box", "Box"]
                            sty = sty.apply(lambda r: ["background-color:#FEF9C3" if (r.name < ultimo and trovata.iloc[r.name]
                                                                                       and c in gialle) else "" for c in r.index], axis=1)
                        with st.expander("Tabella di dettaglio"):
                            # tabella HTML: le intestazioni vanno a capo quando la finestra si restringe (st.dataframe le taglia)
                            st.markdown(
                                "<style>.tab-det{overflow:auto;resize:both;max-height:78vh;min-height:90px;border:1px solid #E4E7EC;border-radius:10px}.tab-det table{border-collapse:collapse;font-size:.82rem;width:100%}"
                                ".tab-det th{position:sticky;top:0;z-index:2;white-space:normal;text-align:center;vertical-align:bottom;background:#F9FAFB;color:#667085;"
                                "font-weight:500;padding:6px 8px;border-bottom:1px solid #E4E7EC;min-width:58px;line-height:1.2}"
                                ".tab-det td{white-space:nowrap;padding:5px 8px;border-bottom:1px solid #F2F4F7;text-align:right}"
                                ".tab-det td.col0,.tab-det td.col2{text-align:left}"
                                ".tab-det td.col1{white-space:normal;min-width:170px;text-align:left}"
                                ".tab-det th.col1{text-align:left}</style><div class='tab-det'>"
                                + sty.hide(axis="index").to_html() + "</div>", unsafe_allow_html=True)
                            scarica_excel.bottone(sty, indice=False, nome=f"spedizione_{scelta}_dettaglio.xlsx")
                            st.caption(
                                "**Box** = pezzi ÷ pezzi per box (per i prodotti a peso: kg ÷ kg per box; dati dall'anagrafica). "
                                "**Volume totale** = box × volume di un box (altezza × lunghezza × profondità). "
                                "**Trasporto atteso** = trasporto totale della spedizione × quota di volume dell'articolo; **Trasporto reale** = "
                                "quanto è stato registrato sulla riga. **Totale** = prezzo di acquisto + trasporto + AIEM + despachos (reale come "
                                "registrato; atteso con trasporto e despachos divisi a volume). **AIEM atteso** = aliquota AIEM dell'articolo × (valore della merce + "
                                "trasporto atteso della riga), come nella fattura delle imposte; **AIEM reale** = imposte registrate sulla riga. "
                                "**Δ** = reale − atteso: positivo = riga sovraccaricata, negativo = sottocaricata.")
                            import pdf_report
                            _intro = [f"<b>Trasporto totale</b> {_eur(tt['trasporto_eur'])} · <b>despachos</b> {_eur(tt['despachos_eur'])} · "
                                      f"<b>imposte</b> {_eur(tt['imposte_eur'])} · <b>valore merce</b> {_eur(tt['merce_eur'])} · "
                                      f"{int(tt['righe'])} righe · ingresso in magazzino " + _date_ingresso(w[w["ext"] == scelta]["data"]),
                                      "Verde = come dovrebbe essere (diviso a volume) · rosso = come è stato registrato, se si scosta dall'atteso."]
                            _nota = ("Box = pezzi ÷ pezzi per box (per i prodotti a peso: kg ÷ kg per box). Volume totale = box × volume di un box. "
                                     "Trasporto atteso = trasporto totale della spedizione × quota di volume dell'articolo; Trasporto reale = quanto è stato "
                                     "registrato sulla riga. Totale = prezzo di acquisto + trasporto + AIEM + despachos. AIEM atteso = aliquota × (valore "
                                     "della merce + trasporto atteso della riga). Δ = reale − atteso: positivo = riga sovraccaricata, negativo = sottocaricata.")
                            st.download_button("Estrai in PDF (per la segnalazione)", pdf_report.tabella_pdf(
                                f"Spedizione {scelta}: dettaglio per articolo", _intro, tab, sty, _nota,
                                a_sinistra=("Articolo", "Descrizione", "Data ingresso", "UM"),
                                larghe={"Articolo": 28, "Descrizione": 82, "Data ingresso": 36, "UM": 22, "Despachos reale": 34, "Despachos atteso": 34}),
                                file_name=f"spedizione_{scelta}.pdf", mime="application/pdf", key="mf_pdf_sped")
                    if len(d_sped) and len(dif_ana):
                        sospetti = dif_ana[dif_ana["item"].isin(d_sped["item"])]
                        for r in sospetti.itertuples():
                            st.markdown(f'<div style="color:#B42318;background:#FEF3F2;border:1px solid #FECDCA;border-radius:10px;'
                                        f'padding:8px 12px;margin:6px 0;font-size:.93rem">Attenzione: per l\'articolo {r.item} il dato '
                                        f'«{r.campo}» nel modello ({r.modello}) è diverso dall\'anagrafica ({r.anagrafica}): '
                                        f'il volume di questa riga potrebbe essere sbagliato.</div>', unsafe_allow_html=True)
                    # cosa non va: spiegazione in rosso e testo per la segnalazione
                    problemi = mf.analisi_spedizione(tt, d_sped, w[w["ext"] == scelta], dati_art)
                    for p in problemi:
                        righe_es = "".join(f"<li>{e}</li>" for e in p["esempio"])
                        st.markdown(
                            '<div style="background:#FEF3F2;border:1px solid #FECDCA;border-radius:12px;padding:12px 16px;margin:8px 0;'
                            'color:#14213D;font-size:.93rem;line-height:1.45">'
                            f'<div style="color:#B42318;font-weight:600;font-size:1rem;margin-bottom:4px">⚠ {p["titolo"]}</div>'
                            f'<b>Cosa dovrebbe succedere.</b> {p["atteso"]}<br>'
                            f'<b>Cosa abbiamo trovato.</b> {p["trovato"]}'
                            + (f'<ul style="margin:4px 0 4px 18px">{righe_es}</ul>' if righe_es else '<br>')
                            + f'<b>Effetto.</b> {p["effetto"]}<br><b>Cosa chiedere.</b> {p["richiesta"]}</div>',
                            unsafe_allow_html=True)
                    if problemi:
                        with st.expander("Testo pronto per la segnalazione " + ("a COMIT" if not hk else "al cliente")):
                            st.code(mf.testo_segnalazione(tt, problemi, hk=hk, tasso=mf.leggi_tassi().get(str(scelta)) if hk else None), language=None)
                    else:
                        st.caption("Nessuna anomalia rilevata su questa spedizione.")

    # ================== trasporto =========================================================
    with _cx(t_tr):
        if trasp is None:
            st.info(needs)
        else:
            st.markdown("#### Trasporto: oscillazioni e scostamenti dallo standard")
            tol_cv, tol_r = 0.40, 2.0        # soglie fisse: oscillazione tra entrate 40%, scostamento dallo standard 2x
            tr = mf.controllo_trasporto(med, ifb, non, tol_cv, tol_r)
            st.caption("Piccole differenze sono normali (mix di merce diverso nella spedizione, due camion invece di uno, tratta più "
                       "cara): si segnalano solo oscillazioni oltre la soglia o un trasporto reale molto lontano dalla logistica dello "
                       "standard. «Reale» = trasporto + despachos medi per unità nelle entrate del periodo.")
            if tr.empty:
                st.success("Nessun articolo sopra le soglie.")
            else:
                nm = a.set_index("item")["desc"].to_dict()
                tr["desc"] = tr["item"].map(nm)
                tr["vend"] = tr["item"].map(a.set_index("item")["vendite"])
                tr = tr.sort_values("vend", ascending=False, na_position="last")
                tt_ = tr[["item", "desc", "vend", "entrate", "trasp_min", "trasp_max", "reale", "std_logistica", "rapporto", "segnale"]].rename(
                    columns={"item": "Articolo", "desc": "Descrizione", "vend": "Vendite mese", "entrate": "Entrate",
                             "trasp_min": "Trasp. min", "trasp_max": "Trasp. max", "reale": "Reale",
                             "std_logistica": "Logistica STD", "rapporto": "Reale / STD", "segnale": "Segnale"})
                for c in ("Vendite mese", "Trasp. min", "Trasp. max", "Reale", "Logistica STD", "Reale / STD"):
                    tt_[c] = pd.to_numeric(tt_[c], errors="coerce")
                tt_ = filtri.filtra(tt_, "trasp")
                sty_t = tt_.style.format({"Vendite mese": "{:,.0f} €", "Trasp. min": "{:.3f} €", "Trasp. max": "{:.3f} €",
                                          "Reale": "{:.3f} €", "Logistica STD": "{:.3f} €", "Reale / STD": "{:.1f}x"}, na_rep="")
                # verde = come dovrebbe essere (standard); rosso = com'e' stato registrato; ambra = trasporto instabile tra le entrate
                sty_t = sty_t.map(lambda v: VERDE_C, subset=["Logistica STD"])
                sty_t = sty_t.map(lambda v: ROSSO_C, subset=["Reale", "Reale / STD"])
                sty_t = sty_t.apply(lambda r: [AMBRA_C if (c in ("Trasp. min", "Trasp. max") and "instabile" in str(r["Segnale"]))
                                               else "" for c in r.index], axis=1)
                _tabella_html(sty_t, a_sinistra=("Articolo", "Descrizione", "Segnale"), altezza=620, a_capo=("Descrizione", "Segnale"))

    # ================== un articolo =======================================================
    with _cx(t_art):
        q = st.text_input("Articolo", placeholder="codice o parte della descrizione", key="mf_art")
        cand = a if not q else a[a["item"].str.lower().str.contains(q.strip().lower(), regex=False)
                                 | a["desc"].astype(str).str.lower().str.contains(q.strip().lower(), regex=False)]
        if cand.empty:
            st.info("Scrivi un codice o una parte della descrizione (articoli presenti nel report dei margini).")
        else:
            cand = cand.sort_values("diff", key=abs, ascending=False)
            sel = st.selectbox("Risultati", list(cand["item"]), format_func=lambda k: f"{k} · {cand.set_index('item').loc[k, 'desc']}")
            r = a[a["item"] == sel].iloc[0]
            _kpi_cols = st.columns(4)
            _kpi(_kpi_cols[0], "Vendite", _eur(r["vendite"]), _it(f"{r['qty']:,.0f}") + f" {r.get('uom', '')} · {r['perimetro']}")
            _kpi(_kpi_cols[1], "STD / FIFO", _it(f"{r['c_std']:.3f} / {r['c_fifo']:.3f}"), "costo unitario nel mese")
            _kpi(_kpi_cols[2], "Margine STD / FIFO", _it(f"{r['m_std'] / r['vendite']:.1%} / {r['m_fifo'] / r['vendite']:.1%}"), "")
            _kpi(_kpi_cols[3], "Δ margine", _eur(r["diff"], True), _pct(r["diff_pct"]), "#B42318" if r["diff"] > 0 else "#027A48")
            st.write("")
            if ent is None:
                st.info(needs)
            else:
                e_it = ent[ent["item"] == sel].sort_values("data", ascending=False)
                if e_it.empty:
                    st.warning("Nessuna entrata di questo articolo nel file Cost Analysis.")
                else:
                    dg1 = mf.diagnosi(pd.DataFrame([r]), med, aiem, trasp, det, ifb, non)
                    if sel in dg1.index and dg1.loc[sel, "cause"]:
                        st.markdown("**Cause probabili:** " + dg1.loc[sel, "cause"].replace(" | ", " · "))
                    isola = e_it["isola"].mode().iat[0] if e_it["isola"].notna().any() else None
                    s = mf.std_di(sel, isola, ifb, non)
                    if s is not None and sel in med.index:
                        m = med.loc[sel]
                        comp = pd.DataFrame({
                            "Componente": ["Prezzo di acquisto", "Logistica (trasporto + despachos)", "AIEM / imposte", "Totale"],
                            "Standard": [s["prezzo"], s["logistica"], s["aiem"], s["std"]],
                            "Entrate (media)": [m["e_prezzo"], m["e_trasp"] + m["e_desp"], m["e_imp"], m["e_tot"]]})
                        comp["Differenza"] = comp["Entrate (media)"] - comp["Standard"]
                        st.dataframe(comp.style.format({"Standard": "{:.3f} €", "Entrate (media)": "{:.3f} €",
                                                        "Differenza": "{:+.3f} €"}).map(_evidenzia, subset=["Differenza"]),
                                     hide_index=True, width="stretch")
                        st.caption("Lo standard totale include la maggiorazione concordata (+2%), quindi supera la somma delle componenti.")
                    st.markdown("**Entrate** (più recenti in alto). «Pendiente» = quantità della partita non ancora consumata dal FIFO.")
                    show = e_it.rename(columns={"data": "Data", "ubic": "Ubicazione", "ext": "Spedizione (fattura)", "doc": "Doc.",
                                                "qty": "Q.tà", "pend": "Pendiente", "uom": "UM", "desp": "Despachos", "imp": "Imposte",
                                                "prezzo": "Prezzo", "trasp": "Trasporto", "tot": "Totale"})
                    st.dataframe(show[["Data", "Ubicazione", "Spedizione (fattura)", "Q.tà", "Pendiente", "UM", "Prezzo", "Trasporto",
                                       "Despachos", "Imposte", "Totale"]].style.format(
                        {"Data": "{:%d/%m/%Y}", "Q.tà": "{:,.0f}", "Pendiente": "{:,.0f}", "Prezzo": "{:.3f} €", "Trasporto": "{:.3f} €",
                         "Despachos": "{:.3f} €", "Imposte": "{:.3f} €", "Totale": "{:.3f} €"}), hide_index=True, width="stretch", height=360)

    # ================== storico ===========================================================
    if t_hk is not None:
        with t_hk:
            st.markdown("#### Costi d'entrata per articolo")
            if ent_hk is None:
                st.info("Carica il Cost Analysis nella pagina «File da caricare» per vedere il costo reale di entrata di ogni articolo.")
            else:
                ce = mf.costi_entrata_hk(ent_hk, mese_chiave)
                if ce.empty:
                    st.warning(f"Nel Cost Analysis non ci sono entrate di {mf.etichetta_mese(mese_chiave)}.")
                else:
                    st.caption(f"Costo medio per UM delle entrate di {mf.etichetta_mese(mese_chiave)}, ponderato sulle quantità. "
                               "Totale = prezzo di acquisto + FOB + certificato sanitario + altri oneri + dogana + pallet.")
                    ce = ce.join(a.set_index("item")[["c_std", "c_fifo"]], on="item")
                    # prezzo di acquisto registrato contro il listino del mese (foglio CURRENT PRICELIST del modello)
                    lst = _listino(p_modello, os.path.getmtime(p_modello)) if (p_modello and os.path.exists(p_modello)) else {}
                    ce["listino"] = ce["item"].map(lambda i: lst.get(i, (np.nan, ""))[0]).where(lambda s_: s_ > 0)
                    um_l = ce["item"].map(lambda i: lst.get(i, (np.nan, ""))[1])
                    # il prezzo registrato e' in valuta, il listino in euro: il rapporto e' il cambio usato alla registrazione
                    ce["d_lst"] = ce["prezzo"] / ce["listino"]
                    ce["nota_lst"] = np.where(ce["listino"].isna(), "non nel listino",
                                              np.where((um_l != "") & (um_l != ce["uom"].astype(str).str.upper()),
                                                       "UM diversa dal listino: confronto non attendibile", ""))
                    ce.loc[ce["nota_lst"].str.startswith("UM diversa"), "d_lst"] = np.nan
                    _cm = float(ce["d_lst"].median()) if ce["d_lst"].notna().any() else np.nan
                    ce["d_lst_pct"] = ce["d_lst"] / _cm - 1 if _cm == _cm else np.nan
                    ce["delta_std"] = ce["tot"] - ce["c_std"]
                    ce["delta_pct"] = ce["delta_std"] / ce["c_std"].where(ce["c_std"] > 0)
                    ce = ce.sort_values("delta_std", key=lambda s: s.abs(), ascending=False, na_position="last")
                    tb_h = ce.rename(columns={"item": "Articolo", "desc": "Descrizione", "uom": "UM", "entrate": "Entrate",
                                              "qty": "Q.tà", "prezzo": "Prezzo acquisto", "fob": "FOB", "salute": "Cert. sanitario",
                                              "altro": "Altri oneri", "dogana": "Dogana", "pallet": "Pallet", "tot": "Totale entrata",
                                              "c_std": "STD per UM", "c_fifo": "FIFO per UM", "delta_std": "Entrata − STD",
                                              "delta_pct": "Δ %", "ultima": "Ultima entrata", "listino": "Prezzo listino (€)", "d_lst": "Cambio implicito",
                                              "d_lst_pct": "Scost. dal cambio mediano", "nota_lst": "Nota listino"})[
                        ["Articolo", "Descrizione", "UM", "Entrate", "Q.tà", "Prezzo acquisto", "Prezzo listino (€)", "Cambio implicito",
                         "Scost. dal cambio mediano", "Nota listino", "FOB", "Cert. sanitario", "Altri oneri",
                         "Dogana", "Pallet", "Totale entrata", "STD per UM", "FIFO per UM", "Entrata − STD", "Δ %", "Ultima entrata"]]
                    tb_h = filtri.filtra(tb_h.reset_index(drop=True), "hk_entrate")
                    eur = lambda v: "" if pd.isna(v) else f"{v:,.3f} €"
                    sty_h = tb_h.style.format({"Q.tà": "{:,.0f}", "Prezzo acquisto": eur, "FOB": eur, "Cert. sanitario": eur,
                                               "Altri oneri": eur, "Dogana": eur, "Pallet": eur, "Totale entrata": eur,
                                               "STD per UM": eur, "FIFO per UM": eur, "Entrata − STD": lambda v: "" if pd.isna(v) else f"{v:+,.3f} €",
                                               "Δ %": lambda v: "" if pd.isna(v) else f"{v:+.1%}", "Ultima entrata": "{:%d/%m/%Y}",
                                               "Prezzo listino (€)": lambda v: "" if pd.isna(v) else f"{v:,.3f} EUR", "Cambio implicito": lambda v: "" if pd.isna(v) else f"{v:.4f}",
                                               "Scost. dal cambio mediano": lambda v: "" if pd.isna(v) else f"{v:+.1%}"})
                    # cambio implicito lontano dalla mediana degli articoli (oltre il 3%): prezzo registrato diverso dal listino
                    sty_h = sty_h.map(lambda v: ROSSO_C if pd.notna(v) and abs(v) >= 0.03 else "", subset=["Scost. dal cambio mediano"])
                    sty_h = sty_h.map(lambda v: VERDE_C if pd.notna(v) else "", subset=["Prezzo listino (€)"])
                    if _cm == _cm:
                        st.markdown(f"**Cambio mediano del mese: {_cm:.4f}** (prezzo registrato ÷ prezzo di listino in euro, sugli articoli "
                                    "con listino e stessa UM). In rosso gli articoli il cui rapporto si discosta di oltre il 3%: prezzo "
                                    "registrato diverso dal listino, o listino non aggiornato.")
                    sty_h = sty_h.map(lambda v: VERDE_C, subset=["STD per UM"]).map(lambda v: ROSSO_C, subset=["Totale entrata"])
                    _tabella_html(sty_h, a_sinistra=("Articolo", "Descrizione", "UM"), altezza=620, a_capo=("Descrizione",))
    with t_sto:
        sto = mf.leggi_storico()
        if not sto:
            st.info("Ancora nessun mese: lo storico si riempie da solo a ogni caricamento.")
        else:
            df = pd.DataFrame([dict(Mese=k, Vendite=v["vendite"], **{"Margine STD": v["p_std"], "Margine FIFO": v["p_fifo"],
                                                                        "Scostamento": v["diff_pct"], "Δ margine": v["diff_eur"]},
                                    **{f"Scost. {o}": x for o, x in v.get("perimetri", {}).items()}) for k, v in sto.items()])
            st.dataframe(df.style.format({"Vendite": "{:,.0f} €", "Margine STD": "{:.1%}", "Margine FIFO": "{:.1%}", "Scostamento": "{:+.2%}",
                                          "Δ margine": "{:+,.0f} €", **{c: "{:+.2%}" for c in df.columns if c.startswith("Scost. ")}}, na_rep=""),
                         hide_index=True, width="stretch")
            band = pd.DataFrame({"y1": [mf.BANDA_MIN], "y2": [mf.BANDA_MAX]})
            base = alt.Chart(df)
            fascia = alt.Chart(band).mark_rect(opacity=0.12, color="#027A48").encode(y="y1:Q", y2="y2:Q")
            linea = base.mark_line(point=True, color="#14213D").encode(
                x=alt.X("Mese:N", title=None), y=alt.Y("Scostamento:Q", axis=alt.Axis(format="%"), title="Scostamento STD − FIFO"),
                tooltip=["Mese", alt.Tooltip("Scostamento:Q", format="+.2%")])
            st.altair_chart((fascia + linea).properties(height=260), width="stretch")
            st.caption("Fascia verde = scostamento buono (tra −3% e −0,5%).")
        # correzione di un caricamento sbagliato
        salvati_ora = mf.mesi_salvati()
        if salvati_ora:
            with st.expander("Correggi i file di un mese"):
                k_corr = st.selectbox("Mese da correggere", salvati_ora, format_func=mf.etichetta_mese, key="mf_corr_mese")
                u1, u2, u3 = st.columns(3)
                n_marg = u1.file_uploader("Margini STD vs FIFO", type=["xlsx"], key=f"mf_corr_marg_{k_corr}")
                n_cost = u2.file_uploader("Cost Analysis", type=["xlsx"], key=f"mf_corr_cost_{k_corr}")
                n_non = u3.file_uploader("Standard non IFB", type=["xlsx"], key=f"mf_corr_non_{k_corr}")
                if st.button("Sostituisci i file caricati", disabled=not (n_marg or n_cost or n_non)):
                    mf.sostituisci_file_mese(k_corr, n_marg.getvalue() if n_marg else None,
                                             n_cost.getvalue() if n_cost else None, n_non.getvalue() if n_non else None)
                    st.session_state[k_mese] = k_corr
                    st.rerun()
                st.caption("Si sostituiscono solo i file caricati qui, gli altri restano; lo storico si aggiorna da solo.")
