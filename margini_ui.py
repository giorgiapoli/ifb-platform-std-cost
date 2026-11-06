"""Pagina 'Margini vs FIFO' dell'app (solo Canarie / COMIT). La logica sta in margini_fifo.py."""
import hashlib
import os
import re

import altair as alt
import numpy as np
import pandas as pd
import streamlit as st

import margini_fifo as mf

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


LUNGHE = ("Cause probabili", "Segnale", "Anomalia", "Motivo", "Esito", "Motivazione")
VERDE_C = "background-color:#D1FADF;color:#05603A"
ROSSO_C = "background-color:#FEE4E2;color:#B42318;font-weight:600"
AMBRA_C = "background-color:#FEF0C7;color:#93370D"


def _tabella_html(sty, a_sinistra=(), altezza=520, a_capo=()):
    """Tabella HTML: le intestazioni vanno a capo quando la finestra si restringe (st.dataframe le taglia).
    a_sinistra = nomi delle colonne di testo da allineare a sinistra; a_capo = colonne di testo lungo che vanno a capo;
    scorre in orizzontale e in verticale."""
    cols = list(sty.data.columns)
    cid = "t" + hashlib.md5("|".join(map(str, cols)).encode("utf-8")).hexdigest()[:8]      # le regole valgono solo per questa tabella
    reg = "".join(f".tab-h.{cid} td:nth-child({i + 1}),.tab-h.{cid} th:nth-child({i + 1}){{text-align:left}}" for i, c in enumerate(cols)
                  if c in a_sinistra)
    reg += "".join(f".tab-h.{cid} td:nth-child({i + 1}){{white-space:normal;min-width:{300 if c in LUNGHE else 170}px;"
                   f"max-width:{520 if c in LUNGHE else 260}px;line-height:1.3;overflow-wrap:anywhere}}" for i, c in enumerate(cols) if c in a_capo)
    st.markdown(
        "<style>.tab-h{overflow:auto;resize:both;height:" + str(min(altezza, 130 + 29 * len(sty.data))) + "px;min-height:90px;"
        "border:1px solid #E4E7EC;border-radius:10px}"
        ".tab-h table{border-collapse:collapse;font-size:.82rem;width:100%}"
        ".tab-h th{position:sticky;top:0;white-space:normal;text-align:center;vertical-align:bottom;background:#F9FAFB;color:#667085;"
        "font-weight:500;padding:6px 8px;border-bottom:1px solid #E4E7EC;min-width:76px;line-height:1.2}"
        ".tab-h td{white-space:nowrap;padding:5px 8px;border-bottom:1px solid #F2F4F7;text-align:right}" + reg
        + f"</style><div class='tab-h {cid}'>" + sty.hide(axis="index").to_html() + "</div>", unsafe_allow_html=True)


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
    if not m and r["esito"].startswith("Non verificabile"):
        m.append("volume non verificabile (meno di 3 righe con le misure del box): nessuna anomalia rilevata")
    return " · ".join(m)


def _date_ingresso(date):
    """Data (o intervallo) di registrazione dell'entrata: '15/10/2026' oppure '15/10/2026 – 16/10/2026'."""
    if date is None or len(date) == 0:
        return ""
    a, b = date.min(), date.max()
    return f"{a:%d/%m/%Y}" if a == b else f"{a:%d/%m/%Y} – {b:%d/%m/%Y}"


def _disegno_container(d, totale_trasporto):
    """Due container affiancati, divisi in tanti blocchi quanti sono gli articoli: in verde com'e' giusto dividere il
    trasporto (per volume), in rosso com'e' stato diviso davvero. Stesso colore/posizione = stesso articolo."""
    g = d.groupby("item").agg(desc=("desc", "first"), vol=("v_ok", "sum"), reg=("reg", "sum"), att=("atteso", "sum")).reset_index()
    g = g.sort_values("vol", ascending=False).reset_index(drop=True)
    g["q_ok"] = g["vol"] / g["vol"].sum()
    g["q_reg"] = g["reg"] / g["reg"].sum()
    verdi = ["#027A48", "#12B76A", "#05603A", "#32D583"]
    rossi = ["#B42318", "#F04438", "#7A271A", "#F97066"]

    def barra(col, palette, valore):
        blocchi = []
        for i, r in g.iterrows():
            q = r[col]
            testo = (f'<b>{r["item"]}</b><br>{q:.0%}' if q >= 0.055 else (f'{q:.0%}' if q >= 0.03 else ""))
            tip = f'{r["item"]} · {str(r["desc"]).strip()} · {q:.1%} · {valore(r):,.0f} €'.replace(",", ".")
            blocchi.append(f'<div title="{tip}" style="flex:{max(q, 0.004):.4f} 1 0;background:{palette[i % len(palette)]};'
                           'color:#fff;font-size:.74rem;line-height:1.15;text-align:center;display:flex;align-items:center;'
                           'justify-content:center;border-right:2px solid #fff;min-width:2px;overflow:hidden">'
                           f'<span>{testo}</span></div>')
        return ('<div style="display:flex;align-items:stretch">'
                '<div style="flex:1;display:flex;height:66px;border:3px solid #475467;border-right:0;border-radius:8px 0 0 8px;'
                'overflow:hidden;background:#fff">' + "".join(blocchi) + '</div>'
                '<div style="width:14px;border:3px solid #475467;border-radius:0 8px 8px 0;background:'
                'repeating-linear-gradient(90deg,#98A2B3 0 3px,#D0D5DD 3px 6px)"></div></div>')

    righe_leg = "".join(
        f'<span style="display:inline-block;margin:2px 10px 2px 0;font-size:.82rem"><span style="display:inline-block;width:10px;height:10px;'
        f'background:{verdi[i % len(verdi)]};margin-right:4px"></span><span style="display:inline-block;width:10px;height:10px;'
        f'background:{rossi[i % len(rossi)]};margin-right:5px"></span><b>{r["item"]}</b> {str(r["desc"]).strip()[:34]}: '
        f'{r["q_ok"]:.0%} → <span style="color:#B42318;font-weight:600">{r["q_reg"]:.0%}</span></span>'
        for i, r in g.iterrows())
    st.markdown(
        '<div style="background:#fff;border:1px solid #E4E7EC;border-radius:12px;padding:12px 16px;margin:8px 0">'
        '<div style="font-weight:600;margin-bottom:4px">Il trasporto ' + f'({totale_trasporto:,.0f} €)'.replace(",", ".") + ' dentro la spedizione</div>'
        '<div style="color:#027A48;font-weight:600;font-size:.88rem;margin:6px 0 3px 0">Come dovrebbe essere diviso (in base al volume dei box)</div>'
        + barra("q_ok", verdi, lambda r: r["att"]) +
        '<div style="color:#B42318;font-weight:600;font-size:.88rem;margin:10px 0 3px 0">Come è stato diviso davvero</div>'
        + barra("q_reg", rossi, lambda r: r["reg"]) +
        '<div style="margin-top:6px">' + righe_leg + '</div></div>', unsafe_allow_html=True)


def render(p_modello):
    st.markdown("#### Margini COMIT: standard cost contro FIFO")
    st.caption("Carica ogni mese i report: i file vengono letti in sola lettura e non escono dal PC. "
               "Differenza margine = margine STD − margine FIFO: negativa = standard prudente, "
               "positiva = i costi reali superano lo standard. Fascia buona: tra −3% e −0,5% delle vendite.")
    # ---- mesi salvati: un pulsante per mese ---------------------------------------------------
    salvati = mf.mesi_salvati()
    sel = st.session_state.get("mf_mese")
    if sel not in salvati:
        sel = None
    if salvati:
        st.markdown("**Mesi in memoria**")
        cols = st.columns([1] * len(salvati) + [1, max(8 - len(salvati), 1)])
        for col, k in zip(cols, salvati):
            if col.button(mf.etichetta_mese(k), key=f"mf_btn_{k}", type="primary" if k == sel else "secondary",
                          width="stretch"):
                st.session_state["mf_mese"] = k
                st.rerun()
        if cols[len(salvati)].button("＋ Nuovo mese", key="mf_nuovo", type="primary" if sel is None else "secondary",
                                     width="stretch"):
            st.session_state["mf_mese"] = None
            st.rerun()
    b_marg = b_cost = b_non = None
    giorni_rec = 30          # entrate recenti a costi accessori zero = fatture probabilmente non ancora registrate
    if sel:
        b_marg, b_cost, b_non = mf.leggi_file_mese(sel)
    else:
        with st.expander("File del mese", expanded=not st.session_state.get("mf_ok", False)):
            c1, c2, c3 = st.columns(3)
            f_marg = c1.file_uploader("1 · Margini STD vs FIFO (obbligatorio)", type=["xlsx"], key="mf_marg",
                                      help="Export 'Difference margin' per articolo, con riga Total.")
            f_cost = c2.file_uploader("2 · Cost Analysis (entrate di magazzino)", type=["xlsx"], key="mf_cost",
                                      help="Una riga per entrata: serve per capire PERCHE' lo standard e il FIFO si allontanano.")
            f_non = c3.file_uploader("3 · Standard cost prodotti NON IFB", type=["xlsx"], key="mf_non",
                                     help="Report che calcola in automatico lo standard dei prodotti acquistati direttamente da COMIT.")
            b_marg = f_marg.getvalue() if f_marg else None
            b_cost = f_cost.getvalue() if f_cost else None
            b_non = f_non.getvalue() if f_non else None
    if not b_marg:
        st.info("Carica il report dei margini per iniziare. Con il Cost Analysis (e il report non IFB) l'app spiega anche "
                "perché gli articoli si allontanano dallo standard.")
        st.session_state["mf_ok"] = False
        return
    st.session_state["mf_ok"] = True
    mt = _motore()
    try:
        art, tot = _margini(b_marg, mt)
    except Exception as ex:
        st.error(f"Non riesco a leggere il report dei margini: {ex}")
        return
    ent = _entrate(b_cost, mt) if b_cost else None
    non = _non_ifb(b_non, mt) if b_non else {}
    if sel is None:        # file appena caricati: li si puo' tenere in memoria con il nome del mese
        mese_nuovo_def = f"{ent['data'].max():%Y-%m}" if ent is not None and len(ent) else ""
        s1, s2, s3 = st.columns([1, 1, 3])
        chiave_nuova = s1.text_input("Mese da salvare", value=mese_nuovo_def, placeholder="es. 2026-10",
                                     help="Anno-mese: il pulsante prendera' il nome del mese (es. 2026-10 = Ottobre 2026).")
        s2.write("")
        if s2.button("Tieni in memoria", disabled=not chiave_nuova.strip(), width="stretch"):
            k = mf.salva_file_mese(chiave_nuova, b_marg, b_cost, b_non)
            st.session_state["mf_mese"] = k
            st.rerun()
        s3.caption("Salva i file caricati: dalla prossima volta basta cliccare il pulsante del mese.")
    ifb = {}
    if p_modello and os.path.exists(p_modello):
        with st.spinner("Leggo gli standard cost IFB dal modello…"):
            ifb = _std_ifb(p_modello, os.path.getmtime(p_modello), mt)
    a = mf.perimetri(art, ifb, non)
    perim = mf.tabella_perimetri(a)
    # lo storico si aggiorna da solo a ogni caricamento (mese in memoria o file appena caricati con un mese valido)
    chiave_auto = sel if sel is not None else chiave_nuova.strip()
    if re.fullmatch(r"\d{4}-\d{2}", chiave_auto):
        mf.salva_mese(chiave_auto, tot, perim)

    # dati di imballo di tutti gli articoli (modello per gli IFB, anagrafica COMIT per gli altri e per i mancanti)
    p_ana = mf.cerca_anagrafica()
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
        trasp = mf.controllo_trasporto(med, ifb, non)

    if ent is not None and len(ent) and len(provv):
        with st.expander(f"{len(provv)} entrate recenti con costi accessori a zero: fatture probabilmente non ancora registrate "
                         f"(escluse dai controlli)"):
            st.dataframe(provv.sort_values("data", ascending=False)[["data", "item", "desc", "ubic", "ext", "qty", "prezzo"]].rename(
                columns={"data": "Data", "item": "Articolo", "desc": "Descrizione", "ubic": "Ubicazione",
                         "ext": "Spedizione (fattura)", "qty": "Q.tà", "prezzo": "Prezzo"}).style.format(
                {"Data": "{:%d/%m/%Y}", "Q.tà": "{:,.0f}", "Prezzo": "{:.3f} €"}), hide_index=True, width="stretch")
    t_esito, t_cont, t_art, t_aiem, t_tr, t_sto = st.tabs(
        ["Esito e outlier", "Spedizioni", "Un articolo", "AIEM", "Trasporto", "Storico mesi"])

    # ================== esito e outlier =====================================================
    with t_esito:
        scelta = st.segmented_control("Mostra", ["Tutti", "IFB", "Non IFB"], default="Tutti", key="mf_perim",
                                      label_visibility="collapsed",
                                      help="IFB = standard dal modello Canarie. Non IFB = prodotti che COMIT compra direttamente.")
        scelta = scelta or "Tutti"
        a_v = a if scelta == "Tutti" else a[a["perimetro"] == scelta]
        tot_v = tot if scelta == "Tutti" else mf.totali(a_v)
        cod, tit, txt = mf.esito(tot_v["diff_pct"])
        _banner(cod, f"{tit}: scostamento {_pct(tot_v['diff_pct'], True, 2)} delle vendite", txt)
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
        o1, o2, o3, o5, o4 = st.columns([1, 1, 1.3, 1.6, 2])
        min_pct = o1.number_input("Scostamento minimo (% vendite)", 0.0, 100.0, 10.0, 1.0) / 100
        min_eur = o2.number_input("Scostamento minimo (€)", 0.0, 100000.0, 50.0, 10.0)
        verso = o3.selectbox("Verso", ["Tutti", "FIFO peggiore dello STD", "STD troppo alto"])
        cerca = o4.text_input("Cerca", placeholder="codice o parte della descrizione")
        entrate = o5.selectbox("Entrate nel periodo", ["Tutti", "Nascondi quelli senza entrate", "Solo quelli senza entrate"],
                               help="«Senza entrate» = nessuna entrata nel periodo analizzato: il FIFO arriva da giacenza più vecchia (danno del passato che si trascina).")
        out = mf.outlier(a_v, min_pct, min_eur)
        if verso == "FIFO peggiore dello STD":
            out = out[out["diff"] > 0]
        elif verso == "STD troppo alto":
            out = out[out["diff"] < 0]
        if cerca:
            q = cerca.strip().lower()
            out = out[out["item"].str.lower().str.contains(q, regex=False) | out["desc"].astype(str).str.lower().str.contains(q, regex=False)]
        dg = None
        if med is not None:
            dg = mf.diagnosi(out, med, aiem, trasp, det, ifb, non)
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
            cols = ["item", "desc", "perimetro", "qty", "uom", "prezzo", "c_std", "c_fifo", "vendite", "diff", "diff_pct"]
            nomi = {"item": "Articolo", "desc": "Descrizione", "perimetro": "Origine", "qty": "Q.tà venduta", "uom": "UM",
                    "prezzo": "Prezzo vendita per UM", "c_std": "STD per UM", "c_fifo": "FIFO per UM", "vendite": "Vendite",
                    "diff": "Δ margine", "diff_pct": "Δ % vendite"}
            if dg is not None:
                cols += ["cause", "d_prezzo", "d_logistica", "d_aiem", "d_altro"]
                nomi.update({"d_prezzo": "Δ prezzo acquisto", "d_logistica": "Δ logistica", "d_aiem": "Δ AIEM",
                             "d_altro": "Non spiegato", "cause": "Cause probabili"})
            tb = out[cols].rename(columns=nomi).reset_index(drop=True)
            fm = {"Q.tà venduta": "{:,.0f}", "Prezzo vendita per UM": "{:.3f} €", "Vendite": "{:,.0f} €", "STD per UM": "{:.3f} €",
                  "FIFO per UM": "{:.3f} €", "Δ margine": "{:+,.0f} €", "Δ % vendite": "{:+.1%}",
                  "Δ prezzo acquisto": "{:+.3f} €", "Δ logistica": "{:+.3f} €", "Δ AIEM": "{:+.3f} €", "Non spiegato": "{:+.3f} €"}
            fmt = lambda f: (lambda v: "" if v is None or pd.isna(v) else f.format(v))
            sty = tb.style.format({k: fmt(v) for k, v in fm.items() if k in tb.columns})
            sty = sty.map(_evidenzia, subset=[c for c in ("Δ margine", "Δ % vendite") if c in tb.columns])
            sty = sty.map(lambda v: VERDE_C, subset=["STD per UM"]).map(lambda v: ROSSO_C, subset=["FIFO per UM"])   # verde = standard, rosso = costo reale
            _tabella_html(sty, a_sinistra=("Articolo", "Descrizione", "Origine", "UM", "Cause probabili"), altezza=640,
                          a_capo=("Descrizione", "Cause probabili"))
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
    with t_aiem:
        if aiem is None:
            st.info(needs)
        else:
            st.markdown("#### AIEM registrato contro anagrafica")
            st.caption(f"Entrate dal {inizio:%d/%m/%Y} al {fine:%d/%m/%Y}. L'AIEM atteso viene dall'anagrafica articolo "
                       "(modello IFB / report non IFB): una voce IMPUESTOS su un articolo senza AIEM, o l'assenza su un articolo "
                       "che ce l'ha, è una registrazione ambigua da segnalare a COMIT.")
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
    with t_cont:
        if cont is None:
            st.info(needs)
        else:
            st.markdown("#### Ripartizione dei costi sulle spedizioni")
            st.caption("Il container si isola con il «External Document No_» (il numero fattura IFB: tutte le righe che lo condividono "
                       "viaggiano insieme). Si confronta il trasporto registrato su ogni riga con la quota di volume attesa "
                       "(box × volume di un box, con box = pezzi ÷ pezzi per box). Verificabile per gli articoli con le misure "
                       "del box nel modello (IFB) o nell'anagrafica COMIT (tutti gli altri).")
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
                c1, c2, c3, c4 = st.columns(4)
                _kpi(c1, "Spedizioni nel periodo", len(cont), f"{int(cont['stessa_chiave'].sum())} con imposte ripartite come il trasporto")
                _kpi(c2, "A volume corretto", n_ok, "ripartizione attesa", "#027A48")
                _kpi(c3, "Errore «collo × pezzi»", n_bug, "manca la divisione per i pezzi per collo", "#B42318" if n_bug else INK)
                _kpi(c4, "Non coerenti", n_nc, "né volume corretto né errore noto", "#B54708" if n_nc else INK)
                st.write("")
                f = st.selectbox("Mostra", ["Tutte le spedizioni", "Spedizioni con anomalie"])
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
                tb = vis.rename(columns={"ext": "Spedizione (fattura)", "data": "Ultima entrata", "ubic": "Ubicazione",
                                         "righe": "Righe", "trasporto_eur": "Trasporto totale", "esito": "Esito",
                                         "despachos_eur": "Despachos totali", "imposte_eur": "Imposte totali",
                                         "merce_eur": "Valore merce", "trasp_pct": "Trasporto su merce",
                                         "stessa_chiave": "Imposte ripartite come il trasporto",
                                         "imposte_su_senza_aiem": "Righe con imposte su articoli senza AIEM",
                                         "impatto_eur": "Impatto stimato", "imposte_indebite_eur": "Imposte su articoli senza AIEM",
                                         "sbilancio_eur": "Costo spostato tra righe", "motivo": "Motivo", "q_art": "Q.tà articolo", "um_art": "UM",
                                         "t_art": "Trasporto sull'articolo", "t_unit": "Trasporto per UM"})
                with st.expander("Elenco delle spedizioni", expanded=True):
                    tb = tb.assign(**{"Imposte ripartite come il trasporto": tb["Imposte ripartite come il trasporto"].map(
                        lambda v: "Sì" if v else "")})
                    sty_el = tb[colonne].style.format(
                        {"Ultima entrata": "{:%d/%m/%Y}", "Valore merce": "{:,.0f} €", "Trasporto totale": "{:,.0f} €",
                         "Trasporto su merce": "{:.1%}", "Despachos totali": "{:,.0f} €", "Imposte totali": "{:,.0f} €",
                         "Costo spostato tra righe": "{:,.0f} €", "Impatto stimato": "{:,.0f} €",
                         "Imposte su articoli senza AIEM": "{:,.0f} €", "Q.tà articolo": "{:,.0f}",
                         "Trasporto sull'articolo": "{:,.0f} €", "Trasporto per UM": "{:.3f} €"}, na_rep="")
                    _tabella_html(sty_el, a_sinistra=("Spedizione (fattura)", "Ubicazione", "Esito", "Motivo", "UM"),
                                  a_capo=("Spedizione (fattura)", "Esito", "Motivo"))
                if len(vis):
                    scelta = st.selectbox("Dettaglio di una spedizione", list(vis["ext"]))
                    tt = cont[cont["ext"] == scelta].iloc[0]
                    st.markdown(f"**Spedizione {scelta}**: trasporto totale **{_eur(tt['trasporto_eur'])}** · despachos "
                                f"{_eur(tt['despachos_eur'])} · imposte {_eur(tt['imposte_eur'])} · valore merce "
                                f"{_eur(tt['merce_eur'])} · {int(tt['righe'])} righe · ingresso in magazzino "
                                + _date_ingresso(w[w["ext"] == scelta]["data"]))
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
                        with st.expander("Disegno del container: come dovrebbe essere diviso il trasporto e come è stato diviso", expanded=True):
                            _disegno_container(d_sped, tt["trasporto_eur"])
                    if len(d_sped):
                        d = d_sped.copy()
                        d["pz_scatola"] = d["qty"] / d["scatole"]
                        d["vol_scatola"] = d["v_ok"] / d["scatole"]          # m³
                        d["quota_vol"] = d["v_ok"] / d["v_ok"].sum()
                        d["quota_reg"] = d["reg"] / d["reg"].sum()
                        d["tr_reg_u"] = d["reg"] / d["qty"]
                        d["tr_att_u"] = d["atteso"] / d["qty"]
                        d["aliq"] = d["aliquota"].where(d["aliquota"] >= 0)
                        d = d.rename(columns={"item": "Articolo", "desc": "Descrizione", "uom": "UM", "qty": "Q.tà",
                                              "pz_scatola": "UM x box", "scatole": "Box", "vol_scatola": "Volume box (m³)",
                                              "v_ok": "Volume totale (m³)", "quota_vol": "Quota di volume", "data": "Data ingresso",
                                              "quota_reg": "Quota registrata", "reg": "Trasporto reale", "atteso": "Trasporto atteso",
                                              "sbilancio": "Δ trasporto", "aliq": "AIEM %", "imp_reg": "AIEM reale",
                                              "aiem_att": "AIEM atteso", "aiem_diff": "Δ AIEM",
                                              "tot_reale": "Totale reale", "tot_atteso": "Totale atteso",
                                              "tr_reg_u": "Trasporto reale per UM", "tr_att_u": "Trasporto atteso per UM"})
                        d = d.sort_values("Δ trasporto", key=abs, ascending=False)
                        cols = ["Articolo", "Descrizione", "Data ingresso", "UM", "Q.tà", "UM x box", "Box", "Volume box (m³)", "Volume totale (m³)",
                                "Quota di volume", "Quota registrata", "Trasporto reale", "Trasporto atteso", "Δ trasporto",
                                "AIEM %", "AIEM reale", "AIEM atteso", "Δ AIEM", "Totale reale", "Totale atteso", "Trasporto reale per UM", "Trasporto atteso per UM"]
                        tab = d[cols].copy()
                        tot_riga = {c: None for c in cols}
                        tot_riga.update({"Articolo": "TOTALE", "Box": d["Box"].sum(), "Volume totale (m³)": d["Volume totale (m³)"].sum(),
                                         "Quota di volume": d["Quota di volume"].sum(), "Quota registrata": d["Quota registrata"].sum(),
                                         "Trasporto reale": d["Trasporto reale"].sum(), "Trasporto atteso": d["Trasporto atteso"].sum(),
                                         "Δ trasporto": d["Δ trasporto"].clip(lower=0).sum(), "AIEM reale": d["AIEM reale"].sum(),
                                         "AIEM atteso": d["AIEM atteso"].sum(), "Δ AIEM": d["Δ AIEM"].clip(lower=0).sum(),
                                         "Totale reale": d["Totale reale"].sum(), "Totale atteso": d["Totale atteso"].sum()})
                        tab = pd.concat([tab, pd.DataFrame([tot_riga])], ignore_index=True)
                        for c in tab.columns:                        # niente "None" nella riga dei totali
                            if c == "Data ingresso":
                                tab[c] = tab[c].map(lambda v: "" if pd.isna(v) else f"{v:%d/%m/%Y}")
                            elif c in ("Articolo", "Descrizione", "UM"):
                                tab[c] = tab[c].fillna("")
                            else:
                                tab[c] = pd.to_numeric(tab[c], errors="coerce")
                        sty = tab.style.format(
                            {"Q.tà": "{:,.0f}", "UM x box": "{:,.0f}", "Box": "{:,.0f}", "Volume box (m³)": "{:.4f}",
                             "Volume totale (m³)": "{:.2f}", "Quota di volume": "{:.1%}", "Quota registrata": "{:.1%}",
                             "Trasporto reale": "{:,.0f} €", "Trasporto atteso": "{:,.0f} €", "Δ trasporto": "{:+,.0f} €",
                             "AIEM %": "{:.0f}%", "AIEM reale": "{:,.0f} €", "AIEM atteso": "{:,.0f} €", "Δ AIEM": "{:+,.0f} €",
                             "Totale reale": "{:,.0f} €", "Totale atteso": "{:,.0f} €",
                             "Trasporto reale per UM": "{:.3f} €", "Trasporto atteso per UM": "{:.3f} €"}, na_rep="")
                        sty = sty.map(_evidenzia, subset=pd.IndexSlice[tab.index[:-1], ["Δ trasporto", "Δ AIEM"]])
                        # verde = come dovrebbe essere; rosso = com'e' stato registrato, se si scosta dall'atteso
                        verdi_c = ["Quota di volume", "Trasporto atteso", "AIEM atteso", "Totale atteso", "Trasporto atteso per UM"]
                        rossi_t = ["Quota registrata", "Trasporto reale", "Trasporto reale per UM"]
                        rossi_a = ["AIEM reale"]
                        soglia_t = 0.5        # rosso per qualsiasi differenza dall'atteso (sotto 0,5 € sono arrotondamenti)
                        soglia_a = 0.5
                        ultimo = tab.index[-1]

                        def colora(r):
                            if r.name == ultimo:
                                return [""] * len(r)
                            err_t = abs(r["Δ trasporto"]) >= soglia_t
                            err_tot = abs(r["Totale reale"] - r["Totale atteso"]) >= soglia_t
                            err_a = pd.notna(r["Δ AIEM"]) and abs(r["Δ AIEM"]) >= soglia_a
                            rosso = "background-color:#FEE4E2;color:#B42318;font-weight:600"
                            ok = "background-color:#F6FEF9;color:#05603A"
                            out = []
                            for c in r.index:
                                if c in verdi_c:
                                    out.append("background-color:#D1FADF;color:#05603A")
                                elif c in rossi_t:
                                    out.append(rosso if err_t else ok)
                                elif c == "Totale reale":
                                    out.append(rosso if err_tot else ok)
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
                        with st.expander("Tabella di dettaglio", expanded=True):
                            # tabella HTML: le intestazioni vanno a capo quando la finestra si restringe (st.dataframe le taglia)
                            st.markdown(
                                "<style>.tab-det{overflow:auto;resize:both;min-height:90px;border:1px solid #E4E7EC;border-radius:10px}.tab-det table{border-collapse:collapse;font-size:.82rem;width:100%}"
                                ".tab-det th{white-space:normal;text-align:center;vertical-align:bottom;background:#F9FAFB;color:#667085;"
                                "font-weight:500;padding:6px 8px;border-bottom:1px solid #E4E7EC;min-width:58px;line-height:1.2}"
                                ".tab-det td{white-space:nowrap;padding:5px 8px;border-bottom:1px solid #F2F4F7;text-align:right}"
                                ".tab-det td:nth-child(1),.tab-det td:nth-child(3),.tab-det td:nth-child(4){text-align:left}"
                                ".tab-det td:nth-child(2){white-space:normal;min-width:170px;text-align:left}"
                                ".tab-det th:nth-child(2){text-align:left}</style><div class='tab-det'>"
                                + sty.hide(axis="index").to_html() + "</div>", unsafe_allow_html=True)
                            st.caption(
                                "**Box** = pezzi ÷ pezzi per box (per i prodotti a peso: kg ÷ kg per box; dati dall'anagrafica). "
                                "**Volume totale** = box × volume di un box (altezza × lunghezza × profondità). "
                                "**Trasporto atteso** = trasporto totale della spedizione × quota di volume dell'articolo; **Trasporto reale** = "
                                "quanto è stato registrato sulla riga. **Totale** = prezzo di acquisto + trasporto + AIEM + despachos (reale come "
                                "registrato; atteso con trasporto e despachos divisi a volume). **AIEM atteso** = aliquota AIEM dell'articolo × (valore della merce + "
                                "trasporto atteso della riga), come nella fattura delle imposte; **AIEM reale** = imposte registrate sulla riga. "
                                "**Δ** = reale − atteso: positivo = riga sovraccaricata, negativo = sottocaricata.")
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
                        with st.expander("Testo pronto per la segnalazione a COMIT"):
                            st.code(mf.testo_segnalazione(tt, problemi), language=None)
                    else:
                        st.caption("Nessuna anomalia rilevata su questa spedizione.")

    # ================== trasporto =========================================================
    with t_tr:
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
                sty_t = tt_.style.format({"Vendite mese": "{:,.0f} €", "Trasp. min": "{:.3f} €", "Trasp. max": "{:.3f} €",
                                          "Reale": "{:.3f} €", "Logistica STD": "{:.3f} €", "Reale / STD": "{:.1f}x"}, na_rep="")
                # verde = come dovrebbe essere (standard); rosso = com'e' stato registrato; ambra = trasporto instabile tra le entrate
                sty_t = sty_t.map(lambda v: VERDE_C, subset=["Logistica STD"])
                sty_t = sty_t.map(lambda v: ROSSO_C, subset=["Reale", "Reale / STD"])
                sty_t = sty_t.apply(lambda r: [AMBRA_C if (c in ("Trasp. min", "Trasp. max") and "instabile" in str(r["Segnale"]))
                                               else "" for c in r.index], axis=1)
                _tabella_html(sty_t, a_sinistra=("Articolo", "Descrizione", "Segnale"), altezza=620, a_capo=("Descrizione", "Segnale"))

    # ================== un articolo =======================================================
    with t_art:
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
                    st.session_state["mf_mese"] = k_corr
                    st.rerun()
                st.caption("Si sostituiscono solo i file caricati qui, gli altri restano; lo storico si aggiorna da solo.")
