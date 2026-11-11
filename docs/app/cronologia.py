"""
Pagine "Logistica", "Anagrafica" e "Listini": cosa e' cambiato nel tempo, di quanto e quando, confrontando
tutti i modelli mensili disponibili (storico + mese in corso). Sola lettura.

Inoltre: grafico dell'andamento con la freccia sui punti in cui lo standard e' stato caricato dalla macro
(prodotto acquistato nel mese).
"""
import hashlib
import json
import os
import pickle
import re
from datetime import datetime

import altair as alt
import pandas as pd
import streamlit as st
from openpyxl.utils.cell import coordinate_to_tuple

import caricamenti
import costi_log
import spiega_variazioni as sv

HERE = os.path.dirname(os.path.abspath(__file__))
CACHE_DIR = os.path.join(HERE, "cache_std")
INK, MUTED, UP, DOWN, ACCENT, AMBRA = "#14213D", "#667085", "#B42318", "#027A48", "#0F766E", "#B54708"
PAL = ["#0F766E", "#2563EB", "#7C3AED", "#D97706"]
TIPI_ANA = {"Unità di misura": "#B54708", "Pallettizzazione/imballo": "#2563EB", "AIEM": "#DB2777",
            "Duties alcolici": "#7C3AED", "Resa": "#0F766E", "Modalità di trasporto": "#D97706",
            "Destinazioni": "#0891B2", "Fornitore": "#667085"}


# ---- versioni e mesi ---------------------------------------------------------------------------
def mesi_in(testo):
    """Mesi (1-12) citati in un testo, nell'ordine: 'SET-OTT 2026' -> [9, 10]; '6- SETTEMBRE' -> [9]."""
    out = []
    for t in re.findall(r"[a-z]+", testo.lower()):
        m = caricamenti.MESI.get(t[:3])
        if m and m not in out:
            out.append(m)
    return out


def etichetta(nome):
    """'Storico · 7-OTTOBRE' -> '7-OTTOBRE'; 'Mese in corso · INPUT' -> 'OGGI'."""
    if nome.startswith("Mese in corso"):
        return "OGGI"
    return re.sub(r"^Storico · ", "", nome).strip()


def versioni_per_pagine(sorgente, modelli, p_old, p_new):
    """[(etichetta, path, mtime)] in ordine cronologico."""
    if sorgente == "Dalle cartelle" and modelli:
        return [(etichetta(n), p, os.path.getmtime(p)) for n, p in modelli.items()]
    return [("PRIMA", p_old, os.path.getmtime(p_old)), ("OGGI", p_new, os.path.getmtime(p_new))]


def _motore():
    return os.path.getmtime(os.path.join(HERE, "spiega_variazioni.py"))


# ---- fotografia di un modello -------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def snapshot(path, mtime, filiale, motore, _carica):
    """Dati di ogni articolo (prezzo, unita' di misura, anagrafica, STD) e tariffe usate di COSTS (LOG).
    Salvato su disco: un modello gia' letto non va riletto da Excel (la chiave include la versione del motore)."""
    key = hashlib.md5(f"snap|{filiale}|{path}|{mtime}|{motore}".encode("utf-8")).hexdigest()
    f = os.path.join(CACHE_DIR, key + ".pkl")
    if os.path.exists(f):
        with open(f, "rb") as fh:
            return pickle.load(fh)
    sv.set_branch(filiale)
    shs = _carica(path, mtime, filiale)
    items = {}
    for sn, sh in shs.items():
        for code, row in sh.items.items():
            d = sv.drivers(sh, row)
            items[(sn, code)] = dict(
                desc=str(sh.val(row, "description") or ""), ifb=str(sh.val(row, "ifbitem") or ""),
                price=sh.n(row, sv._price_name(sh)),
                s2={dest: (sv.std_finale(sh, row, dest) or None) for dest in sv.ISLANDS},
                um=d["um"], dati=d["dati"], aiem=d["aiem_rate"], grado=d["grado"], loc=d["loc"],
                modo=d["modo"], area=d["area"], dest=d["dest"], forn=d["fornitore"])
    log = {}
    prima = next(iter(shs.values()), None)
    if prima is not None:
        usate, lg = getattr(prima, "log_used", set()), getattr(prima, "log", {})
        log = {rc: lg[rc] for rc in usate if lg.get(rc) is not None}
    out = dict(items=items, log=log)
    os.makedirs(CACHE_DIR, exist_ok=True)
    with open(f, "wb") as fh:
        pickle.dump(out, fh)
    return out


_MESI_IT = ["Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre",
            "Dicembre"]
_OGGI = {"nome": "Oggi"}


def nome_mese(label):
    """Solo il nome del mese: '7-OTTOBRE' -> 'Ottobre', '4 - LUGLIO' -> 'Luglio'; la versione del mese in corso (OGGI)
    prende il nome del mese in cui e' stato salvato il modello."""
    t = str(label).strip()
    if t.upper() == "OGGI":
        return _OGGI["nome"]
    n = re.sub(r"^\s*\d+\s*[-–.]?\s*", "", t).strip()
    return n.capitalize() if n else t


def carica_snaps(versioni, filiale, carica):
    prog = st.progress(0.0, text="Leggo le versioni…")
    snaps = []
    for k, (lab, p, mt) in enumerate(versioni):
        prog.progress(k / len(versioni), text=f"Leggo {lab} ({k + 1} di {len(versioni)}): i mesi chiusi si leggono una volta sola")
        snaps.append({**snapshot(p, mt, filiale, _motore(), carica), "data": datetime.fromtimestamp(mt)})
    prog.empty()
    for v, sn in zip(versioni, snaps):
        if str(v[0]).strip().upper() == "OGGI":
            _OGGI["nome"] = _MESI_IT[sn["data"].month - 1]
    return [v[0] for v in versioni], snaps


# ---- eventi: cosa e' cambiato tra un mese e il successivo ---------------------------------------
def _v(x):
    return "–" if x is None or (isinstance(x, float) and pd.isna(x)) else costi_log._txt(x)


def _incompleto(x):
    """Valore mancante o parziale ('–', '#N/A', 'GOMMA/–'): l'articolo non era calcolato (bloccato, delistato...)
    quindi non e' un cambio di valore."""
    t = str(x)
    if "#N/A" in t or "#REF" in t or "#DIV" in t:       # errori di Excel (contengono una "/": controllo prima di dividere)
        return True
    return any(p.strip() in ("", "–", "-", "None", "nan") for p in t.split("/"))


def _data(snap):
    return snap["data"].strftime("%d/%m/%Y")


def _progressivo(df, chiavi):
    """Nessuna colonna di conteggio: le righe della stessa voce restano comunque vicine, in ordine di tempo."""
    return df


def eventi_listini(labels, snaps):
    righe = []
    for i in range(1, len(snaps)):
        A, B = snaps[i - 1]["items"], snaps[i]["items"]
        for k, b in B.items():
            a = A.get(k)
            if not a or not a["price"] or not b["price"] or a["um"] != b["um"]:
                continue                       # prezzo mancante o cambio di unita' (lo vedi in Anagrafica)
            d = b["price"] - a["price"]
            if abs(d) > sv.TOL:
                righe.append({"Mese": labels[i], "Articolo": k[1], "IFB": b["ifb"],
                              "Descrizione": b["desc"][:60], "Prima": f"{costi_log._txt(a['price'])} fino a {nome_mese(labels[i - 1])}",
                              "Dopo": f"{costi_log._txt(b['price'])} da {nome_mese(labels[i])}",
                              "Δ": d, "Δ %": d / a["price"]})
    return _progressivo(pd.DataFrame(righe), ["Articolo"])


def eventi_anagrafica(labels, snaps):
    righe = []
    for i in range(1, len(snaps)):
        A, B = snaps[i - 1]["items"], snaps[i]["items"]
        for k, b in B.items():
            a = A.get(k)
            if not a:
                continue
            cambi = []
            if a["um"] != b["um"] and a["um"][0] and b["um"][0]:
                cambi.append(("Unità di misura", "unità di misura (vendita/base)", "/".join(a["um"]), "/".join(b["um"])))
            # pallettizzazione: solo valori presenti in entrambi i mesi e per articoli calcolati in entrambi
            # (molti campi sono formule del modello: compaiono/spariscono quando l'articolo entra o esce dal calcolo)
            calcolato = any(a["s2"].values()) and any(b["s2"].values())
            for n, x, y in zip(sv.DATA_DRIVERS, a["dati"], b["dati"]):
                if calcolato and x is not None and y is not None and abs(x - y) > 1e-9:
                    cambi.append(("Pallettizzazione/imballo", sv.DATA_LABELS[n], _v(x), _v(y)))
            if a["aiem"] != b["aiem"]:
                cambi.append(("AIEM", "aliquota", _v(a["aiem"]), _v(b["aiem"])))
            if a["grado"] != b["grado"]:
                cambi.append(("Duties alcolici", "grado alcolico", _v(a["grado"]), _v(b["grado"])))
            if a["loc"] != b["loc"]:
                cambi.append(("Resa", "resa (MTS/MTO/FOR)", a["loc"] or "–", b["loc"] or "–"))
            if (a["modo"], a["area"]) != (b["modo"], b["area"]):
                cambi.append(("Modalità di trasporto", "modalità/area", f"{a['modo'] or '–'}/{a['area'] or '–'}",
                              f"{b['modo'] or '–'}/{b['area'] or '–'}"))
            if a["dest"] != b["dest"] and any(a["dest"]) and any(b["dest"]):
                cambi.append(("Destinazioni", "destinazioni", sv._dest_txt(a["dest"]), sv._dest_txt(b["dest"])))
            if a["forn"] != b["forn"]:
                cambi.append(("Fornitore", "fornitore", a["forn"] or "–", b["forn"] or "–"))
            for tipo, campo, pa, pb in cambi:
                if _incompleto(pa) or _incompleto(pb):
                    continue                # un lato mancante: non e' un cambio di valore
                righe.append({"Mese": labels[i], "Articolo": k[1], "IFB": b["ifb"],
                              "Descrizione": b["desc"][:60], "Tipo": tipo, "Campo": campo,
                              "Prima": f"{pa} fino a {nome_mese(labels[i - 1])}", "Dopo": f"{pb} da {nome_mese(labels[i])}"})
    df = pd.DataFrame(righe)
    if not df.empty:
        df["_k"] = df["Articolo"] + "|" + df["Campo"]
        df = _progressivo(df, ["_k"]).drop(columns="_k")
    return df


def _mappa_log(filiale):
    """{(riga, colonna): (voce, colonna)} dai blocchi della tabella dei costi logistici."""
    out = {}
    for titolo, colonne, righe in costi_log.BLOCCHI.get(filiale, []):
        for et, celle in righe:
            for col, addr in zip(colonne, celle):
                if addr:
                    out[coordinate_to_tuple(addr)] = (f"{titolo} · {et}", col if len(colonne) > 1 else "")
    return out


def eventi_logistica(filiale, labels, snaps):
    mappa = _mappa_log(filiale)
    righe = []
    for i in range(1, len(snaps)):
        A, B = snaps[i - 1]["log"], snaps[i]["log"]
        for rc in set(A) | set(B):
            a, b = A.get(rc), B.get(rc)
            if a is None or b is None:
                continue                        # voce presente in un solo modello: non e' un cambio di valore
            num = isinstance(a, (int, float)) and isinstance(b, (int, float))
            if num and abs(a - b) <= max(1e-9, 0.001 * abs(a)):
                continue                       # differenza irrilevante (es. 3.521,44 -> 3.521)
            if not num and costi_log._txt(a) == costi_log._txt(b):
                continue
            voce, col = mappa.get(rc, (f"COSTS (LOG) riga {rc[0]} colonna {rc[1]}", ""))
            righe.append({"Mese": labels[i], "Voce": voce, "Colonna": col,
                          "Prima": f"{_v(a)} fino a {nome_mese(labels[i - 1])}", "Dopo": f"{_v(b)} da {nome_mese(labels[i])}", "Δ": (b - a) if num else None,
                          "Δ %": ((b - a) / a) if num and a else None, "_rc": rc})
    df = pd.DataFrame(righe)
    if not df.empty:
        df["_k"] = df["Voce"] + "|" + df["Colonna"]
        df = _progressivo(df, ["_k"]).drop(columns="_k")
    return df


# ---- grafico con la freccia sui caricamenti -----------------------------------------------------
def marche_caricamento(filiale, base, code, ifb, labels):
    """Punti in cui l'articolo e' stato caricato dalla macro (output SC_Analisi): una riga per caricamento."""
    out = []
    for lab_file, p, mt in caricamenti.trova_file(filiale, base):
        periodo = lab_file.split("  ·  ")[0]
        mesi = mesi_in(periodo)
        if len(mesi) < 2:
            continue
        et = next((l for l in labels if mesi[-1] in mesi_in(l)), None)      # versione del secondo mese del periodo
        if not et:
            continue
        _, _, df = caricamenti.leggi(p, mt)
        if df.empty or "Code" not in df.columns:
            continue
        mask = df["Code"].astype(str).str.strip() == str(code)
        if ifb and "IFB" in df.columns:
            mask = mask | (df["IFB"].astype(str).str.strip() == str(ifb))
        if not mask.any():
            continue
        r = df[mask].iloc[0]
        acq = pd.to_datetime(r.get("Last Date Invoiced"), errors="coerce")
        out.append({"Versione": et, "Caricamento": periodo, "File del": datetime.fromtimestamp(mt).strftime("%d/%m/%Y"),
                    "Acquistato il": "" if pd.isna(acq) else acq.strftime("%d/%m/%Y"),
                    "Quantità": "" if r.get("Quantity") is None else str(r.get("Quantity")),
                    "Azione": str(r.get("ACTION", ""))})
    return pd.DataFrame(out)


def grafico_andamento(serie, ordine, y, colore, marc=None, altezza=300):
    """Linee nel tempo per colore; ▼ sui punti in cui lo standard e' stato caricato (marc)."""
    base = alt.Chart(serie).encode(
        x=alt.X("Versione:N", sort=ordine, title=None, axis=alt.Axis(labelAngle=0, labelLimit=160)),
        y=alt.Y(f"{y}:Q", scale=alt.Scale(zero=False), title=y),
        color=alt.Color(f"{colore}:N", scale=alt.Scale(range=PAL), legend=alt.Legend(orient="top", title=None)),
        tooltip=["Versione", colore, alt.Tooltip(f"{y}:Q", format=",.2f")])
    layer = base.mark_line(strokeWidth=2.5) + base.mark_point(size=80, filled=True)
    if marc is not None and not marc.empty:
        ymax = serie.groupby("Versione")[y].max()
        m = marc.copy()
        m["y"] = m["Versione"].map(ymax)
        m = m.dropna(subset=["y"])
        if not m.empty:
            layer = layer + alt.Chart(m).mark_text(text="▼", dy=-18, size=22, color=AMBRA).encode(
                x=alt.X("Versione:N", sort=ordine), y="y:Q",
                tooltip=["Caricamento", "File del", "Acquistato il", "Quantità", "Azione"])
    return layer.properties(height=altezza, padding={"top": 28, "left": 6, "right": 12, "bottom": 6})


# ---- elementi comuni delle pagine ---------------------------------------------------------------
def _kpi(col, label, value, sub=""):
    col.markdown(f'<div class="kpi"><div class="l">{label}</div><div class="v">{value}</div>'
                 f'<div class="s">{sub}</div></div>', unsafe_allow_html=True)


def _tabella(df, altezza=420, delta_cols=(), euro=False):
    if "Mese" in df.columns:
        df = df.assign(Mese=df["Mese"].map(nome_mese))
    sty = df.reset_index(drop=True).style
    fm = {}
    for c in df.columns:
        if c in ("Prima", "Dopo") and pd.api.types.is_numeric_dtype(df[c]):
            fm[c] = lambda v: "" if pd.isna(v) else f"{v:,.2f}" + (" €" if euro else "")
        elif c == "Δ":
            fm[c] = lambda v: "" if v is None or pd.isna(v) else f"{v:+,.2f}" + (" €" if euro else "")
        elif c == "Δ %":
            fm[c] = lambda v: "" if v is None or pd.isna(v) else f"{v:+.1%}"
    sty = sty.format(fm, na_rep="")
    for c in delta_cols:
        if c in df.columns and df[c].notna().any():
            vmax = float(df[c].abs().quantile(0.95))
            sty = sty.map(lambda v, vm=vmax: caricamenti._cella_delta(v, vm), subset=[c])
    st.dataframe(sty, hide_index=True, width="stretch", height=min(altezza, 40 + 29 * max(len(df), 3)))


def _note(filiale, labels):
    """Eventi noti (annotazioni.json, modificabile): spiegano i salti nei grafici, es. un aumento concordato."""
    try:
        with open(os.path.join(HERE, "annotazioni.json"), encoding="utf-8") as fh:
            note = json.load(fh).get(filiale, {})
    except (OSError, ValueError):
        return
    righe = [f"**{m}**: {t}" for m, t in note.items() if m in labels]
    if righe:
        st.info("Eventi noti · " + " · ".join(righe))


def _intestazione(titolo, testo):
    st.markdown(f"#### {titolo}")
    st.caption(testo)


# ---- pagina LISTINI -----------------------------------------------------------------------------
def pagina_listini(filiale, versioni, carica):
    _intestazione("Listini", "Quando e di quanto e' cambiato il prezzo di acquisto di un articolo, mese per mese "
                  "(dal listino del modello). I cambi di unità di misura non sono aumenti: sono in Anagrafica.")
    if len(versioni) < 2:
        st.info("Servono almeno due versioni del modello.")
        return
    labels, snaps = carica_snaps(versioni, filiale, carica)
    _note(filiale, labels)
    ev = eventi_listini(labels, snaps)
    if ev.empty:
        st.info("Nessuna variazione di prezzo tra le versioni disponibili.")
        return
    ult = labels[-1]
    eu = ev[ev["Mese"] == ult]
    k = st.columns(4)
    _kpi(k[0], "Variazioni di prezzo", len(ev), f"in {len(labels) - 1} confronti mensili")
    _kpi(k[1], "Articoli coinvolti", ev["Articolo"].nunique(), "con almeno una variazione")
    _kpi(k[2], f"Aumenti · {ult}", int((eu["Δ"] > 0).sum()), f"ultimo confronto ({labels[-2]} → {ult})")
    _kpi(k[3], f"Cali · {ult}", int((eu["Δ"] < 0).sum()), f"ultimo confronto ({labels[-2]} → {ult})")
    st.write("")
    c1, c2, c3 = st.columns([3, 2, 3])
    mesi = c1.multiselect("Mese", labels[1:], key=f"lis_m_{filiale}", placeholder="Tutti i mesi")
    verso = c2.radio("Direzione", ["Tutti", "Aumenti", "Cali"], horizontal=True, key=f"lis_v_{filiale}")
    q = c3.text_input("Cerca", placeholder="codice, IFB, descrizione", key=f"lis_q_{filiale}")
    f = ev
    if mesi:
        f = f[f["Mese"].isin(mesi)]
    if verso == "Aumenti":
        f = f[f["Δ"] > 0]
    elif verso == "Cali":
        f = f[f["Δ"] < 0]
    if q.strip():
        t = q.strip().lower()
        f = f[f[["Articolo", "IFB", "Descrizione"]].astype(str).apply(
            lambda s: s.str.lower().str.contains(t, regex=False)).any(axis=1)]
    st.caption(f"{len(f)} variazioni su {len(ev)} · rosso = prezzo in aumento, verde = in calo")
    _tabella(f.sort_values(["Mese", "Δ %"], key=lambda s: s.map({m: i for i, m in enumerate(labels)}) if s.name == "Mese" else s,
                           ascending=[False, False]), 460, ("Δ", "Δ %"), euro=True)
    # prezzo di un articolo nel tempo
    st.divider()
    st.markdown("##### Prezzo di un articolo nel tempo")
    qa = st.text_input("Articolo", placeholder="codice, IFB o descrizione", key=f"lis_a_{filiale}",
                       label_visibility="collapsed").strip().lower()
    if qa:
        ult_items = snaps[-1]["items"]
        match = [(k, v) for k, v in ult_items.items() if qa in k[1].lower() or qa == v["ifb"].lower() or qa in v["desc"].lower()][:40]
        if not match:
            st.info("Nessun articolo trovato.")
            return
        k, v = st.selectbox("Articolo trovato", match, format_func=lambda kv: f"{kv[0][1]} · {kv[1]['ifb']} · {kv[1]['desc'][:60]}",
                            key=f"lis_s_{filiale}") if len(match) > 1 else match[0]
        serie = pd.DataFrame([{"Versione": lab, "Articolo": k[1], "Prezzo di acquisto (€)": s["items"][k]["price"]}
                              for lab, s in zip(labels, snaps) if k in s["items"] and s["items"][k]["price"]])
        if serie["Versione"].nunique() < 2:
            st.info("Servono almeno due versioni con un prezzo per questo articolo.")
            return
        marc = marche_caricamento(filiale, os.path.join(os.path.dirname(HERE), "") and
                                  os.path.join(os.path.dirname(os.path.dirname(HERE)), sv.BRANCH["folder"]),
                                  k[1], v["ifb"], labels)
        st.altair_chart(grafico_andamento(serie, labels, "Prezzo di acquisto (€)", "Articolo", marc, 280), width="stretch")
        um = {lab: "/".join(s["items"][k]["um"]) for lab, s in zip(labels, snaps) if k in s["items"]}
        if len(set(um.values())) > 1:
            st.warning("L'unità di misura è cambiata nel periodo: " + " · ".join(f"{a}: {b}" for a, b in um.items())
                       + ". I prezzi prima e dopo non sono confrontabili direttamente.")


# ---- pagina ANAGRAFICA --------------------------------------------------------------------------
def pagina_anagrafica(filiale, versioni, carica):
    _intestazione("Anagrafica", "Cosa e' cambiato nei dati degli articoli: unità di misura, pallettizzazione, AIEM, duties "
                  "alcolici, resa (MTS/MTO/FOR), modalità di trasporto, destinazioni, fornitore.")
    if len(versioni) < 2:
        st.info("Servono almeno due versioni del modello.")
        return
    labels, snaps = carica_snaps(versioni, filiale, carica)
    _note(filiale, labels)
    ev = eventi_anagrafica(labels, snaps)
    if ev.empty:
        st.info("Nessun cambio di anagrafica tra le versioni disponibili.")
        return
    c1, c2, c3 = st.columns([3, 3, 3])
    tipi = c1.multiselect("Tipo di cambio", sorted(ev["Tipo"].unique()), key=f"ana_t_{filiale}", placeholder="Tutti i tipi")
    mesi = c2.multiselect("Mese", labels[1:], key=f"ana_m_{filiale}", placeholder="Tutti i mesi")
    q = c3.text_input("Cerca", placeholder="codice, IFB, descrizione", key=f"ana_q_{filiale}")
    f = ev
    if tipi:
        f = f[f["Tipo"].isin(tipi)]
    if mesi:
        f = f[f["Mese"].isin(mesi)]
    if q.strip():
        t = q.strip().lower()
        f = f[f[["Articolo", "IFB", "Descrizione"]].astype(str).apply(
            lambda s: s.str.lower().str.contains(t, regex=False)).any(axis=1)]
    st.caption(f"{len(f)} cambi su {len(ev)}")
    f = f.assign(_o=f["Mese"].map({m: i for i, m in enumerate(labels)})).sort_values("_o", ascending=False).drop(columns="_o")
    vista = f.reset_index(drop=True).copy()
    # niente colonna "Campo": il nome del campo si legge nel Tipo solo dove serve (es. quale dato di pallettizzazione)
    vista["Tipo"] = [f"{t} · {c}" if t == "Pallettizzazione/imballo" and c else t
                     for t, c in zip(vista["Tipo"], vista["Campo"])]
    vista = vista.drop(columns="Campo")

    def _colore_tipo(v):
        for k, c in TIPI_ANA.items():
            if str(v).startswith(k):
                return f"background-color: {caricamenti._mix(c, 0.16)}; color: {INK}; font-weight: 600"
        return ""

    sty = vista.style.map(_colore_tipo, subset=["Tipo"])
    st.dataframe(sty, hide_index=True, width="stretch", height=min(520, 40 + 29 * max(len(f), 3)))


# ---- pagina LOGISTICA ---------------------------------------------------------------------------
def pagina_logistica(filiale, versioni, carica):
    _intestazione("Logistica · storico delle tariffe", "Quando e' cambiata una tariffa di COSTS (LOG) usata dalle formule "
                  "del modello, di quanto, e come e' andata nel tempo.")
    if len(versioni) < 2:
        st.info("Servono almeno due versioni del modello.")
        return
    labels, snaps = carica_snaps(versioni, filiale, carica)
    _note(filiale, labels)
    ev = eventi_logistica(filiale, labels, snaps)
    if ev.empty:
        st.success("**Nessuna tariffa è cambiata** tra le versioni disponibili.")
        return
    f = ev.assign(_o=ev["Mese"].map({m: i for i, m in enumerate(labels)})).sort_values(
        ["Voce", "Colonna", "_o"]).drop(columns=["_o", "_rc"])
    _tabella(f, 460, ("Δ %",))
