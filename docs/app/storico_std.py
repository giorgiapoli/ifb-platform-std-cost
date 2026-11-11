"""
Storico giornaliero dello standard cost (Canarie): si carica l'export "item statistics" (una riga per articolo e per giorno)
e si risponde a "come e' andato lo standard di questo articolo e quando e' cambiato?".

Colonne dell'export: Execution Date, Item No_, IFB Item, Description, Section Description, Standard Cost (prodotti NON IFB,
standard della filiale), SC GRANCANARIA, SC TENERIFE, SC FUERTE, SC LANZAROTE (prodotti IFB, uno per isola).

L'export di Business Central si ferma a 150.000 righe ("Exported data exceeded the allowed volume"): se succede, mancano gli
articoli in fondo all'elenco. I file caricati si salvano e si uniscono tra loro, quindi si puo' rifare l'export a pezzi
(per periodo o per gruppo) e caricare tutto.
"""
import io
import os
import pickle
import re
from datetime import datetime

import altair as alt
import numpy as np
import pandas as pd
import streamlit as st

HERE = os.path.dirname(os.path.abspath(__file__))
import percorsi
CARTELLA = os.path.join(percorsi.DATI, "STORICO_STD")
COLONNE = ["data", "item", "ifb", "desc", "sez", "std", "gc", "tf", "fue", "lan"]
NOMI = {"std": "Standard Cost", "gc": "Gran Canaria", "tf": "Tenerife", "fue": "Fuerteventura", "lan": "Lanzarote"}
COLORI = {"Standard Cost": "#0F766E", "Gran Canaria / Tenerife": "#2563EB", "Gran Canaria": "#2563EB", "Tenerife": "#7C3AED",
          "Lanzarote / Fuerteventura": "#D97706", "Lanzarote": "#D97706", "Fuerteventura": "#DB2777"}
INK, MUTED = "#14213D", "#667085"

_FIL = {"k": "CANARIE", "v": "€"}          # filiale di cui si leggono/salvano gli export (ognuna ha la sua cartella) e valuta


def imposta_filiale(k, valuta="€"):
    _FIL["k"] = k
    _FIL["v"] = valuta          # Hong Kong: lo standard cost e' in HKD


def _cartella():
    return CARTELLA if _FIL["k"] == "CANARIE" else os.path.join(CARTELLA, _FIL["k"])


def _norm(c):
    return re.sub(r"[^a-z0-9]", "", str(c).lower())


# ---- lettura e archivio -----------------------------------------------------------------------------
def leggi_export(dati):
    """Legge un export. Ritorna (tabella normalizzata, troncato, ultimo articolo letto)."""
    grezzo = pd.read_excel(io.BytesIO(dati))
    nomi = {_norm(c): c for c in grezzo.columns}
    if "executiondate" in nomi and "standardcost" in nomi and "sclanzarote" not in nomi:
        # formato Hong Kong: un solo standard cost (executiondate, mastercode, Item No, ifbitem, Description, transportation,
        # Section description, Standard cost)
        col = lambda k: grezzo[nomi[k]] if k in nomi else pd.Series([np.nan] * len(grezzo), index=grezzo.index)
        sez = col("sectiondescription").astype("object")
        if "transportation" in nomi:
            sez = (sez.fillna("").astype(str) + " · " + col("transportation").fillna("").astype(str)).str.strip(" ·")
        base = pd.DataFrame({"data": col("executiondate"), "item": col("itemno"), "ifb": col("ifbitem"),
                             "desc": col("description"), "sez": sez, "std": col("standardcost"),
                             "gc": np.nan, "tf": np.nan, "fue": np.nan, "lan": np.nan})
        grezzo = base[COLONNE]
    else:
        if grezzo.shape[1] < len(COLONNE):
            raise ValueError("Il file ha meno colonne del previsto (servono le 10 dell'export item statistics).")
        grezzo = grezzo.iloc[:, :len(COLONNE)]
        grezzo.columns = COLONNE
    data = pd.to_datetime(grezzo["data"], errors="coerce")
    righe_note = grezzo[data.isna()]["data"].dropna().astype(str)
    troncato = bool(righe_note.str.contains("exceeded|omitted", case=False).any())
    df = grezzo[data.notna()].copy()
    df["data"] = data[data.notna()]
    df["item"] = df["item"].astype(str).str.strip()
    for c in ("std", "gc", "tf", "fue", "lan"):
        df[c] = pd.to_numeric(df[c], errors="coerce")
    ultimo = df["item"].iloc[-1] if len(df) else ""
    # filtri applicati all'export (riga di nota in fondo al file): servono a capire perche' un articolo puo' mancare
    fil = sorted({r.strip() for t in righe_note for r in str(t).splitlines()
                  if r.strip() and "filtri applicati" not in r.lower() and "exported data" not in r.lower()})
    out = df.reset_index(drop=True)
    out.attrs["filtri"] = "; ".join(fil)
    return out, troncato, ultimo


def salva_export(nome, dati):
    """Legge e salva l'export nell'archivio. Ritorna le informazioni sul file."""
    df, troncato, ultimo = leggi_export(dati)
    os.makedirs(_cartella(), exist_ok=True)
    info = dict(nome=nome, caricato=datetime.now(), righe=len(df), articoli=int(df["item"].nunique()), filtri=df.attrs.get("filtri", ""),
                dal=df["data"].min(), al=df["data"].max(), troncato=troncato, ultimo=ultimo)
    f = os.path.join(_cartella(), f"{datetime.now():%Y%m%d_%H%M%S}_{os.path.splitext(os.path.basename(nome))[0][:40]}.pkl")
    with open(f, "wb") as fh:
        pickle.dump(dict(info=info, df=df), fh, protocol=pickle.HIGHEST_PROTOCOL)
    return info


def file_salvati():
    if not os.path.isdir(_cartella()):
        return []
    return sorted(os.path.join(_cartella(), x) for x in os.listdir(_cartella()) if x.endswith(".pkl"))


def _file_tolti():
    return os.path.join(_cartella(), "mesi_tolti.json")


def mesi_tolti():
    """Mesi (AAAA-MM) che l'utente ha tolto dal calcolo: i file restano, ma quei mesi non si usano."""
    try:
        import json
        with open(_file_tolti(), encoding="utf-8") as fh:
            return sorted(json.load(fh))
    except Exception:
        return []


def imposta_mese(chiave, tolto):
    import json
    t = set(mesi_tolti())
    (t.add if tolto else t.discard)(chiave)
    os.makedirs(_cartella(), exist_ok=True)
    with open(_file_tolti(), "w", encoding="utf-8") as fh:
        json.dump(sorted(t), fh)


def _firma():
    # la firma comprende i file e i mesi tolti: quando cambiano, i calcoli si rifanno
    return tuple((f, os.path.getmtime(f)) for f in file_salvati()) + (("__tolti__" + ",".join(mesi_tolti()), 0.0),)


def _file_firma(firma):
    return [(f, m) for f, m in firma if not f.startswith("__tolti__")]


@st.cache_data(show_spinner=False)
def _carica_grezzo(firma):
    """Unisce tutti i file salvati (anche i mesi tolti): per lo stesso articolo e giorno vale il file piu' recente."""
    pezzi, infos = [], []
    for f, _ in _file_firma(firma):
        with open(f, "rb") as fh:
            c = pickle.load(fh)
        c["info"]["file"] = f
        pezzi.append(c["df"])
        infos.append(c["info"])
    if not pezzi:
        return pd.DataFrame(columns=COLONNE), []
    df = pd.concat(pezzi, ignore_index=True).drop_duplicates(["data", "item"], keep="last")
    return df.sort_values(["item", "data"]).reset_index(drop=True), infos


@st.cache_data(show_spinner=False)
def _carica_tutti(firma):
    """Come _carica_grezzo, ma senza i mesi che l'utente ha tolto."""
    df, infos = _carica_grezzo(firma)
    tolti = [x for f, _ in firma if f.startswith("__tolti__") for x in f[9:].split(",") if x]
    if tolti and len(df):
        df = df[~df["data"].dt.strftime("%Y-%m").isin(tolti)].reset_index(drop=True)
    return df, infos


def elimina_file(f):
    os.remove(f)


# ---- serie ed eventi ------------------------------------------------------------------------------------
def serie_articolo(df, item):
    """Serie giornaliere dell'articolo, pronte per il grafico: zero = standard non calcolato (lascia un buco).
    Le isole con lo stesso valore in tutti i giorni si uniscono (Gran Canaria / Tenerife, Lanzarote / Fuerteventura)."""
    x = df[df["item"] == item].sort_values("data")
    if x.empty:
        return pd.DataFrame(columns=["data", "serie", "valore"])
    colonne = {"std": "Standard Cost"}
    colonne["gc"] = "Gran Canaria / Tenerife" if np.allclose(x["gc"].fillna(0), x["tf"].fillna(0)) else "Gran Canaria"
    if colonne["gc"] == "Gran Canaria":
        colonne["tf"] = "Tenerife"
    colonne["fue"] = "Lanzarote / Fuerteventura" if np.allclose(x["fue"].fillna(0), x["lan"].fillna(0)) else "Fuerteventura"
    if colonne["fue"] == "Fuerteventura":
        colonne["lan"] = "Lanzarote"
    out = []
    for c, nome in colonne.items():
        v = x[c].where(x[c].fillna(0) != 0)
        if v.notna().any():
            out.append(pd.DataFrame({"data": x["data"], "serie": nome, "valore": v}))
    return pd.concat(out, ignore_index=True) if out else pd.DataFrame(columns=["data", "serie", "valore"])


@st.cache_data(show_spinner=False)
def eventi_tutti(firma):
    """Tutti i cambi di standard: una riga per articolo, giorno e destinazione (le isole con lo stesso cambio si uniscono)."""
    df, _ = _carica_tutti(firma)
    if df.empty:
        return pd.DataFrame(columns=["data", "item", "desc", "serie", "prima", "dopo"])
    righe = []
    for c in ("std", "gc", "tf", "fue", "lan"):
        v = df[c].fillna(0)
        prec = v.groupby(df["item"]).shift(1)
        gg = df["item"] == df["item"].shift(1)
        ch = prec.notna() & gg & ((v - prec).abs() > 1e-9)
        e = df.loc[ch, ["data", "item", "desc"]].copy()
        e["col"], e["prima"], e["dopo"] = c, prec[ch].values, v[ch].values
        righe.append(e)
    ev = pd.concat(righe, ignore_index=True)
    # unisco le isole che cambiano insieme
    gruppi = {"gc": "G", "tf": "G", "fue": "L", "lan": "L", "std": "S"}
    ev["g"] = ev["col"].map(gruppi)
    chiavi = ["data", "item", "desc", "g", "prima", "dopo"]
    ev = ev.groupby(chiavi, dropna=False)["col"].apply(lambda s: tuple(sorted(set(s)))).reset_index()

    def nome(t):
        if t == ("std",):
            return "Standard Cost"
        if set(t) == {"gc", "tf"}:
            return "Gran Canaria / Tenerife"
        if set(t) == {"fue", "lan"}:
            return "Lanzarote / Fuerteventura"
        return " / ".join(NOMI[x] for x in t)

    ev["serie"] = ev["col"].map(nome)
    ev["delta"] = np.where((ev["prima"] > 0) & (ev["dopo"] > 0), ev["dopo"] - ev["prima"], np.nan)
    ev["delta_pct"] = np.where((ev["prima"] > 0) & (ev["dopo"] > 0), ev["dopo"] / ev["prima"] - 1, np.nan)
    ev["nota"] = np.where(ev["prima"] == 0, "nuovo standard (prima non calcolato)",
                          np.where(ev["dopo"] == 0, "standard azzerato (non piu' calcolato)", ""))
    return ev.drop(columns=["g", "col"]).sort_values(["data", "item"]).reset_index(drop=True)


# ---- interfaccia ----------------------------------------------------------------------------------------
def _it(txt):
    return txt.replace(",", "§").replace(".", ",").replace("§", ".")


def _kpi(col, label, value, sub=""):
    col.markdown(f'<div class="kpi"><div class="l">{label}</div><div class="v">{value}</div><div class="s">{sub}</div></div>',
                 unsafe_allow_html=True)


def a_sistema(codici, destinazione):
    """Standard di un articolo realmente a sistema (export dello storico), per destinazione ('GC-TF' o 'LAN-FUE').
    Ritorna {valore, data, ultimo_cambio} oppure None se l'articolo non e' nell'export."""
    df, _ = _carica_tutti(_firma())
    col = {"GC-TF": "gc", "LAN-FUE": "lan", "HK": "std"}.get(destinazione)
    if df.empty or col is None:
        return None
    cod = {str(c) for c in codici if c}
    x = df[df["item"].astype(str).isin(cod) | df["ifb"].astype(str).isin(cod)].sort_values("data")
    if x.empty:
        return None
    v = x[col].fillna(0)
    cambi = x.loc[(v != v.shift(1)) & v.shift(1).notna(), "data"]
    return {"valore": float(v.iloc[-1]), "data": x["data"].iloc[-1], "ultimo_cambio": cambi.iloc[-1] if len(cambi) else None}


SOGLIA_FORTE = 0.15          # variazione oltre il 15%: evidenziata come anomala


def _e4(v):
    return _it(f"{v:.4f}") + " " + _FIL["v"]


def storia_reale(codici):
    """Storia dell'articolo a sistema, in poche righe: da quanto partiva, quando e di quanto e' cambiato, quando si e' bloccato.
    Ritorna [(tipo, data, testo)] (tipo: up/down/stop/info/ok) oppure None se l'articolo non e' nell'export."""
    firma = _firma()
    df, _ = _carica_tutti(firma)
    if df.empty:
        return None
    cod = {str(c) for c in codici if c}
    x = df[df["item"].astype(str).isin(cod) | df["ifb"].astype(str).isin(cod)]
    if x.empty:
        return None
    item = x["item"].iloc[0]
    serie = serie_articolo(df, item)
    ev = eventi_tutti(firma)
    e = ev[ev["item"] == item].sort_values("data")
    g0, g1 = x["data"].min(), x["data"].max()
    righe = []
    ini = []
    for nome, g in serie.sort_values("data").groupby("serie", sort=False):
        v = g.dropna(subset=["valore"])
        if not v.empty:
            ini.append((nome, v.iloc[0]))
    if ini:
        righe.append(("info", g0, "Inizio dei dati: " + " · ".join(f"{n} <b>{_e4(r['valore'])}</b>" for n, r in ini)))
    # cambi: le destinazioni che cambiano insieme si uniscono in una riga
    gruppi = {}
    for r in e.itertuples():
        gruppi.setdefault((r.data, r.prima, r.dopo, r.nota), []).append(r.serie)
    voci = []
    for (data, prima, dopo, nota), nomi in gruppi.items():
        n = " + ".join(nomi)
        if str(nota).startswith("standard azzerato"):
            voci.append(("stop", data, f"{n}: standard <b>azzerato</b> (non più calcolato), ultimo valore {_e4(prima)}"))
        elif str(nota).startswith("nuovo standard"):
            voci.append(("info", data, f"{n}: <b>nuovo standard {_e4(dopo)}</b> (prima non calcolato)"))
        else:
            pct = dopo / prima - 1 if prima else 0
            voci.append(("alert" if abs(pct) >= SOGLIA_FORTE else ("up" if dopo > prima else "down"), data,
                         f"{n}: da {_e4(prima)} a <b>{_e4(dopo)}</b> ({_it(f'{pct:+.1%}')})"))
    if not voci:
        righe.append(("ok", None, "Standard <b>sempre uguale</b> in tutto il periodo coperto dai dati."))
    else:
        if len(voci) > 6:
            righe.append(("info", None, f"… e altri {len(voci) - 6} cambi precedenti"))
        righe += voci[-6:]
    # situazione oggi
    attuali, fermi = [], []
    for nome, g in serie.sort_values("data").groupby("serie", sort=False):
        ultimo = g.iloc[-1]
        if pd.notna(ultimo["valore"]):
            attuali.append(f"{nome} <b>{_e4(ultimo['valore'])}</b>")
        else:
            valido = g.dropna(subset=["valore"])
            if not valido.empty:
                dopo_ultimo = g[g["data"] > valido.iloc[-1]["data"]]["data"].min()
                fermi.append(f"{nome}: non calcolato dal {dopo_ultimo:%d/%m/%Y} (ultimo valore {_e4(valido.iloc[-1]['valore'])})")
    if attuali:
        righe.append(("ok", g1, "Oggi: " + " · ".join(attuali)))
        # da quanto tempo lo standard non si aggiorna
        ultimo_cambio = e["data"].max() if len(e) else None
        if ultimo_cambio is not None and pd.notna(ultimo_cambio):
            n = (g1 - ultimo_cambio).days
            righe.append(("stop" if n >= 60 else "info", None,
                          f"Non si aggiorna da <b>{n} giorni</b> (ultimo cambio il {ultimo_cambio:%d/%m/%Y})"))
        else:
            n = (g1 - g0).days
            righe.append(("stop" if n >= 60 else "info", None,
                          f"Non è mai cambiato nei <b>{n} giorni</b> coperti dai dati (dal {g0:%d/%m/%Y})"))
    for t in fermi:
        righe.append(("stop", None, t))
    return righe


_COL_TIPO = {"up": "#B42318", "down": "#027A48", "stop": "#B54708", "info": "#667085", "ok": "#027A48", "alert": "#B42318",
             "carica": "#0E7490"}


def scheda_storia(tipo, titolo, sottotitolo, righe):
    """Riquadro con la storia di un articolo: tipo 'reale' (macchina, viola) o 'simulato' (modello, blu)."""
    col, bg, nome = (("#5925DC", "#F4F3FF", "REALE") if tipo == "reale" else ("#175CD3", "#EFF8FF", "SIMULATO"))
    corpo = ""
    n_att = sum(1 for t, _, _ in (righe or []) if t in ("alert", "stop"))
    for t, d, testo in (righe or []):
        data = f"{d:%d/%m/%Y}" if d is not None and not isinstance(d, str) else (d or "")
        if t == "alert":
            stile = "background:#FEF3F2;border-left:4px solid #B42318;border-radius:8px;padding:5px 8px;"
            testo = ('<span style="background:#B42318;color:#fff;border-radius:6px;padding:0 7px;font-size:.68rem;font-weight:800;'
                     'letter-spacing:.04em;margin-right:6px">VARIAZIONE FORTE</span>' + testo)
        elif t == "carica":          # richiesta di caricamento inviata (articolo presente negli output della macro)
            stile = "background:#ECFEFF;border-left:4px solid #0E7490;border-radius:8px;padding:5px 8px;"
            testo = ('<span style="background:#0E7490;color:#fff;border-radius:6px;padding:0 7px;font-size:.68rem;font-weight:800;'
                     'letter-spacing:.04em;margin-right:6px">CARICAMENTO</span>' + testo)
        elif t == "stop":
            stile = "background:#FFFAEB;border-left:4px solid #B54708;border-radius:8px;padding:5px 8px;"
            testo = ('<span style="background:#B54708;color:#fff;border-radius:6px;padding:0 7px;font-size:.68rem;font-weight:800;'
                     'letter-spacing:.04em;margin-right:6px">ATTENZIONE</span>' + testo)
        else:
            stile = ""
        corpo += (f'<div style="display:grid;grid-template-columns:86px 1fr;gap:8px;margin:5px 0;font-size:.88rem;line-height:1.35;{stile}">'
                  f'<div style="color:#667085;font-variant-numeric:tabular-nums"><span style="display:inline-block;width:8px;height:8px;'
                  f'border-radius:50%;background:{_COL_TIPO.get(t, "#667085")};margin-right:6px"></span>{data}</div>'
                  f'<div style="color:#14213D">{testo}</div></div>')
    if not corpo:
        corpo = '<div style="color:#667085;font-size:.88rem;margin-top:6px">Nessun dato.</div>'
    return (f'<div style="background:#fff;border:1px solid #E4E7EC;border-top:5px solid {col};border-radius:14px;padding:14px 16px;'
            f'height:100%"><div style="display:flex;align-items:center;gap:10px"><span style="background:{bg};color:{col};'
            f'border:1px solid {col}33;border-radius:999px;padding:2px 12px;font-size:.78rem;font-weight:800;letter-spacing:.06em">{nome}</span>'
            f'<span style="font-weight:700;color:#14213D">{titolo}</span>'
            + (f'<span style="margin-left:auto;background:#B42318;color:#fff;border-radius:999px;padding:2px 11px;font-size:.78rem;font-weight:700">'
               f'⚠ {n_att} da controllare</span>' if n_att else '') + '</div>'
            f'<div style="color:#667085;font-size:.8rem;margin:4px 0 8px 0">{sottotitolo}</div>{corpo}</div>')


def elenco_articoli():
    """(codice, codice IFB, descrizione) di tutti gli articoli dell'export caricato: servono per la ricerca."""
    df, _ = _carica_tutti(_firma())
    if df.empty:
        return []
    a = df.drop_duplicates("item", keep="last")
    return [(r.item, "" if pd.isna(r.ifb) else str(r.ifb), str(r.desc)) for r in a.itertuples()]


def _tab_eventi(ev):
    t = ev.rename(columns={"data": "Data", "item": "Articolo", "desc": "Descrizione", "serie": "Destinazione", "prima": "Prima",
                           "dopo": "Dopo", "delta": "Δ", "delta_pct": "Δ %", "nota": "Nota"})
    cols = ["Data", "Articolo", "Descrizione", "Destinazione", "Prima", "Dopo", "Δ", "Δ %", "Nota"]
    cols = [c for c in cols if c in t.columns and (c != "Articolo" or t["Articolo"].nunique() > 1)
            and (c != "Descrizione" or t["Descrizione"].nunique() > 1)]
    _fm = {"Data": "{:%d/%m/%Y}", "Prima": "{:.4f} " + _FIL["v"], "Dopo": "{:.4f} " + _FIL["v"], "Δ": "{:+.4f} " + _FIL["v"], "Δ %": "{:+.1%}"}
    sty = t[cols].style.format({k: v for k, v in _fm.items() if k in cols}, na_rep="")

    def colore(v):
        if pd.isna(v) or v == 0:
            return ""
        return "color:#B42318;font-weight:600" if v > 0 else "color:#027A48;font-weight:600"

    sty = sty.map(colore, subset=[c for c in ("Δ", "Δ %") if c in cols])
    st.dataframe(sty, hide_index=True, width="stretch", height=min(520, 46 + 29 * max(len(t), 3)))


@st.dialog("Articoli cambiati a sistema", width="large")
def _giorno_dettaglio(e_giorno, giorno):
    """Sottopagina: gli articoli il cui standard e' cambiato in macchina in quel giorno."""
    n = e_giorno["item"].nunique()
    st.markdown(f"**{giorno:%d/%m/%Y}** · {n} articoli con lo standard cambiato a sistema (export del report)")
    _tab_eventi(e_giorno.sort_values("delta_pct", key=lambda x: x.abs(), ascending=False, na_position="last").drop(columns=["data"]))


def unisci_caricamenti(storia, car):
    """Inserisce le righe dei caricamenti (richieste inviate) nella storia a sistema, in ordine di data."""
    if not car:
        return storia
    st_ = storia or []
    testa = [x for x in st_ if x[1] is None and x[0] == "info"]
    datate = sorted([x for x in st_ if x[1] is not None and x[0] != "ok"] + car, key=lambda x: pd.Timestamp(x[1]))
    coda = [x for x in st_ if x[0] == "ok" or (x[1] is None and x[0] != "info")]
    return testa + datate + coda


def nota_mancante():
    """Perche' un articolo puo' non comparire nello storico caricato: filtri dell'export e/o troncamento."""
    _, infos = _carica_tutti(_firma())
    m = []
    fil = sorted({i.get("filtri") for i in infos if i.get("filtri")})
    if fil:
        m.append("l'export caricato ha un filtro («" + "; ".join(fil) + "»): gli articoli fuori dal filtro non ci sono")
    if any(i.get("troncato") for i in infos):
        m.append("l'export si è fermato al limite di 150.000 righe")
    return " · ".join(m)


_MESI_IT = ["gennaio", "febbraio", "marzo", "aprile", "maggio", "giugno", "luglio", "agosto", "settembre", "ottobre", "novembre",
            "dicembre"]


def copertura_mesi(df, infos):
    """Mese per mese: giorni coperti dall'export, articoli, file e se il file e' incompleto. Serve a caricare un file per mese."""
    import calendar
    oggi = pd.Timestamp.today().to_period("M")
    vuoto = df is None or df.empty
    inizio = min(df["data"].min().to_period("M"), oggi - 5) if not vuoto else oggi - 5
    fine = max(df["data"].max().to_period("M"), oggi) if not vuoto else oggi
    out = []
    for p in pd.period_range(inizio, fine, freq="M"):
        d = df[df["data"].dt.to_period("M") == p] if not vuoto else pd.DataFrame(columns=["data", "item"])
        giorni_tot = calendar.monthrange(p.year, p.month)[1] if p != oggi else int(pd.Timestamp.today().day)
        da, a = p.start_time, p.end_time
        file_ = [i for i in infos if i["dal"] <= a and i["al"] >= da]
        out.append(dict(chiave=str(p), tolto=str(p) in mesi_tolti(), mese=f"{_MESI_IT[p.month - 1].capitalize()} {p.year}",
                        giorni=int(d["data"].nunique()), tot=giorni_tot,
                        articoli=int(d["item"].nunique()), file=[i["nome"] for i in file_], troncato=any(i["troncato"] for i in file_)))
    return out


def corpo_caricamento(infos, df=None):
    """Caricamento dei file dello standard giornaliero e elenco di quelli gia' salvati (scheda «File da caricare»)."""
    files = st.file_uploader("Un file unico con più mesi, oppure più file insieme", type=["xlsx"],
                             accept_multiple_files=True, key="ss_file",
                             help="Più file si uniscono: per lo stesso articolo e giorno vale l'ultimo caricato.")
    import assistente
    assistente.bolla("ss_file", "Lo storico dello standard in macchina",
                     "Carica l'export del report (una riga per articolo e giorno). Puoi caricare più file: per lo stesso articolo e giorno "
                     "vale l'ultimo. Se l'export si ferma per il limite di righe, rifallo a pezzi e carica tutto.", larghezza=460)
    if files and st.button("Salva i file caricati", type="primary"):
        esiti = []
        for f in files:
            try:
                i = salva_export(f.name, f.getvalue())
                esiti.append(f"{f.name}: {i['righe']:,} righe, {i['articoli']} articoli, dal {i['dal']:%d/%m/%Y} al {i['al']:%d/%m/%Y}"
                             .replace(",", "."))
            except Exception as ex:
                esiti.append(f"{f.name}: non letto ({ex})")
        st.session_state["ss_esiti"] = esiti
        st.rerun()
    for e in st.session_state.pop("ss_esiti", []):
        st.caption(e)
    dg, ig = _carica_grezzo(_firma())          # anche i mesi tolti: servono per poterli rimettere
    cop = copertura_mesi(dg, ig)
    if cop:
        st.markdown("**Mesi** · carica un file per mese qui sotto, oppure un file unico con più mesi (o più file insieme) nel campo "
                    "sopra: li assegno io al mese giusto. Un export copre al massimo circa un mese e mezzo.")
        for n, c in enumerate(cop, 1):
            if c["tolto"]:
                chip, col, bg = "tolto dal calcolo", "#475467", "#F2F4F7"
            elif c["giorni"] >= c["tot"] and not c["troncato"]:
                chip, col, bg = "completo", "#027A48", "#ECFDF3"
            elif c["giorni"] == 0:
                chip, col, bg = "mancante", "#B42318", "#FEF3F2"
            else:
                chip, col, bg = ("incompleto (troncato)" if c["troncato"] else f"parziale: {c['giorni']}/{c['tot']} giorni"), "#B54708", "#FFFAEB"
            k1, k2, k3, k4, kf, k5 = st.columns([0.45, 1.8, 2.2, 3.0, 3.2, 1.2], vertical_alignment="center")
            k1.markdown(f"<span style='color:#667085'>{n})</span>", unsafe_allow_html=True)
            k2.markdown(f"**{c['mese']}**")
            k3.markdown(f"<span style='background:{bg};color:{col};border-radius:999px;padding:1px 10px;font-size:.78rem;font-weight:600'>{chip}</span>",
                        unsafe_allow_html=True)
            k4.markdown(f"<span style='color:#667085;font-size:.88rem'>{c['articoli']} articoli · {', '.join(c['file']) or '—'}</span>",
                        unsafe_allow_html=True)
            nk = st.session_state.get("ss_nk", 0)
            fm = kf.file_uploader(f"Carica {c['mese']}", type=["xlsx"], key=f"ss_up_{c['chiave']}_{_FIL['k']}_{nk}",
                                  label_visibility="collapsed")
            if fm is not None:          # il file del mese si salva subito; il campo si svuota
                try:
                    i_ = salva_export(fm.name, fm.getvalue())
                    st.session_state["ss_esiti"] = [f"{fm.name} ({c['mese']}): {i_['righe']:,} righe, {i_['articoli']} articoli, dal "
                                                    f"{i_['dal']:%d/%m/%Y} al {i_['al']:%d/%m/%Y}".replace(",", ".")]
                except Exception as ex:
                    st.session_state["ss_esiti"] = [f"{fm.name}: non letto ({ex})"]
                st.session_state["ss_nk"] = nk + 1
                st.rerun()
            if c["giorni"]:          # un mese si toglie dal calcolo e si rimette quando serve (i file restano salvati)
                if k5.button("Rimetti" if c["tolto"] else "Togli", key=f"ss_mese_{c['chiave']}_{_FIL['k']}",
                             type="primary" if c["tolto"] else "secondary"):
                    imposta_mese(c["chiave"], not c["tolto"])
                    st.rerun()
    if infos:
        st.markdown("**File salvati**")
        for i in infos:
            c1, c2 = st.columns([8, 1])
            _m = ["gennaio", "febbraio", "marzo", "aprile", "maggio", "giugno", "luglio", "agosto", "settembre", "ottobre",
                  "novembre", "dicembre"]
            _p = f"{i['dal'].day} {_m[i['dal'].month - 1]} {i['dal'].year} → {i['al'].day} {_m[i['al'].month - 1]} {i['al'].year}"
            c1.markdown(f"**{i['nome']}** · periodo coperto: **{_p}** · {i['articoli']} articoli<br>"
                        f"<span style='color:#667085;font-size:.85rem'>caricato il {i['caricato']:%d/%m/%Y alle %H:%M}</span>"
                        + (" <span style='background:#FFFAEB;color:#B54708;border-radius:999px;padding:1px 9px;font-size:.78rem;"
                           "font-weight:600'>incompleto (troncato)</span>" if i["troncato"] else "")
                        + (f" <span style='background:#F2F4F7;color:#475467;border-radius:999px;padding:1px 9px;font-size:.78rem;"
                           f"font-weight:500'>filtro dell'export: {i['filtri']}</span>" if i.get("filtri") else ""), unsafe_allow_html=True)
            if c2.button("Togli", key="ss_del_" + os.path.basename(i["file"])):
                elimina_file(i["file"])
                st.rerun()


def _confronto(df, items, d_da, d_a):
    """Tabella di tutti gli articoli trovati: standard al primo e all'ultimo giorno del periodo, differenza e numero di cambi."""
    anag = df.drop_duplicates("item", keep="last").set_index("item")
    righe = []
    for it in items:
        if it not in anag.index:
            continue
        for nome, g in serie_articolo(df, it).groupby("serie", sort=False):
            g = g[(g["data"].dt.date >= d_da) & (g["data"].dt.date <= d_a)].dropna(subset=["valore"]).sort_values("data")
            if g.empty:
                continue
            v0, v1 = g["valore"].iloc[0], g["valore"].iloc[-1]
            righe.append({"Articolo": it, "Descrizione": anag.loc[it, "desc"], "Destinazione": nome,
                          "Primo giorno": g["data"].iloc[0], "Standard iniziale": v0,
                          "Ultimo giorno": g["data"].iloc[-1], "Standard attuale": v1, "Δ": v1 - v0,
                          "Δ %": (v1 / v0 - 1) if v0 else np.nan, "Cambi": int((g["valore"].diff().abs() > 1e-9).sum())})
    return pd.DataFrame(righe)


def render(filiale, item=None, trovati=None, modello=None):
    """Storico giornaliero dello standard. item = articolo scelto nella ricerca in alto (se c'e', si mostra il suo andamento)."""
    st.markdown("#### Storico giornaliero dello standard cost")
    firma = _firma()
    item_scelto = item                               # il modello (punti) vale solo per l'articolo scelto in alto
    df, infos = _carica_tutti(firma)
    if df.empty:
        st.info("Carica l'export dello standard cost giornaliero nella pagina «File da caricare» (pulsante in alto a destra) per vedere come è cambiato ogni articolo nel tempo.")
        return
    ev = eventi_tutti(firma)
    g0, g1 = df["data"].min(), df["data"].max()

    # ---- periodo: filtro compatto a cursore (vale per tutto quello che c'e' sotto) -------------------
    col_filtro, _vuoto = st.columns([2, 3])
    d_da, d_a = col_filtro.slider("Periodo", min_value=g0.date(), max_value=g1.date(), value=(g0.date(), g1.date()),
                                  format="DD/MM/YYYY", key="ss_range")

    # ---- gli articoli trovati nella ricerca in alto: tutti insieme -----------------------------------
    if trovati and len(trovati) > 1:
        conf = _confronto(df, trovati[:60], d_da, d_a)
        if not conf.empty:
            import filtri
            conf = filtri.filtra(conf, "conf").reset_index(drop=True)
            st.markdown(f"**{len(trovati)} articoli trovati**: standard al primo e all'ultimo giorno del periodo")
            sty = conf.style.format({"Primo giorno": "{:%d/%m/%Y}", "Ultimo giorno": "{:%d/%m/%Y}", "Standard iniziale": "{:.4f} " + _FIL["v"],
                                     "Standard attuale": "{:.4f} " + _FIL["v"], "Δ": "{:+.4f} " + _FIL["v"], "Δ %": "{:+.1%}"}, na_rep="")
            sty = sty.map(lambda v: "" if pd.isna(v) or v == 0 else ("color:#B42318;font-weight:600" if v > 0 else "color:#027A48;font-weight:600"),
                          subset=["Δ", "Δ %"])
            sel = st.dataframe(sty, hide_index=True, width="stretch", height=min(420, 46 + 29 * max(len(conf), 3)),
                               on_select="rerun", selection_mode="single-row", key="ss_conf_tab")
            st.caption("Clicca una riga per vedere sotto il grafico di quell'articolo.")
            if sel.selection.rows:
                item = conf.iloc[sel.selection.rows[0]]["Articolo"]

    # ---- l'articolo scelto in alto ---------------------------------------------------------------------
    if item:
        anag = df.drop_duplicates("item", keep="last").set_index("item")
        if item in anag.index:
            _scheda_articolo(df, ev, item, anag.loc[item], d_da, d_a, modello if item == item_scelto else None)
        else:
            st.caption("Questo articolo non è nell'export caricato (se l'export è stato troncato potrebbe mancare).")

    if item:           # con un articolo scelto si vede solo il suo grafico: il quadro generale compare solo all'inizio
        return
    n_art = df["item"].nunique()
    st.caption(f"Dati dal {g0:%d/%m/%Y} al {g1:%d/%m/%Y} · {n_art} articoli · {len(ev)} cambi di standard")

    # ---- quando cambiano gli standard -----------------------------------------------------------------
    st.markdown("##### Quando sono cambiati gli standard")
    if ev.empty:
        st.info("Nel periodo caricato nessuno standard è cambiato.")
    else:
        e = ev[(ev["data"].dt.date >= d_da) & (ev["data"].dt.date <= d_a)]
        if e.empty:
            st.info("Nessun cambio in questo periodo.")
        else:
            per_giorno = e.groupby("data")["item"].nunique().rename("Articoli cambiati").reset_index()
            sel_g = alt.selection_point(name="giorno", fields=["data"], on="click", clear="dblclick")
            ch = alt.Chart(per_giorno).mark_bar(color="#0F766E", cursor="pointer").add_params(sel_g).encode(
                x=alt.X("data:T", title=None, axis=alt.Axis(format="%d/%m")),
                y=alt.Y("Articoli cambiati:Q", title="Articoli con standard cambiato"),
                opacity=alt.condition(sel_g, alt.value(1), alt.value(0.85)),
                tooltip=[alt.Tooltip("data:T", title="Giorno", format="%d/%m/%Y"), "Articoli cambiati"]).properties(height=200)
            ev_g = st.altair_chart(ch, width="stretch", on_select="rerun", key="ss_chart_giorni")
            import assistente
            assistente.bolla("ss_chart_giorni", "Clicca una barra",
                             "Si apre la sottopagina con gli articoli il cui standard è cambiato quel giorno a sistema, con prima, dopo e variazione.",
                             larghezza=360)
            punti = (ev_g.selection or {}).get("giorno", []) if ev_g is not None else []
            if punti:
                _v = punti[0].get("data")
                if isinstance(_v, (int, float)):         # il browser manda la mezzanotte locale in millisecondi: scelgo il giorno piu' vicino
                    _ms = per_giorno["data"].astype("datetime64[ms]").astype("int64")
                    g = per_giorno["data"].iloc[int((_ms - _v).abs().argmin())]
                else:
                    g = pd.to_datetime(_v).normalize()
                if st.session_state.get("ss_giorno_aperto") != g:          # si apre una volta sola per clic
                    st.session_state["ss_giorno_aperto"] = g
                    _eg = e[e["data"].dt.normalize() == g]
                    if len(_eg):
                        _giorno_dettaglio(_eg, g)
            else:
                st.session_state.pop("ss_giorno_aperto", None)



def _scheda_articolo(df, ev, item, riga, d_da=None, d_a=None, modello=None):
    serie = serie_articolo(df, item)
    ec = ev[ev["item"] == item]
    if d_da is not None:                                   # solo il periodo scelto nel filtro
        serie = serie[(serie["data"].dt.date >= d_da) & (serie["data"].dt.date <= d_a)]
        ec = ec[(ec["data"].dt.date >= d_da) & (ec["data"].dt.date <= d_a)]
    st.markdown(f"**{item} · {riga['desc']}**" + (f" · IFB {riga['ifb']}" if pd.notna(riga["ifb"]) else "")
                + (f" · {riga['sez']}" if pd.notna(riga["sez"]) else ""))
    if serie.empty:
        st.info("Per questo articolo lo standard non risulta calcolato (zero) nel periodo scelto.")
        return
    schede = []
    for nome, g in serie.sort_values("data").groupby("serie", sort=False):
        ultima = g.iloc[-1]
        if pd.notna(ultima["valore"]):
            schede.append((nome, _it(f"{ultima['valore']:.4f} {_FIL['v']}"), f"al {ultima['data']:%d/%m/%Y}"))
        else:       # standard non piu' calcolato (zero) negli ultimi giorni
            valida = g.dropna(subset=["valore"])
            ultimo_ok = valida.iloc[-1]
            schede.append((nome, "non calcolato", f"ultimo valore {_it(f'{ultimo_ok['valore']:.4f}')} {_FIL['v']} il {ultimo_ok['data']:%d/%m/%Y}"))
    k = st.columns(min(len(schede), 4) + 1)
    for col, (nome, valore, sotto) in zip(k, schede):
        _kpi(col, nome, valore, sotto)
    _kpi(k[-1], "Cambi nel periodo", len(ec), (f"ultimo il {ec['data'].max():%d/%m/%Y}" if len(ec) else "standard sempre uguale"))
    st.write("")
    punti = serie.merge(ec.assign(serie_ev=ec["serie"])[["data", "serie"]], on=["data", "serie"], how="inner")
    # stessi colori in tutta la pagina: REALE (macchina) = viola, SIMULATO (modello) = blu; le destinazioni si distinguono per tonalita'
    viola = {"Standard Cost": "#3E1C96", "Gran Canaria": "#5925DC", "Tenerife": "#7A5AF8", "Lanzarote": "#9E77ED", "Fuerteventura": "#B692F6"}
    blu = {"GC": "#175CD3", "LAN": "#53B1FD"}

    def col_reale(nome):
        if nome.startswith("Gran Canaria"):
            return viola["Gran Canaria"]
        if nome.startswith(("Lanzarote", "Fuerteventura")) and "/" in nome:
            return "#9E77ED"
        return viola.get(nome, "#6941C6")

    nomi = list(serie["serie"].unique())
    base = alt.Chart(serie).encode(
        x=alt.X("data:T", title=None, axis=alt.Axis(format="%d/%m")),
        y=alt.Y("valore:Q", title=f"Standard cost ({_FIL['v']})", scale=alt.Scale(zero=False)),
        color=alt.Color("serie:N", title=None, scale=alt.Scale(domain=nomi, range=[col_reale(n) for n in nomi]), legend=None))
    linea = base.mark_line(interpolate="step-after", strokeWidth=2.5)
    marche = alt.Chart(punti).mark_point(filled=True, size=70).encode(
        x="data:T", y="valore:Q", color=alt.Color("serie:N", legend=None, scale=alt.Scale(domain=nomi, range=[col_reale(n) for n in nomi])),
        tooltip=[alt.Tooltip("data:T", title="Dal", format="%d/%m/%Y"), "serie:N", alt.Tooltip("valore:Q", title="Standard", format=".4f")])
    # passando (o cliccando) sulla linea compare il giorno, il valore e da quando vale: punti invisibili lungo tutta la linea
    sh = serie.dropna(subset=["valore"]).sort_values(["serie", "data"]).copy()
    sh["cambio"] = sh.groupby("serie")["valore"].transform(lambda v: (v != v.shift()).cumsum())
    sh["dal"] = sh.groupby(["serie", "cambio"])["data"].transform("min")
    hover = alt.Chart(sh).mark_point(filled=True, size=220, opacity=0.01).encode(
        x="data:T", y="valore:Q",
        tooltip=[alt.Tooltip("serie:N", title="Destinazione"), alt.Tooltip("data:T", title="Giorno", format="%d/%m/%Y"),
                 alt.Tooltip("valore:Q", title=f"Standard ({_FIL['v']})", format=".4f"), alt.Tooltip("dal:T", title="Vale dal", format="%d/%m/%Y")])
    grafico = linea + marche + hover
    chip = lambda c, t: (f'<span style="display:inline-block;width:11px;height:11px;border-radius:3px;background:{c};margin:0 5px 0 12px;'
                         f'vertical-align:-1px"></span>{t}')
    leg_reale = ('<b style="color:#5925DC">REALE</b> (macchina, linea continua)'
                 + "".join(chip(col_reale(n), n) for n in nomi))
    leg_sim = ""
    if modello is not None and len(modello):
        # standard nei modelli Excel: blu, linea tratteggiata con punti, sopra quello a sistema
        m = modello.copy()
        m["serie"] = m["dest"].map({"GC-TF": "Gran Canaria / Tenerife", "LAN-FUE": "Lanzarote / Fuerteventura", "HK": "Standard Cost"})
        if d_da is not None:
            m = m[(m["data"].dt.date >= d_da) & (m["data"].dt.date <= d_a)]
        if len(m):
            dom = list(m["serie"].unique())
            grafico = grafico + alt.Chart(m).mark_line(strokeDash=[6, 4], strokeWidth=2, point=alt.OverlayMarkDef(filled=True, size=70)).encode(
                x="data:T", y="valore:Q",
                color=alt.Color("serie:N", legend=None, scale=alt.Scale(domain=dom, range=[blu["LAN"] if n.startswith(("Lanz", "Fuert")) else blu["GC"] for n in dom])),
                tooltip=[alt.Tooltip("modello:N", title="Modello"), alt.Tooltip("data:T", title="Salvato il", format="%d/%m/%Y"),
                         "serie:N", alt.Tooltip("valore:Q", title="Standard", format=".4f")])
            leg_sim = ('<br><b style="color:#175CD3">SIMULATO</b> (modello Excel, tratteggiata con punti'
                       + ")"
                       + "".join(chip(blu["LAN"] if n.startswith(("Lanz", "Fuert")) else blu["GC"], n.replace(" / ", "/")) for n in dom))
    st.markdown(f'<div style="font-size:.84rem;color:#344054;margin:2px 0 4px 0;line-height:1.7">{leg_reale}{leg_sim}</div>',
                unsafe_allow_html=True)
    st.altair_chart(grafico.resolve_scale(color="independent").properties(height=300), width="stretch")   # colori indipendenti: viola = reale, blu = simulato
    if ec.empty:
        st.caption("Nessun cambio di standard nel periodo caricato.")
    else:
        st.markdown("**Quando è cambiato** (la data è il primo giorno con il nuovo valore)")
        _tab_eventi(ec.sort_values("data", ascending=False))
