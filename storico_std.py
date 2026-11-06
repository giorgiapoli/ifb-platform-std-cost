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
from datetime import datetime

import altair as alt
import numpy as np
import pandas as pd
import streamlit as st

HERE = os.path.dirname(os.path.abspath(__file__))
CARTELLA = os.path.join(HERE, "STORICO_STD")
COLONNE = ["data", "item", "ifb", "desc", "sez", "std", "gc", "tf", "fue", "lan"]
NOMI = {"std": "Standard Cost", "gc": "Gran Canaria", "tf": "Tenerife", "fue": "Fuerteventura", "lan": "Lanzarote"}
COLORI = {"Standard Cost": "#0F766E", "Gran Canaria / Tenerife": "#2563EB", "Gran Canaria": "#2563EB", "Tenerife": "#7C3AED",
          "Lanzarote / Fuerteventura": "#D97706", "Lanzarote": "#D97706", "Fuerteventura": "#DB2777"}
INK, MUTED = "#14213D", "#667085"


# ---- lettura e archivio -----------------------------------------------------------------------------
def leggi_export(dati):
    """Legge un export. Ritorna (tabella normalizzata, troncato, ultimo articolo letto)."""
    grezzo = pd.read_excel(io.BytesIO(dati))
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
    return df.reset_index(drop=True), troncato, ultimo


def salva_export(nome, dati):
    """Legge e salva l'export nell'archivio. Ritorna le informazioni sul file."""
    df, troncato, ultimo = leggi_export(dati)
    os.makedirs(CARTELLA, exist_ok=True)
    info = dict(nome=nome, caricato=datetime.now(), righe=len(df), articoli=int(df["item"].nunique()),
                dal=df["data"].min(), al=df["data"].max(), troncato=troncato, ultimo=ultimo)
    f = os.path.join(CARTELLA, f"{datetime.now():%Y%m%d_%H%M%S}_{os.path.splitext(os.path.basename(nome))[0][:40]}.pkl")
    with open(f, "wb") as fh:
        pickle.dump(dict(info=info, df=df), fh, protocol=pickle.HIGHEST_PROTOCOL)
    return info


def file_salvati():
    if not os.path.isdir(CARTELLA):
        return []
    return sorted(os.path.join(CARTELLA, x) for x in os.listdir(CARTELLA) if x.endswith(".pkl"))


def _firma():
    return tuple((f, os.path.getmtime(f)) for f in file_salvati())


@st.cache_data(show_spinner=False)
def _carica_tutti(firma):
    """Unisce tutti i file salvati: per lo stesso articolo e giorno vale il file piu' recente."""
    pezzi, infos = [], []
    for f, _ in firma:
        with open(f, "rb") as fh:
            c = pickle.load(fh)
        c["info"]["file"] = f
        pezzi.append(c["df"])
        infos.append(c["info"])
    if not pezzi:
        return pd.DataFrame(columns=COLONNE), []
    df = pd.concat(pezzi, ignore_index=True).drop_duplicates(["data", "item"], keep="last")
    return df.sort_values(["item", "data"]).reset_index(drop=True), infos


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
    sty = t[cols].style.format({"Data": "{:%d/%m/%Y}", "Prima": "{:.4f} €", "Dopo": "{:.4f} €", "Δ": "{:+.4f} €", "Δ %": "{:+.1%}"},
                               na_rep="")

    def colore(v):
        if pd.isna(v) or v == 0:
            return ""
        return "color:#B42318;font-weight:600" if v > 0 else "color:#027A48;font-weight:600"

    sty = sty.map(colore, subset=[c for c in ("Δ", "Δ %") if c in cols])
    st.dataframe(sty, hide_index=True, width="stretch", height=min(520, 46 + 29 * max(len(t), 3)))


def _caricamento(df, infos, aperto):
    """Riquadro per caricare i file e vedere quelli gia' salvati."""
    with st.expander("Carica o aggiorna i dati (export «item statistics»)", expanded=aperto):
        files = st.file_uploader("Export giornaliero dello standard per articolo (uno o più file)", type=["xlsx"],
                                 accept_multiple_files=True, key="ss_file",
                                 help="Una riga per articolo e giorno. Più file si uniscono: per lo stesso articolo e giorno vale l'ultimo caricato.")
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
        if infos:
            st.markdown("**File salvati**")
            for i in infos:
                c1, c2 = st.columns([8, 1])
                c1.caption(f"{i['nome']} · caricato il {i['caricato']:%d/%m/%Y %H:%M} · {i['articoli']} articoli · "
                           f"dal {i['dal']:%d/%m/%Y} al {i['al']:%d/%m/%Y}" + (" · TRONCATO" if i["troncato"] else ""))
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


def riquadro_caricamento():
    """Riquadro «Carica o aggiorna i dati»: va in cima alla scheda, sopra la ricerca."""
    df, infos = _carica_tutti(_firma())
    _caricamento(df, infos, df.empty)


def render(filiale, item=None, trovati=None):
    """Storico giornaliero dello standard. item = articolo scelto nella ricerca in alto (se c'e', si mostra il suo andamento)."""
    st.markdown("#### Storico giornaliero dello standard cost")
    firma = _firma()
    df, infos = _carica_tutti(firma)
    if df.empty:
        st.info("Carica l'export dello standard cost giornaliero per vedere come è cambiato ogni articolo nel tempo.")
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
            st.markdown(f"**{len(trovati)} articoli trovati**: standard al primo e all'ultimo giorno del periodo")
            sty = conf.style.format({"Primo giorno": "{:%d/%m/%Y}", "Ultimo giorno": "{:%d/%m/%Y}", "Standard iniziale": "{:.4f} €",
                                     "Standard attuale": "{:.4f} €", "Δ": "{:+.4f} €", "Δ %": "{:+.1%}"}, na_rep="")
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
            _scheda_articolo(df, ev, item, anag.loc[item], d_da, d_a)
        else:
            st.caption("Questo articolo non è nell'export caricato (se l'export è stato troncato potrebbe mancare).")

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
            ch = alt.Chart(per_giorno).mark_bar(color="#0F766E").encode(
                x=alt.X("data:T", title=None, axis=alt.Axis(format="%d/%m")),
                y=alt.Y("Articoli cambiati:Q", title="Articoli con standard cambiato"),
                tooltip=[alt.Tooltip("data:T", title="Giorno", format="%d/%m/%Y"), "Articoli cambiati"]).properties(height=200)
            st.altair_chart(ch, width="stretch")



def _scheda_articolo(df, ev, item, riga, d_da=None, d_a=None):
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
            schede.append((nome, _it(f"{ultima['valore']:.4f} €"), f"al {ultima['data']:%d/%m/%Y}"))
        else:       # standard non piu' calcolato (zero) negli ultimi giorni
            valida = g.dropna(subset=["valore"])
            ultimo_ok = valida.iloc[-1]
            schede.append((nome, "non calcolato", f"ultimo valore {_it(f'{ultimo_ok['valore']:.4f}')} € il {ultimo_ok['data']:%d/%m/%Y}"))
    k = st.columns(min(len(schede), 4) + 1)
    for col, (nome, valore, sotto) in zip(k, schede):
        _kpi(col, nome, valore, sotto)
    _kpi(k[-1], "Cambi nel periodo", len(ec), (f"ultimo il {ec['data'].max():%d/%m/%Y}" if len(ec) else "standard sempre uguale"))
    st.write("")
    punti = serie.merge(ec.assign(serie_ev=ec["serie"])[["data", "serie"]], on=["data", "serie"], how="inner")
    base = alt.Chart(serie).encode(
        x=alt.X("data:T", title=None, axis=alt.Axis(format="%d/%m")),
        y=alt.Y("valore:Q", title="Standard cost (€)", scale=alt.Scale(zero=False)),
        color=alt.Color("serie:N", title=None, scale=alt.Scale(domain=list(serie["serie"].unique()),
                                                              range=[COLORI.get(s, "#667085") for s in serie["serie"].unique()]),
                        legend=alt.Legend(orient="bottom")))
    linea = base.mark_line(interpolate="step-after", strokeWidth=2.5)
    marche = alt.Chart(punti).mark_point(filled=True, size=70).encode(
        x="data:T", y="valore:Q", color=alt.Color("serie:N", legend=None),
        tooltip=[alt.Tooltip("data:T", title="Dal", format="%d/%m/%Y"), "serie:N", alt.Tooltip("valore:Q", title="Standard", format=".4f")])
    st.altair_chart((linea + marche).properties(height=300), width="stretch")
    if ec.empty:
        st.caption("Nessun cambio di standard nel periodo caricato.")
    else:
        st.markdown("**Quando è cambiato** (la data è il primo giorno con il nuovo valore)")
        _tab_eventi(ec.sort_values("data", ascending=False))
