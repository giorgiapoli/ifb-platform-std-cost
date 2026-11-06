"""
Tabella dei costi logistici considerati dal modello, letta dal foglio "COSTS (LOG)" (sola lettura).
Per ogni filiale una lista di blocchi: titolo, nomi delle colonne e righe (etichetta + celle da leggere).
Confronta la versione nuova con la precedente ed evidenzia le tariffe cambiate.
"""
import hashlib
import os
import pickle

import openpyxl
import pandas as pd
import streamlit as st
from openpyxl.utils.cell import coordinate_to_tuple

import spiega_variazioni as sv

AMBRA = "#FEF0C7"
SECCO_FRESCO_GELO = ["Secco", "Fresco", "Gelo"]
NCS = ["Nord", "Centro", "Sud"]

# (titolo, colonne, [(etichetta, [celle...]), ...]) - le celle sono quelle del foglio COSTS (LOG); None = vuoto
BLOCCHI = {
    "CANARIE": [
        ("Gomma · Verona → Barcellona", ["Valore"], [
            ("Costo tratta (€ per camion)", ["C3"]), ("Pallet per camion", ["F3"]),
            ("Pallet per container", ["G3"])]),
        ("Gomma · costo pallet Verona → Barcellona (€)", ["Secco", "Fresco", "Gelo"], [
            ("Costo pallet", ["D5", "E5", "F5"])]),
        ("Gomma · Barcellona → Canarie (€ per pallet)", ["Secco", "Fresco", "Gelo"], [
            ("Gran Canaria", ["D9", "E9", "F9"]), ("Tenerife", ["D10", "E10", "F10"]),
            ("Lanzarote", ["D11", "E11", "F11"]), ("Fuerteventura", ["D12", "E12", "F12"])]),
        ("Gomma · costi accessori", ["Valore"], [
            ("Fuel (su costo pallet)", ["F14"]), ("Assicurazione (su valore merce)", ["F15"]),
            ("Dogana export (€ a spedizione)", ["F16"])]),
        ("Mare · costo per container (€)", NCS, [
            ("Gran Canaria", ["D20", "E20", "F20"]), ("Tenerife", ["D21", "E21", "F21"]),
            ("Lanzarote", ["D22", "E22", "F22"]), ("Fuerteventura", ["D23", "E23", "F23"])]),
        ("Mare · inland + sdoganamento (€ per container)", ["Secco", "Fresco/Frozen"], [
            ("Gran Canaria", ["D27", "E27"]), ("Tenerife", ["D28", "E28"]),
            ("Lanzarote", ["D29", "E29"]), ("Fuerteventura", ["D30", "E30"])]),
        ("Tasse e dogana (€)", ["Valore"], [("REA", ["D35"]), ("Despacho de Aduanas", ["D36"])]),
        ("Unifreddo · MTS (deposito)", ["Secco", "Gelo", "Fresco"], [
            ("Deposito (€ pallet/quindicina)", ["E56", "F56", "G56"]), ("Inbound (€ pallet)", ["E57", "F57", "G57"]),
            ("Picking (€ collo)", ["E59", "F59", "G59"])]),
        ("Unifreddo · MTO (cross docking)", ["Secco", "Gelo", "Fresco"], [
            ("In + stock + out (€ pallet)", ["E62", "F62", "G62"]), ("Giorni inclusi", ["E63", "F63", "G63"]),
            ("Stock oltre periodo (€/giorno/pallet)", ["E64", "F64", "G64"])]),
        ("Altro", ["Valore"], [("Costo bancale (€)", ["I1"]), ("Assicurazione stock (quota)", ["E71"])]),
    ],
    "HONG KONG": [
        ("Parametri", ["Valore"], [
            ("Cambio EUR/HKD", ["C2"]), ("Intercompany", ["F2"]), ("Costo bancale (€)", ["I2"]),
            ("Carriage vino", ["C4"])]),
        ("Mare · FOB + nolo per container (€)", NCS, [
            ("DRY", ["D23", "E23", "F23"]), ("FRESH", ["D24", "E24", "F24"]), ("FROZEN", ["D25", "E25", "F25"])]),
        ("Mare · oneri di import (€)", NCS, [
            ("Local import charges", ["D29", "E29", "F29"]), ("Demurrage", ["D32", "E32", "F32"]),
            ("Health certificate", ["D33", "E33", "F33"])]),
        ("Kin Keung · inland e terminal", ["HKD", "EUR"], [
            ("Inland (con pallet)", ["C13", "D13"]), ("Terminal", ["C14", "D14"]), ("Sdoganamento", [None, "D15"])]),
        ("VGM (€)", ["Valore"], [("VGM", ["C35"])]),
        ("MTS (deposito)", ["Secco", "Gelo", "Fresco"], [
            ("Deposito (€ pallet/quindicina)", ["D43", "E43", "F43"]), ("Inbound (€ pallet)", ["D44", "E44", "F44"]),
            ("Picking (€ collo)", ["D46", "E46", "F46"])]),
        ("MTO (cross docking)", ["Secco", "Gelo", "Fresco"], [
            ("In + stock + out (€ pallet)", ["D49", "E49", "F49"]), ("Giorni inclusi", ["D50", "E50", "F50"]),
            ("Stock oltre periodo (€/giorno/pallet)", ["D51", "E51", "F51"])]),
        ("Etichettatura (€)", ["Secco", "Gelo", "Fresco"], [
            ("Lavorazione a collo", ["D54", "E54", "F54"]), ("Lavorazione a pezzo", ["D55", "E55", "F55"])]),
        ("Via aerea · tariffe (HKD)", ["Valore"], [
            ("Terminal (HKD/kg)", ["C68"]), ("Doc + handling (HKD/spedizione)", ["C69"]),
            ("Sdoganamento (HKD/pratica)", ["C70"]), ("Movimentazione (HKD/kg)", ["C71"]),
            ("Inland refrigerato (HKD/kg)", ["C72"]), ("Repacking (HKD/pallet)", ["C73"])]),
    ],
}


_CACHE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cache_std")


def _su_disco(nome, path, mtime, calcola):
    """Risultato salvato su disco (chiave = file, data del file, versione del motore): evita di riaprire Excel
    ad ogni avvio dell'app."""
    motore = os.path.getmtime(sv.__file__)
    chiave = hashlib.md5(f"{nome}|{path}|{mtime}|{motore}".encode("utf-8")).hexdigest()
    f = os.path.join(_CACHE, f"{nome}_{chiave}.pkl")
    if os.path.exists(f):
        try:
            with open(f, "rb") as fh:
                return pickle.load(fh)
        except Exception:
            pass
    out = calcola()
    try:
        os.makedirs(_CACHE, exist_ok=True)
        with open(f + ".tmp", "wb") as fh:
            pickle.dump(out, fh)
        os.replace(f + ".tmp", f)
        for v in sorted((x for x in os.listdir(_CACHE) if x.startswith(nome + "_")),
                        key=lambda x: os.path.getmtime(os.path.join(_CACHE, x)))[:-16]:
            os.remove(os.path.join(_CACHE, v))
    except Exception:
        pass
    return out


@st.cache_data(show_spinner=False)
def leggi_log(path, mtime):
    return _su_disco("log", path, mtime, lambda: _leggi_log(path))


def _leggi_log(path):
    """Celle del foglio COSTS (LOG): {(riga, colonna): valore}. Vuoto se il foglio non esiste."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    out = {}
    if "COSTS (LOG)" in wb.sheetnames:
        for r, row in enumerate(wb["COSTS (LOG)"].iter_rows(min_row=1, max_row=90, max_col=14, values_only=True), 1):
            for c, v in enumerate(row, 1):
                if v not in (None, ""):
                    out[(r, c)] = v
    wb.close()
    return out


def _val(cells, addr):
    if not addr:
        return None
    v = cells.get(coordinate_to_tuple(addr))
    return v.strip() if isinstance(v, str) else v


def _txt(v):
    if v is None:
        return ""
    if isinstance(v, (int, float)):
        s = f"{v:.4g}" if abs(v) < 1 else f"{v:,.2f}"
        if "." in s:
            s = s.rstrip("0").rstrip(".")
        return s.replace(",", "§").replace(".", ",").replace("§", ".")
    return str(v).replace("\xa0", " ").strip()


def _uguali(a, b):
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(a - b) < 1e-9
    return _txt(a) == _txt(b)


CATEGORIE = {"Gomma": "#EAB308", "Mare": "#06B6D4", "Via aerea": "#7C3AED", "Unifreddo · logistica MTS/MTO": "#DC2626",
             "Altro": "#98A2B3"}


def _categoria(titolo):
    t = titolo.lower()
    if t.startswith("gomma"):
        return "Gomma"
    if t.startswith("mare") or "kin keung" in t or t.startswith("vgm"):
        return "Mare"
    if "aerea" in t:
        return "Via aerea"
    if "mts" in t or "mto" in t or "unifreddo" in t:
        return "Unifreddo · logistica MTS/MTO"
    return "Altro"


def _titolo_breve(titolo):
    """Toglie il prefisso della categoria ('Gomma · Verona → Barcellona' -> 'Verona → Barcellona')."""
    for pre in ("Gomma · ", "Mare · ", "Unifreddo · ", "Via aerea · "):
        if titolo.startswith(pre):
            return titolo[len(pre):]
    return titolo


def render(filiale, p_old, p_new, mt_old, mt_new):
    blocchi = BLOCCHI.get(filiale, [])
    st.markdown("#### Costi logistici considerati")
    new = leggi_log(p_new, mt_new)
    old = leggi_log(p_old, mt_old)
    if not new or not blocchi:
        st.info("Non trovo il foglio COSTS (LOG) nel modello nuovo.")
        return
    usate = _su_disco("usate", p_new, mt_new, lambda: sv.celle_log_usate(p_new))      # solo le celle richiamate dalle formule del modello
    cambi, tabelle = [], []
    for titolo, colonne, righe in blocchi:
        dati, stili, etichette = [], [], []
        for etichetta, celle in righe:
            if not any(a and coordinate_to_tuple(a) in usate for a in celle):
                continue                    # voce non usata da nessuna formula: non la mostro
            riga, sriga = [], []
            for col, addr in zip(colonne, celle):
                if not addr or coordinate_to_tuple(addr) not in usate:
                    riga.append("–")        # cella di questa riga non usata dal modello
                    sriga.append("color: #98A2B3")
                    continue
                vn, vo = _val(new, addr), _val(old, addr)
                diverso = bool(old) and not _uguali(vn, vo)
                riga.append(_txt(vn))
                sriga.append(f"background-color: {AMBRA}; font-weight: 600" if diverso else "")
                if diverso:
                    cambi.append({"Voce": f"{titolo} · {etichetta}", "Colonna": col if len(colonne) > 1 else "",
                                  "Precedente": _txt(vo), "Nuovo": _txt(vn)})
            dati.append(riga + ["–"] * (len(colonne) - len(riga)))
            stili.append(sriga + ["color: #98A2B3"] * (len(colonne) - len(sriga)))
            etichette.append(etichetta)
        if not dati:
            continue
        df = pd.DataFrame(dati, index=etichette, columns=colonne)
        sty = df.style.apply(lambda _, s=pd.DataFrame(stili, index=df.index, columns=df.columns): s, axis=None)
        tabelle.append((titolo, sty))
    if not old:
        st.caption("Valori della versione nuova. (Nella versione precedente il foglio COSTS (LOG) non c'e': nessun confronto.)")
    elif cambi:
        st.warning(f"**{len(cambi)} tariffe cambiate** rispetto alla versione precedente (evidenziate in giallo nelle tabelle).")
        st.dataframe(pd.DataFrame(cambi), hide_index=True, width="stretch")
    else:
        st.success("**Nessuna tariffa cambiata** rispetto alla versione precedente.")
    # raggruppo i blocchi per categoria: Gomma, Mare, Via aerea, Unifreddo, altro
    gruppi = {}
    for titolo, sty in tabelle:
        gruppi.setdefault(_categoria(titolo), []).append((titolo, sty))
    pallino = {"Gomma": "🟡", "Mare": "🔵", "Via aerea": "🟣", "Unifreddo · logistica MTS/MTO": "🔴", "Altro": "⚪"}
    con_cambi = {_categoria(c["Voce"]) for c in cambi}
    for cat in [c for c in CATEGORIE if c in gruppi]:
        # chiuso di default; si apre da solo se dentro c'e' una tariffa cambiata
        n_cambi = sum(1 for c in cambi if _categoria(c["Voce"]) == cat)
        etichetta = f"{pallino.get(cat, '')} **{cat}**" + (f" · {n_cambi} cambiate" if n_cambi else "")
        with st.expander(etichetta, expanded=cat in con_cambi):
            cols = st.columns(2)
            for i, (titolo, sty) in enumerate(gruppi[cat]):
                with cols[i % 2]:
                    st.markdown(f"**{_titolo_breve(titolo)}**")
                    st.dataframe(sty, width="stretch")
