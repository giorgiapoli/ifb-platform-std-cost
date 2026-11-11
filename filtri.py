"""
Filtri per colonna su una tabella: un pulsante «Filtra» sopra la tabella, si scelgono le colonne e i valori che si vogliono
vedere (elenco di valori per le colonne di testo, intervallo per quelle numeriche). Si applica ai dati PRIMA di mostrarli,
quindi anche il pulsante «Excel» scarica solo le righe filtrate.
"""
import pandas as pd
import streamlit as st


def filtra(df, chiave, escludi=()):
    """Mostra il pulsante «Filtra» e ritorna il DataFrame con le righe che rispettano i filtri scelti."""
    if df is None or len(df) == 0:
        return df
    colonne = [c for c in df.columns if c not in escludi]
    kc = f"flt_{chiave}"
    scelte = st.session_state.get(f"{kc}_cols", [])
    n_att = len([c for c in scelte if c in colonne])
    with st.popover(f"Filtra{f' ({n_att})' if n_att else ''}", icon=":material/filter_alt:"):
        scelte = st.multiselect("Colonne da filtrare", colonne, key=f"{kc}_cols",
                                placeholder="Scegli una o più colonne")
        maschera = pd.Series(True, index=df.index)
        for c in scelte:
            s = df[c]
            if pd.api.types.is_numeric_dtype(s) and not pd.api.types.is_bool_dtype(s):
                v = pd.to_numeric(s, errors="coerce")
                if v.notna().any() and v.min() < v.max():
                    lo, hi = float(v.min()), float(v.max())
                    a, b = st.slider(str(c), lo, hi, (lo, hi), step=(hi - lo) / 200, key=f"{kc}_n_{c}")
                    if (a, b) != (lo, hi):
                        maschera &= v.between(a, b)
            else:
                valori = sorted(s.dropna().astype(str).unique())
                if len(valori) > 1500:
                    testo = st.text_input(str(c), key=f"{kc}_t_{c}", placeholder="contiene…")
                    if testo:
                        maschera &= s.astype(str).str.lower().str.contains(testo.lower(), regex=False)
                else:
                    sel = st.multiselect(str(c), valori, key=f"{kc}_v_{c}", placeholder="tutti i valori")
                    if sel:
                        maschera &= s.astype(str).isin(sel)
        if scelte:
            st.caption(f"{int(maschera.sum())} righe su {len(df)}")
    return df[maschera]
