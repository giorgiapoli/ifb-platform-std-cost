"""
Scheda «File da caricare»: tutti i file che si caricano, in un posto solo, ognuno con lo stato e con l'indicazione di
DOVE si esporta. I file dei mesi di margini, lo storico giornaliero dello standard e i modelli restano separati e ben distinti.
"""
import glob
import os
import shutil
from datetime import datetime

import streamlit as st

import margini_fifo as mf
import storico_std

REPORT = "Standard Cost - Canarie"          # report Power BI da cui si esportano i dati
# dove si esporta ogni file: report Power BI e pagina del report
FONTI = {
    "storico": f"Report Power BI «{REPORT}» → pagina «Standard Cost» → esporta i dati (item statistics, una riga per articolo e giorno).",
    "non_ifb": f"Report Power BI «{REPORT}» → pagina «Standard Cost PROVEED…» → esporta i dati.",
    "margini": f"Report Power BI «{REPORT}».",
    "cost": "Report Power BI «COST Analisys - Cnr» → esporta i dati (una riga per entrata di magazzino).",
}

CSS = """<style>
.fc-fonte{background:#F9FAFB;border:1px solid #E4E7EC;border-radius:8px;padding:7px 11px;margin:2px 0 10px 0;
          font-size:.84rem;color:#344054}
.fc-fonte b{color:#14213D}
.fc-chip{display:inline-block;border-radius:999px;padding:2px 10px;font-size:.78rem;font-weight:600;margin-left:8px}
.fc-ok{background:#ECFDF3;color:#027A48}.fc-no{background:#FEF3F2;color:#B42318}.fc-man{background:#FFFAEB;color:#B54708}
.fc-tit{font-size:1.05rem;font-weight:700;color:#14213D}
</style>"""


def _titolo(num, testo, chip=None, classe="fc-ok"):
    c = f'<span class="fc-chip {classe}">{chip}</span>' if chip else ""
    st.markdown(f'<div class="fc-tit">{num} · {testo}{c}</div>', unsafe_allow_html=True)


def _file_mese(chiave, etichetta, aiuto, destra=False):
    """Uploader di uno dei 3 file del mese: il file resta in memoria anche quando si chiude questa pagina."""
    f = st.file_uploader(etichetta, type=["xlsx"], key=chiave, label_visibility="collapsed", help=aiuto)
    import assistente
    assistente.bolla(chiave, etichetta, aiuto, larghezza=340, destra=destra)
    import memoria_file
    if f is not None:
        st.session_state["dato_" + chiave] = f
        memoria_file.salva(chiave, f)                 # resta sul PC anche ricaricando la pagina o riaprendo l'app
    tenuto = st.session_state.get("dato_" + chiave) or memoria_file.carica(chiave)
    if tenuto is not None:
        c1, c2 = st.columns([3, 1])
        c1.caption(f"In uso: {tenuto.name}")
        if c2.button("Togli", key="togli_" + chiave):
            st.session_state.pop("dato_" + chiave, None)
            memoria_file.togli(chiave)
            st.rerun()


def _fonte(testo):
    st.markdown(f'<div class="fc-fonte"><b>Dove si esporta:</b> {testo}</div>', unsafe_allow_html=True)


def _eta(mtime):
    g = (datetime.now() - datetime.fromtimestamp(mtime)).days
    return "oggi" if g <= 0 else "ieri" if g == 1 else f"{g} giorni fa"


def _sostituisci(chiave, etichetta, target, tipi):
    """Caricamento facoltativo di un file della cartella INPUT: il file attuale non si perde, finisce in INPUT/_precedenti
    (con data e ora) e il nuovo prende il suo posto. target = percorso del file da sostituire (o da creare)."""
    with st.expander(f"Carica un nuovo file: {etichetta}"):
        f = st.file_uploader(etichetta, type=tipi, key=f"sost_{chiave}", label_visibility="collapsed")
        if f is not None:
            esiste = os.path.exists(target)
            st.caption((f"Sostituisce «{os.path.basename(target)}»: quello attuale viene tenuto nella cartella _precedenti."
                        if esiste else f"Viene salvato come «{os.path.basename(target)}» nella cartella INPUT."))
            if st.button("Salva nella cartella INPUT", type="primary", key=f"sost_ok_{chiave}"):
                try:
                    if esiste:
                        prec = os.path.join(os.path.dirname(target), "_precedenti")
                        os.makedirs(prec, exist_ok=True)
                        nome, est = os.path.splitext(os.path.basename(target))
                        shutil.copy2(target, os.path.join(prec, f"{nome}_{datetime.now():%Y%m%d_%H%M}{est}"))
                    with open(target, "wb") as fh:
                        fh.write(f.getbuffer())
                    st.session_state["sost_esito"] = f"«{os.path.basename(target)}» aggiornato."
                    st.rerun()
                except Exception as ex:
                    st.error(f"Non riesco a salvare il file: {ex}")


def pannello_filiali(sv, DATA_ROOT, filiale):
    """Pannello di controllo delle filiali: aggiungere una filiale (con il modello fatto come Canarie o Hong Kong) o toglierne una."""
    with st.expander("⚙ Filiali · pannello di controllo"):          # chiuso di partenza: lo apre l'utente
        st.caption("Aggiungi una filiale con il modello Excel fatto come quello di Canarie o di Hong Kong. "
                   "Se il modello ha una struttura diversa, serve prima adattare l'app.")
        _nome_f = st.text_input("Nome della filiale", key="nf_nome", placeholder="es. Madeira")
        _tipo_f = st.radio("Il modello Excel è fatto come quello di", ["Canarie", "Hong Kong"], horizontal=True, key="nf_tipo")
        _cart_f = st.text_input("Cartella della filiale", key="nf_cartella",
                                help="Percorso completo della cartella (o solo il nome, se sta accanto alle altre filiali).")
        _mod_f = st.text_input("Modello del mese in corso", key="nf_modello", placeholder="INPUT/05_Modello_Standard_Cost.xlsm",
                               help="Percorso del file, relativo alla cartella della filiale (o completo).")
        _sto_f = st.text_input("Modelli dei mesi chiusi (facoltativo)", key="nf_storico", placeholder="STORICO MESI/**/*Modello*.xls*",
                               help="Schema dei file dei mesi passati, relativo alla cartella della filiale. Lascia vuoto se non ci sono.")
        if st.button("Aggiungi", key="nf_ok", type="primary", disabled=not (_nome_f.strip() and _cart_f.strip() and _mod_f.strip())):
            _chiave = _nome_f.strip().upper()
            _base = _cart_f.strip() if os.path.isabs(_cart_f.strip()) else os.path.join(DATA_ROOT, _cart_f.strip())
            _file = _mod_f.strip() if os.path.isabs(_mod_f.strip()) else os.path.join(_base, _mod_f.strip())
            if _chiave in sv.BRANCHES:
                st.error("Esiste già una filiale con questo nome.")
            elif not os.path.isdir(_base):
                st.error(f"Non trovo la cartella: {_base}")
            elif not os.path.exists(_file):
                st.error(f"Non trovo il file del modello: {_file}")
            else:
                sv.salva_filiale_utente(_chiave, _nome_f.strip(), "CANARIE" if _tipo_f == "Canarie" else "HONG KONG",
                                        _cart_f.strip(), _mod_f.strip(), _sto_f.strip() or "__nessuno__/*.xlsm")
                st.session_state["filiale_sel"] = _nome_f.strip()
                st.rerun()
        if sv.BRANCHES.get(filiale, {}).get("utente"):
            if st.button(f"Rimuovi la filiale «{sv.BRANCHES[filiale]['label']}»", key="nf_rimuovi"):
                sv.rimuovi_filiale_utente(filiale)
                st.session_state["filiale_sel"] = list(sv.BRANCHES)[0]
                st.rerun()
            st.caption("Rimuove solo la voce dall'elenco: i file della filiale non vengono toccati.")


def render(filiale, cartella_input, modelli, p_old, p_new, sorgente, sv=None, data_root=None):
    st.markdown(CSS, unsafe_allow_html=True)
    st.markdown("#### File da caricare")
    st.caption("Tutti i file che servono all'app, divisi per tipo. I file non escono dal PC.")
    if st.session_state.get("sost_esito"):
        st.success(st.session_state.pop("sost_esito"))
    if sv is not None:
        pannello_filiali(sv, data_root, filiale)

    sfx = "" if filiale == "CANARIE" else "_" + filiale.replace(" ", "_")          # file dei margini separati per filiale
    hk = filiale != "CANARIE"
    if True:          # storico a sistema e margini: per ogni filiale (ognuna ha i suoi file)
        # ---- 1. storico standard cost in macchina -----------------------------------------------------
        with st.container(border=True):
            df, infos = storico_std._carica_tutti(storico_std._firma())
            if df.empty:
                _titolo("1", "Standard cost in macchina, storico giornaliero · " + sv.BRANCHES[filiale]["label"], "da caricare", "fc-no")
            else:
                _titolo("1", "Standard cost in macchina, storico giornaliero · " + sv.BRANCHES[filiale]["label"],
                        f"caricato fino al {df['data'].max():%d/%m/%Y}")
            _fonte(FONTI["storico"] if not hk else
               f"Report Power BI «Standard Cost - {filiale.title()}» → esporta i dati (item statistics: executiondate, mastercode, "
               "Item No, ifbitem, Description, transportation, Section description, Standard cost).")
            storico_std.corpo_caricamento(infos, df)

        # ---- 2. margini del mese -----------------------------------------------------------------
        with st.container(border=True):
            salvati = mf.mesi_salvati()
            _titolo("2", "Margini del mese" + (" COMIT (3 file)" if not hk else " (2 file)"),
                    ("in memoria: " + ", ".join(mf.etichetta_mese(k) for k in salvati)) if salvati else "nessun mese in memoria",
                    "fc-ok" if salvati else "fc-man")
            st.caption("Si caricano insieme, una volta al mese. Poi in «Margini vs FIFO» si dà il nome al mese con «Tieni in memoria».")
            c1, c2, c3 = st.columns(3) if not hk else (*st.columns(2), None)
            with c1:
                st.markdown("**Margini STD vs FIFO** (obbligatorio)")
                _fonte(FONTI["margini"] if not hk else "Stesso formato del report dei margini di Canarie.")
                _file_mese("mf_marg" + sfx, "Margini STD vs FIFO", "Export 'Difference margin' per articolo, con riga Total.")
            with c2:
                st.markdown("**Cost Analysis** (entrate di magazzino)")
                _fonte(FONTI["cost"] if not hk else "Stesso formato del Cost Analysis di Canarie (una riga per entrata di magazzino).")
                _file_mese("mf_cost" + sfx, "Cost Analysis", "Una riga per entrata: serve per capire perché lo standard e il FIFO si allontanano.", destra=hk)
            if not hk:          # il report dei prodotti non IFB c'e' solo per Canarie
                with c3:
                    st.markdown("**Standard cost prodotti NON IFB**")
                    _fonte(FONTI["non_ifb"])
                    _file_mese("mf_non" + sfx, "Standard cost prodotti NON IFB", "Prodotti acquistati direttamente da COMIT: serve per controllare anche i loro margini.", destra=True)

    # ---- 3. modelli standard cost (Excel) -----------------------------------------------------------
    with st.container(border=True):
        if p_new and os.path.exists(p_new):
            _titolo("3", "Modelli standard cost (Excel)", f"in uso: {os.path.basename(p_new)}", "fc-ok")
            st.caption(f"Salvato il {datetime.fromtimestamp(os.path.getmtime(p_new)):%d/%m/%Y alle %H:%M} · "
                       f"{len(modelli)} versioni trovate nelle cartelle.")
        else:
            _titolo("3", "Modelli standard cost (Excel)", "non trovati", "fc-no")
        st.markdown('<div class="fc-fonte"><b>Dove sono:</b> il modello del mese in corso è nella cartella <b>INPUT</b>, i mesi chiusi '
                    'in <b>STORICO MESI</b>. Per usare un file diverso: barra a sinistra → «Sorgente dei file» → «Carica i file».</div>',
                    unsafe_allow_html=True)
        _mod = modelli.get("Mese in corso · INPUT") or os.path.join(cartella_input, "05_Modello_Standard_Cost" + ("_COMIT" if not hk else "") + ".xlsm")
        _sostituisci("modello", "modello del mese in corso", _mod, ["xlsm", "xlsx"])

    # ---- 4. file nella cartella INPUT (solo sul PC: nella versione web non c'e' nessuna cartella) --------
    import sys
    if sys.platform == "emscripten":
        return
    with st.container(border=True):
        _titolo("4", "File nella cartella INPUT")
        st.caption("Quelli automatici li aggiorna Power Automate ogni giorno: caricarli qui serve solo se vuoi forzare un aggiornamento.")
        st.caption("Anagrafica e Pricelist del mese sono già dentro i modelli Excel: non serve caricarli.")
        voci = [("02_Ultime_Fatture.xlsx", "automatico, ogni giorno"), ("03_Nuovi_Item.xlsx", "automatico, ogni giorno")]
        for nome, come in voci:
            trovati = sorted(glob.glob(os.path.join(cartella_input, nome)), key=os.path.getmtime)
            if hk and not trovati and ("Anagrafica" in nome or "Pricelist" in nome):
                continue                  # file che Hong Kong non ha: non lo segno come mancante
            if trovati:
                p = trovati[-1]
                vecchio = (datetime.now() - datetime.fromtimestamp(os.path.getmtime(p))).days > (3 if "automatico" in come else 35)
                chip = f'<span class="fc-chip {"fc-man" if vecchio else "fc-ok"}">aggiornato {_eta(os.path.getmtime(p))}</span>'
                st.markdown(f"**{os.path.basename(p)}** · {come}{chip}", unsafe_allow_html=True)
            else:
                st.markdown(f'**{nome.replace("*", "")}** · {come}<span class="fc-chip fc-no">non trovato</span>', unsafe_allow_html=True)
            _sostituisci(nome.split(".")[0].split("*")[0].strip("_ "), nome.replace("*", "").replace(".xlsx", ""),
                         trovati[-1] if trovati else os.path.join(cartella_input, nome.replace("*", "")), ["xlsx"])
