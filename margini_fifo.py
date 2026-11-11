"""
Controllo margini COMIT: margine a STANDARD COST contro margine a FIFO, mese per mese.

Cosa fa (tutto in sola lettura, niente esce dal PC):
  1. Esito del mese: margine STD, margine FIFO e scostamento in punti di vendite. Fascia "buona" = tra -3% e -0,5%.
     Sopra -0,5% il costo reale e' troppo vicino (o oltre) lo standard: si sta andando male.
  2. Outlier: gli articoli che si allontanano dallo standard (per scostamento % e per euro).
  3. Perche': per ogni articolo confronta le entrate di magazzino (Cost Analysis) con le componenti dello standard
     (listino, logistica, AIEM) e segnala cio' che e' registrato in modo ambiguo:
       - AIEM registrato su articoli che in anagrafica non lo hanno (o non registrato dove c'e');
       - container ripartiti male: il container si isola con il "External Document No_" (lo stesso numero fattura IFB),
         e si controlla se il trasporto e' ripartito per volume corretto (volume collo / pezzi per collo)
         oppure per "volume del collo x pezzi" (errore gia' segnalato a COMIT);
       - imposte ripartite con la stessa chiave del trasporto;
       - trasporto che oscilla molto da un'entrata all'altra o che si scosta dallo standard.
Convenzione: Differenza margine = Margine STD - Margine FIFO.  Negativa = il FIFO ha un margine migliore dello STD
(standard prudente), positiva = il costo reale e' superiore allo standard (margine reale peggiore).
"""
import glob
import io
import json
import os
import re

import numpy as np
import pandas as pd

import spiega_variazioni as sv

HERE = os.path.dirname(os.path.abspath(__file__))
import percorsi
STORICO = os.path.join(percorsi.DATI, "margini_storico.json")

BANDA_MIN, BANDA_MAX = -0.03, -0.005        # fascia di scostamento considerata buona (in % delle vendite)
ISOLA_DI = {"GRACAN": "GC-TF", "TENERIFE": "GC-TF", "LANZAR": "LAN-FUE", "FUERTE": "LAN-FUE", "LAPALM": "GC-TF"}

COLONNE_MARGINI = {
    "Item No_": "item", "Description": "desc", "Quantity": "qty", "UOM": "uom", "Price": "prezzo",
    "Cost unit Standard": "cstd", "Cost unit FIFO": "cfifo", "Sales Amount": "vendite",
    "Cost Amount Standard": "costo_std", "Cost Amount FIFO": "costo_fifo", "Margin Standard": "m_std",
    "Margin FIFO": "m_fifo", "Margin Standard %": "p_std", "Margin FIFO %": "p_fifo",
    "Difference Margin": "diff", "Difference Margin %": "diff_pct"}
COLONNE_ENTRATE = ["item", "desc", "data", "ubic", "ext", "doc", "lotto", "qty", "pend", "uom", "altro", "abono",
                   "desp", "imp", "prezzo", "trasp", "tot"]


# ---- lettura dei file ---------------------------------------------------------------------------
def _num(s):
    return pd.to_numeric(s, errors="coerce").replace([np.inf, -np.inf], np.nan)


def leggi_margini(data):
    """Report margini per articolo (export 'Difference STD FIFO'). Ritorna (articoli, totale)."""
    df = pd.read_excel(io.BytesIO(data))
    # le intestazioni cambiano un po' da un report all'altro (Canarie: «Cost unit Standard», Hong Kong: «Unit cost std»): si
    # riconoscono ignorando maiuscole, spazi e simboli
    sin = {"itemno": "item", "no": "item", "description": "desc", "quantity": "qty", "uom": "uom", "price": "prezzo",
           "costunitstandard": "cstd", "unitcoststd": "cstd", "unitcoststandard": "cstd",
           "costunitfifo": "cfifo", "unitcostfifo": "cfifo", "salesamount": "vendite", "salesamountactual": "vendite",
           "costamountstandard": "costo_std", "costamountfifo": "costo_fifo", "costamountactual": "costo_fifo",
           "marginstandard": "m_std", "marginfifo": "m_fifo", "margin": "m_fifo", "marginstandardpct": "p_std",
           "marginfifopct": "p_fifo", "marginpct": "p_fifo", "differencemargin": "diff", "differencemarginpct": "diff_pct",
           "ifbitem": "ifb", "sectiondescription": "sez"}
    _n = lambda c: re.sub(r"[^a-z0-9]", "", str(c).lower().replace("%", "pct"))
    df = df.rename(columns={c: sin.get(_n(c), COLONNE_MARGINI.get(str(c).strip(), c)) for c in df.columns})
    mancanti = [v for v in ("item", "vendite", "cstd", "cfifo", "m_std", "m_fifo", "diff") if v not in df.columns]
    if mancanti:
        raise ValueError("Nel report dei margini mancano colonne: " + ", ".join(mancanti))
    df["item"] = df["item"].astype(str).str.strip()
    for c in ("qty", "prezzo", "cstd", "cfifo", "vendite", "costo_std", "costo_fifo", "m_std", "m_fifo", "diff"):
        if c in df.columns:
            df[c] = _num(df[c])
    # nota dei filtri applicati all'export (in fondo al file): lo scostamento vale solo per gli articoli che il filtro lascia passare
    _nota = df.loc[df["item"].str.lower().str.startswith("filtri"), "item"]
    filtri = ""
    if len(_nota):
        _tr = next((r for r in str(_nota.iloc[0]).splitlines() if r.lower().startswith("transportation")), "")
        _tag = re.findall(r"(\+?\d*\S?C|CH|DRY|FR)\s+(SEA|AIR|ALTRO)", _tr)
        filtri = ("Trasporto: " + ", ".join(sorted({t[1] for t in _tag}))) if _tag else             re.sub(r"\s+", " ", _tr.replace("�", " ")).strip()
    riga_tot = df[df["item"].str.lower() == "total"]
    art = df[~df["item"].str.lower().isin(["total"]) & ~df["item"].str.lower().str.startswith("filtri")].copy()
    art = art[art["vendite"].notna() & (art["vendite"] != 0)]
    if len(riga_tot):
        t = riga_tot.iloc[0]
        tot = dict(vendite=float(t["vendite"]), m_std=float(t["m_std"]), m_fifo=float(t["m_fifo"]))
    else:
        tot = dict(vendite=float(art["vendite"].sum()), m_std=float(art["m_std"].sum()), m_fifo=float(art["m_fifo"].sum()))
    tot["filtri"] = filtri
    tot["diff"] = tot["m_std"] - tot["m_fifo"]
    tot["p_std"] = tot["m_std"] / tot["vendite"]
    tot["p_fifo"] = tot["m_fifo"] / tot["vendite"]
    tot["diff_pct"] = tot["diff"] / tot["vendite"]
    # costi unitari positivi (nel report sono negativi)
    art["c_std"] = art["cstd"].abs()
    art["c_fifo"] = art["cfifo"].abs()
    art["diff_pct"] = art["diff"] / art["vendite"]
    return art.reset_index(drop=True), tot


def leggi_entrate_hk(data):
    """Cost Analysis di Hong Kong: una riga per entrata, con i costi unitari scomposti (FOB, certificato sanitario, altri oneri,
    dogana, pallet, prezzo di acquisto, totale). Non c'e' il numero di spedizione ne' l'AIEM."""
    raw = pd.read_excel(io.BytesIO(data), header=None)
    if raw.shape[1] < 15:
        raise ValueError("Il Cost Analysis di Hong Kong ha meno colonne del previsto (servono 15).")
    df = raw.iloc[2:, :15].copy()
    df.columns = ["item", "desc", "data", "doc", "lotto", "qty", "uom", "x", "fob", "salute", "altro", "dogana", "pallet",
                  "prezzo", "tot"]
    df = df[df["item"].notna()].copy()
    df["item"] = df["item"].astype(str).str.strip()
    df["data"] = pd.to_datetime(df["data"], errors="coerce")
    for c in ("qty", "fob", "salute", "altro", "dogana", "pallet", "prezzo", "tot"):
        df[c] = _num(df[c]).fillna(0.0)
    return df[df["data"].notna() & (df["qty"] > 0)].drop(columns=["x"]).reset_index(drop=True)


def leggi_listino(path):
    """Listino di acquisto del mese dal foglio «CURRENT PRICELIST (ACQ)» del modello: {codice: (prezzo, UM)}.
    Prezzo = FCA scontato se c'e', altrimenti DAP scontato, altrimenti prezzo MTS (per UM del listino)."""
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    nome = next((n for n in wb.sheetnames if n.strip().upper().startswith("CURRENT PRICELIST")), None)
    out = {}
    if nome:
        it = wb[nome].iter_rows(min_row=1, max_col=20, values_only=True)
        hdr = None
        for r in it:
            if hdr is None:
                if r and str(r[0]).strip().lower() == "no_":
                    hdr = [re.sub(r"[^a-z0-9]", "", str(c).lower()) for c in r]
                continue
            if r[0] in (None, ""):
                continue
            g = lambda n: r[hdr.index(n)] if n in hdr and hdr.index(n) < len(r) else None
            num = lambda v: float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else 0.0
            prezzo = next((p for p in (num(g("fcadiscounted")), num(g("dapdiscounted")), num(g("mtsprice"))) if p > 0), 0.0)
            out[str(r[0]).strip()] = (prezzo, str(g("unitofmeasure") or "").strip().upper())
    wb.close()
    return out


def peso_riga(item, qty, uom, ana):
    """Peso (kg) di una riga: se l'UM e' KG la quantita' stessa; per pezzi e colli, quantita' x peso netto del pezzo (x pezzi per collo)."""
    a = ana.get(item) or {}
    um = (a.get("_um") or str(uom)).upper()
    if um == "KG":
        return float(qty)
    pw = a.get("_pcsw")
    if not pw:
        return np.nan
    return float(qty) * pw * ((a.get("Quantity x Packaging") or 1.0) if um == "BOX" else 1.0)


def classifica_peso_hk(w, ana, tol=0.10):
    """Hong Kong: le spedizioni via aerea ripartiscono il costo a PESO (kg), non a volume. Per ogni spedizione confronta la
    ripartizione registrata con quella a peso. Ritorna per spedizione: e_kg (distanza dalla ripartizione a peso), n (righe con peso),
    ibridi (articoli etichettati via mare, «SEA», in una spedizione ripartita a peso: viaggiano in aereo)."""
    righe = []
    for ext, g in w.groupby("ext"):
        g = g.assign(reg=g["trasp"] * g["qty"])
        g = g.assign(peso=[peso_riga(r.item, r.qty, r.uom, ana) for r in g.itertuples()])
        a = g.groupby("item").agg(reg=("reg", "sum"), peso=("peso", "sum"))
        a = a[(a["peso"] > 0)]
        if len(a) < 3 or a["reg"].sum() <= 0:
            continue
        e_kg = float(0.5 * (a["reg"] / a["reg"].sum() - a["peso"] / a["peso"].sum()).abs().sum())
        ib = [i for i in a.index if "SEA" in (ana.get(i, {}).get("_trasp") or "")]
        righe.append(dict(ext=str(ext), e_kg=e_kg, n_peso=len(a), a_peso=e_kg <= tol, ibridi=ib))
    return pd.DataFrame(righe, columns=["ext", "e_kg", "n_peso", "a_peso", "ibridi"])


def adatta_entrate_hk(ent):
    """Entrate di Hong Kong nello stesso formato di quelle di Canarie, cosi' i controlli sulle spedizioni valgono anche per loro:
    spedizione = documento di acquisto, trasporto = FOB, altri oneri (certificato, dogana, pallet...) come «despachos», niente AIEM."""
    out = pd.DataFrame({
        "item": ent["item"], "desc": ent["desc"], "data": ent["data"], "ubic": "HONG KONG", "ext": ent["doc"].astype(str).str.strip(),
        "doc": ent["doc"].astype(str), "lotto": ent["lotto"], "qty": ent["qty"], "pend": 0.0, "uom": ent["uom"], "altro": 0.0,
        "abono": 0.0, "desp": ent["salute"] + ent["altro"] + ent["dogana"] + ent["pallet"], "imp": 0.0, "prezzo": ent["prezzo"],
        "trasp": ent["fob"], "tot": ent["tot"]})
    out["isola"] = "HK"
    out["c_salute"], out["c_altro"], out["c_dogana"], out["c_pallet"] = ent["salute"], ent["altro"], ent["dogana"], ent["pallet"]
    return out.reset_index(drop=True)


def costi_entrata_hk(ent, mese):
    """Per articolo, nel mese (AAAA-MM; se non valido tutte le entrate): costo medio di entrata per UM, ponderato sulle quantita'."""
    e = ent
    if re.fullmatch(r"\d{4}-\d{2}", str(mese).strip()):
        e = ent[ent["data"].dt.strftime("%Y-%m") == mese.strip()]
    if e.empty:
        return e
    cols = ["fob", "salute", "altro", "dogana", "pallet", "prezzo", "tot"]
    pesata = e[cols].mul(e["qty"], axis=0).groupby(e["item"]).sum().div(e.groupby("item")["qty"].sum(), axis=0)
    out = pesata.join(e.groupby("item").agg(desc=("desc", "first"), uom=("uom", "first"), entrate=("qty", "count"),
                                            qty=("qty", "sum"), ultima=("data", "max")))
    return out.reset_index()


def leggi_entrate(data):
    """Cost Analysis: una riga per entrata di magazzino, con il costo unitario scomposto."""
    df = pd.read_excel(io.BytesIO(data), header=1)
    if df.shape[1] < len(COLONNE_ENTRATE):
        raise ValueError("Il file Cost Analysis ha meno colonne del previsto (servono 17).")
    df = df.iloc[:, :len(COLONNE_ENTRATE)]
    df.columns = COLONNE_ENTRATE
    df["item"] = df["item"].astype(str).str.strip()
    df["ext"] = df["ext"].astype(str).str.strip()
    df["data"] = pd.to_datetime(df["data"], errors="coerce")
    for c in ("qty", "pend", "altro", "abono", "desp", "imp", "prezzo", "trasp", "tot"):
        df[c] = _num(df[c]).fillna(0.0)
    df["isola"] = df["ubic"].map(ISOLA_DI)
    return df[df["data"].notna() & (df["qty"] > 0)].reset_index(drop=True)


def leggi_non_ifb(data):
    """Report 'standard cost prodotti non IFB' (acquistati direttamente da COMIT): componenti dello standard."""
    df = pd.read_excel(io.BytesIO(data))
    df.columns = [re.sub(r"\s+", " ", str(c)).strip() for c in df.columns]
    pick = lambda *nomi: next((c for c in df.columns if c.lower() in [n.lower() for n in nomi]), None)
    c_item, c_aiem = pick("No_"), pick("AIEM")
    c_pr, c_log = pick("Prezzo Netto per Sales UOM", "Prezzo Netto"), pick("Costo Logistico per UOM")
    c_ca, c_alc = pick("Costo AIEM"), pick("Costo Imposta Alcol Totale")
    c_std, c_forn = pick("NEW Standard Cost"), pick("Vendor Name")
    c_kg, c_plt, c_div = pick("Kg per Box"), pick("Quantity x PLT"), pick("Divisore_Collo")
    out = {}
    z = lambda r, c: float(r[c]) if c and pd.notna(r[c]) and np.isfinite(float(r[c])) else 0.0
    for _, r in df.iterrows():
        rate = z(r, c_aiem)
        out[str(r[c_item]).strip()] = dict(
            rate=rate if 0 <= rate <= 100 else -1.0,   # aliquota fuori scala (dato sporco): non verificabile
            prezzo=z(r, c_pr), logistica=z(r, c_log),
            aiem=z(r, c_ca) + z(r, c_alc), std=z(r, c_std), fornitore=str(r[c_forn]) if c_forn else "",
            kgbox=z(r, c_kg), qplt=z(r, c_plt), div=z(r, c_div))
    return out


def leggi_std_ifb(path):
    """Componenti dello standard degli articoli IFB dal modello Canarie (per coppia di isole)."""
    sv.set_branch("CANARIE")
    out = {}
    for sh in sv.load(path).values():
        for k, row in sh.items.items():
            try:
                d = sv.drivers(sh, row)
                rec = dict(rate=d["aiem_rate"] or 0.0, fornitore=d["fornitore"], um=d["um"][0])
                for isl in sv.ISLANDS:
                    c = sv.components(sh, row, isl)
                    s2 = c["s2"] or 0.0
                    price = c["price"] or 0.0
                    aiem = c["parts"].get("AIEM", 0.0)
                    rec[isl] = dict(prezzo=price, aiem=aiem, logistica=max(s2 - price - aiem, 0.0) if s2 else 0.0,
                                    std=(s2 * c["f"]) if s2 else 0.0)
                g = lambda n: sh.n(row, n)
                h, l, p = g("boxheightmm"), g("boxlengthmm"), g("boxdepthmm")
                rec["vbox"] = (h * l * p / 1e9) if h and l and p else None
                rec["qpk"], rec["kgbox"], rec["qplt"] = g("quantityxpackaging"), g("kgperbox"), g("quantityxplt")
                rec["pkplt"], rec["h"], rec["l"], rec["p"] = g("packagingxpallet"), h, l, p
                out[k] = rec
            except Exception:
                continue
    return out


def cerca_anagrafica():
    """Ultimo file 'Anagrafica ...xlsx' nella cartella INPUT (scaricato a mano ogni mese)."""
    trovati = glob.glob(os.path.join(HERE, "INPUT", "Anagrafica*.xlsx"))
    return max(trovati, key=os.path.getmtime) if trovati else None


CAMPI_ANAGRAFICA = [("Quantity x Packaging", "qpk", "pezzi per box"), ("Packaging x Pallet", "pkplt", "box per pallet"),
                    ("Box Height mm", "h", "altezza box (mm)"), ("Box Length mm", "l", "lunghezza box (mm)"),
                    ("Box Depth mm", "p", "profondita' box (mm)")]


def leggi_anagrafica(path):
    """Dati di imballo dall'anagrafica articoli: {codice: {campo: valore}}. Se path e' il modello Excel, si legge il suo foglio
    «Anagrafica Articoli» (l'anagrafica del mese e' gia' dentro il modello: nessun altro file da cercare)."""
    if str(path).lower().endswith((".xlsm", ".xlsx")):
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        nomi = wb.sheetnames
        wb.close()
        foglio = next((n for n in nomi if n.strip().lower() == "anagrafica articoli"), None)
        df = pd.read_excel(path, sheet_name=foglio) if foglio else pd.read_excel(path)
        if "No_" not in df.columns:          # modello di Hong Kong: intestazioni senza spazi e magari una riga di titolo sopra
            grezzo = pd.read_excel(path, sheet_name=foglio, header=None)
            _n = lambda c: re.sub(r"[^a-z0-9]", "", str(c).lower())
            r0 = next((i for i in range(min(6, len(grezzo))) if _n(grezzo.iloc[i, 0]) == "no"), 0)
            nomi = {"no": "No_", "quantityxpackaging": "Quantity x Packaging", "packagingxpallet": "Packaging x Pallet",
                    "boxheightmm": "Box Height mm", "boxlengthmm": "Box Length mm", "boxdepthmm": "Box Depth mm",
                    "salesunitofmeasure": "Sales Unit of Measure", "netweight": "PCS Net Weight", "pcsnetweight": "PCS Net Weight",
                    "aiem": "AIEM _", "transportation": "Transportation"}
            df = grezzo.iloc[r0 + 1:].copy()
            df.columns = [nomi.get(_n(c), str(c)) for c in grezzo.iloc[r0]]
            df = df.loc[:, ~df.columns.duplicated()]
    else:
        df = pd.read_excel(path)
    df.columns = [re.sub(r"\s+", " ", str(c)).strip() for c in df.columns]
    out = {}
    for _, r in df.iterrows():
        k = str(r["No_"]).strip()
        out[k] = {c: (float(r[c]) if c in df.columns and pd.notna(r[c]) else None) for c, _, _ in CAMPI_ANAGRAFICA}
        out[k]["_um"] = str(r["Sales Unit of Measure"]).strip().upper() if "Sales Unit of Measure" in df.columns else ""
        out[k]["_pcsw"] = float(r["PCS Net Weight"]) if "PCS Net Weight" in df.columns and pd.notna(r["PCS Net Weight"]) else None
        out[k]["_aiem"] = float(r["AIEM _"]) if "AIEM _" in df.columns and pd.notna(r["AIEM _"]) else None
        out[k]["_trasp"] = str(r["Transportation"]).strip().upper() if "Transportation" in df.columns and pd.notna(r["Transportation"]) else ""
    return out


def dati_articoli(std_ifb, std_non, ana):
    """Dati di imballo e aliquota AIEM di tutti gli articoli (IFB e non IFB), per verificare il volume delle spedizioni.
    Fonti: modello Canarie per gli IFB; anagrafica COMIT (tutto cio' che c'e' in NAV) per gli altri e per gli IFB senza misure.
    Per i prodotti a peso i kg per box sono pezzi per box x peso netto del pezzo."""
    out = {}
    for k, r in std_ifb.items():
        out[k] = dict(um=r.get("um"), qpk=r.get("qpk"), kgbox=r.get("kgbox"), vbox=r.get("vbox"), rate=r["rate"], fonte="modello")
    for k, a in ana.items():
        h, l, p = a.get("Box Height mm"), a.get("Box Length mm"), a.get("Box Depth mm")
        vb = h * l * p / 1e9 if h and l and p else None
        qpk, pw, um = a.get("Quantity x Packaging"), a.get("_pcsw"), a.get("_um")
        kgbox = qpk * pw if um == "KG" and qpk and pw else None
        if k in out:
            if not out[k].get("vbox") and vb:          # il modello non ha le misure: uso l'anagrafica
                out[k].update(vbox=vb, qpk=qpk or out[k].get("qpk"), kgbox=kgbox or out[k].get("kgbox"), fonte="anagrafica")
            continue
        rate = std_non[k]["rate"] if k in std_non and std_non[k]["rate"] >= 0 else a.get("_aiem")
        out[k] = dict(um=um, qpk=qpk, kgbox=kgbox, vbox=vb, rate=rate if rate is not None and 0 <= rate <= 100 else -1.0,
                      fonte="anagrafica")
    for k, r in std_non.items():                       # non IFB che non sono nell'anagrafica: solo l'aliquota
        if k not in out:
            out[k] = dict(um=None, qpk=None, kgbox=r.get("kgbox") or None, vbox=None, rate=r["rate"], fonte="report non IFB")
    return out


def confronta_anagrafica(std_ifb, ana):
    """Confronta i dati di imballo del modello con l'anagrafica. Ritorna (differenze, confrontati, assenti)."""
    righe, confrontati, assenti = [], 0, 0
    for k, r in std_ifb.items():
        if k not in ana:
            assenti += 1
            continue
        confrontati += 1
        for campo, chiave, nome in CAMPI_ANAGRAFICA:
            m, a = r.get(chiave), ana[k].get(campo)
            if not a:
                continue        # anagrafica vuota o a zero: l'export non ha il dato, non contraddice il modello
            if m is None or a is None or abs(m - a) > 1e-6 * max(1.0, abs(m)):
                righe.append(dict(item=k, campo=nome, modello=m, anagrafica=a))
    return pd.DataFrame(righe, columns=["item", "campo", "modello", "anagrafica"]), confrontati, assenti


# ---- esito del mese -----------------------------------------------------------------------------
def esito(diff_pct):
    """(codice, titolo, spiegazione) in base alla fascia di scostamento."""
    if diff_pct > BANDA_MAX:
        return ("rosso", "Si sta andando male",
                "Il margine a FIFO e' pari o peggiore di quello a standard: i costi reali sono troppo vicini "
                "(o sopra) lo standard. Fascia buona: tra -3% e -0,5%.")
    if diff_pct < BANDA_MIN:
        return ("giallo", "Standard troppo prudente",
                "Il FIFO costa molto meno dello standard: il margine reale e' piu' alto del previsto. "
                "Non e' un rischio di margine, ma lo standard si e' allontanato dalla realta'.")
    return ("verde", "Margine in linea",
            "Lo scostamento tra margine standard e FIFO e' nella fascia buona (tra -3% e -0,5%).")


def perimetri(art, std_ifb, std_non):
    a = art.copy()
    a["perimetro"] = np.where(a["item"].isin(std_ifb), "IFB", np.where(a["item"].isin(std_non), "Non IFB", "Fuori modello"))
    return a


def totali(a):
    """Totali del mese (vendite, margini, scostamento) per un sottoinsieme di articoli."""
    v, ms, mf_ = float(a["vendite"].sum()), float(a["m_std"].sum()), float(a["m_fifo"].sum())
    return dict(vendite=v, m_std=ms, m_fifo=mf_, diff=ms - mf_, p_std=ms / v if v else 0.0,
                p_fifo=mf_ / v if v else 0.0, diff_pct=(ms - mf_) / v if v else 0.0)


def tabella_perimetri(a):
    g = a.groupby("perimetro").agg(articoli=("item", "count"), vendite=("vendite", "sum"), m_std=("m_std", "sum"),
                                   m_fifo=("m_fifo", "sum"), diff=("diff", "sum")).reset_index()
    g["p_std"], g["p_fifo"] = g["m_std"] / g["vendite"], g["m_fifo"] / g["vendite"]
    g["diff_pct"] = g["diff"] / g["vendite"]
    return g.sort_values("diff_pct", ascending=False).reset_index(drop=True)


# ---- entrate: finestra e medie ponderate ---------------------------------------------------------
def finestra(e, mesi):
    """Entrate degli ultimi `mesi` mesi (quelli che alimentano il FIFO delle vendite del mese)."""
    fine = e["data"].max()
    inizio = (fine.to_period("M") - (mesi - 1)).to_timestamp()
    return e[e["data"] >= inizio].copy(), inizio, fine


def finestra_mese(e, chiave):
    """Entrate registrate nel mese scelto ('2026-10'). Se la chiave non e' un anno-mese si usa l'ultimo mese del file.
    Ritorna (entrate, primo giorno del mese, ultima entrata del mese)."""
    if not re.fullmatch(r"\d{4}-\d{2}", str(chiave).strip()):
        chiave = f"{e['data'].max():%Y-%m}"
    p = pd.Period(str(chiave).strip(), "M")
    w = e[e["data"].dt.to_period("M") == p].copy()
    return w, p.to_timestamp(), (w["data"].max() if len(w) else p.to_timestamp())


def escludi_provvisorie(w, fine, giorni):
    """Le entrate recentissime senza alcun costo accessorio (trasporto, despachos, imposte tutti a zero) di solito
    non hanno ancora la fattura registrata: non sono un'anomalia e falserebbero le medie. Ritorna (entrate valide,
    entrate escluse, data da cui un'entrata e' 'recente')."""
    recenti_dal = fine - pd.Timedelta(days=giorni)
    prov = (w["data"] > recenti_dal) & (w["trasp"] == 0) & (w["imp"] == 0) & (w["desp"] == 0)
    return w[~prov].copy(), w[prov].copy(), recenti_dal


def medie_entrate(w):
    """Costo unitario medio (ponderato sulla quantita') per articolo e componente, piu' oscillazione del trasporto."""
    def f(d):
        q = d["qty"]
        wa = lambda c: float((d[c] * q).sum() / q.sum())
        tr = d["trasp"][d["trasp"] > 0]
        return pd.Series(dict(
            e_prezzo=wa("prezzo"), e_trasp=wa("trasp"), e_desp=wa("desp"), e_imp=wa("imp"), e_tot=wa("tot"),
            e_altro=wa("tot") - wa("prezzo") - wa("trasp") - wa("desp") - wa("imp"),
            n_entrate=len(d), ultima=d["data"].max(), qty_entrata=float(q.sum()),
            tr_min=float(tr.min()) if len(tr) else np.nan, tr_max=float(tr.max()) if len(tr) else np.nan,
            tr_cv=float(tr.std() / tr.mean()) if len(tr) > 2 and tr.mean() else np.nan,
            isola=d["isola"].mode().iat[0] if d["isola"].notna().any() else None))
    return w.groupby("item").apply(f, include_groups=False)


def std_di(item, isola, std_ifb, std_non):
    """Componenti dello standard (prezzo, logistica, aiem, std finale, aliquota AIEM) per un articolo/isola."""
    if item in std_ifb:
        r = std_ifb[item]
        c = r.get(isola if isola in r else "GC-TF")
        return dict(prezzo=c["prezzo"], logistica=c["logistica"], aiem=c["aiem"], std=c["std"], rate=r["rate"])
    if item in std_non:
        r = std_non[item]
        return dict(prezzo=r["prezzo"], logistica=r["logistica"], aiem=r["aiem"], std=r["std"], rate=r["rate"])
    return None


# ---- controllo 1: AIEM --------------------------------------------------------------------------
def controllo_aiem(w, std_ifb, std_non, recenti_dal=None):
    """Per ogni entrata: imposte registrate contro l'AIEM atteso dall'anagrafica. Ritorna le entrate anomale."""
    righe = []
    for r in w.itertuples():
        s = std_di(r.item, r.isola, std_ifb, std_non)
        if s is None:
            continue
        atteso = s["aiem"]
        if s["rate"] < 0:
            continue
        if s["rate"] == 0 and r.imp > 0.002:
            tipo = "AIEM registrato, ma l'articolo non ha AIEM in anagrafica"
        elif s["rate"] > 0 and r.imp <= 0.0005:
            if recenti_dal is not None and r.data > recenti_dal:
                continue                       # entrata recente: la fattura delle imposte puo' non essere ancora registrata
            tipo = "AIEM non registrato, ma l'articolo ha AIEM in anagrafica"
        elif s["rate"] > 0 and atteso > 0 and not (0.5 <= r.imp / atteso <= 2.0):
            tipo = "AIEM registrato molto diverso dall'atteso"
        else:
            continue
        righe.append(dict(item=r.item, desc=r.desc, data=r.data, ubic=r.ubic, ext=r.ext, doc=r.doc, qty=r.qty,
                          uom=r.uom, aliquota=s["rate"], imp_reg=r.imp, imp_att=atteso,
                          extra_eur=r.qty * (r.imp - atteso), tipo=tipo))
    return pd.DataFrame(righe)


def riepilogo_aiem(an, w):
    if an.empty:
        return an
    n_tot = w.groupby("item").size()
    g = an.groupby(["item", "desc", "tipo"]).agg(entrate_anomale=("doc", "count"), extra_eur=("extra_eur", "sum"),
                                                 ultima=("data", "max"), aliquota=("aliquota", "first"),
                                                 spedizioni=("ext", lambda s: ", ".join(sorted(set(s))))).reset_index()
    g["entrate_totali"] = g["item"].map(n_tot)
    return g.sort_values("extra_eur", key=lambda s: s.abs(), ascending=False)


# ---- controllo 2: container ---------------------------------------------------------------------
def _tvd(a, b):
    return float(np.abs(a - b).sum() / 2)


def controllo_container(w, std_ifb, min_righe=3):
    """Isola ogni spedizione, container o camion (stesso External Document No_) e verifica come e' ripartito il trasporto.
    Atteso: quota di volume = box x volume del box, con box = pezzi / pezzi per box.
    Errore noto di COMIT: volume del box x pezzi (senza dividere per i pezzi per box)."""
    contenitori, dettagli = [], []
    for ext, d in w[w["trasp"] > 0].groupby("ext"):
        # imposte ripartite con la stessa chiave del trasporto (stesso rapporto imposte/trasporto su tutte le righe)
        x = d[(d["imp"] > 0) & (d["trasp"] > 0)]
        stessa = bool(len(x) >= 3 and (x["imp"] / x["trasp"]).std() / (x["imp"] / x["trasp"]).mean() < 0.02)
        n_senza_aiem, indebite = 0, 0.0
        for r in x.itertuples():
            if r.item in std_ifb and std_ifb[r.item]["rate"] == 0:
                n_senza_aiem += 1
                indebite += r.imp * r.qty          # imposte caricate su un articolo che non ha AIEM
        righe = []
        for r in d.itertuples():
            s = std_ifb.get(r.item)
            if not s or not s.get("vbox"):
                continue
            if r.uom == "BOX":
                scatole = r.qty
            elif r.uom == "KG" and s.get("kgbox"):
                scatole = r.qty / s["kgbox"]
            elif s.get("qpk"):
                scatole = r.qty / s["qpk"]
            else:
                continue
            righe.append(dict(item=r.item, desc=r.desc, uom=r.uom, qty=r.qty, scatole=scatole,
                              reg=r.trasp * r.qty, v_ok=scatole * s["vbox"], v_err=r.qty * s["vbox"],
                              merce=r.prezzo * r.qty, imp_reg=r.imp * r.qty, aliquota=s["rate"], data=r.data,
                              desp_reg=r.desp * r.qty))
        base = dict(ext=ext, data=d["data"].max(), ubic=d["ubic"].iat[0], righe=len(d), righe_volume=len(righe),
                    stessa_chiave=stessa, imposte_su_senza_aiem=n_senza_aiem, imposte_indebite_eur=indebite, trasporto_eur=float((d["trasp"] * d["qty"]).sum()),
                    despachos_eur=float((d["desp"] * d["qty"]).sum()), imposte_eur=float((d["imp"] * d["qty"]).sum()),
                    merce_eur=float((d["prezzo"] * d["qty"]).sum()))
        if len(righe) < min_righe:
            base.update(esito="Non verificabile (pochi articoli con dati di volume)", err_ok=np.nan, err_bug=np.nan, sbilancio_eur=np.nan)
            contenitori.append(base)
            continue
        t = pd.DataFrame(righe)
        tot = t["reg"].sum()
        if tot <= 0:
            continue
        s_reg = t["reg"] / tot
        s_ok, s_err = t["v_ok"] / t["v_ok"].sum(), t["v_err"] / t["v_err"].sum()
        e_ok, e_bug = _tvd(s_reg, s_ok), _tvd(s_reg, s_err)
        if e_ok <= 0.10:
            es = "Ripartito a volume correttamente"
        elif e_bug < e_ok and e_bug <= 0.10:
            es = "Volume del collo x pezzi (manca la divisione per i pezzi per collo)"
        else:
            es = "Ripartizione non coerente con il volume"
        t["atteso"] = tot * s_ok
        t["sbilancio"] = t["reg"] - t["atteso"]
        # AIEM atteso: aliquota dell'articolo x (merce + trasporto atteso della riga), come nella fattura delle imposte
        t["aiem_att"] = t["aliquota"].clip(lower=0) / 100 * (t["merce"] + t["atteso"])
        t["aiem_diff"] = t["imp_reg"] - t["aiem_att"]
        # despachos atteso: stessa quota di volume del trasporto; totale = prezzo di acquisto + trasporto + AIEM + despachos
        t["desp_att"] = t["desp_reg"].sum() * s_ok
        t["tot_reale"] = t["merce"] + t["reg"] + t["imp_reg"] + t["desp_reg"]
        t["tot_atteso"] = t["merce"] + t["atteso"] + t["aiem_att"] + t["desp_att"]
        base.update(esito=es, err_ok=e_ok, err_bug=e_bug, sbilancio_eur=float(t["sbilancio"].abs().sum() / 2))
        contenitori.append(base)
        t.insert(0, "ext", ext)          # le righe si tengono per tutte le spedizioni verificabili (anche quelle corrette)
        dettagli.append(t)
    cont = pd.DataFrame(contenitori)
    if len(cont):
        cont["trasp_pct"] = cont["trasporto_eur"] / cont["merce_eur"].replace(0, np.nan)
        # gravita': costo spostato tra le righe (volume sbagliato) + imposte caricate su articoli senza AIEM
        cont["impatto_eur"] = cont["sbilancio_eur"].fillna(0) + cont["imposte_indebite_eur"]
    det = pd.concat(dettagli, ignore_index=True) if dettagli else pd.DataFrame()
    return cont, det


def _it(x, dec=0):
    """Numero all'italiana: 1.896 / 3,45."""
    s = f"{abs(x):,.{dec}f}".replace(",", "§").replace(".", ",").replace("§", ".")
    return ("-" if x < 0 else "") + s


def _eur(x):
    return _it(x) + " \u20ac"


def _n(x):
    """Numero senza decimali inutili: 99 invece di 99,0."""
    return _it(x, 0 if abs(x - round(x)) < 0.05 else 1)


def analisi_spedizione(tt, d, w_ext, std_ifb):
    """Spiegazione chiara e analitica dei problemi di una spedizione (container o camion = stessa fattura IFB).
    Ritorna una lista di problemi: titolo, atteso, trovato, esempio, effetto, richiesta (testi pronti per la lettura
    e per la segnalazione a COMIT)."""
    um = {"PCS": "pezzi", "KG": "kg", "BOX": "box"}
    problemi = []
    es = tt["esito"]
    tot_tr = tt["trasporto_eur"]
    if es.startswith("Volume del collo") and len(d):
        o = d.sort_values("sbilancio", ascending=False)
        su, giu = o.iloc[0], o.iloc[-1]
        n_su, n_giu = int((d["sbilancio"] > 1).sum()), int((d["sbilancio"] < -1).sum())

        def riga(r):
            pps = r["qty"] / r["scatole"] if r["scatole"] else 0
            return (f"{str(r['desc']).strip()}: {_it(r['qty'])} {um.get(r['uom'], r['uom'])} in {_n(r['scatole'])} box "
                    f"({_n(pps)} per box) -> registrati {_eur(r['reg'])} di trasporto, a volume ne spetterebbero "
                    f"{_eur(r['atteso'])}")
        problemi.append(dict(
            titolo="Trasporto diviso con il volume sbagliato",
            atteso="Il trasporto di una spedizione va diviso tra gli articoli in base al volume che occupano: numero di box "
                   "x volume di un box.",
            trovato="La divisione e' stata fatta con il volume del box x i pezzi, senza dividere prima i pezzi per quanti "
                    "ne contiene un box. Gli articoli con tanti pezzi per box risultano molto piu' ingombranti di quanto sono.",
            esempio=[riga(su), riga(giu)],
            effetto=f"{_eur(tt['sbilancio_eur'])} di trasporto (su {_eur(tot_tr)}) sono finiti sugli articoli sbagliati: "
                    f"{n_su} {'sovraccaricato' if n_su == 1 else 'sovraccaricati'}, "
                    f"{n_giu} {'sottocaricato' if n_giu == 1 else 'sottocaricati'}. Il totale della spedizione non cambia, ma cambiano costo "
                    "FIFO e margine di ogni articolo.",
            richiesta="Ripartire il trasporto in base al numero di box x volume del box (pezzi diviso pezzi per box)."))
    elif es.startswith("Ripartizione non coerente") and len(d):
        o = d.reindex(d["sbilancio"].abs().sort_values(ascending=False).index)
        r1 = o.iloc[0]
        problemi.append(dict(
            titolo="Trasporto diviso con un criterio non chiaro",
            atteso="Il trasporto di una spedizione va diviso tra gli articoli in base al volume che occupano: numero di box "
                   "x volume di un box.",
            trovato="Le quote registrate non corrispondono al volume dei colli, nemmeno con l'errore gia' segnalato "
                    "(pezzi x volume del box): potrebbe essere stato usato un altro criterio (valore, peso, quantita') "
                    "o una ripartizione manuale.",
            esempio=[f"{str(r1['desc']).strip()}: registrati {_eur(r1['reg'])} di trasporto, a volume ne spetterebbero "
                     f"{_eur(r1['atteso'])} ({'+' if r1['sbilancio'] > 0 else '-'}{_eur(abs(r1['sbilancio']))})"],
            effetto=f"Circa {_eur(tt['sbilancio_eur'])} di trasporto (su {_eur(tot_tr)}) risultano assegnati agli articoli "
                    "sbagliati; il totale non cambia ma i margini per articolo si'.",
            richiesta="Chiarire quale criterio e' stato usato per dividere il trasporto di questa fattura."))
    elif pd.notna(tt.get("trasp_pct")) and tt["trasp_pct"] > 0.35:
        problemi.append(dict(
            titolo="Trasporto molto alto rispetto alla merce",
            atteso="Il costo del trasporto e' di solito una frazione contenuta del valore della merce trasportata.",
            trovato=f"Il trasporto registrato e' {_eur(tot_tr)} su una merce da {_eur(tt['merce_eur'])}, cioe' il "
                    f"{_it(tt['trasp_pct'] * 100)}%.",
            esempio=[],
            effetto="Tutti gli articoli della spedizione hanno un costo FIFO molto piu' alto dello standard.",
            richiesta="Verificare l'importo con la fattura del trasportatore e che non sia stata registrata due volte o su "
                      "piu' spedizioni."))
    # imposte (AIEM)
    senza = {}
    for r in w_ext[w_ext["imp"] > 0].itertuples():
        if r.item in std_ifb and std_ifb[r.item]["rate"] == 0:
            nome = str(r.desc).strip()
            senza[nome] = senza.get(nome, 0.0) + r.imp * r.qty
    rapp = tt["imposte_eur"] / tot_tr if tot_tr else 0
    if tt["imposte_indebite_eur"] >= 1:
        elenco = sorted(senza.items(), key=lambda kv: -kv[1])
        trovato = (f"{len(senza)} {'articolo' if len(senza) == 1 else 'articoli'} che in anagrafica non ha"
                   f"{'' if len(senza) == 1 else 'nno'} AIEM ha"
                   f"{'' if len(senza) == 1 else 'nno'} ricevuto {_eur(tt['imposte_indebite_eur'])} di imposte.")
        if tt["stessa_chiave"]:
            trovato += (f" Le imposte sono divise con lo stesso criterio del trasporto (il rapporto imposte/trasporto e' lo "
                        f"stesso su tutte le righe, circa {_it(rapp * 100)}%) e non in base all'aliquota di ogni articolo.")
        problemi.append(dict(
            titolo="Imposte (AIEM) caricate su articoli che non le prevedono",
            atteso="L'AIEM si applica solo agli articoli che hanno un'aliquota in anagrafica, ed e' proporzionale all'aliquota.",
            trovato=trovato,
            esempio=[f"{n}: {_eur(v)} di imposte pur non avendo AIEM in anagrafica" for n, v in elenco[:3]],
            effetto="Costo FIFO e margine di quegli articoli sono falsati; agli articoli che hanno davvero AIEM ne arriva meno "
                    "del dovuto.",
            richiesta="Assegnare le imposte solo agli articoli con AIEM in anagrafica, in base alla loro aliquota."))
    elif tt["stessa_chiave"]:
        problemi.append(dict(
            titolo="Imposte divise come il trasporto",
            atteso="Le imposte (AIEM) vanno divise in base all'aliquota di ogni articolo.",
            trovato=f"Il rapporto imposte/trasporto e' lo stesso su tutte le righe (circa {_it(rapp * 100)}%): le imposte sono "
                    "state divise con la stessa chiave del trasporto.",
            esempio=[],
            effetto="Gli articoli con aliquota alta ricevono meno imposte del dovuto e quelli con aliquota bassa piu' del dovuto.",
            richiesta="Dividere le imposte in base all'aliquota AIEM di ogni articolo."))
    return problemi


def _file_tassi():
    return os.path.join(percorsi.DATI, f"tassi_cambio_{_FIL['k'].replace(' ', '_')}.json")


def leggi_tassi():
    """{spedizione: tasso di cambio usato alla registrazione dell'entrata}, inserito a mano (Hong Kong)."""
    try:
        with open(_file_tassi(), encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return {}


def salva_tasso(spedizione, valore):
    t = leggi_tassi()
    if valore and valore > 0:
        t[str(spedizione)] = float(valore)
    else:
        t.pop(str(spedizione), None)
    with open(_file_tassi(), "w", encoding="utf-8") as fh:
        json.dump(t, fh, ensure_ascii=False, indent=1)


def testo_segnalazione(tt, problemi, hk=False, tasso=None):
    """Testo pronto da incollare in una mail a COMIT (hk=True: Hong Kong, senza imposte e despachos, con FOB al posto del trasporto)."""
    data = tt["data"].strftime("%d/%m/%Y") if pd.notna(tt["data"]) else ""
    t = [f"Oggetto: Ripartizione costi accessori - fattura/spedizione {tt['ext']}", "",
         (f"Buongiorno, sulla spedizione {tt['ext']} (ultima entrata del {data}) risulta: FOB totale "
          f"{_eur(tt['trasporto_eur'])}, valore merce {_eur(tt['merce_eur'])}. Abbiamo rilevato quanto segue."
          + (f" Tasso di cambio usato alla registrazione dell'entrata: {tasso:.4f}." if tasso else "")
          if hk else
          f"Buongiorno, sulla spedizione {tt['ext']} (ultima entrata del {data}, {tt['ubic']}) risulta: trasporto totale "
          f"{_eur(tt['trasporto_eur'])}, imposte {_eur(tt['imposte_eur'])}, despachos {_eur(tt['despachos_eur'])}, valore merce "
          f"{_eur(tt['merce_eur'])}. Abbiamo rilevato quanto segue."), ""]
    for i, p in enumerate(problemi, 1):
        t.append(f"{i}) {p['titolo']}")
        t.append(f"   Cosa ci aspettiamo: {p['atteso']}")
        t.append(f"   Cosa abbiamo trovato: {p['trovato']}")
        for e in p["esempio"]:
            t.append(f"   Esempio - {e}")
        t.append(f"   Effetto: {p['effetto']}")
        t.append(f"   Richiesta: {p['richiesta']}")
        t.append("")
    if hk:          # in testo: il costo e' il FOB
        t = [x.replace("trasporto", "FOB").replace("Trasporto", "FOB") for x in t]
    t.append("Potete verificare e, se confermato, correggere la registrazione? Grazie.")
    return "\n".join(t)


# ---- controllo 3: trasporto nel tempo e rispetto allo standard -----------------------------------
def controllo_trasporto(med, std_ifb, std_non, tol_cv=0.40, tol_rapporto=2.0):
    """Articoli il cui trasporto registrato oscilla molto tra le entrate o e' lontano dalla logistica dello standard
    (trasporto + despachos reali contro logistica dello standard). Le differenze piccole sono normali e non si segnalano."""
    righe = []
    for item, r in med.iterrows():
        s = std_di(item, r["isola"], std_ifb, std_non)
        reale = r["e_trasp"] + r["e_desp"]
        std_log = s["logistica"] if s else np.nan
        rap = (reale / std_log) if pd.notna(std_log) and std_log > 0 else np.nan
        segnali = []
        if pd.notna(r["tr_cv"]) and r["n_entrate"] >= 3 and r["tr_cv"] > tol_cv:
            segnali.append(f"Trasporto instabile tra le entrate (da {r['tr_min']:.2f} a {r['tr_max']:.2f} per unita')")
        if pd.notna(std_log) and std_log >= 0.05 and reale <= 0.001:
            pass        # trasporto non registrato nelle entrate: non e' un'anomalia di ripartizione, l'articolo non si mostra
        elif pd.notna(rap) and abs(reale - std_log) >= 0.05 and (rap > tol_rapporto or rap < 1 / tol_rapporto):
            segnali.append(f"Trasporto reale {rap:.1f}x lo standard")
        if not segnali:
            continue
        righe.append(dict(item=item, entrate=int(r["n_entrate"]), trasp_min=r["tr_min"], trasp_max=r["tr_max"],
                          oscillazione=r["tr_cv"], reale=reale, std_logistica=std_log, rapporto=rap,
                          segnale=" · ".join(segnali)))
    return pd.DataFrame(righe)


# ---- outlier e diagnosi per articolo -------------------------------------------------------------
def outlier(a, min_pct, min_eur):
    x = a[(a["diff"].abs() >= min_eur) & (a["diff_pct"].abs() >= min_pct)].copy()
    return x.sort_values("diff", ascending=False)      # dal peggiore (FIFO sopra lo standard) al migliore


def diagnosi_hk(a, ent, mese, in_modello):
    """Cause probabili dello scostamento STD-FIFO per Hong Kong (a = articoli, ent = entrate nel formato Canarie, mese 'AAAA-MM',
    in_modello = codici degli articoli gestiti nel modello). Si guarda cosa e' cambiato nelle entrate del mese rispetto allo storico:
    - FOB (trasporto) molto piu' alto: e' arrivata via aerea una merce etichettata via mare («ibrido»);
    - prezzo di acquisto diverso: possibile tasso di cambio (la merce e' registrata al cambio del giorno di entrata; via mare l'arrivo
      richiede circa un mese, via aerea pochi giorni)."""
    m = ent["data"].dt.strftime("%Y-%m") == str(mese).strip()
    e_m, e_h = ent[m], ent[~m & (ent["data"] < ent[m]["data"].min())] if m.any() else ent.iloc[0:0]
    e_h = e_h[e_h["data"] >= (e_h["data"].max() - pd.Timedelta(days=200))] if len(e_h) else e_h

    def medie(e, col):
        v = e[col].mul(e["qty"]).groupby(e["item"]).sum() / e.groupby("item")["qty"].sum()
        return v
    p_m, f_m = medie(e_m, "prezzo"), medie(e_m, "trasp")
    p_h = e_h.groupby("item")["prezzo"].median()
    f_h = e_h.groupby("item")["trasp"].median()
    out = []
    for r in a.itertuples():
        it = r.item
        cause, d_p, d_f = [], np.nan, np.nan
        if it in p_m.index and it in p_h.index and p_h[it] > 0:
            d_p = float(p_m[it] - p_h[it])
            pct = p_m[it] / p_h[it] - 1
            if abs(pct) >= 0.03:
                cause.append(f"prezzo di acquisto {pct:+.1%} rispetto allo storico: possibile tasso di cambio (registrato al cambio del "
                             "giorno di entrata; via mare l'arrivo richiede circa un mese)")
        if it in f_m.index and it in f_h.index and f_h[it] > 0:
            d_f = float(f_m[it] - f_h[it])
            if f_m[it] >= 2 * f_h[it] and d_f > 0.05:
                cause.append(f"FOB {f_m[it] / f_h[it]:.1f}x lo storico: probabile spedizione via aerea su un articolo via mare")
        elif it in f_m.index and f_m[it] > 0 and it not in f_h.index:
            cause.append("nessuna entrata precedente per confrontare il FOB")
        if it not in in_modello:
            cause.append("non è nel modello: standard gestito da altri (se è anomalo, segnalare: c'è da aggiornare lo standard cost)")
        if it not in e_m["item"].values:
            cause = ["nessuna entrata nel periodo analizzato: il FIFO arriva da giacenza piu' vecchia"] + cause
        gap = float(getattr(r, "c_fifo") - getattr(r, "c_std"))
        d_altro = gap - (0 if pd.isna(d_p) else d_p) - (0 if pd.isna(d_f) else d_f)
        out.append(dict(item=it, cause=" | ".join(cause) if cause else "", d_prezzo=d_p, d_logistica=d_f, d_aiem=0.0, d_altro=d_altro))
    return pd.DataFrame(out, columns=["item", "cause", "d_prezzo", "d_logistica", "d_aiem", "d_altro"]).set_index("item")


def diagnosi(a, med, an_aiem, trasp, det_cont, std_ifb, std_non):
    """Aggiunge ad ogni articolo la scomposizione del divario unitario FIFO - STD e le cause probabili."""
    aiem_per_item = {}
    if len(an_aiem):
        for item, g in an_aiem.groupby("item"):
            aiem_per_item[item] = "; ".join(sorted(set(g["tipo"])))
    tr_per_item = trasp.set_index("item")["segnale"].to_dict() if len(trasp) else {}
    cont_per_item = {}
    if len(det_cont):
        for item, g in det_cont.groupby("item"):
            peggiore = g.loc[g["sbilancio"].abs().idxmax()]
            if abs(peggiore["sbilancio"]) >= 20:
                cont_per_item[item] = f"spedizione {peggiore['ext']}: trasporto {peggiore['sbilancio']:+.0f} € vs ripartizione a volume"
    righe = []
    for r in a.itertuples():
        cause = []
        gap = r.c_fifo - r.c_std
        d_prezzo = d_log = d_aiem = d_altro = np.nan
        s = None
        m = med.loc[r.item] if r.item in med.index else None
        if m is not None:
            s = std_di(r.item, m["isola"], std_ifb, std_non)
        if m is not None and s is not None and s["prezzo"] > 0:
            d_prezzo = m["e_prezzo"] - s["prezzo"]
            d_log = (m["e_trasp"] + m["e_desp"]) - s["logistica"]
            d_aiem = m["e_imp"] - s["aiem"]
            buffer = r.c_std - (s["prezzo"] + s["logistica"] + s["aiem"])      # maggiorazione concordata (+2%)
            d_altro = gap - (d_prezzo + d_log + d_aiem - buffer)               # non spiegato dalle entrate della finestra
            if s["prezzo"] and abs(d_prezzo) / s["prezzo"] > 0.03:
                cause.append(f"prezzo di acquisto {d_prezzo / s['prezzo']:+.0%} rispetto allo standard")
        elif m is None:
            cause.append("nessuna entrata nel periodo analizzato: il FIFO arriva da giacenza piu' vecchia")
        if r.item in aiem_per_item:
            cause.append(aiem_per_item[r.item])
        if r.item in tr_per_item and tr_per_item[r.item]:
            cause.append(tr_per_item[r.item])
        if r.item in cont_per_item:
            cause.append(cont_per_item[r.item])
        if r.c_std == 0:
            cause.append("standard cost a zero: articolo senza standard nel modello")
        if r.c_std > 0 and (r.c_fifo / r.c_std > 3 or r.c_fifo / r.c_std < 1 / 3):
            cause.append("costo FIFO e STD di ordini di grandezza diversi: controllare l'unita' di misura")
        righe.append(dict(item=r.item, gap_unit=gap, d_prezzo=d_prezzo, d_logistica=d_log, d_aiem=d_aiem,
                          d_altro=d_altro, cause=" | ".join(cause)))
    return pd.DataFrame(righe).set_index("item")


# ---- storico mensile ----------------------------------------------------------------------------
def leggi_storico():
    try:
        with open(_file_storico(), encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def salva_mese(mese, tot, perim):
    """Registra nello storico il riepilogo del mese (si aggiorna da solo a ogni caricamento; non riscrive se invariato)."""
    s = leggi_storico()
    nuovo = dict(vendite=tot["vendite"], p_std=tot["p_std"], p_fifo=tot["p_fifo"], diff_pct=tot["diff_pct"],
                 diff_eur=tot["diff"], perimetri={r.perimetro: r.diff_pct for r in perim.itertuples()})
    if s.get(mese) == nuovo:
        return s
    s[mese] = nuovo
    with open(_file_storico(), "w", encoding="utf-8") as f:
        json.dump(dict(sorted(s.items())), f, ensure_ascii=False, indent=1)
    return s


def excel_report(tot, perim, out, aiem, cont, det_cont, trasp):
    """Excel con le tabelle principali, da girare a chi deve correggere le registrazioni."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        pd.DataFrame([tot]).to_excel(xw, sheet_name="Esito mese", index=False)
        perim.to_excel(xw, sheet_name="Per perimetro", index=False)
        out.to_excel(xw, sheet_name="Outlier", index=False)
        for nome, df in (("AIEM", aiem), ("Spedizioni", cont), ("Spedizioni righe", det_cont), ("Trasporto", trasp)):
            if df is not None and len(df):
                df.to_excel(xw, sheet_name=nome, index=False)
    return buf.getvalue()


# ---- mesi salvati (i 3 caricamenti di un mese restano in memoria) --------------------------------
CARTELLA_MESI = os.path.join(percorsi.DATI, "MARGINI_MESI")
_FIL = {"k": "CANARIE"}          # filiale: Canarie resta com'era, le altre hanno una sottocartella e un file di storico propri


def imposta_filiale(k):
    _FIL["k"] = k


def _dir_mesi():
    return CARTELLA_MESI if _FIL["k"] == "CANARIE" else os.path.join(CARTELLA_MESI, _FIL["k"])


def _file_storico():
    return STORICO if _FIL["k"] == "CANARIE" else os.path.join(percorsi.DATI, f"margini_storico_{_FIL['k']}.json")
NOMI_MESI = ["Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", "Luglio", "Agosto", "Settembre", "Ottobre",
             "Novembre", "Dicembre"]
_FILE_MESE = {"marg": "margini.xlsx", "cost": "cost_analysis.xlsx", "non": "standard_non_ifb.xlsx"}


def etichetta_mese(chiave):
    """'2026-10' -> 'Ottobre 2026' (altre diciture restano com'e')."""
    m = re.fullmatch(r"(\d{4})-(\d{2})", chiave.strip())
    return f"{NOMI_MESI[int(m.group(2)) - 1]} {m.group(1)}" if m and 1 <= int(m.group(2)) <= 12 else chiave


def mesi_salvati():
    """Chiavi dei mesi salvati, dal piu' vecchio al piu' recente."""
    if not os.path.isdir(_dir_mesi()):
        return []
    return sorted(d for d in os.listdir(_dir_mesi())
                  if os.path.exists(os.path.join(_dir_mesi(), d, _FILE_MESE["marg"])))


def salva_file_mese(chiave, marg, cost=None, non=None):
    """Salva i caricamenti del mese (byte dei file). Sovrascrive il mese se esiste gia'."""
    chiave = re.sub(r'[\/:*?"<>|]', "-", chiave.strip())
    d = os.path.join(_dir_mesi(), chiave)
    os.makedirs(d, exist_ok=True)
    for k, dati in (("marg", marg), ("cost", cost), ("non", non)):
        f = os.path.join(d, _FILE_MESE[k])
        if dati:
            with open(f, "wb") as fh:
                fh.write(dati)
        elif os.path.exists(f):
            os.remove(f)
    return chiave


def sostituisci_file_mese(chiave, marg=None, cost=None, non=None):
    """Corregge un caricamento sbagliato: sostituisce solo i file passati, gli altri restano."""
    d = os.path.join(_dir_mesi(), chiave)
    for k, dati in (("marg", marg), ("cost", cost), ("non", non)):
        if dati:
            with open(os.path.join(d, _FILE_MESE[k]), "wb") as fh:
                fh.write(dati)


def leggi_file_mese(chiave):
    """(marg, cost, non) come byte; None dove il file non c'e'."""
    d = os.path.join(_dir_mesi(), chiave)
    out = []
    for k in ("marg", "cost", "non"):
        f = os.path.join(d, _FILE_MESE[k])
        out.append(open(f, "rb").read() if os.path.exists(f) else None)
    return tuple(out)
