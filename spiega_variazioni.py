"""
spiega_variazioni.py  --  "Perche' e' cambiato lo standard cost?"

Confronta due versioni del modello Standard Cost (.xlsm/.xlsx, SOLA LETTURA, valori
calcolati salvati da Excel) e scompone la variazione di ogni articolo, isola per isola,
nelle sue cause: listino, trasporto, assicurazione, pallet, AIEM, MTS/MTO, altro.

Le colonne vengono cercate per NOME dell'intestazione, non per posizione, cosi' funziona anche
sul foglio BEVERAGE (colonne spostate) e su MEAT-FISH (intestazione a riga 7).
Gli articoli si cercano con il No_ della filiale oppure con il codice IFB della sede.

Uso:
  python spiega_variazioni.py VECCHIO.xlsm NUOVO.xlsm --articolo 14623
  python spiega_variazioni.py VECCHIO.xlsm NUOVO.xlsm --articolo 14623 14627 --isola GC-TF
  python spiega_variazioni.py VECCHIO.xlsm NUOVO.xlsm            # report Excel di tutti i cambiati

Nota: il "costo standard" confrontato e' lo STANDARD Cost Step 2 (prima del +2% concordato).
Il +2% e' mostrato a parte nel report. Le cause "tariffe" sono dedotte per esclusione
(nessun driver dell'articolo e' cambiato), non lette da COSTS (LOG).
"""
import argparse
import os
import re
import sys
from collections import defaultdict

import openpyxl
from openpyxl.utils.cell import column_index_from_string, coordinate_to_tuple, range_boundaries

TOL = 0.005  # sotto questa differenza (EUR) consideriamo il valore invariato


# ---- Configurazione per filiale ---------------------------------------------------------------
# Token delle colonne (nome intestazione normalizzato):  "nome" = colonna esatta,
# "prefisso*" = prima colonna che inizia cosi', "aiem#k" = k-esima colonna AIEM per unita'.
def _canarie_dest():
    # Standard unico per coppia di isole: Gran Canaria + Tenerife (colonne di Gran Canaria; in tutti gli
    # articoli coincidono) e Lanzarote + Fuerteventura (colonne di Lanzarote).
    isole = [("GC-TF", "Gran Canaria / Tenerife", "grancanaria", "grancan", 0),
             ("LAN-FUE", "Lanzarote / Fuerteventura", "lanzarote", "lanzarote", 1)]
    return {
        cod: dict(
            nome=nome, s1=f"standardcoststep1{step}", s2=f"standardcoststep2{step}",
            parts={
                "Trasporto": [key, "inlandcustomclearence" + key, "veronabarcellona*", "barcellona" + key],
                "Assicurazione": ["assicurazione*"],
                "Pallet": ["palletcostperunit"],
                "AIEM": [f"aiem#{aiem}"],
            })
        for cod, nome, key, step, aiem in isole}


BRANCHES = {
    "CANARIE": dict(
        label="Canarie", sheets={"ITEM STANDARD COST": 5, "ITEM STANDARD COST_BEVERAGE": 5,
                                 "ITEM STANDARD COST_MEAT-FISH": 7},
        dest=_canarie_dest(), delist_row=3,
        groups=["Listino", "Trasporto", "Assicurazione", "Pallet", "AIEM", "MTS/MTO", "Maggiorazione +2%", "Altro"],
        kinds={"Trasporto": "transport", "Assicurazione": "insurance", "Pallet": "pallet", "AIEM": "aiem"},
        mode_col="truckocontainer", area_col="area",
        model_glob="05_Modello_Standard_Cost_COMIT*.xlsm",
        folder="Data Analysis - 066 STANDARD COST CANARIE",
        current="STANDARD_COST_AUTOMATION/INPUT/05_Modello_Standard_Cost_COMIT.xlsm",
        storico="IFB/STORICO MESI/**/*Modello*Standard*Cost*.xls*",
        dest_label="Isole",
        log={"transport": ["C3", "F3", "G3", "D5:F5", "D9:F12", "F14:F16", "D20:F23", "D27:E30"],
             "pallet": ["I1"], "mtsmto": ["E56:G59", "E62:G64"], "insurance": ["E71"]}),
    "HONG KONG": dict(
        label="Hong Kong", sheets={"ITEM STDC_SEA_FOOD": 5, "ITEM STDC_SEA_WINE": 5, "ITEM STDC_SEA_MEAT-FISH": 5},
        dest={"HK": dict(nome="Hong Kong", s1="standardcoststep1euro", s2="standardcoststep2euro",
                         parts={"FOB": ["fobcostperunit"],
                                "Oneri import": ["localimportchargesperunit"],
                                "VGM": ["vgmchargesperunit"],
                                "Cert. sanitario": ["healthcertificateperunit"],
                                "Pallet": ["palletcostperunit"],
                                "Imposta alcolici": ["alcolctaxperunit"]})},
        delist_row=1,
        groups=["Listino", "FOB", "Oneri import", "VGM", "Cert. sanitario", "Pallet", "Imposta alcolici",
                "MTS/MTO", "Maggiorazione +2%", "Altro"],
        kinds={"FOB": "transport", "Oneri import": "tariff", "VGM": "tariff", "Cert. sanitario": "tariff",
               "Pallet": "pallet", "Imposta alcolici": "alcool"},
        mode_col="", area_col="areanordcentrosud",
        model_glob="05_Modello_Standard_Cost*.xls*",
        folder="Data Analysis - 068 STANDARD COST HONG KONG",
        current="STANDARD_COST_AUTOMATION/INPUT/05_Modello_Standard_Cost_HK.xlsm",
        storico="IFB/**/05_Modello_Standard_Cost*.xlsm",
        dest_label="Destinazione",
        log={"transport": ["D23:F25", "D29:F29", "D32:F33", "C13:D15", "C35", "C2"],
             "pallet": ["I2"], "mtsmto": ["D43:F46", "D49:F51"]}),
}

BRANCH = BRANCHES["CANARIE"]
HEADER_ROWS, SHEETS, ISLANDS, GROUPS = {}, [], {}, []


def set_branch(name):
    """Seleziona la filiale (CANARIE / HONG KONG): fogli, destinazioni e colonne di costo."""
    global BRANCH, HEADER_ROWS, SHEETS, ISLANDS, GROUPS
    BRANCH = BRANCHES[name]
    HEADER_ROWS = BRANCH["sheets"]
    SHEETS = list(HEADER_ROWS)
    ISLANDS = BRANCH["dest"]
    GROUPS = BRANCH["groups"]


set_branch("CANARIE")

# Dati di pallettizzazione/imballo per articolo (nome colonna normalizzato -> etichetta leggibile).
# Se uno di questi cambia, la variazione di trasporto/pallet viene attribuita ai dati logistici.
DATA_LABELS = {
    "quantityxplt": "unita' per pallet",
    "packagingxpallet": "colli per pallet",
    "quantityxpackaging": "unita' per collo",
    "layerxpallet": "strati per pallet",
    "packagingxlayer": "colli per strato",
    "kgxplt": "kg per pallet",
    "divisorecollo": "divisore collo",
    "kgperbox": "kg per collo",
    "numeropalletpercontainerreale": "pallet per container (reale)",
}
DATA_DRIVERS = list(DATA_LABELS)


def norm(s):
    return re.sub(r"[^a-z0-9]", "", str(s).lower()) if s is not None else ""


def num(x):
    return float(x) if isinstance(x, (int, float)) and not isinstance(x, bool) else None


def fmt(x, signed=True):
    s = f"{x:+.2f}" if signed else f"{x:.2f}"
    return s.replace(".", ",")


class Sheet:
    """Un foglio ITEM STANDARD COST* letto in memoria, con accesso per nome colonna."""

    def __init__(self, ws, header_row):
        rows = list(ws.iter_rows(min_row=header_row, values_only=True))
        self.header = rows[0]
        self.idx = defaultdict(list)
        for i, h in enumerate(self.header):
            n = norm(h)
            if n:
                self.idx[n].append(i)
        self.items = {}
        for r in rows[1:]:
            k = r[0]
            if k in (None, ""):
                continue
            self.items[str(k).strip()] = r
        # No_ = codice della filiale; "IFB Item" = codice della sede: permetto la ricerca per entrambi
        self.ifb = {}
        c_ifb = self.col("ifbitem")
        if c_ifb is not None:
            for k, r in self.items.items():
                v = r[c_ifb] if c_ifb < len(r) else None
                if v not in (None, "") and not str(v).startswith("#"):
                    self.ifb.setdefault(str(v).strip(), []).append(k)
        aiem_cols = sorted(i for n, l in self.idx.items() if n.startswith("aiem") and n != "aiem" for i in l)
        self.aiem_cols = aiem_cols[:2]

    def col(self, name, occ=0):
        l = self.idx.get(name)
        if not l or occ >= len(l):
            return None
        return l[occ]

    def val(self, row, name, occ=0):
        c = self.col(name, occ)
        return row[c] if c is not None and c < len(row) else None

    def n(self, row, name, occ=0):
        return num(self.val(row, name, occ))


_REF_LOG = re.compile(r"'COSTS \(LOG\)'!\$?([A-Z]{1,2})\$?(\d+)(?::\$?([A-Z]{1,2})\$?(\d+))?")
_REF_LOC = re.compile(r"(?<![A-Za-z0-9_'!.])\$?([A-Z]{1,2})\$?(\d+)(?::\$?([A-Z]{1,2})\$?(\d+))?(?![A-Za-z0-9_(])")


def _celle_da_match(m):
    c1, r1, c2, r2 = m.groups()
    a, b = column_index_from_string(c1), int(r1)
    if c2:
        return [(r, c) for r in range(b, int(r2) + 1) for c in range(a, column_index_from_string(c2) + 1)]
    return [(b, a)]


def celle_log_usate(path):
    """Celle (riga, colonna) del foglio COSTS (LOG) richiamate dalle formule degli altri fogli del modello,
    comprese quelle da cui dipendono (formule interne a COSTS (LOG)). Il resto sono note/valori non usati."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=False)
    usate = set()
    for ws in wb.worksheets:
        if ws.title == "COSTS (LOG)":
            continue
        for row in ws.iter_rows(min_row=1, max_row=40, max_col=120, values_only=True):
            for v in row:
                if isinstance(v, str) and v.startswith("=") and "COSTS (LOG)" in v:
                    for m in _REF_LOG.finditer(v):
                        usate.update(_celle_da_match(m))
    formule = {}
    if "COSTS (LOG)" in wb.sheetnames:
        for r, row in enumerate(wb["COSTS (LOG)"].iter_rows(min_row=1, max_row=90, max_col=14, values_only=True), 1):
            for c, v in enumerate(row, 1):
                if isinstance(v, str) and v.startswith("="):
                    formule[(r, c)] = v
    wb.close()
    da_fare = list(usate)
    while da_fare:
        v = formule.get(da_fare.pop())
        if not v:
            continue
        for m in _REF_LOC.finditer(re.sub(r"'[^']+'![$A-Z0-9:]+", "", v)):   # senza i riferimenti ad altri fogli
            for rc in _celle_da_match(m):
                if rc not in usate:
                    usate.add(rc)
                    da_fare.append(rc)
    return usate


def load(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    out = {}
    for sn in SHEETS:
        if sn in wb.sheetnames:
            out[sn] = Sheet(wb[sn], HEADER_ROWS[sn])
    delist = set()
    if "DELISTATI" in wb.sheetnames:  # Canarie: riga 1 titolo, riga 2 intestazioni, dati da riga 3
        for r in wb["DELISTATI"].iter_rows(min_row=BRANCH["delist_row"], max_col=1, values_only=True):
            if r[0] not in (None, ""):
                delist.add(str(r[0]).strip())
    log_used = celle_log_usate(path)   # celle di COSTS (LOG) davvero usate dalle formule
    log = {}                      # tariffe del foglio COSTS (LOG): servono a capire se e' cambiata una tariffa
    if "COSTS (LOG)" in wb.sheetnames:
        for r, row in enumerate(wb["COSTS (LOG)"].iter_rows(min_row=1, max_row=90, max_col=14, values_only=True), 1):
            for c, v in enumerate(row, 1):
                if v not in (None, ""):
                    log[(r, c)] = v
    for sh in out.values():
        sh.delist = delist
        sh.log = log
        sh.log_used = log_used
    wb.close()
    return out


def stato(so, ro, sn_, rn, code, isl):
    """Perche' un articolo non e' confrontabile. Restituisce (stato, std_a_sistema) o None
    se non c'e' nulla da segnalare (es. isola non servita: STD = 0 in entrambi i file)."""
    c = ISLANDS[isl]
    a, b = std_finale(so, ro, isl), std_finale(sn_, rn, isl)
    if a is not None and b is not None and a == 0 and b == 0:
        return None
    a = None if a == 0 else a   # STD = 0: il modello non ha calcolato nulla (prezzo mancante)
    b = None if b == 0 else b
    bloccato = str(sn_.val(rn, "blocked")).strip() == "1"
    delist = code in sn_.delist
    motivo = "BLOCCATO" if bloccato else ("DELISTATO" if delist else "prezzo/anagrafica mancante")
    if a is not None and b is None:
        return (f"USCITO DAL CALCOLO ({motivo}): a sistema resta lo STD precedente, non verra' aggiornato", a)
    if a is None and b is not None:
        return ("ORA CALCOLABILE (prima senza STD): e' un nuovo STD, non una variazione", None)
    if a is None and b is None:
        return (f"NESSUNO STD CALCOLATO in entrambi i file ({motivo})", None)
    return ("Dati incompleti (Step 1 non numerico): controllare il modello", None)


def comp_value(sh, row, token):
    if token.startswith("aiem#"):
        k = int(token[5:])
        return num(row[sh.aiem_cols[k]]) if len(sh.aiem_cols) > k else None
    if token.endswith("*"):
        return _by_prefix(sh, row, token[:-1])
    return sh.n(row, token)


def std_finale(sh, row, isl):
    """Standard cost vero: Step 2 x (1 + maggiorazione concordata) dove il modello ha la colonna finale
    (da luglio); altrimenti Step 2. E' il valore che la macro carica (New_SC)."""
    c = ISLANDS[isl]
    s2, s2f = sh.n(row, c["s2"]), sh.n(row, c["s2"], occ=1)
    if s2 is None:
        return None
    return s2f if (s2 and s2f) else s2


def components(sh, row, isl):
    """Valori numerici delle componenti per una destinazione. None dove non calcolabile."""
    c = ISLANDS[isl]
    z = lambda v: v or 0.0
    parts = {g: sum(z(comp_value(sh, row, t)) for t in toks) for g, toks in c["parts"].items()}
    s2, s2f = sh.n(row, c["s2"]), sh.n(row, c["s2"], occ=1)
    return dict(price=sh.n(row, _price_name(sh)), parts=parts, s1=sh.n(row, c["s1"]), s2=s2, s2f=s2f,
                f=(s2f / s2) if s2 and s2f else 1.0)   # f = 1,02 dove il modello ha la maggiorazione concordata


def _price_name(sh):
    for n in sh.idx:
        if n.startswith("prezzoacquisto"):
            return n
    return "prezzoacquisto"


def _by_prefix(sh, row, prefix):
    for n, l in sh.idx.items():
        if n.startswith(prefix):
            return num(row[l[0]])
    return None


def drivers(sh, row):
    d = {
        "modo": str(sh.val(row, BRANCH["mode_col"]) or "").strip().upper() if BRANCH["mode_col"] else "",
        "area": str(sh.val(row, BRANCH["area_col"]) or "").strip().upper(),
        "grado": sh.n(row, "gradoalcolico30"),
        "dest": tuple(sh.n(row, n) for n in ("gc", "tf", "fue", "lan")),
        "um": (str(sh.val(row, "salesunitofmeasure") or "").strip().upper(),
               str(sh.val(row, "baseunitofmeasure") or "").strip().upper()),
        "pcsxkg": sh.n(row, "pcsxkg"),
        "loc": str(sh.val(row, "ubicazioneformtsmto") or "").strip().upper(),
        "aiem_rate": sh.n(row, "aiem"),
        "fornitore": str(sh.val(row, "vendorname") or "").strip(),
    }
    d["dati"] = tuple(sh.n(row, n) for n in DATA_DRIVERS)
    return d


def diff(a, b):
    return (b - a) if a is not None and b is not None else None


def _celle(spec):
    out = []
    for x in spec:
        if ":" in x:
            c1, r1, c2, r2 = range_boundaries(x)
            out += [(r, c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)]
        else:
            out.append(coordinate_to_tuple(x))
    return out


def log_changed(sh_o, sh_n, kind):
    """True/False se le tariffe di COSTS (LOG) di quel tipo sono cambiate tra i due modelli; None se non verificabile."""
    lo, ln = getattr(sh_o, "log", {}), getattr(sh_n, "log", {})
    spec = BRANCH.get("log", {}).get(kind)
    if not lo or not ln or not spec:
        return None
    usate = getattr(sh_n, "log_used", None)
    for rc in _celle(spec):
        if usate is not None and rc not in usate:
            continue                  # cella non usata dalle formule: un suo cambio non puo' spiegare nulla
        a, b = lo.get(rc), ln.get(rc)
        if isinstance(a, (int, float)) and isinstance(b, (int, float)):
            if abs(a - b) > 1e-9:
                return True
        elif (a is None) != (b is None) or str(a).strip() != str(b).strip():
            return True
    return False


def _dest_txt(flags):
    nomi = [n for n, v in zip(("GC", "TF", "FUE", "LAN"), flags) if v]
    if not nomi:
        return "nessuna"
    return "tutte e 4" if len(nomi) == 4 else "+".join(nomi)


def explain(sh_old, r_old, sh_new, r_new, isl):
    co, cn = components(sh_old, r_old, isl), components(sh_new, r_new, isl)
    if None in (co["s1"], co["s2"], cn["s1"], cn["s2"]):
        return None
    if co["s2"] == 0 or cn["s2"] == 0:   # STD = 0 significa "senza prezzo/non calcolato": non e' una variazione
        return None
    fo, fn = co["f"], cn["f"]
    do, dn = drivers(sh_old, r_old), drivers(sh_new, r_new)
    mode_changed = (do["modo"], do["area"]) != (dn["modo"], dn["area"])

    def _num_txt(v):
        return "n.d." if v is None else (f"{v:.2f}".rstrip("0").rstrip(".")).replace(".", ",")

    cambi = [f"{DATA_LABELS[n]} {_num_txt(x)} -> {_num_txt(y)}"
             for n, x, y in zip(DATA_DRIVERS, do["dati"], dn["dati"])
             if (x is None) != (y is None) or (x is not None and abs(x - y) > 1e-9)]
    data_changed = bool(cambi)
    dati_txt = "pallettizzazione/dati logistici: " + "; ".join(cambi)
    loc_changed = do["loc"] != dn["loc"]
    rate_changed = do["aiem_rate"] != dn["aiem_rate"]
    grado_changed = do["grado"] != dn["grado"]
    dest_o, dest_n = _dest_txt(do["dest"]), _dest_txt(dn["dest"])
    dest_changed = any(do["dest"]) and any(dn["dest"]) and dest_o != dest_n
    vendor_changed = do["fornitore"] != dn["fornitore"]
    # Cambio di unita' di misura (es. PCS -> KG): il prezzo non e' salito, e' riferito a un'altra unita'.
    # Se si puo' convertire (pezzi per kg), calcolo lo standard nuovo espresso nella vecchia unita'.
    um_info, um_tag = None, None
    if do["um"][0] and dn["um"][0] and do["um"] != dn["um"]:
        ua = do["um"][0] if do["um"][0] != dn["um"][0] else do["um"][1]
        ub = dn["um"][0] if do["um"][0] != dn["um"][0] else dn["um"][1]
        conv = dn["pcsxkg"] or do["pcsxkg"]
        eq = None
        if conv and conv > 0:
            if (ua, ub) == ("PCS", "KG"):
                eq = cn["s2"] * fn / conv
            elif (ua, ub) == ("KG", "PCS"):
                eq = cn["s2"] * fn * conv
        um_info = dict(da=ua, a=ub, eq=eq, delta_eq=None if eq is None else eq - co["s2"] * fo)
        um_tag = (f"cambio unita' di misura {ua} -> {ub}"
                  + (f" (a parita' di unita': {fmt(co['s2'] * fo, False)} -> {fmt(eq, False)} EUR)" if eq is not None else ""))
    tar_tr = log_changed(sh_old, sh_new, "transport")
    tar_pal = log_changed(sh_old, sh_new, "pallet")
    tar_mm = log_changed(sh_old, sh_new, "mtsmto")

    g = {k: 0.0 for k in GROUPS}
    # valori finali (con la maggiorazione concordata): contributi scalati per il fattore del modello nuovo; se il
    # fattore cambia tra i due modelli (es. introduzione del +2% a luglio) la differenza e' una componente a parte
    g["Listino"] = (diff(co["price"], cn["price"]) or 0.0) * fn
    for k in co["parts"]:
        g[k] = (cn["parts"][k] - co["parts"][k]) * fn
    g["MTS/MTO"] = ((cn["s2"] - cn["s1"]) - (co["s2"] - co["s1"])) * fn
    g["Maggiorazione +2%"] = (fn - fo) * co["s2"]
    old_f, new_f = co["s2"] * fo, cn["s2"] * fn
    tot = new_f - old_f
    g["Altro"] = tot - sum(g[k] for k in GROUPS if k != "Altro")

    # cambi "a monte": anagrafica, work tab (resa, modalita', destinazioni), tariffe di COSTS (LOG)
    a_monte = []
    if loc_changed:
        a_monte.append(f"resa {do['loc'] or '-'} -> {dn['loc'] or '-'}")
    if mode_changed:
        a_monte.append(f"modalita'/area {do['modo'] or '-'}/{do['area'] or '-'} -> {dn['modo'] or '-'}/{dn['area'] or '-'}")
    if dest_changed:
        a_monte.append(f"destinazioni {dest_o} -> {dest_n}")
    if data_changed:
        a_monte.append(dati_txt)
    if rate_changed:
        a_monte.append(f"AIEM aliquota {_num_txt(do['aiem_rate'])} -> {_num_txt(dn['aiem_rate'])}")
    if grado_changed:
        a_monte.append(f"grado alcolico {_num_txt(do['grado'])} -> {_num_txt(dn['grado'])}")
    if vendor_changed:
        a_monte.append(f"fornitore {do['fornitore'] or '?'} -> {dn['fornitore'] or '?'}")
    if tar_mm:
        a_monte.append("tariffe Unifreddo (MTS/MTO) cambiate in COSTS (LOG)")
    if tar_tr:
        a_monte.append("tariffe di trasporto cambiate in COSTS (LOG)")
    if tar_pal:
        a_monte.append("costo bancale cambiato in COSTS (LOG)")

    tag = {}
    if vendor_changed:
        tag["Listino"] = f"fornitore {do['fornitore'] or '?'} -> {dn['fornitore'] or '?'}"
    elif loc_changed:
        tag["Listino"] = f"cambio resa {do['loc'] or '-'} -> {dn['loc'] or '-'}"
    else:
        tag["Listino"] = "prezzo di acquisto da listino"
    for gname, kind in BRANCH["kinds"].items():
        if kind == "transport":
            if mode_changed:
                tag[gname] = f"modalita'/area {do['modo']}/{do['area']} -> {dn['modo']}/{dn['area']}"
            elif dest_changed:
                tag[gname] = f"cambio destinazioni {dest_o} -> {dest_n}"
            elif data_changed:
                tag[gname] = dati_txt
            elif tar_tr:
                tag[gname] = "tariffe trasporto cambiate (COSTS LOG)"
            else:
                tag[gname] = "tariffe trasporto"
        elif kind == "pallet":
            if data_changed:
                tag[gname] = dati_txt
            elif dest_changed:
                tag[gname] = f"cambio destinazioni {dest_o} -> {dest_n}"
            elif tar_pal:
                tag[gname] = "costo bancale cambiato (COSTS LOG)"
            else:
                tag[gname] = "tariffa pallet"
        elif kind == "insurance":
            tag[gname] = "su nuovo prezzo/trasporto" if abs(g["Listino"]) > TOL else "tariffa"
        elif kind == "aiem":
            tag[gname] = (f"aliquota {_num_txt(do['aiem_rate'])} -> {_num_txt(dn['aiem_rate'])}" if rate_changed
                          else "effetto sulla nuova base (listino/trasporto)")
        elif kind == "alcool":
            tag[gname] = (f"grado alcolico {_num_txt(do['grado'])} -> {_num_txt(dn['grado'])}" if grado_changed
                          else "effetto sul nuovo prezzo/aliquota")
        else:  # tariff (es. Hong Kong: oneri di import, VGM, certificato)
            if data_changed:
                tag[gname] = dati_txt
            elif tar_tr:
                tag[gname] = "tariffe cambiate (COSTS LOG)"
            else:
                tag[gname] = "tariffe/costi di filiale"
    if loc_changed:
        tag["MTS/MTO"] = f"cambio resa {do['loc'] or '-'} -> {dn['loc'] or '-'}"
    elif tar_mm:
        tag["MTS/MTO"] = "tariffe Unifreddo cambiate (COSTS LOG)"
    else:
        tag["MTS/MTO"] = "tariffe logistica MTS/MTO"
    tag["Maggiorazione +2%"] = f"maggiorazione concordata {fmt((fo - 1) * 100, False)}% -> {fmt((fn - 1) * 100, False)}%"
    tag["Altro"] = "residuo/arrotondamenti"

    if um_info:                      # tutte le componenti sono riferite a un'altra unita': stessa spiegazione per tutte
        a_monte.insert(0, um_tag)
        tag = {k: um_tag for k in tag}
    parts = [(k, v) for k, v in g.items() if abs(v) > TOL]
    parts.sort(key=lambda kv: -abs(kv[1]))
    frase_parts = [f"{k} {fmt(v)} ({tag[k]})" for k, v in parts]
    delta_pct = (tot / old_f * 100) if old_f else None
    return dict(old=old_f, new=new_f, delta=tot, pct=delta_pct, g=g, tag=tag, a_monte=a_monte, um=um_info,
                final_old=co["s2f"], final_new=cn["s2f"], parts=parts, frase_parts=frase_parts)


def sentence(code, desc, isl, e):
    if abs(e["delta"]) <= TOL:
        return f"Art. {code} {desc} [{isl}]: invariato ({fmt(e['new'], False)} EUR)"
    pct = f" ({fmt(e['pct'])}%)" if e["pct"] is not None else ""
    if e.get("um"):
        u = e["um"]
        eq = (f"; a parita' di unita': {fmt(e['old'], False)} -> {fmt(u['eq'], False)} EUR ({fmt(u['delta_eq'])})"
              if u["eq"] is not None else "")
        return (f"Art. {code} {desc} [{isl}]: {fmt(e['old'], False)} -> {fmt(e['new'], False)} EUR per CAMBIO UNITA' DI MISURA "
                f"{u['da']} -> {u['a']} (il listino non e' aumentato: e' riferito a un'altra unita'){eq}")
    body = "; ".join(e["frase_parts"]) if e["frase_parts"] else "variazione sotto soglia"
    monte = f" | a monte: {'; '.join(e['a_monte'])}" if e.get("a_monte") else ""
    return (f"Art. {code} {desc} [{isl}]: {fmt(e['old'], False)} -> {fmt(e['new'], False)} EUR, "
            f"{fmt(e['delta'])}{pct} = {body}{monte}")


def analizza(old, new, articoli, isole, soglia_report):
    """Cuore del confronto (nessuna stampa): usato dalla riga di comando e dall'app web.
    Ritorna (rows_out, stati, notes, msgs); msgs = messaggi per la modalita' --articolo."""
    rows_out, notes, stati, msgs = [], [], [], []
    for sn in SHEETS:
        if sn not in old or sn not in new:
            notes.append(f"Foglio {sn} non presente in entrambi i file: saltato")
            continue
        so, sn_ = old[sn], new[sn]
        if articoli:  # ogni codice puo' essere il No_ della filiale oppure il codice IFB della sede
            keys = []
            for a in articoli:
                if a in sn_.items or a in so.items:
                    keys.append(a)
                else:
                    keys.extend(sn_.ifb.get(a, []) or so.ifb.get(a, []))
        else:
            keys = list(sn_.items.keys())
        for code in keys:
            if code not in sn_.items and code not in so.items:
                continue
            if code not in so.items:
                if articoli:
                    msgs.append(f"Art. {code}: NUOVO articolo (non presente nel file vecchio)")
                else:
                    stati.append((sn, code, str(sn_.val(sn_.items[code], "description") or "")[:40], "",
                                  "NUOVO ARTICOLO nel modello (non c'era nel file vecchio)", {}))
                continue
            if code not in sn_.items:
                ro = so.items[code]
                msg = ("RIGA SPARITA dal modello (delistato/rimosso?): a sistema resta lo STD precedente, "
                       "non verra' aggiornato")
                if articoli:
                    msgs.append(f"Art. {code}: {msg}")
                else:
                    std_old = {i: std_finale(so, ro, i) for i in isole}
                    stati.append((sn, code, str(so.val(ro, "description") or "")[:40],
                                  str(so.val(ro, "vendorname") or ""), msg, std_old))
                continue
            ro, rn = so.items[code], sn_.items[code]
            desc = str(sn_.val(rn, "description") or "")[:40]
            forn = str(sn_.val(rn, "vendorname") or "")
            ifb = str(sn_.val(rn, "ifbitem") or "")
            if ifb and ifb != code and not ifb.startswith("#"):
                desc = f"{desc} (IFB {ifb})"
            in_delist = code in sn_.delist
            isl_stati = []
            for isl in isole:
                e = explain(so, ro, sn_, rn, isl)
                if e is None:
                    st = stato(so, ro, sn_, rn, code, isl)
                    if st:
                        isl_stati.append((isl, st))
                        if articoli:
                            msgs.append(f"Art. {code} {desc} [{isl}]: {st[0]}"
                                  + (f" (STD a sistema: {fmt(st[1], False)} EUR)" if st[1] is not None else ""))
                    elif articoli:
                        msgs.append(f"Art. {code} {desc} [{isl}]: isola non servita (STD = 0)")
                    continue
                if articoli:
                    msgs.append(sentence(code, desc, isl, e) + (" [articolo presente in DELISTATI]" if in_delist else ""))
                if abs(e["delta"]) > soglia_report or articoli:
                    if in_delist:
                        e["frase_parts"].append("ATTENZIONE: articolo presente nel foglio DELISTATI")
                    rows_out.append((sn, code, desc, forn, isl, e))
            if isl_stati and not articoli:
                std_old = {i: std_finale(so, ro, i) for i in isole}
                for testo in dict.fromkeys(s[0] for _, s in isl_stati):
                    isole_txt = [i for i, s in isl_stati if s[0] == testo]
                    stati.append((sn, code, desc, forn,
                                  testo + (f" [isole: {', '.join(isole_txt)}]" if len(isole_txt) < len(isole) else ""),
                                  std_old))
    return rows_out, stati, notes, msgs


def run(old_path, new_path, articoli, isole, out_path, soglia_report):
    print("Leggo", os.path.basename(old_path), "...")
    old = load(old_path)
    print("Leggo", os.path.basename(new_path), "...")
    new = load(new_path)
    rows_out, stati, notes, msgs = analizza(old, new, articoli, isole, soglia_report)
    if articoli:
        for m in msgs:
            print(m)
        return
    if not rows_out and not stati:
        print("Nessuna variazione sopra soglia.")
        return
    write_report(out_path, rows_out, notes, os.path.basename(old_path), os.path.basename(new_path),
                 stati, isole)




def write_report(out_path, rows_out, notes, f_old, f_new, stati=(), isole=()):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Variazioni"
    head = ["Foglio", "Articolo", "Descrizione", "Fornitore", "Isola", "STD vecchio", "STD nuovo",
            "Delta EUR", "Delta %"] + GROUPS + ["Causa principale", "Spiegazione", "Finale +2% nuovo"]
    ws.append(head)
    summary = defaultdict(int)
    for sn, code, desc, forn, isl, e in rows_out:
        main = e["parts"][0][0] if e["parts"] else ""
        if main:
            summary[main + " - " + e["tag"][main]] += 1
        ws.append([sn, code, desc, forn, isl, round(e["old"], 4), round(e["new"], 4),
                   round(e["delta"], 4), None if e["pct"] is None else round(e["pct"] / 100, 4)]
                  + [round(e["g"][k], 4) for k in GROUPS]
                  + [main, sentence(code, desc, isl, e), e["final_new"]])
    for c in range(1, len(head) + 1):
        ws.cell(row=1, column=c).font = openpyxl.styles.Font(bold=True)
    ws.freeze_panes = "F2"
    ws.auto_filter.ref = ws.dimensions
    for r in range(2, ws.max_row + 1):
        ws.cell(row=r, column=9).number_format = "0.0%"
    for col, w in zip("ABCDEFGHIJKLMNOPQRS", [18, 10, 36, 28, 7, 11, 11, 10, 9, 10, 10, 13, 9, 9, 9, 9, 22, 90, 12]):
        ws.column_dimensions[col].width = w
    w2 = wb.create_sheet("Riepilogo")
    w2.append(["File vecchio", f_old])
    w2.append(["File nuovo", f_new])
    w2.append([])
    w2.append(["Causa principale", "N. casi (articolo x isola)"])
    for k, v in sorted(summary.items(), key=lambda kv: -kv[1]):
        w2.append([k, v])
    for n in notes:
        w2.append([n])
    w2.column_dimensions["A"].width = 70
    # foglio "Non aggiornati": articoli che escono dal calcolo / bloccati / delistati ecc.
    w3 = wb.create_sheet("Non aggiornati")
    w3.append(["Foglio", "Articolo", "Descrizione", "Fornitore", "Stato"]
              + [f"STD precedente {i}" for i in isole])
    for c in range(1, 6 + len(isole)):
        w3.cell(row=1, column=c).font = openpyxl.styles.Font(bold=True)
    ordine = lambda s: (0 if s[4].startswith(("USCITO", "RIGA SPARITA")) else 1 if s[4].startswith("ORA") else 2, s[1])
    cont = defaultdict(int)
    for sn, code, desc, forn, testo, std_old in sorted(stati, key=ordine):
        w3.append([sn, code, desc, forn, testo] + [std_old.get(i) for i in isole])
        cont[testo.split(" (")[0].split(" [")[0]] += 1
    w3.freeze_panes = "C2"
    w3.auto_filter.ref = w3.dimensions
    for col, w in zip("ABCDE", [18, 10, 38, 28, 100]):
        w3.column_dimensions[col].width = w
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    wb.save(out_path)
    print(f"\nReport salvato: {out_path}")
    print(f"Casi con variazione: {len(rows_out)}")
    print("Cause principali:")
    for k, v in sorted(summary.items(), key=lambda kv: -kv[1])[:8]:
        print(f"  {v:5d}  {k}")
    if cont:
        print("Articoli non confrontabili / non aggiornati (foglio 'Non aggiornati'):")
        for k, v in sorted(cont.items(), key=lambda kv: -kv[1]):
            print(f"  {v:5d}  {k}")


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("vecchio")
    ap.add_argument("nuovo")
    ap.add_argument("--filiale", choices=list(BRANCHES), default="CANARIE", help="filiale (default CANARIE)")
    ap.add_argument("--articolo", nargs="*", help="codice/i articolo (colonna No_)")
    ap.add_argument("--isola", nargs="*", help="destinazione/i (default tutte)")
    ap.add_argument("--soglia", type=float, default=TOL, help="minima variazione in EUR per il report")
    ap.add_argument("--out", help="file Excel di output (default: cartella OUTPUT)")
    a = ap.parse_args()
    set_branch(a.filiale)
    here = os.path.dirname(os.path.abspath(__file__))
    out = a.out or os.path.join(here, "OUTPUT", "Spiegazione_variazioni_"
                                + os.path.splitext(os.path.basename(a.vecchio))[0][:30].replace(" ", "_")
                                + "_vs_" + os.path.splitext(os.path.basename(a.nuovo))[0][:30].replace(" ", "_") + ".xlsx")
    a.isola = a.isola or list(ISLANDS)
    run(a.vecchio, a.nuovo, [str(x) for x in a.articolo] if a.articolo else None, a.isola, out, a.soglia)


if __name__ == "__main__":
    main()
