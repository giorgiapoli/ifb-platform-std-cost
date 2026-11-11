"""
Memoria dei file caricati: l'ultimo file di ogni tipo (e di ogni filiale) resta salvato sul PC, cosi' ricaricando la pagina o
riaprendo l'app non si deve caricare di nuovo. Si toglie con il pulsante «Togli».
"""
import os

DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ULTIMI_CARICATI")


class FileTenuto:
    """Si comporta come un file caricato: .name e .getvalue()."""

    def __init__(self, nome, dati):
        self.name, self._d = nome, dati

    def getvalue(self):
        return self._d


def salva(chiave, f):
    os.makedirs(DIR, exist_ok=True)
    with open(os.path.join(DIR, chiave + ".bin"), "wb") as fh:
        fh.write(f.getvalue())
    with open(os.path.join(DIR, chiave + ".txt"), "w", encoding="utf-8") as fh:
        fh.write(getattr(f, "name", chiave))


def carica(chiave):
    p = os.path.join(DIR, chiave + ".bin")
    if not os.path.exists(p):
        return None
    try:
        nome = open(os.path.join(DIR, chiave + ".txt"), encoding="utf-8").read()
    except OSError:
        nome = chiave
    with open(p, "rb") as fh:
        return FileTenuto(nome, fh.read())


def togli(chiave):
    for est in (".bin", ".txt"):
        p = os.path.join(DIR, chiave + est)
        if os.path.exists(p):
            os.remove(p)
