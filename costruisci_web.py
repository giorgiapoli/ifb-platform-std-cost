"""Prepara la versione web (stlite): copia i file dell'app in docs/ e crea docs/index.html.
Uso: python costruisci_web.py   ->  poi pubblicare la cartella docs con GitHub Pages."""
import glob
import json
import os
import shutil

QUI = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(QUI, "docs")
STLITE = "1.12.1"
os.makedirs(os.path.join(OUT, "app"), exist_ok=True)
files = {}
for f in sorted(glob.glob(os.path.join(QUI, "*.py"))):
    n = os.path.basename(f)
    if n == "costruisci_web.py":
        continue
    shutil.copy(f, os.path.join(OUT, "app", n))
    files[n] = {"url": "./app/" + n}
for sub in ("assets", ".streamlit"):
    for f in glob.glob(os.path.join(QUI, sub, "*")):
        d = os.path.join(OUT, "app", sub)
        os.makedirs(d, exist_ok=True)
        shutil.copy(f, d)
        files[f"{sub}/{os.path.basename(f)}"] = {"url": f"./app/{sub}/{os.path.basename(f)}"}
html = f"""<!doctype html>
<html lang="it"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Standard Cost</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/@stlite/browser@{STLITE}/build/style.css">
</head><body>
<div id="root"></div>
<script type="module">
import {{ mount }} from "https://cdn.jsdelivr.net/npm/@stlite/browser@{STLITE}/build/stlite.js";
mount({{
  requirements: ["openpyxl", "xlsxwriter", "reportlab", "xlrd"],
  idbfsMountpoints: ["/home/pyodide/dati"],
  entrypoint: "app.py",
  files: {json.dumps(files)}
}}, document.getElementById("root"));
</script></body></html>"""
open(os.path.join(OUT, "index.html"), "w", encoding="utf-8").write(html)
print("ok", len(files), "file")
