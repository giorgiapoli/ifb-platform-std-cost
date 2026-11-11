"""
Punto di ingresso dell'app. Il motore di calcolo tiene la filiale scelta in variabili condivise da tutto il programma: se due
pagine aperte insieme (per esempio una su Canarie e una su Hong Kong) girano nello stesso istante, si calpestano e compare un
errore (KeyError). Per questo le esecuzioni dell'app si mettono in fila, una alla volta; il codice vero e' in app_main.py.
"""
import os
import runpy
import sys
import threading

_serratura = getattr(sys, "_stdc_serratura", None)
if _serratura is None:
    _serratura = sys._stdc_serratura = threading.RLock()

with _serratura:
    runpy.run_path(os.path.join(os.path.dirname(os.path.abspath(__file__)), "app_main.py"), run_name="__main__")
