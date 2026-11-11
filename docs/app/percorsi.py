"""Dove l'app salva i dati caricati e le impostazioni (mesi in memoria, file caricati, filiale scelta, tassi di cambio).
Sul PC e' la cartella dell'app; nella versione web (browser) e' una cartella che il browser conserva tra una visita e l'altra."""
import os

DATI = os.environ.get("STDC_DATI") or os.path.dirname(os.path.abspath(__file__))
