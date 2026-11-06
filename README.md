# Standard Cost Canarie

App Streamlit per spiegare le variazioni dello standard cost (Canarie), controllare i margini COMIT STD vs FIFO e consultare lo storico giornaliero dello standard.

I dati non sono nel repository: l'app li legge in locale dai file Excel caricati dall'utente (cartelle `INPUT`, `MARGINI_MESI`, `STORICO_STD`, create/riempite in locale e ignorate da git).

## Avvio

```
pip install -r requirements.txt
streamlit run app.py
```

Su Windows: `avvia_app.ps1` avvia l'app sulla porta 8777 e la apre nel browser, `ferma_app.ps1` la ferma.
