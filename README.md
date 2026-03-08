# Piano Annuale dei Flussi di Cassa — Scuole Italiane

App Streamlit per la generazione del Piano Annuale dei Flussi di Cassa (art. 6, D.L. 155/2024).

## Funzionalità

- Caricamento PDF Modello H+ e Modello L
- Estrazione dati tramite AI (Claude Sonnet)
- Modalità automatica (percentuali configurabili) e manuale
- Generazione Excel formato ministeriale (3 fogli: Entrate, Spese, Piano Flussi MIM)
- Generazione Nota di accompagnamento (Word)
- Generazione Decreto di adozione DS (Word)

## Setup su Streamlit Cloud

1. Fork o carica questo repository su GitHub
2. Crea una nuova app su [share.streamlit.io](https://share.streamlit.io)
3. Collega il repository GitHub
4. Imposta il file principale: `app.py`
5. Nella sezione **Secrets**, aggiungi:

```toml
ANTHROPIC_API_KEY = "sk-ant-..."
```

## Setup locale

```bash
pip install -r requirements.txt
# Crea .streamlit/secrets.toml con la chiave API
streamlit run app.py
```

## Struttura

```
app.py                  # Entry point
pages/
  1_istruzioni.py       # Pagina istruzioni
  2_dati_scuola.py      # Dati istituto
  3_caricamento.py      # Upload PDF + config percentuali
  4_genera.py           # Genera e scarica documenti
utils/
  extractor.py          # Estrazione PDF via Claude API
  excel_generator.py    # Generazione Excel (openpyxl)
  doc_generator.py      # Generazione Word (python-docx)
requirements.txt
```

## Normativa di riferimento

- Art. 6, D.L. 22 ottobre 2024, n. 155 (conv. L. 189/2024)
- D.I. n. 129 del 28/08/2018 — Regolamento gestione amm.-contabile istituzioni scolastiche
- Nota MIM n. 2284/2025 — Piano dei Flussi di Cassa (formato semplificato)
