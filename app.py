import streamlit as st
import pandas as pd
import pdfplumber
from docx import Document
from io import BytesIO
import datetime

# --- CONFIGURAZIONE ---
st.set_page_config(page_title="FlussoFacile 2026 - FIX", layout="wide")

# --- FUNZIONE DI PULIZIA E RICERCA NUMERI ---
def estrai_numeri_da_riga(riga):
    """Cerca tutti i valori numerici in una riga e li restituisce come lista di float."""
    numeri = []
    for cella in riga:
        if cella:
            # Pulizia: togliamo simboli e punti delle migliaia
            s = str(cella).replace('€', '').replace('.', '').replace(',', '.').replace('\n', '').strip()
            try:
                val = float(s)
                if val != 0: numeri.append(val)
            except:
                continue
    return numeri

# --- MOTORE DI ANALISI DINAMICO ---
def analizza_pdf_intelligente(file_pdf, perc):
    rows = []
    with pdfplumber.open(file_pdf) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if table:
                for r in table:
                    # Una riga è valida se ha un codice (r[0]) e dei numeri
                    valori = estrai_numeri_da_riga(r)
                    if r[0] and len(valori) >= 2:
                        # LOGICA: Di solito l'ultimo numero è il riscosso/pagato, 
                        # il penultimo (o quello più grande) è il budget totale.
                        budget = max(valori)
                        riscosso = valori[-1] if valori[-1] < budget else valori[0]
                        if len(valori) == 1: riscosso = 0 # Se c'è solo un numero, non è stato ancora mosso nulla
                        
                        # Se è una riga di intestazione (es. "2026"), la saltiamo
                        if r[0].isdigit() and len(r[0]) == 4 and budget > 2026: 
                             pass # Anno del residuo, ok
                        
                        differenza = budget - riscosso
                        prev_1 = riscosso + (differenza * (perc/100))
                        prev_2 = differenza * (1 - (perc/100))
                        
                        rows.append([r[0], r[1][:50] if r[1] else "Voce contabile", budget, riscosso, prev_1, prev_2])
    return pd.DataFrame(rows, columns=['Codice', 'Descrizione', 'Budget', 'Già Mosso', 'Prev. Gen-Ago', 'Prev. Set-Dic'])

# --- (Il resto dell'interfaccia rimane uguale, usa analizza_pdf_intelligente) ---