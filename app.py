import streamlit as st
import pandas as pd
import pdfplumber
from docx import Document
from io import BytesIO
import datetime
import re

# --- CONFIGURAZIONE ---
st.set_page_config(page_title="FlussoFacile 2026 - Versione Integrale", layout="wide")

with st.sidebar:
    st.title("📌 Area Tecnica")
    st.header("A cura di: [IL TUO NOME]") 
    st.markdown("---")
    fondo_cassa = st.number_input("1. Fondo Cassa al 01/01/2026 (€)", min_value=0.0, format="%.2f")
    data_cut_off = st.date_input("2. Data Situazione Contabile", datetime.date(2026, 3, 7))
    modalita = st.radio("3. Modalità:", ["Automatica", "Manuale"])
    
    p_e, p_s, p_r = 100, 66, 33 
    if modalita == "Automatica":
        p_e = st.slider("% Entrate (Gen-Ago)", 0, 100, 100)
        p_s = st.slider("% Spese (Gen-Ago)", 0, 100, 66)
        p_r = st.slider("% Residui (Gen-Ago)", 0, 100, 33)

# --- FUNZIONI DI PULIZIA DATI ---
def pulisci_numero(testo):
    """Trasforma '1.250,50 €' in 1250.50"""
    if not testo: return 0.0
    s = str(testo).replace('€', '').replace('.', '').replace(',', '.').strip()
    try: return float(s)
    except: return 0.0

# --- MOTORE DI ESTRAZIONE E CALCOLO ---
def analizza_e_calcola(file_pdf, tipo, perc):
    rows = []
    with pdfplumber.open(file_pdf) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if table:
                for r in table:
                    # Filtriamo le righe che hanno un codice contabile (es. A01 o 01/02)
                    if r[0] and (len(r[0]) >= 2):
                        descr = r[1] if len(r) > 1 else ""
                        budget = pulisci_numero(r[2]) if len(r) > 2 else 0.0
                        riscosso_pagato = pulisci_numero(r[3]) if len(r) > 3 else 0.0
                        
                        differenza = budget - riscosso_pagato
                        prev_1 = riscosso_pagato + (differenza * (perc/100))
                        prev_2 = differenza * (1 - (perc/100))
                        
                        rows.append([r[0], descr, budget, riscosso_pagato, prev_1, prev_2])
    return pd.DataFrame(rows, columns=['Codice', 'Descrizione', 'Budget', 'Già Mosso', 'Previsione Gen-Ago', 'Previsione Set-Dic'])

def genera_excel_completo(df_h, df_l, f_cassa):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Foglio Entrate e Spese Competenza
        df_h.to_excel(writer, sheet_name='COMPETENZA (Mod. H)', index=False)
        # Foglio Residui
        df_l.to_excel(writer, sheet_name='RESIDUI (Mod. L)', index=False)
        
        # Foglio di Riepilogo Saldi
        tot_entr = df_h['Previsione Gen-Ago'].sum()
        tot_uscit = df_h['Previsione Set-Dic'].sum()
        df_riepilogo = pd.DataFrame({
            'Voce': ['Fondo Cassa Iniziale', 'Totale Entrate Previste', 'Totale Uscite Previste', 'Saldo Finale Stimato'],
            'Valore (€)': [f_cassa, tot_entr, tot_uscit, f_cassa + tot_entr - tot_uscit]
        })
        df_riepilogo.to_excel(writer, sheet_name='RIEPILOGO FINALE', index=False)
    return output.getvalue()

# --- INTERFACCIA ---
st.title("🚀 FlussoFacile 2026 - Kit Completo")

file_h = st.file_uploader("Carica Modello H (Competenza)", type="pdf")
file_l = st.file_uploader("Carica Modello L (Residui)", type="pdf")

if file_h and file_l:
    if fondo_cassa <= 0:
        st.warning("⚠️ Inserisci il Fondo Cassa per attivare i calcoli reali.")
    else:
        # Calcolo Reale
        df_h_calc = analizza_e_calcola(file_h, "H", p_s)
        df_l_calc = analizza_e_calcola(file_l, "L", p_r)
        
        st.success(f"Analisi completata con successo! Generati i flussi per {len(df_h_calc) + len(df_l_calc)} voci contabili.")
        
        c1, c2, c3 = st.columns(3)
        with c1:
            excel_data = genera_excel_completo(df_h_calc, df_l_calc, fondo_cassa)
            st.download_button("📊 Scarica Excel con DATI REALI", excel_data, file_name="Piano_Flussi_REALE.xlsx")
        # ... (Pulsanti Word rimangono come prima)