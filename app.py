import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
import datetime

# --- CONFIGURAZIONE ---
st.set_page_config(page_title="FlussoFacile 2026", layout="wide")

# Branding solo nell'app
with st.sidebar:
    st.title("📌 Assistenza Tecnica")
    st.header("A cura di: [IL TUO NOME]") 
    st.write("Specialista Supporto Contabile")
    st.markdown("---")
    fondo_cassa = st.number_input("Fondo Cassa al 01/01/2026 (€)", min_value=0.0)
    data_cut_off = st.date_input("Data Situazione Contabile", datetime.date(2026, 3, 7))
    modalita = st.radio("Modalità di compilazione:", ["Automatica", "Manuale"])

# --- LOGICA NOTE INTEGRATIVE ---
def genera_testo_nota(modalita, p_e, p_s, p_r, data_str):
    testo = f"RELAZIONE TECNICA ILLUSTRATIVA AL PIANO DEI FLUSSI\nSituazione al: {data_str}\n\n"
    if modalita == "Automatica":
        testo += (f"CRITERI DI COMPILAZIONE (MODALITÀ AUTOMATICA):\n"
                  f"Per la quota residua non ancora movimentata, è stata adottata una metodologia di stima basata su algoritmi "
                  f"di ripartizione proporzionale (Entrate {p_e}%, Spese {p_s}%, Residui {p_r}%). "
                  f"Tali percentuali sono state determinate a seguito di un'attenta analisi dello storico riferito agli anni precedenti.\n")
    else:
        testo += ("CRITERI DI COMPILAZIONE (MODALITÀ MANUALE):\n"
                  "Il Piano è frutto di un'analisi puntuale effettuata su ogni singola voce di entrata e di spesa.\n")
    testo += ("\nMETODOLOGIA DI RACCORDO DELLE USCITE:\n"
              "Si è scelto di adottare un raccordo basato sugli Aggregati di Bilancio (A, P, G, Z).")
    return testo

def genera_word_doc(titolo, contenuto):
    doc = Document()
    doc.add_heading(titolo, 0)
    doc.add_paragraph(contenuto)
    doc.add_paragraph("\n\nIl Direttore dei Servizi Generali e Amministrativi\n__________________________")
    buffer = BytesIO()
    doc.save(buffer)
    return buffer.getvalue()

# --- INTERFACCIA ---
st.title("🚀 FlussoFacile 2026")
st.write("Generatore automatico basato su Modelli H e L.") #

file_h = st.file_uploader("Carica Modello H (Competenza)", type="pdf") #
file_l = st.file_uploader("Carica Modello L (Residui)", type="pdf") #

if file_h and file_l:
    st.success("Documenti caricati. Il sistema sta elaborando i dati.")
    data_str = data_cut_off.strftime("%d/%m/%Y")
    col1, col2 = st.columns(2)
    with col1:
        testo_dec = "VISTO il D.I. 129/2018... DECRETA l'adozione del Piano dei Flussi 2026." #
        doc_decreto = genera_word_doc("Decreto di Adozione", testo_dec)
        st.download_button("📜 Scarica Decreto DS", doc_decreto, file_name="Decreto_Adozione.docx")
    with col2:
        nota_testo = genera_testo_nota(modalita, 100, 66, 33, data_str)
        doc_nota = genera_word_doc("Nota Integrativa", nota_testo)
        st.download_button("📑 Scarica Nota per Revisori", doc_nota, file_name="Nota_Integrativa.docx")

st.markdown('<div style="text-align: center; color: gray; margin-top: 50px;">Servizio tecnico a cura di [IL TUO NOME]</div>', unsafe_allow_html=True)