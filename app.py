import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
import datetime

# --- CONFIGURAZIONE ---
st.set_page_config(page_title="FlussoFacile 2026", layout="wide")

with st.sidebar:
    st.title("📌 Area Tecnica")
    st.header("A cura di: [IL TUO NOME]") 
    st.markdown("---")
    
    # FONDO CASSA OBBLIGATORIO
    fondo_cassa = st.number_input("1. Inserisci Fondo Cassa al 01/01/2026 (€)", min_value=0.0, step=100.0, format="%.2f")
    data_cut_off = st.date_input("2. Data Situazione Contabile", datetime.date(2026, 3, 7))
    
    st.markdown("---")
    modalita = st.radio("3. Modalità di compilazione:", ["Automatica", "Manuale"])
    
    # PERCENTUALI SCELTE DALL'UTENTE
    p_e, p_s, p_r = 100, 66, 33 # Valori di default
    if modalita == "Automatica":
        st.subheader("Parametri Automazione")
        p_e = st.slider("% Entrate (Gen-Ago)", 0, 100, 100)
        p_s = st.slider("% Spese (Gen-Ago)", 0, 100, 66)
        p_r = st.slider("% Residui (Gen-Ago)", 0, 100, 33)

# --- FUNZIONI ---
def genera_testo_nota(mod, pe, ps, pr, data_str):
    testo = f"RELAZIONE TECNICA ILLUSTRATIVA AL PIANO DEI FLUSSI\nSituazione al: {data_str}\n\n"
    if mod == "Automatica":
        testo += (f"CRITERI DI COMPILAZIONE (MODALITÀ AUTOMATICA):\n"
                  f"Per la quota residua, è stata adottata una ripartizione prudenziale basata sullo storico degli anni precedenti: "
                  f"Entrate {pe}%, Spese {ps}%, Residui {pr}%.\n")
    else:
        testo += "CRITERI DI COMPILAZIONE (MODALITÀ MANUALE): Analisi puntuale delle singole voci.\n"
    return testo

def genera_word_doc(titolo, contenuto):
    doc = Document()
    doc.add_heading(titolo, 0)
    doc.add_paragraph(contenuto)
    doc.add_paragraph("\n\nIl Direttore dei Servizi Generali e Amministrativi\n__________________________")
    buffer = BytesIO()
    doc.save(buffer)
    return buffer.getvalue()

# --- INTERFACCIA CENTRALE ---
st.title("🚀 FlussoFacile 2026")
st.write("Generatore basato su Modelli H e L.")

file_h = st.file_uploader("Carica Modello H (Competenza)", type="pdf") #
file_l = st.file_uploader("Carica Modello L (Residui)", type="pdf") #

if file_h and file_l:
    # CONTROLLO FONDO CASSA
    if fondo_cassa <= 0:
        st.warning("⚠️ Per procedere, inserisci l'importo del Fondo Cassa nella barra laterale.")
    else:
        st.success("✅ Documenti pronti per il download!")
        data_str = data_cut_off.strftime("%d/%m/%Y")
        
        col1, col2 = st.columns(2)
        with col1:
            doc_decreto = genera_word_doc("Decreto di Adozione", "VISTO il D.I. 129/2018...") #
            st.download_button("📜 Scarica Decreto DS", doc_decreto, file_name="Decreto_Adozione.docx")
            
        with col2:
            nota_testo = genera_testo_nota(modalita, p_e, p_s, p_r, data_str)
            doc_nota = genera_word_doc("Nota Integrativa", nota_testo)
            st.download_button("📑 Scarica Nota per Revisori", doc_nota, file_name="Nota_Integrativa.docx")

st.markdown('<div style="text-align: center; color: gray; margin-top: 50px;">Servizio tecnico a cura di [IL TUO NOME]</div>', unsafe_allow_html=True)