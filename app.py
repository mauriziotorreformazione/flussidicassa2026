import streamlit as st
import os

st.set_page_config(
    page_title="Piano Annuale dei Flussi di Cassa",
    page_icon="🏫",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Base directory — works both locally and on Streamlit Cloud
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Sidebar navigation
st.sidebar.title("📋 Navigazione")
st.sidebar.markdown("---")

pages = {
    "📖 Istruzioni": "pages/1_istruzioni.py",
    "🏫 Dati Scuola": "pages/2_dati_scuola.py",
    "📄 Caricamento PDF": "pages/3_caricamento.py",
    "📊 Genera Documenti": "pages/4_genera.py",
}

# Initialize session state
if "page" not in st.session_state:
    st.session_state.page = "📖 Istruzioni"

for label in pages:
    if st.sidebar.button(label, use_container_width=True,
                         type="primary" if st.session_state.page == label else "secondary"):
        st.session_state.page = label
        st.rerun()

st.sidebar.markdown("---")
st.sidebar.caption("Piano Flussi di Cassa · Scuole · 2026")

# Route to correct page using absolute paths
page = st.session_state.page
page_file = os.path.join(BASE_DIR, pages[page])

with open(page_file, encoding="utf-8") as f:
    exec(f.read(), {"__file__": page_file})
