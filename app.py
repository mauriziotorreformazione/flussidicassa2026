import streamlit as st

st.set_page_config(
    page_title="Piano Annuale dei Flussi di Cassa",
    page_icon="🏫",
    layout="wide",
    initial_sidebar_state="expanded"
)

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

# Route to correct page
page = st.session_state.page

if page == "📖 Istruzioni":
    exec(open("pages/1_istruzioni.py").read())
elif page == "🏫 Dati Scuola":
    exec(open("pages/2_dati_scuola.py").read())
elif page == "📄 Caricamento PDF":
    exec(open("pages/3_caricamento.py").read())
elif page == "📊 Genera Documenti":
    exec(open("pages/4_genera.py").read())
