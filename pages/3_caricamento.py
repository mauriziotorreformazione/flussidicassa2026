import streamlit as st
import base64

st.title("📄 Caricamento PDF e Modalità")
st.markdown("---")

# Initialize session state
if "modello_h_bytes" not in st.session_state:
    st.session_state.modello_h_bytes = None
if "modello_l_bytes" not in st.session_state:
    st.session_state.modello_l_bytes = None
if "modalita" not in st.session_state:
    st.session_state.modalita = "automatica"
if "percentuali" not in st.session_state:
    st.session_state.percentuali = {
        "entrate_competenza_gen_ago": 100,
        "entrate_competenza_set_dic": 0,
        "spese_competenza_gen_ago": 67,
        "spese_competenza_set_dic": 33,
        "residui_attivi_gen_ago": 33,
        "residui_attivi_set_dic": 67,
        "residui_passivi_gen_ago": 33,
        "residui_passivi_set_dic": 67,
    }
if "dati_estratti" not in st.session_state:
    st.session_state.dati_estratti = None

# ─────────────────────────────────────────────
st.markdown("### 1️⃣ Carica i PDF")
col1, col2 = st.columns(2)

with col1:
    st.markdown("**Modello H+ — Conto Finanziario**")
    st.caption("Contiene entrate e spese in competenza (Programmazione Definitiva, Somme Riscosse/Pagate)")
    file_h = st.file_uploader("Carica Modello H", type=["pdf"], key="upload_h")
    if file_h:
        st.session_state.modello_h_bytes = file_h.read()
        st.success(f"✅ {file_h.name} caricato ({len(st.session_state.modello_h_bytes)//1024} KB)")

with col2:
    st.markdown("**Modello L — Elenco Residui**")
    st.caption("Contiene residui attivi (crediti) e residui passivi (debiti) da anni precedenti")
    file_l = st.file_uploader("Carica Modello L", type=["pdf"], key="upload_l")
    if file_l:
        st.session_state.modello_l_bytes = file_l.read()
        st.success(f"✅ {file_l.name} caricato ({len(st.session_state.modello_l_bytes)//1024} KB)")

st.markdown("---")

# ─────────────────────────────────────────────
st.markdown("### 2️⃣ Scegli la modalità di compilazione")

col1, col2 = st.columns(2)
with col1:
    if st.button("🤖 Modalità Automatica", use_container_width=True,
                 type="primary" if st.session_state.modalita == "automatica" else "secondary"):
        st.session_state.modalita = "automatica"
        st.rerun()
with col2:
    if st.button("📝 Modalità Manuale", use_container_width=True,
                 type="primary" if st.session_state.modalita == "manuale" else "secondary"):
        st.session_state.modalita = "manuale"
        st.rerun()

if st.session_state.modalita == "automatica":
    st.info("🤖 **Automatica**: le percentuali vengono applicate automaticamente. Puoi modificarle qui sotto.")
else:
    st.info("📝 **Manuale**: riceverai il file Excel con i valori già riscossi/pagati precompilati. Inserisci tu le previsioni.")

st.markdown("---")

# ─────────────────────────────────────────────
if st.session_state.modalita == "automatica":
    st.markdown("### 3️⃣ Percentuali di ripartizione")
    st.caption("Modifica le percentuali se necessario. La somma Gen-Ago + Set-Dic deve essere ≤ 100% per ogni categoria.")

    p = st.session_state.percentuali

    def percentuale_row(label, key_ga, key_sd, help_text=""):
        col1, col2, col3 = st.columns([3, 1.5, 1.5])
        with col1:
            st.markdown(f"**{label}**")
            if help_text:
                st.caption(help_text)
        with col2:
            val_ga = st.number_input(
                "Gen-Ago %", min_value=0, max_value=100,
                value=p[key_ga], key=f"{key_ga}_input",
                label_visibility="collapsed"
            )
        with col3:
            val_sd = st.number_input(
                "Set-Dic %", min_value=0, max_value=100,
                value=p[key_sd], key=f"{key_sd}_input",
                label_visibility="collapsed"
            )
        # Validate
        if val_ga + val_sd > 100:
            st.error(f"⚠️ {label}: Gen-Ago + Set-Dic = {val_ga+val_sd}% (max 100%)")
        elif val_ga + val_sd < 100:
            st.warning(f"ℹ️ {label}: residuo non ripartito = {100-val_ga-val_sd}%")
        else:
            st.empty()
        p[key_ga] = val_ga
        p[key_sd] = val_sd

    # Header
    col1, col2, col3 = st.columns([3, 1.5, 1.5])
    with col2:
        st.markdown("**Gen-Ago %**")
    with col3:
        st.markdown("**Set-Dic %**")

    st.markdown("**📥 ENTRATE**")
    percentuale_row(
        "Entrate in competenza",
        "entrate_competenza_gen_ago", "entrate_competenza_set_dic",
        "Base: Programmaz. Definitiva − Già riscosso"
    )
    percentuale_row(
        "Residui attivi (da Modello L)",
        "residui_attivi_gen_ago", "residui_attivi_set_dic",
        "Base: Importo residuo"
    )

    st.markdown("**📤 SPESE**")
    percentuale_row(
        "Spese in competenza",
        "spese_competenza_gen_ago", "spese_competenza_set_dic",
        "Base: Programmaz. Definitiva − Già pagato"
    )
    percentuale_row(
        "Residui passivi (da Modello L)",
        "residui_passivi_gen_ago", "residui_passivi_set_dic",
        "Base: Importo residuo"
    )

    st.session_state.percentuali = p

    if st.button("↩️ Ripristina percentuali predefinite"):
        st.session_state.percentuali = {
            "entrate_competenza_gen_ago": 100, "entrate_competenza_set_dic": 0,
            "spese_competenza_gen_ago": 67, "spese_competenza_set_dic": 33,
            "residui_attivi_gen_ago": 33, "residui_attivi_set_dic": 67,
            "residui_passivi_gen_ago": 33, "residui_passivi_set_dic": 67,
        }
        st.rerun()

    st.markdown("---")

# ─────────────────────────────────────────────
st.markdown("### 4️⃣ Estrai i dati dai PDF")

pdf_pronti = st.session_state.modello_h_bytes and st.session_state.modello_l_bytes

if not pdf_pronti:
    st.warning("⚠️ Carica entrambi i PDF prima di procedere.")
else:
    if st.button("🔍 Estrai dati con AI", use_container_width=True, type="primary"):
        with st.spinner("Estrazione in corso... (può richiedere 30-60 secondi)"):
            try:
                from utils.extractor import estrai_dati_pdf
                dati = estrai_dati_pdf(
                    st.session_state.modello_h_bytes,
                    st.session_state.modello_l_bytes
                )
                st.session_state.dati_estratti = dati

                # Auto-populate school data if extracted
                if "dati_scuola" not in st.session_state:
                    st.session_state.dati_scuola = {}
                ds = st.session_state.dati_scuola
                if dati.get("nome_istituto") and not ds.get("nome_istituto"):
                    ds["nome_istituto"] = dati["nome_istituto"]
                if dati.get("cf") and not ds.get("cf"):
                    ds["cf"] = dati["cf"]
                if dati.get("codice_mecc") and not ds.get("codice_mecc"):
                    ds["codice_mecc"] = dati["codice_mecc"]
                if dati.get("indirizzo") and not ds.get("indirizzo"):
                    ds["indirizzo"] = dati["indirizzo"]
                if dati.get("citta") and not ds.get("citta"):
                    ds["citta"] = dati["citta"]

                st.success("✅ Estrazione completata!")
                st.rerun()

            except Exception as e:
                st.error(f"❌ Errore durante l'estrazione: {str(e)}")

    if st.session_state.dati_estratti:
        dati = st.session_state.dati_estratti
        st.success("✅ Dati estratti — anteprima:")

        tab1, tab2, tab3, tab4 = st.tabs(["📥 Entrate", "📤 Spese", "📋 Residui Attivi", "📋 Residui Passivi"])

        with tab1:
            if dati.get("entrate"):
                import pandas as pd
                df = pd.DataFrame(dati["entrate"])
                st.dataframe(df, use_container_width=True)
            else:
                st.info("Nessuna voce di entrata estratta.")

        with tab2:
            if dati.get("spese"):
                import pandas as pd
                df = pd.DataFrame(dati["spese"])
                st.dataframe(df, use_container_width=True)
            else:
                st.info("Nessuna voce di spesa estratta.")

        with tab3:
            if dati.get("residui_attivi"):
                import pandas as pd
                df = pd.DataFrame(dati["residui_attivi"])
                st.dataframe(df, use_container_width=True)
            else:
                st.info("Nessun residuo attivo estratto.")

        with tab4:
            if dati.get("residui_passivi"):
                import pandas as pd
                df = pd.DataFrame(dati["residui_passivi"])
                st.dataframe(df, use_container_width=True)
            else:
                st.info("Nessun residuo passivo estratto.")

        st.success("✅ Tutto pronto! Procedi con **📊 Genera Documenti**.")
