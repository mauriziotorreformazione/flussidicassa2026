import streamlit as st
import base64

st.title("📄 Caricamento PDF e Modalità")
st.markdown("---")

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

PERCENTUALI_DEFAULT = {
    "entrate_competenza_gen_ago": 100,
    "entrate_competenza_set_dic": 0,
    "spese_competenza_gen_ago": 67,
    "spese_competenza_set_dic": 33,
    "residui_attivi_gen_ago": 33,
    "residui_attivi_set_dic": 67,
    "residui_passivi_gen_ago": 33,
    "residui_passivi_set_dic": 67,
}

# ── SEZIONE 1: Upload PDF ──────────────────────────────
st.markdown("### 1️⃣ Carica i PDF")
col1, col2 = st.columns(2)

with col1:
    st.markdown("**Modello H 2026 — Conto Finanziario (secondo livello)**")
    st.caption("Contiene entrate e spese in competenza")
    file_h = st.file_uploader("Carica Modello H", type=["pdf"], key="upload_h")
    if file_h:
        st.session_state.modello_h_bytes = file_h.read()
        st.success(f"✅ {file_h.name} ({len(st.session_state.modello_h_bytes)//1024} KB)")

with col2:
    st.markdown("**Modello L 2025 — Elenco Residui**")
    st.caption("Contiene residui attivi e passivi da anni precedenti")
    file_l = st.file_uploader("Carica Modello L", type=["pdf"], key="upload_l")
    if file_l:
        st.session_state.modello_l_bytes = file_l.read()
        st.success(f"✅ {file_l.name} ({len(st.session_state.modello_l_bytes)//1024} KB)")

st.markdown("---")

# ── SEZIONE 2: Modalità ───────────────────────────────
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
    st.info("🤖 **Automatica**: percentuali applicate automaticamente. Puoi modificarle qui sotto.")
else:
    st.info("📝 **Manuale**: riceverai il file Excel precompilato con i già riscossi/pagati. Inserisci tu le previsioni.")

st.markdown("---")

# ── SEZIONE 3: Percentuali (solo automatica) ──────────
if st.session_state.modalita == "automatica":
    st.markdown("### 3️⃣ Percentuali di ripartizione")
    st.caption("La somma Gen–Ago + Set–Dic deve essere ≤ 100% per ogni categoria.")

    p = st.session_state.percentuali.copy()

    # Header
    col1, col2, col3 = st.columns([3, 1.5, 1.5])
    with col2:
        st.markdown("**Gen–Ago %**")
    with col3:
        st.markdown("**Set–Dic %**")

    def percentuale_row(label, key_ga, key_sd, help_text=""):
        col1, col2, col3 = st.columns([3, 1.5, 1.5])
        with col1:
            st.markdown(f"**{label}**")
            if help_text:
                st.caption(help_text)
        with col2:
            val_ga = st.number_input(
                "Gen-Ago %", min_value=0, max_value=100,
                value=int(p[key_ga]), key=f"pct_{key_ga}",
                label_visibility="collapsed"
            )
        with col3:
            val_sd = st.number_input(
                "Set-Dic %", min_value=0, max_value=100,
                value=int(p[key_sd]), key=f"pct_{key_sd}",
                label_visibility="collapsed"
            )
        if val_ga + val_sd > 100:
            st.error(f"⚠️ {label}: Gen–Ago + Set–Dic = {val_ga+val_sd}% (max 100%)")
        p[key_ga] = val_ga
        p[key_sd] = val_sd

    st.markdown("**📥 ENTRATE**")
    percentuale_row("Entrate in competenza", "entrate_competenza_gen_ago", "entrate_competenza_set_dic",
                    "Base: Programmaz. Definitiva − Già riscosso")
    percentuale_row("Residui attivi (Modello L)", "residui_attivi_gen_ago", "residui_attivi_set_dic",
                    "Base: Importo residuo")
    st.markdown("**📤 SPESE**")
    percentuale_row("Spese in competenza", "spese_competenza_gen_ago", "spese_competenza_set_dic",
                    "Base: Programmaz. Definitiva − Già pagato")
    percentuale_row("Residui passivi (Modello L)", "residui_passivi_gen_ago", "residui_passivi_set_dic",
                    "Base: Importo residuo")

    st.session_state.percentuali = p

    if st.button("↩️ Ripristina percentuali predefinite", key="btn_ripristina"):
        for k, v in PERCENTUALI_DEFAULT.items():
            st.session_state[f"pct_{k}"] = v
        st.session_state.percentuali = PERCENTUALI_DEFAULT.copy()
        st.rerun()

    st.markdown("---")

# ── SEZIONE 4: Estrazione ─────────────────────────────
st.markdown("### 4️⃣ Estrai i dati dai PDF")

pdf_pronti = st.session_state.modello_h_bytes and st.session_state.modello_l_bytes

if not pdf_pronti:
    st.warning("⚠️ Carica entrambi i PDF prima di procedere.")
else:
    if st.button("🔍 Estrai dati con AI", use_container_width=True, type="primary"):
        with st.spinner("Estrazione in corso... (può richiedere 30–60 secondi)"):
            try:
                from utils.extractor import estrai_dati_pdf
                dati = estrai_dati_pdf(
                    st.session_state.modello_h_bytes,
                    st.session_state.modello_l_bytes
                )
                st.session_state.dati_estratti = dati

                # Auto-popola dati scuola se estratti
                if "dati_scuola" not in st.session_state:
                    st.session_state.dati_scuola = {}
                ds = st.session_state.dati_scuola
                for campo in ["nome_istituto", "cf", "codice_mecc", "indirizzo", "citta"]:
                    if dati.get(campo) and not ds.get(campo):
                        ds[campo] = dati[campo]

                if dati.get("_parse_error"):
                    st.warning("⚠️ Estrazione completata ma con possibili imprecisioni. Verifica i dati nell'anteprima.")
                else:
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
