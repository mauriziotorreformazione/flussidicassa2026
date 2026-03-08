import streamlit as st

st.title("📊 Genera Documenti")
st.markdown("---")

# Check prerequisites
dati = st.session_state.get("dati_estratti")
dati_scuola = st.session_state.get("dati_scuola", {})
modalita = st.session_state.get("modalita", "automatica")
percentuali = st.session_state.get("percentuali", {})

# Status check
col1, col2, col3 = st.columns(3)
with col1:
    if dati:
        st.success("✅ Dati PDF estratti")
    else:
        st.error("❌ PDF non ancora elaborati")
with col2:
    if dati_scuola.get("nome_istituto"):
        st.success(f"✅ Dati scuola: {dati_scuola.get('nome_istituto', '')[:30]}")
    else:
        st.warning("⚠️ Dati scuola incompleti")
with col3:
    fondo = dati_scuola.get("fondo_cassa", 0)
    if fondo > 0:
        st.success(f"✅ Fondo cassa: € {fondo:,.2f}")
    else:
        st.warning("⚠️ Fondo cassa non inserito")

if not dati:
    st.warning("⚠️ Prima di generare i documenti, carica i PDF e avvia l'estrazione dalla pagina **📄 Caricamento PDF**.")
    st.stop()

st.markdown("---")

# Riepilogo parametri
st.markdown("### 📋 Riepilogo elaborazione")
col1, col2 = st.columns(2)
with col1:
    st.markdown(f"**Istituto:** {dati_scuola.get('nome_istituto', 'N/D')}")
    st.markdown(f"**Anno:** {dati_scuola.get('anno_esercizio', '2026')}")
    st.markdown(f"**Fondo cassa al 01/01:** € {float(dati_scuola.get('fondo_cassa', 0)):,.2f}")
    st.markdown(f"**Modalità:** {'🤖 Automatica' if modalita == 'automatica' else '📝 Manuale'}")
with col2:
    n_entrate = len(dati.get("entrate", []))
    n_spese = len(dati.get("spese", []))
    n_res_att = len(dati.get("residui_attivi", []))
    n_res_pas = len(dati.get("residui_passivi", []))
    st.markdown(f"**Voci entrate (Mod. H):** {n_entrate}")
    st.markdown(f"**Voci spese (Mod. H):** {n_spese}")
    st.markdown(f"**Residui attivi (Mod. L):** {n_res_att}")
    st.markdown(f"**Residui passivi (Mod. L):** {n_res_pas}")

if modalita == "automatica":
    with st.expander("📊 Percentuali applicate"):
        p = percentuali
        col1, col2, col3 = st.columns([3, 1.5, 1.5])
        with col1:
            st.markdown("**Categoria**")
        with col2:
            st.markdown("**Gen-Ago**")
        with col3:
            st.markdown("**Set-Dic**")
        categories = [
            ("Entrate competenza", "entrate_competenza_gen_ago", "entrate_competenza_set_dic"),
            ("Spese competenza", "spese_competenza_gen_ago", "spese_competenza_set_dic"),
            ("Residui attivi", "residui_attivi_gen_ago", "residui_attivi_set_dic"),
            ("Residui passivi", "residui_passivi_gen_ago", "residui_passivi_set_dic"),
        ]
        for cat, k_ga, k_sd in categories:
            col1, col2, col3 = st.columns([3, 1.5, 1.5])
            with col1:
                st.write(cat)
            with col2:
                st.write(f"{p.get(k_ga, 0)}%")
            with col3:
                st.write(f"{p.get(k_sd, 0)}%")

st.markdown("---")

# Generate button
st.markdown("### 🚀 Genera i documenti")

if st.button("⚙️ GENERA TUTTI I DOCUMENTI", use_container_width=True, type="primary"):
    
    progress = st.progress(0)
    status = st.empty()

    try:
        # 1. Excel
        status.info("📊 Generazione file Excel...")
        progress.progress(20)
        from utils.excel_generator import genera_excel
        excel_bytes = genera_excel(dati, dati_scuola, percentuali, modalita)
        st.session_state["excel_bytes"] = excel_bytes
        progress.progress(45)

        # 2. Nota Word
        status.info("📝 Generazione nota di accompagnamento...")
        from utils.doc_generator import genera_nota
        nota_bytes = genera_nota(dati_scuola, percentuali, modalita)
        st.session_state["nota_bytes"] = nota_bytes
        progress.progress(70)

        # 3. Decreto Word
        status.info("📋 Generazione decreto di adozione...")
        from utils.doc_generator import genera_decreto
        decreto_bytes = genera_decreto(dati_scuola, percentuali, modalita)
        st.session_state["decreto_bytes"] = decreto_bytes
        progress.progress(100)

        status.success("✅ Tutti i documenti generati con successo!")

    except Exception as e:
        progress.empty()
        st.error(f"❌ Errore durante la generazione: {str(e)}")
        st.exception(e)

# Download buttons (shown when docs are ready)
if st.session_state.get("excel_bytes"):
    st.markdown("---")
    st.markdown("### 💾 Scarica i documenti")

    nome_istituto = dati_scuola.get("nome_istituto", "Scuola").replace(" ", "_")[:20]
    anno = dati_scuola.get("anno_esercizio", "2026")

    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown("**📊 File Excel**")
        st.caption("Piano dei Flussi di Cassa — formato MIM")
        st.download_button(
            label="⬇️ Scarica Excel",
            data=st.session_state["excel_bytes"],
            file_name=f"Piano_Flussi_Cassa_{nome_istituto}_{anno}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    with col2:
        st.markdown("**📝 Nota accompagnamento**")
        st.caption("Per i Revisori dei Conti — Word")
        st.download_button(
            label="⬇️ Scarica Nota Word",
            data=st.session_state["nota_bytes"],
            file_name=f"Nota_Accompagnamento_Flussi_{nome_istituto}_{anno}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

    with col3:
        st.markdown("**📋 Decreto DS**")
        st.caption("Decreto di adozione — Word")
        st.download_button(
            label="⬇️ Scarica Decreto Word",
            data=st.session_state["decreto_bytes"],
            file_name=f"Decreto_Adozione_Flussi_{nome_istituto}_{anno}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

    st.markdown("---")
    st.markdown("### ✅ Cosa fare dopo")
    st.markdown("""
    1. **Apri il file Excel** e verifica i dati — soprattutto le voci evidenziate in arancione
    2. Se sei in **modalità manuale**, completa le colonne Gen-Ago e Set-Dic nel foglio ENTRATE e SPESE
    3. **Firma e protocolla** il Decreto DS
    4. **Trasmetti ai Revisori** la nota di accompagnamento insieme al file Excel
    5. **Pubblica all'Albo** il decreto entro il 28 febbraio
    """)

    st.info("💡 Per rigenerare con parametri diversi, torna alla pagina **📄 Caricamento PDF** e modifica le impostazioni.")
