import streamlit as st
import json
import io

st.title("📊 Genera Documenti")
st.markdown("---")

# ── SALVATAGGIO / CARICAMENTO SESSIONE ────────────────────
with st.expander("💾 Salva / Carica sessione"):
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**Salva sessione**")
        st.caption("Scarica un file JSON con tutti i dati inseriti per riutilizzarli in futuro.")
        if st.button("⬇️ Salva sessione", use_container_width=True):
            sessione = {
                "dati_scuola": st.session_state.get("dati_scuola", {}),
                "modalita": st.session_state.get("modalita", "automatica"),
                "percentuali": st.session_state.get("percentuali", {}),
                "dati_estratti": st.session_state.get("dati_estratti", None),
            }
            # Converti bytes in None per la serializzazione
            sessione_clean = {}
            for k, v in sessione.items():
                if isinstance(v, dict):
                    sessione_clean[k] = {kk: vv for kk, vv in v.items() if not isinstance(vv, bytes)}
                else:
                    sessione_clean[k] = v

            nome = st.session_state.get("dati_scuola", {}).get("nome_istituto", "scuola").replace(" ", "_")[:20]
            anno = st.session_state.get("dati_scuola", {}).get("anno_esercizio", "2026")
            st.download_button(
                label="📥 Clicca qui per scaricare",
                data=json.dumps(sessione_clean, ensure_ascii=False, indent=2),
                file_name=f"sessione_flussi_{nome}_{anno}.json",
                mime="application/json",
                use_container_width=True
            )

    with col2:
        st.markdown("**Carica sessione**")
        st.caption("Ripristina una sessione precedentemente salvata.")
        file_sessione = st.file_uploader("Carica file JSON sessione", type=["json"], key="upload_sessione")
        if file_sessione:
            try:
                dati_sessione = json.load(file_sessione)
                if st.button("✅ Ripristina sessione", use_container_width=True):
                    if "dati_scuola" in dati_sessione:
                        st.session_state.dati_scuola = dati_sessione["dati_scuola"]
                    if "modalita" in dati_sessione:
                        st.session_state.modalita = dati_sessione["modalita"]
                    if "percentuali" in dati_sessione:
                        st.session_state.percentuali = dati_sessione["percentuali"]
                    if "dati_estratti" in dati_sessione and dati_sessione["dati_estratti"]:
                        st.session_state.dati_estratti = dati_sessione["dati_estratti"]
                    st.success("✅ Sessione ripristinata!")
                    st.rerun()
            except Exception as e:
                st.error(f"❌ Errore nel caricamento: {e}")

st.markdown("---")

# ── CHECK PREREQUISITI ────────────────────────────────────
dati = st.session_state.get("dati_estratti")
dati_scuola = st.session_state.get("dati_scuola", {})
modalita = st.session_state.get("modalita", "automatica")
percentuali = st.session_state.get("percentuali", {})

col1, col2, col3 = st.columns(3)
with col1:
    if dati:
        st.success("✅ Dati PDF estratti")
    else:
        st.error("❌ PDF non ancora elaborati")
with col2:
    if dati_scuola.get("nome_istituto"):
        st.success(f"✅ {dati_scuola.get('nome_istituto', '')[:25]}")
    else:
        st.warning("⚠️ Dati scuola incompleti")
with col3:
    fondo = dati_scuola.get("fondo_cassa", 0)
    if float(fondo) > 0:
        st.success(f"✅ Fondo cassa: € {float(fondo):,.2f}")
    else:
        st.warning("⚠️ Fondo cassa non inserito")

if not dati:
    st.warning("⚠️ Prima di generare i documenti, carica i PDF e avvia l'estrazione dalla pagina **📄 Caricamento PDF**.")
    st.stop()

st.markdown("---")

# ── RIEPILOGO ─────────────────────────────────────────────
st.markdown("### 📋 Riepilogo elaborazione")
col1, col2 = st.columns(2)
with col1:
    st.markdown(f"**Istituto:** {dati_scuola.get('nome_istituto', 'N/D')}")
    st.markdown(f"**Anno:** {dati_scuola.get('anno_esercizio', '2026')}")
    st.markdown(f"**Fondo cassa al 01/01:** € {float(dati_scuola.get('fondo_cassa', 0)):,.2f}")
    minute = float(dati_scuola.get('fondo_minute_spese', 0))
    if minute > 0:
        st.markdown(f"**Fondo minute spese:** € {minute:,.2f}")
    st.markdown(f"**Modalità:** {'🤖 Automatica' if modalita == 'automatica' else '📝 Manuale'}")
with col2:
    st.markdown(f"**Voci entrate (Mod. H):** {len(dati.get('entrate', []))}")
    st.markdown(f"**Voci spese (Mod. H):** {len(dati.get('spese', []))}")
    st.markdown(f"**Residui attivi (Mod. L):** {len(dati.get('residui_attivi', []))}")
    st.markdown(f"**Residui passivi (Mod. L):** {len(dati.get('residui_passivi', []))}")

if modalita == "automatica":
    with st.expander("📊 Percentuali applicate"):
        p = percentuali
        categories = [
            ("Entrate competenza", "entrate_competenza_gen_ago", "entrate_competenza_set_dic"),
            ("Spese competenza", "spese_competenza_gen_ago", "spese_competenza_set_dic"),
            ("Residui attivi", "residui_attivi_gen_ago", "residui_attivi_set_dic"),
            ("Residui passivi", "residui_passivi_gen_ago", "residui_passivi_set_dic"),
        ]
        for cat, k_ga, k_sd in categories:
            col1, col2, col3 = st.columns([3, 1, 1])
            with col1: st.write(cat)
            with col2: st.write(f"**{p.get(k_ga, 0)}%** Gen-Ago")
            with col3: st.write(f"**{p.get(k_sd, 0)}%** Set-Dic")

st.markdown("---")

# ── GENERA ────────────────────────────────────────────────
st.markdown("### 🚀 Genera i documenti")

if st.button("⚙️ GENERA TUTTI I DOCUMENTI", use_container_width=True, type="primary"):
    progress = st.progress(0)
    status = st.empty()
    try:
        status.info("📊 Generazione file Excel...")
        progress.progress(20)
        from utils.excel_generator import genera_excel
        excel_bytes = genera_excel(dati, dati_scuola, percentuali, modalita)
        st.session_state["excel_bytes"] = excel_bytes
        progress.progress(50)

        status.info("📝 Generazione nota di accompagnamento...")
        from utils.doc_generator import genera_nota
        nota_bytes = genera_nota(dati_scuola, percentuali, modalita)
        st.session_state["nota_bytes"] = nota_bytes
        progress.progress(75)

        status.info("📋 Generazione decreto di adozione...")
        from utils.doc_generator import genera_decreto
        decreto_bytes = genera_decreto(dati_scuola, percentuali, modalita)
        st.session_state["decreto_bytes"] = decreto_bytes
        progress.progress(100)

        status.success("✅ Tutti i documenti generati con successo!")

    except Exception as e:
        progress.empty()
        st.error(f"❌ Errore: {str(e)}")
        st.exception(e)

# ── DOWNLOAD ──────────────────────────────────────────────
if st.session_state.get("excel_bytes"):
    st.markdown("---")
    st.markdown("### 💾 Scarica i documenti")

    nome = dati_scuola.get("nome_istituto", "Scuola").replace(" ", "_")[:20]
    anno = dati_scuola.get("anno_esercizio", "2026")

    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("**📊 File Excel**")
        st.caption("Piano dei Flussi — formato MIM")
        st.download_button(
            label="⬇️ Scarica Excel",
            data=st.session_state["excel_bytes"],
            file_name=f"Piano_Flussi_{nome}_{anno}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with col2:
        st.markdown("**📝 Nota accompagnamento**")
        st.caption("Per i Revisori dei Conti — Word")
        st.download_button(
            label="⬇️ Scarica Nota Word",
            data=st.session_state["nota_bytes"],
            file_name=f"Nota_Accompagnamento_{nome}_{anno}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
    with col3:
        st.markdown("**📋 Decreto DS**")
        st.caption("Decreto di adozione — Word")
        st.download_button(
            label="⬇️ Scarica Decreto Word",
            data=st.session_state["decreto_bytes"],
            file_name=f"Decreto_Adozione_{nome}_{anno}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

    st.markdown("---")
    st.markdown("### ✅ Cosa fare dopo")
    st.markdown("""
1. **Apri il file Excel** e verifica i dati — le voci evidenziate in **giallo** hanno programmazione = 0 con movimenti
2. Se sei in **modalità manuale**, completa le colonne Gen–Ago e Set–Dic nei fogli ENTRATE e SPESE
3. **Firma e protocolla** il Decreto DS
4. **Trasmetti ai Revisori** la nota di accompagnamento insieme al file Excel
5. **Pubblica all'Albo** il decreto entro il 28 febbraio
    """)
    st.info("💡 Per rigenerare con parametri diversi, torna alla pagina **📄 Caricamento PDF** e modifica le impostazioni.")
