import streamlit as st

st.title("🏫 Dati della Scuola")
st.markdown("---")
st.info("Inserisci i dati della tua scuola. Alcuni campi verranno precompilati automaticamente dopo il caricamento dei PDF.")

if "dati_scuola" not in st.session_state:
    st.session_state.dati_scuola = {
        "nome_istituto": "I.C. GIOVANNI PASCOLI",
        "indirizzo": "Via Roma, 1",
        "cf": "00000000000",
        "codice_mecc": "RCIC000000",
        "citta": "00100 ROMA",
        "dirigente": "",
        "dsga": "",
        "anno_esercizio": "2026",
        "fondo_cassa": 0.0,
        "fondo_minute_spese": 0.0,
        "data_delibera_ci": "",
        "num_delibera_ci": "",
        "data_decreto": "",
        "num_protocollo": "",
        "email": "",
        "tel": "",
    }

st.markdown("### 📋 Dati istituto")
st.caption("Questi campi vengono precompilati automaticamente dai PDF se disponibili.")

col1, col2 = st.columns(2)
with col1:
    st.session_state.dati_scuola["nome_istituto"] = st.text_input(
        "Nome Istituto *",
        value=st.session_state.dati_scuola["nome_istituto"],
        placeholder="es. I.C. GIOVANNI PASCOLI"
    )
    st.session_state.dati_scuola["indirizzo"] = st.text_input(
        "Indirizzo",
        value=st.session_state.dati_scuola["indirizzo"],
        placeholder="es. Via Roma, 1"
    )
    st.session_state.dati_scuola["citta"] = st.text_input(
        "Città",
        value=st.session_state.dati_scuola["citta"],
        placeholder="es. 00100 ROMA"
    )
with col2:
    st.session_state.dati_scuola["cf"] = st.text_input(
        "Codice Fiscale",
        value=st.session_state.dati_scuola["cf"],
        placeholder="es. 80000000000"
    )
    st.session_state.dati_scuola["codice_mecc"] = st.text_input(
        "Codice Meccanografico",
        value=st.session_state.dati_scuola["codice_mecc"],
        placeholder="es. RCIC000000"
    )
    st.session_state.dati_scuola["email"] = st.text_input(
        "Email istituzionale",
        value=st.session_state.dati_scuola["email"],
        placeholder="es. scuola@istruzione.it"
    )

st.markdown("---")
st.markdown("### 👤 Personale")
col1, col2 = st.columns(2)
with col1:
    st.session_state.dati_scuola["dirigente"] = st.text_input(
        "Dirigente Scolastico *",
        value=st.session_state.dati_scuola["dirigente"],
        placeholder="es. Prof.ssa Maria Rossi"
    )
with col2:
    st.session_state.dati_scuola["dsga"] = st.text_input(
        "Direttore S.G.A. (DSGA) *",
        value=st.session_state.dati_scuola["dsga"],
        placeholder="es. Dott. Mario Bianchi"
    )

st.markdown("---")
st.markdown("### 💰 Dati finanziari")
col1, col2, col3 = st.columns(3)
with col1:
    st.session_state.dati_scuola["anno_esercizio"] = st.text_input(
        "Anno esercizio *",
        value=st.session_state.dati_scuola["anno_esercizio"]
    )
with col2:
    fondo = st.number_input(
        "Fondo Cassa al 01/01 (€) *",
        min_value=0.0,
        value=float(st.session_state.dati_scuola["fondo_cassa"]),
        step=0.01,
        format="%.2f",
        help="Saldo di cassa della scuola al 1° gennaio dell'esercizio."
    )
    st.session_state.dati_scuola["fondo_cassa"] = fondo
with col3:
    minute = st.number_input(
        "Fondo Minute Spese (€)",
        min_value=0.0,
        value=float(st.session_state.dati_scuola.get("fondo_minute_spese", 0.0)),
        step=0.01,
        format="%.2f",
        help="Importo del fondo minute spese. Verrà inserito come partita di giro: in entrata (Set–Dic) e in uscita (Gen–Ago)."
    )
    st.session_state.dati_scuola["fondo_minute_spese"] = minute

if minute > 0:
    st.info(f"💡 Il fondo minute spese di **€ {minute:,.2f}** verrà inserito come partita di giro: **in entrata** nella colonna Set–Dic e **in uscita** nella colonna Gen–Ago.")

st.markdown("---")
st.markdown("### 📅 Atti amministrativi")
st.caption("Necessari per la compilazione del decreto di adozione.")
col1, col2 = st.columns(2)
with col1:
    st.session_state.dati_scuola["num_delibera_ci"] = st.text_input(
        "N° delibera Consiglio d'Istituto",
        value=st.session_state.dati_scuola["num_delibera_ci"],
        placeholder="es. 8"
    )
    st.session_state.dati_scuola["data_delibera_ci"] = st.text_input(
        "Data delibera Consiglio d'Istituto",
        value=st.session_state.dati_scuola["data_delibera_ci"],
        placeholder="es. 03/02/2026"
    )
with col2:
    st.session_state.dati_scuola["num_protocollo"] = st.text_input(
        "N° protocollo decreto DS",
        value=st.session_state.dati_scuola["num_protocollo"],
        placeholder="es. 1363"
    )
    st.session_state.dati_scuola["data_decreto"] = st.text_input(
        "Data decreto DS",
        value=st.session_state.dati_scuola["data_decreto"],
        placeholder="es. 20/02/2026"
    )

st.markdown("---")
campi_obbligatori = [
    st.session_state.dati_scuola["nome_istituto"],
    st.session_state.dati_scuola["dirigente"],
    st.session_state.dati_scuola["dsga"],
]
if all(campi_obbligatori):
    st.success("✅ Dati principali inseriti. Procedi con il **📄 Caricamento PDF**.")
else:
    st.warning("⚠️ Compila almeno i campi obbligatori (*) per procedere.")
