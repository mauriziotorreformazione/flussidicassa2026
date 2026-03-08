import streamlit as st

st.title("📖 Istruzioni per l'utilizzo")
st.markdown("---")

st.info("**Benvenuto** nell'applicazione per la generazione del Piano Annuale dei Flussi di Cassa.\n\nSegui le istruzioni di questa pagina prima di procedere.")

st.markdown("## 🎯 A cosa serve questa app")
st.markdown("""
Questa applicazione ti permette di generare in pochi minuti il **Piano Annuale dei Flussi di Cassa** 
richiesto dalla normativa (art. 6, D.L. 155/2024, convertito dalla Legge 189/2024), partendo 
direttamente dai documenti contabili della tua scuola.

L'app produce **3 documenti**:
- 📊 **File Excel** con il Piano dei Flussi di Cassa (formato ministeriale)
- 📝 **Nota di accompagnamento** per i Revisori dei Conti (Word + PDF)
- 📋 **Decreto di adozione** del Dirigente Scolastico (Word + PDF)
""")

st.markdown("---")

st.markdown("## 📂 Cosa ti serve")

col1, col2 = st.columns(2)
with col1:
    st.markdown("""
    **Documenti da caricare:**
    - 📄 **Modello H+** — Conto Consuntivo (Conto Finanziario) in formato PDF
    - 📄 **Modello L** — Elenco Residui Attivi e Passivi in formato PDF
    """)
with col2:
    st.markdown("""
    **Dati da inserire manualmente:**
    - 💰 Fondo cassa al 01/01/2026
    - 👤 Nome del Dirigente Scolastico
    - 👤 Nome del DSGA
    - 📅 Data e numero delibera del Consiglio d'Istituto
    - 📅 Data e numero protocollo del decreto
    """)

st.markdown("---")

st.markdown("## 🔢 Il Fondo Cassa")
st.markdown("""
Il **Fondo Cassa al 01/01** è il saldo di cassa della scuola all'inizio dell'esercizio finanziario.
Puoi trovarlo nel rendiconto dell'anno precedente o nel sistema contabile (es. ARGO Bilancio 2.0).

Questo dato è obbligatorio: viene inserito nel Piano dei Flussi ministeriale ed entra nel 
calcolo del saldo di cassa stimato al 31/08 e al 31/12.
""")

st.markdown("---")

st.markdown("## ⚙️ Le due modalità di compilazione")

tab1, tab2 = st.tabs(["📝 Modalità Manuale", "🤖 Modalità Automatica"])

with tab1:
    st.markdown("""
    ### Modalità Manuale
    
    Nella modalità manuale, l'app estrae dal Modello H e dal Modello L tutti gli importi 
    (programmazione definitiva, già riscosso/pagato, residui) e ti presenta un file Excel 
    con le colonne **Gennaio–Agosto** e **Settembre–Dicembre** precompilate con i valori 
    già incassati/pagati.
    
    **Tu devi:**
    - Verificare i dati estratti
    - Inserire (o modificare) le previsioni nelle colonne Gen-Ago e Set-Dic
    
    **Vincoli da rispettare:**
    - La colonna Gen-Ago può essere modificata **solo in aumento** rispetto al già riscosso/pagato
    - La somma Gen-Ago + Set-Dic **non può superare** la Programmazione Definitiva
    - Eccezione: voci con Programmazione = 0 ma movimenti già presenti (evidenziate in arancione) — nessun vincolo
    
    > 💡 **Quando usarla:** quando hai una conoscenza dettagliata dei flussi attesi e vuoi 
    > personalizzare voce per voce.
    """)

with tab2:
    st.markdown("""
    ### Modalità Automatica
    
    Nella modalità automatica, l'app ripartisce automaticamente gli importi tra Gen-Ago e 
    Set-Dic applicando **percentuali predefinite** sulla differenza tra Programmazione Definitiva 
    e quanto già riscosso/pagato.
    
    **Percentuali di default:**
    
    | Categoria | Base calcolo | Gen-Ago | Set-Dic |
    |-----------|-------------|---------|---------|
    | Entrate competenza | Programmaz. Definitiva − Già riscosso | 100% | 0% |
    | Spese competenza | Programmaz. Definitiva − Già pagato | 67% | 33% |
    | Residui attivi | Importo residuo | 33% | 67% |
    | Residui passivi | Importo residuo | 33% | 67% |
    
    **Puoi modificare le percentuali** per categoria prima di generare il file.
    
    Le somme già riscosse/pagate vengono sempre aggiunte alla colonna Gen-Ago.
    
    > 💡 **Quando usarla:** per una prima stima veloce, o quando non hai elementi per 
    > prevedere voce per voce.
    
    > ⚠️ **Attenzione:** le percentuali producono una stima. Ti consigliamo di rivedere 
    > il file Excel generato e correggere manualmente le voci che conosci bene.
    """)

st.markdown("---")

st.markdown("## 📊 Struttura del file Excel generato")
st.markdown("""
Il file Excel contiene **3 fogli**:

**Foglio 1 — ENTRATE**
- Prima sezione: voci di entrata in competenza (dal Modello H)
- Seconda sezione: residui attivi (dal Modello L)
- Colonne: Voce | Progr. Definitiva | Già Riscosso | Gen-Ago | Set-Dic | Totale

**Foglio 2 — SPESE**
- Prima sezione: voci di spesa in competenza per aggregato (dal Modello H)
- Seconda sezione: residui passivi (dal Modello L)
- Colonne: Voce | Progr. Definitiva | Già Pagato | Gen-Ago | Set-Dic | Totale

**Foglio 3 — PIANO DEI FLUSSI (MIM)**
- Formato ministeriale ufficiale
- Si compila automaticamente dai dati dei fogli 1 e 2
- Riporta il Fondo Cassa iniziale e i saldi stimati al 31/08 e 31/12
""")

st.markdown("---")

st.markdown("## 🚀 Come procedere")
st.markdown("""
1. Vai alla pagina **🏫 Dati Scuola** e inserisci le informazioni della tua scuola
2. Vai alla pagina **📄 Caricamento PDF** e carica i due modelli
3. Scegli la modalità (manuale o automatica) e configura le percentuali se necessario
4. Vai alla pagina **📊 Genera Documenti** e scarica i 3 file prodotti
""")

st.success("✅ Hai letto le istruzioni? Procedi con la pagina **🏫 Dati Scuola** dalla barra laterale.")
