import streamlit as st

st.title("📖 Istruzioni per l'utilizzo")
st.caption("Web App creata da **Maurizio Torre**")
st.markdown("---")

st.info("**Benvenuto** nell'applicazione per la generazione del Piano Annuale dei Flussi di Cassa.\n\nSegui le istruzioni di questa pagina prima di procedere.")

st.markdown("## 🎯 A cosa serve questa app")
st.markdown("""
Questa applicazione ti permette di generare in pochi minuti il **Piano Annuale dei Flussi di Cassa** 
richiesto dalla normativa (art. 6, D.L. 155/2024, convertito dalla Legge 189/2024), partendo 
direttamente dai documenti contabili della tua scuola.

L'app produce **3 documenti**:
- 📊 **File Excel** con il Piano dei Flussi di Cassa (formato ministeriale)
- 📝 **Nota di accompagnamento** per i Revisori dei Conti (Word)
- 📋 **Decreto di adozione** del Dirigente Scolastico (Word)
""")

st.markdown("---")
st.markdown("## 📂 Cosa ti serve")
col1, col2 = st.columns(2)
with col1:
    st.markdown("""
**Documenti da caricare:**
- 📄 **Modello H 2026** — Conto Consuntivo Conto Finanziario in formato PDF
  *(utilizzare il modello al **secondo livello**)*
- 📄 **Modello L 2025** — Elenco Residui Attivi e Passivi in formato PDF
""")
with col2:
    st.markdown("""
**Dati da inserire manualmente:**
- 💰 Fondo cassa al 01/01/2026
- 💰 Fondo minute spese
- 👤 Nome del Dirigente Scolastico
- 👤 Nome del DSGA
- 📅 Data e numero delibera del Consiglio d'Istituto
- 📅 Data e numero protocollo del decreto
""")

st.markdown("---")
st.markdown("## 🔢 Il Fondo Cassa e le Minute Spese")
col1, col2 = st.columns(2)
with col1:
    st.markdown("""
**Fondo Cassa al 01/01**

È il saldo di cassa della scuola all'inizio dell'esercizio finanziario.
Trovalo nel rendiconto dell'anno precedente o nel tuo sistema contabile.

Viene inserito nel Piano dei Flussi ministeriale ed entra nel 
calcolo del saldo stimato al 31/08 e al 31/12.
""")
with col2:
    st.markdown("""
**Fondo Minute Spese**

È il fondo per piccole spese di cassa. Va indicato come **partita di giro**:
- In **entrata** nella colonna Set–Dic (ricostituzione fondo)
- In **uscita** nella colonna Gen–Ago (utilizzo fondo)

L'importo è quello previsto nella programmazione della scuola.
""")

st.markdown("---")
st.markdown("## ⚙️ Le due modalità di compilazione")

tab1, tab2 = st.tabs(["📝 Modalità Manuale", "🤖 Modalità Automatica"])

with tab1:
    st.markdown("""
### Modalità Manuale

L'app estrae dal Modello H e dal Modello L tutti gli importi e ti presenta 
un file Excel con le colonne **Gen–Ago** e **Set–Dic** precompilate con i 
valori già incassati/pagati. Sei tu a inserire le previsioni.

**Vincoli da rispettare:**
- Gen–Ago non può essere inferiore al già riscosso/pagato
- Gen–Ago + Set–Dic non può superare la Programmazione Definitiva
- Voci con programmazione = 0 ma movimenti presenti → evidenziate in **giallo** → nessun vincolo
""")
    st.success("✅ **Quando usarla:** quando conosci bene i flussi attesi della tua scuola e vuoi personalizzare la distribuzione voce per voce, ad esempio perché sai già quando arriveranno certi finanziamenti o quando scadono certi pagamenti.")

with tab2:
    st.markdown("""
### Modalità Automatica

L'app ripartisce automaticamente gli importi tra Gen–Ago e Set–Dic 
applicando **percentuali predefinite** sulla differenza tra Programmazione 
Definitiva e quanto già riscosso/pagato.

**Percentuali di default (modificabili):**

| Categoria | Base calcolo | Gen–Ago | Set–Dic |
|-----------|-------------|---------|---------|
| Entrate competenza | Programmaz. − Già riscosso | 100% | 0% |
| Spese competenza | Programmaz. − Già pagato | 67% | 33% |
| Residui attivi | Importo residuo | 33% | 67% |
| Residui passivi | Importo residuo | 33% | 67% |

Le somme già riscosse/pagate vengono sempre assegnate a Gen–Ago.
""")
    st.success("✅ **Quando usarla:** per una prima stima veloce, oppure quando non hai elementi sufficienti per prevedere i flussi voce per voce. Le percentuali sono modificabili prima della generazione. Ti consigliamo comunque di rivedere il file Excel generato e correggere le voci che conosci meglio.")

st.markdown("---")
st.markdown("## 📊 Struttura del file Excel generato")
st.markdown("""
Il file Excel contiene **3 fogli**:

**Foglio 1 — ENTRATE**
- Prima sezione: voci di entrata in competenza (dal Modello H 2026, secondo livello)
- Seconda sezione: residui attivi (dal Modello L 2025)
- In fondo: riga partite di giro — Fondo Minute Spese (in entrata, colonna Set–Dic)
- Colonne: Voce | Progr. Definitiva | Già Riscosso | Gen–Ago | Set–Dic | Totale

**Foglio 2 — SPESE**
- Prima sezione: voci di spesa per aggregato e progetto (dal Modello H 2026)
- Seconda sezione: residui passivi (dal Modello L 2025)
- In fondo: riga partite di giro — Fondo Minute Spese (in uscita, colonna Gen–Ago)
- Colonne: Voce | Progr. Definitiva | Già Pagato | Gen–Ago | Set–Dic | Totale

**Foglio 3 — PIANO DEI FLUSSI (MIM)**
- Struttura ministeriale ufficiale (Allegato 3, Nota MIM n. 2284/2025)
- Si compila automaticamente aggregando i dati per codice PDC
- Riporta il Fondo Cassa iniziale e i saldi stimati al 31/08 e al 31/12

> ⚠️ **Nota metodologica:** le previsioni di spesa sono proposte per **aggregato e progetto**
> (A01, A02, P01, P02...) anziché per tipologia di spesa come nel formato ministeriale.
> Questa scelta è intenzionale per rendere la compilazione più semplice e leggibile,
> mantenendo la piena corrispondenza con il Piano dei Conti Integrato tramite la mappatura PDC.
""")

st.markdown("---")
st.markdown("## 💾 Salvataggio dei dati")
st.markdown("""
L'app non salva automaticamente i dati tra una sessione e l'altra.
Nella pagina **📊 Genera Documenti** trovi il pulsante **Salva sessione**
che scarica un file JSON con tutti i dati inseriti.

Per ripristinare una sessione precedente, usa il pulsante **Carica sessione**
nella stessa pagina e seleziona il file JSON salvato in precedenza.
""")

st.markdown("---")
st.markdown("## 🚀 Come procedere")
st.markdown("""
1. Vai alla pagina **🏫 Dati Scuola** e inserisci le informazioni della tua scuola
2. Vai alla pagina **📄 Caricamento PDF** e carica i due modelli
3. Scegli la modalità e configura le percentuali se necessario
4. Vai alla pagina **📊 Genera Documenti** e scarica i 3 file prodotti
""")

st.success("✅ Hai letto le istruzioni? Procedi con la pagina **🏫 Dati Scuola** dalla barra laterale.")
st.markdown("---")
st.caption("Web App creata da **Maurizio Torre** · Piano Annuale dei Flussi di Cassa · Scuole Italiane · 2026")
