"""
Estrazione dati da Modello H e Modello L tramite Claude API.
Restituisce dizionario strutturato con entrate, spese, residui.
"""
import base64
import json
import re
import anthropic
import streamlit as st


def get_client():
    api_key = st.secrets.get("ANTHROPIC_API_KEY") or st.secrets.get("anthropic_api_key")
    if not api_key:
        raise ValueError("Chiave API Anthropic non trovata nei secrets di Streamlit.")
    return anthropic.Anthropic(api_key=api_key)


def pdf_to_base64(pdf_bytes: bytes) -> str:
    return base64.standard_b64encode(pdf_bytes).decode("utf-8")


PROMPT_MODELLO_H = """Sei un esperto di contabilità scolastica italiana.
Analizza questo documento (Modello H – Conto Consuntivo Conto Finanziario) e estrai i dati in JSON.

Il documento ha due sezioni: ENTRATE e SPESE.

════════════════════════════════════════
SEZIONE ENTRATE — REGOLE PRECISE
════════════════════════════════════════
La tabella ha le colonne: Liv.1 | Liv.2 | ENTRATE (descrizione) | Programmaz.definitiva (col.a) | Somme accertate (col.b) | Somme riscosse (col.c) | Somme rimaste da riscuotere (col.d) | differenza (col.e)

REGOLE OBBLIGATORIE PER LE ENTRATE:
1. Estrai SOLO le righe con valore in colonna LIV.2 (es. 01, 02, 03...) — NON estrarre le righe LIV.1 (es. 01, 02, 03 senza sottovoce) perché sono totali aggregati che causerebbero doppi conteggi
2. La colonna "previsione_definitiva" = colonna (a) = "Programmaz. definitiva" — NON confonderla con altre colonne
3. La colonna "somme_riscosse" = colonna (c) = "Somme riscosse" — è la TERZA colonna numerica
4. IGNORA completamente la colonna (b) "Somme accertate" — non ci serve
5. ESCLUDI la voce 01/01 "AVANZO NON VINCOLATO" e 01/02 "VINCOLATO" — l'avanzo di amministrazione non va nel piano flussi
6. Se una cella è vuota o con trattino, usa 0.0
7. Converti importi italiani in float: "11.854,66" → 11854.66

ESEMPIO ENTRATE dal documento:
- Riga Liv1=03, Liv2=01, desc="DOTAZIONE ORDINARIA", col.a=11854.66, col.c=11854.66 → previsione_definitiva=11854.66, somme_riscosse=11854.66
- Riga Liv1=06, Liv2=01, desc="CONTRIBUTI VOLONTARI DA FAMIGLIE", col.a=0.0, col.c=673.00 → previsione_definitiva=0.0, somme_riscosse=673.00

════════════════════════════════════════
SEZIONE SPESE — REGOLE PRECISE
════════════════════════════════════════
La tabella spese ha: Liv.1 | Liv.2 | SPESE | Programmaz.definitiva (col.a) | Somme impegnate (col.b) | Somme pagate (col.c) | Somme rimaste da pagare (col.d) | differenza (col.e)

REGOLE OBBLIGATORIE PER LE SPESE:
1. Estrai SOLO le righe con codice aggregato di SECONDO LIVELLO: A01, A02, A03, A04, A05, A06, P01, P02, P03, P04, P05, R98
2. NON estrarre le righe di primo livello A, P, G, R, D — sono totali aggregati
3. Ogni aggregato (A01, A02, ecc.) deve avere i propri importi dalla propria riga, NON dai totali superiori
4. La colonna "previsione_definitiva" = colonna (a)
5. La colonna "somme_pagate" = colonna (c) = "Somme pagate"
6. IGNORA colonna (b) "Somme impegnate" — non ci serve
7. Se una cella è vuota, usa 0.0
8. Converti importi italiani in float

ESEMPIO SPESE dal documento:
- A01 "FUNZIONAMENTO GENERALE E DECORO DELLA SCUOLA" col.a=32362.05, col.c=2672.90 → previsione_definitiva=32362.05, somme_pagate=2672.90
- A02 "FUNZIONAMENTO AMMINISTRATIVO" col.a=25854.66, col.c=16.00
- P01 "PROGETTI IN AMBITO SCIENTIFICO" col.a=2154.65, col.c=0.0
- P02 "PROGETTI IN AMBITO UMANISTICO" col.a=164170.05, col.c=3452.55
- R98 "FONDO DI RISERVA" col.a=1000.00, col.c=0.0

════════════════════════════════════════
REGOLE JSON
════════════════════════════════════════
- Rispondi SOLO con JSON valido, zero testo prima o dopo, zero markdown
- Nessuna virgola finale prima di } o ]
- Tutti gli importi = numeri float con punto decimale
- Valori assenti = 0.0

Formato esatto da restituire:
{"nome_istituto":"...","cf":"...","codice_mecc":"...","indirizzo":"...","citta":"...","anno":"2026","entrate":[{"codice":"03/01","descrizione":"DOTAZIONE ORDINARIA","previsione_definitiva":11854.66,"somme_riscosse":11854.66,"codice_pdc":"E.2.01.01.01.001"}],"spese":[{"aggregato":"A01","descrizione":"FUNZIONAMENTO GENERALE E DECORO DELLA SCUOLA","previsione_definitiva":32362.05,"somme_pagate":2672.90,"codice_pdc":"U.1.03.02.09.999"}]}

Mappatura codici PDC entrate:
02/01=E.2.01.05.01.005, 02/02=E.2.01.05.01.004, 02/03=E.2.01.05.01.999,
03/01=E.2.01.01.01.001, 03/02=E.2.01.01.01.002, 03/03=E.2.01.01.02.001,
04/01=E.2.01.03.01.001, 05/06=E.2.01.04.01.999,
06/01=E.2.01.02.01.001, 06/02=E.2.01.02.01.001, 06/03=E.2.01.02.01.001,
06/04=E.2.01.02.01.001, 06/10=E.2.01.02.01.002,
12/02=E.3.03.03.04.001, 12/03=E.3.03.99.99.999,
99/01=E.9.01.99.99.999

Mappatura codici PDC spese:
A01=U.1.03.02.09.999, A02=U.1.03.02.09.999, A03=U.1.03.02.09.999,
A04=U.1.03.02.09.999, A05=U.1.03.02.09.999, A06=U.1.03.02.09.999,
P01=U.1.10.99.99.999, P02=U.1.10.99.99.999, P03=U.1.10.99.99.999,
P04=U.1.10.99.99.999, P05=U.1.10.99.99.999,
R98=U.1.99.99.99.999
"""

PROMPT_MODELLO_L = """Sei un esperto di contabilità scolastica italiana.
Analizza questo documento (Modello L – Elenco Residui) e estrai TUTTI i dati in formato JSON.

Il documento contiene due sezioni:
1. RESIDUI ATTIVI: crediti verso terzi (entrate non ancora riscosse da anni precedenti)
2. RESIDUI PASSIVI: debiti verso fornitori (spese impegnate ma non ancora pagate da anni precedenti)

REGOLE CRITICHE PER IL JSON:
- Rispondi SOLO con JSON valido, nessun testo prima o dopo
- NON usare virgolette nei valori numerici
- NON lasciare virgole finali prima di } o ]
- Tutti gli importi devono essere numeri float con il punto decimale (es. 15000.00)
- Se un importo ha la virgola come separatore decimale, convertilo (es. 1.234,56 → 1234.56)

Formato risposta:
{"residui_attivi":[{"anno":"2024","numero":"1","data":"15/03/2024","livello1":"E.2","livello2":"E.2.01","livello3":"E.2.01.05","debitore":"MINISTERO ISTRUZIONE","oggetto":"FSE - PON 2024","importo":15000.00,"codice_pdc":"E.2.01.05.01.005"}],"residui_passivi":[{"anno":"2024","numero":"5","data":"10/06/2024","livello1":"U.1","livello2":"U.1.03","livello3":"U.1.03.02","creditore":"FORNITORE SRL","oggetto":"Acquisto materiale didattico","importo":800.00,"codice_pdc":"U.1.03.02.09.999"}],"totale_residui_attivi":0.0,"totale_residui_passivi":0.0}

Includi TUTTI i residui presenti nel documento.
"""


def parse_json_response(text: str) -> dict:
    """Estrae JSON dalla risposta, rimuovendo eventuali markdown."""
    text = text.strip()
    # Remove markdown code blocks if present
    text = re.sub(r'^```(?:json)?\s*', '', text)
    text = re.sub(r'\s*```$', '', text)
    text = text.strip()
    
    # Try direct parse first
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass
    
    # Try to extract JSON object from text
    try:
        start = text.index('{')
        end = text.rindex('}') + 1
        return json.loads(text[start:end])
    except (ValueError, json.JSONDecodeError):
        pass
    
    # Try json_repair approach: fix common issues
    # Remove trailing commas before } or ]
    text_fixed = re.sub(r',\s*([}\]])', r'\1', text)
    # Fix unescaped newlines in strings
    text_fixed = re.sub(r'(?<!\\)\n', ' ', text_fixed)
    try:
        return json.loads(text_fixed)
    except json.JSONDecodeError:
        pass

    # Last resort: ask Claude to fix it by returning empty structure
    return {
        "nome_istituto": "", "cf": "", "codice_mecc": "",
        "indirizzo": "", "citta": "", "anno": "2026",
        "entrate": [], "spese": [],
        "residui_attivi": [], "residui_passivi": [],
        "totale_residui_attivi": 0.0, "totale_residui_passivi": 0.0,
        "_parse_error": True
    }


def estrai_modello_h(client, pdf_bytes: bytes) -> dict:
    """Estrae dati dal Modello H tramite Claude API."""
    b64 = pdf_to_base64(pdf_bytes)
    
    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=8000,
        messages=[
            {
                "role": "user",
                "content": [
                    {
                        "type": "document",
                        "source": {
                            "type": "base64",
                            "media_type": "application/pdf",
                            "data": b64,
                        }
                    },
                    {
                        "type": "text",
                        "text": PROMPT_MODELLO_H
                    }
                ]
            }
        ]
    )
    
    text = response.content[0].text
    return parse_json_response(text)


def estrai_modello_l(client, pdf_bytes: bytes) -> dict:
    """Estrae dati dal Modello L tramite Claude API."""
    b64 = pdf_to_base64(pdf_bytes)
    
    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=8000,
        messages=[
            {
                "role": "user",
                "content": [
                    {
                        "type": "document",
                        "source": {
                            "type": "base64",
                            "media_type": "application/pdf",
                            "data": b64,
                        }
                    },
                    {
                        "type": "text",
                        "text": PROMPT_MODELLO_L
                    }
                ]
            }
        ]
    )
    
    text = response.content[0].text
    return parse_json_response(text)


def estrai_dati_pdf(modello_h_bytes: bytes, modello_l_bytes: bytes) -> dict:
    """
    Funzione principale: estrae dati da entrambi i PDF e li combina.
    """
    client = get_client()
    
    # Extract from both PDFs
    dati_h = estrai_modello_h(client, modello_h_bytes)
    dati_l = estrai_modello_l(client, modello_l_bytes)
    
    # Combine results
    risultato = {
        # School info from Modello H header
        "nome_istituto": dati_h.get("nome_istituto", ""),
        "cf": dati_h.get("cf", ""),
        "codice_mecc": dati_h.get("codice_mecc", ""),
        "indirizzo": dati_h.get("indirizzo", ""),
        "citta": dati_h.get("citta", ""),
        "anno": dati_h.get("anno", "2026"),
        
        # Financial data
        "entrate": dati_h.get("entrate", []),
        "spese": dati_h.get("spese", []),
        "residui_attivi": dati_l.get("residui_attivi", []),
        "residui_passivi": dati_l.get("residui_passivi", []),
        
        # Totals for verification
        "totale_residui_attivi": dati_l.get("totale_residui_attivi", 0.0),
        "totale_residui_passivi": dati_l.get("totale_residui_passivi", 0.0),
    }
    
    return risultato
