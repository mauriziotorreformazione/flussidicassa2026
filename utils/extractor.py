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
Analizza questo documento (Modello H+ – Conto Finanziario) e estrai TUTTI i dati in formato JSON.

Il documento contiene due sezioni principali:
1. ENTRATE: voci con codice (es. E1/1, E2/1...), descrizione, Previsione Definitiva, Somme Accertate, Somme Riscosse
2. SPESE: voci con aggregato/progetto (es. A01, A02, P01...), descrizione, Previsione Definitiva, Somme Impegnate, Somme Pagate

Estrai anche dall'intestazione: nome istituto, codice fiscale, codice meccanografico, indirizzo, città.

REGOLE CRITICHE PER IL JSON:
- Rispondi SOLO con JSON valido, nessun testo prima o dopo
- NON usare virgolette nei valori numerici
- NON lasciare virgole finali prima di } o ]
- Tutti gli importi devono essere numeri float (es. 45000.00 non "45000,00")
- Se un importo ha la virgola come separatore decimale, convertilo con il punto (es. 1.234,56 → 1234.56)
- Se un valore è assente o trattino, usa 0.0

Formato risposta:
{"nome_istituto":"...","cf":"...","codice_mecc":"...","indirizzo":"...","citta":"...","anno":"...","entrate":[{"codice":"E2/1","descrizione":"Finanziamenti UE - FSE","previsione_definitiva":45000.00,"somme_accertate":45000.00,"somme_riscosse":12000.00,"codice_pdc":"E.2.01.05.01.005"}],"spese":[{"aggregato":"A01","descrizione":"Funzionamento generale","previsione_definitiva":8000.00,"somme_impegnate":7500.00,"somme_pagate":5000.00,"codice_pdc":"U.1.03.02.09.999"}]}

Includi TUTTE le voci presenti, anche quelle con importi zero.
Per il codice PDC usa la mappatura: E1/1=escluso, E2/1=E.2.01.05.01.005, E3/1=E.2.01.01.01.001, E6/x=E.2.01.02.01.001, E12/2=E.3.03.03.04.001, E99/x=E.9.01.99.99.999
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
