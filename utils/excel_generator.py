"""
Generazione file Excel Piano dei Flussi di Cassa.
Produce 3 fogli: ENTRATE, SPESE, PIANO FLUSSI (formato MIM).
"""
import io
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter

# ─── COLORS ──────────────────────────────────────────────
COLOR_HEADER_BLUE = "1F4E79"       # Dark blue header
COLOR_HEADER_LIGHT = "BDD7EE"      # Light blue subheader
COLOR_ENTRATE = "E2EFDA"           # Light green - entrate
COLOR_SPESE = "FCE4D6"             # Light orange - spese
COLOR_RESIDUI = "FFF2CC"           # Light yellow - residui
COLOR_TOTALE = "D6DCE4"            # Grey - totals
COLOR_ANOMALIA = "FF9900"          # Orange - voci programmaz=0 con movimenti
COLOR_PIANO_HEADER = "203864"      # Very dark blue for piano flussi
COLOR_PIANO_INCASSO = "DEEAF1"     # Light for piano flussi rows

FONT_HEADER = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
FONT_SUBHEADER = Font(name="Calibri", bold=True, size=10)
FONT_NORMAL = Font(name="Calibri", size=10)
FONT_BOLD = Font(name="Calibri", bold=True, size=10)
FONT_ANOMALIA = Font(name="Calibri", bold=True, color="C00000", size=10)

BORDER_THIN = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin")
)
BORDER_MEDIUM = Border(
    left=Side(style="medium"), right=Side(style="medium"),
    top=Side(style="medium"), bottom=Side(style="medium")
)

NUM_EURO = '#,##0.00 €'
NUM_PERC = '0%'


def _fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)


def _set_col_widths(ws, widths: dict):
    for col_letter, width in widths.items():
        ws.column_dimensions[col_letter].width = width


def _header_row(ws, row, values, fill_color, font=None, height=18):
    ws.row_dimensions[row].height = height
    for col, val in enumerate(values, 1):
        cell = ws.cell(row=row, column=col, value=val)
        cell.fill = _fill(fill_color)
        cell.font = font or FONT_HEADER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER_THIN


def _data_row(ws, row, values, fill_color=None, font=None, formats=None, anomalia=False):
    ws.row_dimensions[row].height = 16
    for col, val in enumerate(values, 1):
        cell = ws.cell(row=row, column=col, value=val)
        if fill_color:
            cell.fill = _fill(fill_color)
        if anomalia:
            cell.font = FONT_ANOMALIA
        else:
            cell.font = font or FONT_NORMAL
        cell.border = BORDER_THIN
        if formats and col - 1 < len(formats) and formats[col - 1]:
            cell.number_format = formats[col - 1]
        # Align numbers right
        if isinstance(val, (int, float)):
            cell.alignment = Alignment(horizontal="right", vertical="center")
        else:
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=False)


def calcola_flussi(voce: dict, percentuali: dict, tipo: str, modalita: str) -> tuple:
    """
    Calcola Gen-Ago e Set-Dic per una voce.
    tipo: 'entrata_comp' | 'spesa_comp' | 'residuo_attivo' | 'residuo_passivo'
    Restituisce (gen_ago, set_dic)
    """
    if tipo == "entrata_comp":
        programmaz = voce.get("previsione_definitiva", 0.0)
        gia_incassato = voce.get("somme_riscosse", 0.0)
        pct_ga = percentuali.get("entrate_competenza_gen_ago", 100) / 100
        pct_sd = percentuali.get("entrate_competenza_set_dic", 0) / 100
    elif tipo == "spesa_comp":
        programmaz = voce.get("previsione_definitiva", 0.0)
        gia_incassato = voce.get("somme_pagate", 0.0)
        pct_ga = percentuali.get("spese_competenza_gen_ago", 67) / 100
        pct_sd = percentuali.get("spese_competenza_set_dic", 33) / 100
    elif tipo == "residuo_attivo":
        programmaz = voce.get("importo", 0.0)
        gia_incassato = 0.0
        pct_ga = percentuali.get("residui_attivi_gen_ago", 33) / 100
        pct_sd = percentuali.get("residui_attivi_set_dic", 67) / 100
    else:  # residuo_passivo
        programmaz = voce.get("importo", 0.0)
        gia_incassato = 0.0
        pct_ga = percentuali.get("residui_passivi_gen_ago", 33) / 100
        pct_sd = percentuali.get("residui_passivi_set_dic", 67) / 100

    anomalia = (programmaz == 0 and gia_incassato > 0)

    if anomalia:
        return gia_incassato, 0.0, True

    differenza = programmaz - gia_incassato
    if differenza < 0:
        differenza = 0.0

    gen_ago = round(gia_incassato + differenza * pct_ga, 2)
    set_dic = round(differenza * pct_sd, 2)
    return gen_ago, set_dic, False


def crea_foglio_entrate(wb, dati, percentuali, modalita):
    ws = wb.create_sheet("ENTRATE")
    _set_col_widths(ws, {"A": 12, "B": 45, "C": 18, "D": 18, "E": 18, "F": 18, "G": 18})

    row = 1
    # Title
    ws.merge_cells(f"A{row}:G{row}")
    cell = ws.cell(row=row, column=1, value="PIANO DEI FLUSSI DI CASSA — ENTRATE")
    cell.font = Font(name="Calibri", bold=True, size=13, color="FFFFFF")
    cell.fill = _fill(COLOR_HEADER_BLUE)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[row].height = 24

    row += 1
    headers = ["Codice", "Descrizione voce", "Programm. Definitiva", "Già Riscosso",
               "Gen–Ago", "Set–Dic", "TOTALE"]
    _header_row(ws, row, headers, COLOR_HEADER_BLUE)

    fmt = [None, None, NUM_EURO, NUM_EURO, NUM_EURO, NUM_EURO, NUM_EURO]

    # ── SEZIONE 1: Entrate in competenza ──────────────────
    row += 1
    ws.merge_cells(f"A{row}:G{row}")
    cell = ws.cell(row=row, column=1, value="ENTRATE IN COMPETENZA (Modello H)")
    cell.font = FONT_SUBHEADER
    cell.fill = _fill(COLOR_ENTRATE)
    cell.alignment = Alignment(horizontal="left", vertical="center")
    cell.border = BORDER_THIN
    ws.row_dimensions[row].height = 18

    totale_prog_comp = 0
    totale_gia_comp = 0
    totale_ga_comp = 0
    totale_sd_comp = 0

    for voce in dati.get("entrate", []):
        codice = voce.get("codice", "")
        # Skip E1/1 and E1/2 (avanzo di amministrazione - excluded)
        if codice in ["E1/1", "E1/2"]:
            continue

        row += 1
        prog = voce.get("previsione_definitiva", 0.0)
        gia = voce.get("somme_riscosse", 0.0)

        if modalita == "automatica":
            gen_ago, set_dic, anomalia = calcola_flussi(voce, percentuali, "entrata_comp", modalita)
        else:
            # Manual: precompile with already collected, leave rest for user
            gen_ago = gia
            set_dic = 0.0
            anomalia = (prog == 0 and gia > 0)

        totale_row = gen_ago + set_dic

        _data_row(ws, row,
                  [codice, voce.get("descrizione", ""), prog, gia, gen_ago, set_dic, totale_row],
                  fill_color=COLOR_ANOMALIA if anomalia else None,
                  formats=fmt, anomalia=anomalia)

        totale_prog_comp += prog
        totale_gia_comp += gia
        totale_ga_comp += gen_ago
        totale_sd_comp += set_dic

    # Subtotal entrate competenza
    row += 1
    _data_row(ws, row,
              ["", "TOTALE ENTRATE COMPETENZA",
               totale_prog_comp, totale_gia_comp,
               totale_ga_comp, totale_sd_comp,
               totale_ga_comp + totale_sd_comp],
              fill_color=COLOR_TOTALE, font=FONT_BOLD, formats=fmt)

    # ── SEZIONE 2: Residui attivi ──────────────────────────
    row += 2
    ws.merge_cells(f"A{row}:G{row}")
    cell = ws.cell(row=row, column=1, value="RESIDUI ATTIVI — crediti anni precedenti (Modello L)")
    cell.font = FONT_SUBHEADER
    cell.fill = _fill(COLOR_RESIDUI)
    cell.alignment = Alignment(horizontal="left", vertical="center")
    cell.border = BORDER_THIN
    ws.row_dimensions[row].height = 18

    totale_prog_res = 0
    totale_ga_res = 0
    totale_sd_res = 0

    for residuo in dati.get("residui_attivi", []):
        row += 1
        importo = residuo.get("importo", 0.0)

        if modalita == "automatica":
            gen_ago, set_dic, anomalia = calcola_flussi(residuo, percentuali, "residuo_attivo", modalita)
        else:
            gen_ago = 0.0
            set_dic = 0.0
            anomalia = False

        desc = f"{residuo.get('anno', '')} - {residuo.get('oggetto', residuo.get('debitore', ''))}"
        _data_row(ws, row,
                  [residuo.get("codice_pdc", ""), desc, importo, 0.0, gen_ago, set_dic, gen_ago + set_dic],
                  fill_color=COLOR_RESIDUI if not anomalia else COLOR_ANOMALIA,
                  formats=fmt, anomalia=anomalia)

        totale_prog_res += importo
        totale_ga_res += gen_ago
        totale_sd_res += set_dic

    # Subtotal residui attivi
    row += 1
    _data_row(ws, row,
              ["", "TOTALE RESIDUI ATTIVI",
               totale_prog_res, 0.0,
               totale_ga_res, totale_sd_res,
               totale_ga_res + totale_sd_res],
              fill_color=COLOR_TOTALE, font=FONT_BOLD, formats=fmt)

    # ── GRAND TOTAL ENTRATE ────────────────────────────────
    row += 2
    gt_vals = [
        "TOTALE GENERALE ENTRATE",
        totale_prog_comp + totale_prog_res,
        totale_gia_comp,
        totale_ga_comp + totale_ga_res,
        totale_sd_comp + totale_sd_res,
        totale_ga_comp + totale_ga_res + totale_sd_comp + totale_sd_res,
    ]
    gt_cols = [1, 3, 4, 5, 6, 7]
    gt_fmts = [None, NUM_EURO, NUM_EURO, NUM_EURO, NUM_EURO, NUM_EURO]
    ws.row_dimensions[row].height = 18
    for idx, (col, val) in enumerate(zip(gt_cols, gt_vals)):
        cell = ws.cell(row=row, column=col, value=val)
        cell.fill = _fill(COLOR_HEADER_BLUE)
        cell.font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
        cell.border = BORDER_THIN
        cell.number_format = gt_fmts[idx] or "General"
        cell.alignment = Alignment(horizontal="right" if isinstance(val, float) else "left", vertical="center")
    # Merge col 1-2 for label
    ws.merge_cells(f"A{row}:B{row}")

    if modalita == "manuale":
        row += 2
        ws.merge_cells(f"A{row}:G{row}")
        cell = ws.cell(row=row, column=1,
                       value="⚠️ MODALITÀ MANUALE: inserire i valori Gen-Ago e Set-Dic per ogni voce. "
                             "Gen-Ago non può essere inferiore al 'Già Riscosso'. "
                             "Le celle evidenziate in arancione hanno Programmazione=0 con movimenti.")
        cell.font = Font(name="Calibri", bold=True, color="C00000", size=10)
        cell.fill = _fill("FFFFC0")
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws.row_dimensions[row].height = 32

    return {
        "totale_entrate_ga": totale_ga_comp + totale_ga_res,
        "totale_entrate_sd": totale_sd_comp + totale_sd_res,
    }


def crea_foglio_spese(wb, dati, percentuali, modalita):
    ws = wb.create_sheet("SPESE")
    _set_col_widths(ws, {"A": 12, "B": 45, "C": 18, "D": 18, "E": 18, "F": 18, "G": 18})

    row = 1
    ws.merge_cells(f"A{row}:G{row}")
    cell = ws.cell(row=row, column=1, value="PIANO DEI FLUSSI DI CASSA — SPESE")
    cell.font = Font(name="Calibri", bold=True, size=13, color="FFFFFF")
    cell.fill = _fill("C00000")
    cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[row].height = 24

    row += 1
    headers = ["Codice", "Descrizione voce", "Programm. Definitiva", "Già Pagato",
               "Gen–Ago", "Set–Dic", "TOTALE"]
    _header_row(ws, row, headers, "C00000")

    fmt = [None, None, NUM_EURO, NUM_EURO, NUM_EURO, NUM_EURO, NUM_EURO]

    # ── SEZIONE 1: Spese in competenza ────────────────────
    row += 1
    ws.merge_cells(f"A{row}:G{row}")
    cell = ws.cell(row=row, column=1, value="SPESE IN COMPETENZA (Modello H)")
    cell.font = FONT_SUBHEADER
    cell.fill = _fill(COLOR_SPESE)
    cell.alignment = Alignment(horizontal="left", vertical="center")
    cell.border = BORDER_THIN
    ws.row_dimensions[row].height = 18

    totale_prog_comp = 0
    totale_gia_comp = 0
    totale_ga_comp = 0
    totale_sd_comp = 0

    for voce in dati.get("spese", []):
        row += 1
        prog = voce.get("previsione_definitiva", 0.0)
        gia = voce.get("somme_pagate", 0.0)

        if modalita == "automatica":
            gen_ago, set_dic, anomalia = calcola_flussi(voce, percentuali, "spesa_comp", modalita)
        else:
            gen_ago = gia
            set_dic = 0.0
            anomalia = (prog == 0 and gia > 0)

        _data_row(ws, row,
                  [voce.get("aggregato", ""), voce.get("descrizione", ""),
                   prog, gia, gen_ago, set_dic, gen_ago + set_dic],
                  fill_color=COLOR_ANOMALIA if anomalia else None,
                  formats=fmt, anomalia=anomalia)

        totale_prog_comp += prog
        totale_gia_comp += gia
        totale_ga_comp += gen_ago
        totale_sd_comp += set_dic

    row += 1
    _data_row(ws, row,
              ["", "TOTALE SPESE COMPETENZA",
               totale_prog_comp, totale_gia_comp,
               totale_ga_comp, totale_sd_comp,
               totale_ga_comp + totale_sd_comp],
              fill_color=COLOR_TOTALE, font=FONT_BOLD, formats=fmt)

    # ── SEZIONE 2: Residui passivi ─────────────────────────
    row += 2
    ws.merge_cells(f"A{row}:G{row}")
    cell = ws.cell(row=row, column=1, value="RESIDUI PASSIVI — debiti anni precedenti (Modello L)")
    cell.font = FONT_SUBHEADER
    cell.fill = _fill(COLOR_RESIDUI)
    cell.alignment = Alignment(horizontal="left", vertical="center")
    cell.border = BORDER_THIN
    ws.row_dimensions[row].height = 18

    totale_prog_res = 0
    totale_ga_res = 0
    totale_sd_res = 0

    for residuo in dati.get("residui_passivi", []):
        row += 1
        importo = residuo.get("importo", 0.0)

        if modalita == "automatica":
            gen_ago, set_dic, anomalia = calcola_flussi(residuo, percentuali, "residuo_passivo", modalita)
        else:
            gen_ago = 0.0
            set_dic = 0.0
            anomalia = False

        desc = f"{residuo.get('anno', '')} - {residuo.get('oggetto', residuo.get('creditore', ''))}"
        _data_row(ws, row,
                  [residuo.get("codice_pdc", ""), desc, importo, 0.0, gen_ago, set_dic, gen_ago + set_dic],
                  fill_color=COLOR_RESIDUI if not anomalia else COLOR_ANOMALIA,
                  formats=fmt, anomalia=anomalia)

        totale_prog_res += importo
        totale_ga_res += gen_ago
        totale_sd_res += set_dic

    row += 1
    _data_row(ws, row,
              ["", "TOTALE RESIDUI PASSIVI",
               totale_prog_res, 0.0,
               totale_ga_res, totale_sd_res,
               totale_ga_res + totale_sd_res],
              fill_color=COLOR_TOTALE, font=FONT_BOLD, formats=fmt)

    row += 2
    gt_vals = [
        "TOTALE GENERALE SPESE",
        totale_prog_comp + totale_prog_res,
        totale_gia_comp,
        totale_ga_comp + totale_ga_res,
        totale_sd_comp + totale_sd_res,
        totale_ga_comp + totale_ga_res + totale_sd_comp + totale_sd_res,
    ]
    gt_cols = [1, 3, 4, 5, 6, 7]
    gt_fmts = [None, NUM_EURO, NUM_EURO, NUM_EURO, NUM_EURO, NUM_EURO]
    ws.row_dimensions[row].height = 18
    for idx, (col, val) in enumerate(zip(gt_cols, gt_vals)):
        cell = ws.cell(row=row, column=col, value=val)
        cell.fill = _fill("C00000")
        cell.font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
        cell.border = BORDER_THIN
        cell.number_format = gt_fmts[idx] or "General"
        cell.alignment = Alignment(horizontal="right" if isinstance(val, float) else "left", vertical="center")
    ws.merge_cells(f"A{row}:B{row}")

    if modalita == "manuale":
        row += 2
        ws.merge_cells(f"A{row}:G{row}")
        cell = ws.cell(row=row, column=1,
                       value="⚠️ MODALITÀ MANUALE: inserire i valori Gen-Ago e Set-Dic per ogni voce.")
        cell.font = Font(name="Calibri", bold=True, color="C00000", size=10)
        cell.fill = _fill("FFFFC0")
        cell.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[row].height = 22

    return {
        "totale_spese_ga": totale_ga_comp + totale_ga_res,
        "totale_spese_sd": totale_sd_comp + totale_sd_res,
    }


def crea_foglio_piano_flussi(wb, dati_scuola, totali, fondo_cassa):
    ws = wb.create_sheet("PIANO FLUSSI (MIM)")
    _set_col_widths(ws, {"A": 50, "B": 22, "C": 22, "D": 22})

    anno = dati_scuola.get("anno_esercizio", "2026")
    nome = dati_scuola.get("nome_istituto", "ISTITUTO")

    row = 1
    ws.merge_cells(f"A{row}:D{row}")
    cell = ws.cell(row=row, column=1,
                   value=f"PIANO ANNUALE DEI FLUSSI DI CASSA — ESERCIZIO {anno}")
    cell.font = Font(name="Calibri", bold=True, size=14, color="FFFFFF")
    cell.fill = _fill(COLOR_PIANO_HEADER)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[row].height = 28

    row += 1
    ws.merge_cells(f"A{row}:D{row}")
    cell = ws.cell(row=row, column=1, value=nome)
    cell.font = Font(name="Calibri", bold=True, size=11)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[row].height = 20

    row += 2
    headers = ["VOCE", f"Gen–Ago {anno}", f"Set–Dic {anno}", f"TOTALE {anno}"]
    _header_row(ws, row, headers, COLOR_PIANO_HEADER, height=20)

    fmt = [None, NUM_EURO, NUM_EURO, NUM_EURO]

    def piano_row(ws, r, label, ga, sd, fill=None, bold=False, indent=0):
        label_full = ("  " * indent) + label
        totale = ga + sd
        _data_row(ws, r, [label_full, ga, sd, totale],
                  fill_color=fill, font=FONT_BOLD if bold else FONT_NORMAL, formats=fmt)

    # Fondo cassa iniziale
    row += 1
    ws.row_dimensions[row].height = 18
    cell = ws.cell(row=row, column=1, value=f"Fondo cassa al 01/01/{anno}")
    cell.font = FONT_BOLD
    cell.fill = _fill(COLOR_HEADER_LIGHT)
    cell.border = BORDER_THIN
    cell.alignment = Alignment(horizontal="left", vertical="center")
    for col in range(2, 5):
        c = ws.cell(row=row, column=col, value=fondo_cassa if col == 2 else 0.0)
        c.font = FONT_BOLD
        c.fill = _fill(COLOR_HEADER_LIGHT)
        c.border = BORDER_THIN
        c.number_format = NUM_EURO
        c.alignment = Alignment(horizontal="right", vertical="center")

    # ── ENTRATE ───────────────────────────────────────────
    row += 1
    ws.merge_cells(f"A{row}:D{row}")
    cell = ws.cell(row=row, column=1, value="ENTRATE")
    cell.font = Font(name="Calibri", bold=True, size=11, color="FFFFFF")
    cell.fill = _fill(COLOR_HEADER_BLUE)
    cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[row].height = 20

    ent_ga = totali.get("totale_entrate_ga", 0)
    ent_sd = totali.get("totale_entrate_sd", 0)
    row += 1
    piano_row(ws, row, "Totale entrate (competenza + residui)", ent_ga, ent_sd,
              fill=COLOR_ENTRATE, bold=True)

    # ── SPESE ─────────────────────────────────────────────
    row += 1
    ws.merge_cells(f"A{row}:D{row}")
    cell = ws.cell(row=row, column=1, value="SPESE")
    cell.font = Font(name="Calibri", bold=True, size=11, color="FFFFFF")
    cell.fill = _fill("C00000")
    cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[row].height = 20

    spe_ga = totali.get("totale_spese_ga", 0)
    spe_sd = totali.get("totale_spese_sd", 0)
    row += 1
    piano_row(ws, row, "Totale spese (competenza + residui)", spe_ga, spe_sd,
              fill=COLOR_SPESE, bold=True)

    # ── SALDI STIMATI ─────────────────────────────────────
    row += 2
    ws.merge_cells(f"A{row}:D{row}")
    cell = ws.cell(row=row, column=1, value="SALDI STIMATI")
    cell.font = Font(name="Calibri", bold=True, size=11, color="FFFFFF")
    cell.fill = _fill(COLOR_PIANO_HEADER)
    cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[row].height = 20

    saldo_agosto = fondo_cassa + ent_ga - spe_ga
    saldo_dicembre = saldo_agosto + ent_sd - spe_sd

    row += 1
    _data_row(ws, row,
              [f"Saldo di cassa stimato al 31/08/{anno}", saldo_agosto, "", saldo_agosto],
              fill_color=COLOR_HEADER_LIGHT, font=FONT_BOLD, formats=fmt)

    row += 1
    _data_row(ws, row,
              [f"Saldo di cassa stimato al 31/12/{anno}", "", saldo_dicembre, saldo_dicembre],
              fill_color=COLOR_HEADER_LIGHT, font=FONT_BOLD, formats=fmt)

    # Warning if negative balance
    if saldo_agosto < 0 or saldo_dicembre < 0:
        row += 2
        ws.merge_cells(f"A{row}:D{row}")
        cell = ws.cell(row=row, column=1,
                       value="⚠️ ATTENZIONE: uno o più saldi stimati risultano NEGATIVI. Verificare i dati inseriti.")
        cell.font = Font(name="Calibri", bold=True, color="C00000", size=10)
        cell.fill = _fill("FFFFC0")
        ws.row_dimensions[row].height = 22

    row += 2
    ws.merge_cells(f"A{row}:D{row}")
    modalita_label = totali.get("modalita", "automatica").upper()
    cell = ws.cell(row=row, column=1,
                   value=f"Elaborato con modalità {modalita_label} — "
                         f"Piano Annuale dei Flussi di Cassa — Art. 6 D.L. 155/2024")
    cell.font = Font(name="Calibri", italic=True, size=9, color="808080")
    cell.alignment = Alignment(horizontal="center")
    ws.row_dimensions[row].height = 16


def genera_excel(dati, dati_scuola, percentuali, modalita) -> bytes:
    """
    Genera il file Excel completo e lo restituisce come bytes.
    """
    wb = Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    # Create sheets
    tot_entrate = crea_foglio_entrate(wb, dati, percentuali, modalita)
    tot_spese = crea_foglio_spese(wb, dati, percentuali, modalita)

    totali = {**tot_entrate, **tot_spese, "modalita": modalita}
    fondo_cassa = float(dati_scuola.get("fondo_cassa", 0))

    crea_foglio_piano_flussi(wb, dati_scuola, totali, fondo_cassa)

    # Save to bytes
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()
