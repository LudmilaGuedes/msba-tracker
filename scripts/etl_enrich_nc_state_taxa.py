# -*- coding: utf-8 -*-
"""
Data enrichment (isencao da taxa de aplicacao): NC State (Institute for
Advanced Analytics), Raleigh, NC -- MS in Analytics.

Os FAQs de admissao de varios programas de pos da NC State (Economics/
Poole, History, Physics, Design), que seguem a politica da Graduate School,
dizem que existe "a limited number of application fee waivers" por
dificuldade financeira, so para cidadaos americanos e residentes
permanentes, e que "application fee waivers are not available to
international students". A pagina da Graduate School (ServiceNow) nao pode
ser lida por fetch. A taxa para internacionais e US$95.

Uso: python scripts/etl_enrich_nc_state_taxa.py (a partir da raiz).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "NC State",
    "Estado": "NC",
    "Cidade": "Raleigh",
}

UPDATES = {
    "Rounds Aplicação": (
        "Prioritário (entrevistas e triagem técnica em novembro; decisão em 4/dez; "
        "aceite até 15/jan) e regular (rolling até preencher as vagas; turma fechada até "
        "o início de maio). Taxa de US$95 para internacionais, SEM isenção (só "
        "cidadãos/residentes dos EUA podem pedir isenção por dificuldade financeira)"
    ),
}


def main():
    rotate_backup()
    print("Backup rotacionado (política fixa) em: data/msba_programs_database_backup.xlsx")

    wb = openpyxl.load_workbook(DB_PATH)
    ws = wb[SHEET_DATA]

    headers = [cell.value for cell in ws[1]]
    header_to_col = {h: i + 1 for i, h in enumerate(headers)}

    unknown_cols = [h for h in UPDATES if h not in header_to_col]
    if unknown_cols:
        raise ValueError(f"Colunas inexistentes no schema: {unknown_cols}")

    col_uni = header_to_col["Nome da Universidade"]
    col_estado = header_to_col["Estado"]
    col_cidade = header_to_col["Cidade"]

    target_rows = []
    for row in ws.iter_rows(min_row=2):
        if (
            row[col_uni - 1].value == MATCH["Nome da Universidade"]
            and row[col_estado - 1].value == MATCH["Estado"]
            and row[col_cidade - 1].value == MATCH["Cidade"]
        ):
            target_rows.append(row[0].row)

    if len(target_rows) == 0:
        raise ValueError("Nenhuma linha encontrada para NC State (Raleigh, NC). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (NC State, Raleigh, NC)")

    no_fill = PatternFill(fill_type=None)

    before = {}
    after = {}
    for header, new_value in UPDATES.items():
        col_idx = header_to_col[header]
        cell = ws.cell(row=row_idx, column=col_idx)
        before[header] = cell.value
        cell.value = new_value
        cell.fill = no_fill
        after[header] = new_value

    wb.save(DB_PATH)

    print(f"\n{len(UPDATES)} campos atualizados na linha {row_idx} (NC State):\n")
    for header in UPDATES:
        print(f"  - {header}:")
        print(f"      Antes:  {before[header]!r}")
        print(f"      Depois: {after[header]!r}")


if __name__ == "__main__":
    main()
