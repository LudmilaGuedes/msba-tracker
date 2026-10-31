# -*- coding: utf-8 -*-
"""
CORRECAO: a rodada em lote anterior sobrescreveu 'Início' da University
of Washington, que JA TINHA um valor real ("Verão (junho)", nao
"Verificar"), com uma versao reformulada ("Junho (turma única por ano)")
que perdeu a classificacao original de estacao ("Verão"). Mescla as duas
versoes em vez de substituir.
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {"Nome da Universidade": "University of Washington"}

UPDATES = {
    "Início": "Verão (junho) -- turma única por ano",
}


def main():
    rotate_backup()
    print("Backup rotacionado (política fixa) em: data/msba_programs_database_backup.xlsx")

    wb = openpyxl.load_workbook(DB_PATH)
    ws = wb[SHEET_DATA]
    headers = [cell.value for cell in ws[1]]
    header_to_col = {h: i + 1 for i, h in enumerate(headers)}
    col_uni = header_to_col["Nome da Universidade"]

    target_rows = [row[0].row for row in ws.iter_rows(min_row=2) if row[col_uni - 1].value == MATCH["Nome da Universidade"]]
    if len(target_rows) != 1:
        raise ValueError(f"Esperado 1 linha para University of Washington, achou {len(target_rows)}. Abortando.")

    row_idx = target_rows[0]
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
    print("\nCampo corrigido:")
    for header in UPDATES:
        print(f"  - {header}: {before[header]!r} -> {after[header]!r}")


if __name__ == "__main__":
    main()
