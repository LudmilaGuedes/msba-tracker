# -*- coding: utf-8 -*-
"""
Remove a linha da Xavier University (Williams College of Business,
Cincinnati, OH) do msba_programs_database.xlsx.

MOTIVO (decisao de set/2026, depois do enrichment em
etl_enrich_xavier.py):
1. Nao ha confirmacao de que a Xavier aceita Duolingo. As paginas oficiais
   de admissao internacional e de pos citam so TOEFL e IELTS; o "Duolingo
   105" veio de um snippet de busca. A varredura nacional
   (etl_add_nationwide_sweep.py) so admite programas que aceitam Duolingo.
2. O MSBA e hibrido (~30% online), nao totalmente presencial.

Uso: python scripts/etl_remove_xavier.py (a partir da raiz).
"""

import sys
from datetime import datetime

import openpyxl

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"
SHEET_SUMMARY = "Resumo"
TARGET = ("Xavier University", "OH", "Cincinnati")


def main():
    rotate_backup()
    print("Backup rotacionado (política fixa) em: data/msba_programs_database_backup.xlsx")

    wb = openpyxl.load_workbook(DB_PATH)
    ws = wb[SHEET_DATA]
    col = {c.value: i + 1 for i, c in enumerate(ws[1])}
    rows = [
        r[0].row for r in ws.iter_rows(min_row=2)
        if (r[col["Nome da Universidade"] - 1].value, r[col["Estado"] - 1].value, r[col["Cidade"] - 1].value) == TARGET
    ]
    if len(rows) != 1:
        raise ValueError(f"Esperava 1 linha para {TARGET}, encontrou {len(rows)}. Abortando.")

    print(f"Removendo linha {rows[0]}: {TARGET[0]} ({TARGET[2]}, {TARGET[1]})")
    ws.delete_rows(rows[0])
    ws.auto_filter.ref = f"A1:{ws.cell(row=1, column=ws.max_column).column_letter}{ws.max_row}"

    total = ws.max_row - 1
    for srow in wb[SHEET_SUMMARY].iter_rows(min_row=2):
        if srow[0].value in ("Total de Programas", "Programas Presenciais STEM"):
            srow[1].value = total
        if srow[0].value == "Data de Compilação":
            srow[1].value = f"{srow[1].value} - Xavier removida em {datetime.now().strftime('%d/%m/%Y')}"

    wb.save(DB_PATH)
    print(f"Total de programas na base: {total}")


if __name__ == "__main__":
    main()
