# -*- coding: utf-8 -*-
"""
Remove a LSU (E. J. Ourso College of Business, Baton Rouge, LA) -- MS in
Analytics -- do msba_programs_database.xlsx.

Motivo (20/out/2026): a página oficial do Stephenson Department diz "On-campus
Master of Science in Analytics program applications are closed. We are still
accepting applications for the Online MSA." Só o formato online continua
aberto, e ele não serve para visto F-1. A página não informa motivo nem data
de reabertura. A usuária pediu a remoção.

Fonte: https://www.lsu.edu/business/sdeis/ms-analytics.php

Uso: python scripts/etl_remove_lsu.py (a partir da raiz).
"""

import sys

import openpyxl

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"
SHEET_SUMMARY = "Resumo"
MATCH = ("LSU", "LA", "Baton Rouge")


def main():
    rotate_backup()
    print("Backup rotacionado (política fixa) em: data/msba_programs_database_backup.xlsx")

    wb = openpyxl.load_workbook(DB_PATH)
    ws = wb[SHEET_DATA]
    headers = [cell.value for cell in ws[1]]
    col = {h: i + 1 for i, h in enumerate(headers)}

    target_rows = [
        row[0].row
        for row in ws.iter_rows(min_row=2)
        if (
            row[col["Nome da Universidade"] - 1].value,
            row[col["Estado"] - 1].value,
            row[col["Cidade"] - 1].value,
        ) == MATCH
    ]
    if len(target_rows) != 1:
        raise ValueError(f"Esperava 1 linha para {MATCH}, achei {len(target_rows)}. Abortando.")

    row_idx = target_rows[0]
    program = ws.cell(row_idx, col["Nome Oficial do Programa"]).value
    ws.delete_rows(row_idx, 1)
    print(f"Linha {row_idx} removida: LSU (Baton Rouge, LA) -- {program}")

    last_col_letter = ws.cell(row=1, column=len(headers)).column_letter
    ws.auto_filter.ref = f"A1:{last_col_letter}{ws.max_row}"

    total = ws.max_row - 1
    summary = wb[SHEET_SUMMARY]
    for row in summary.iter_rows(min_row=2):
        if row[0].value in ("Total de Programas", "Programas Presenciais STEM"):
            print(f"Resumo '{row[0].value}': {row[1].value} -> {total}")
            row[1].value = total

    wb.save(DB_PATH)
    print(f"\nArquivo salvo. A base agora tem {total} programas.")


if __name__ == "__main__":
    main()
