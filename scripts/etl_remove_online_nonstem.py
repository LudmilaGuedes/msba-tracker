# -*- coding: utf-8 -*-
"""
Remove do msba_programs_database.xlsx os programas confirmados em fonte
oficial (out/2026) como 100% online ou nao-STEM. A base so admite MSBA
presencial e STEM.

- Michigan State (East Lansing, MI): o MS in Business Data Science &
  Analytics da Broad e "fully online, STEM-designated" (online.msu.edu) e
  "fully online and there is no required on campus time" (broad.msu.edu).
  Nao existe versao presencial.
- Virginia Tech (Blacksburg, VA): a pagina oficial de admissao do MSBA-BA
  (msba.vt.edu/ba/admission.html) diz "This program is not STEM certified".

Mantidos por falta de certeza: VT Global BA (STEM nao declarado em nenhuma
pagina oficial acessivel), Texas A&M (aulas ao vivo presenciais ou online),
Boston College (M.S. Applied Analytics, STEM, presencial ou online) e
UCF (Professional) (admissao suspensa, mas nao e online).

Uso: python scripts/etl_remove_online_nonstem.py (a partir da raiz).
"""

import sys
from datetime import datetime

import openpyxl

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"
SHEET_SUMMARY = "Resumo"
TARGETS = [
    ("Michigan State", "MI", "East Lansing"),
    ("Virginia Tech", "VA", "Blacksburg"),
]


def main():
    rotate_backup()
    print("Backup rotacionado (política fixa) em: data/msba_programs_database_backup.xlsx")

    wb = openpyxl.load_workbook(DB_PATH)
    ws = wb[SHEET_DATA]
    col = {c.value: i + 1 for i, c in enumerate(ws[1])}

    def key(r):
        return (r[col["Nome da Universidade"] - 1].value, r[col["Estado"] - 1].value, r[col["Cidade"] - 1].value)

    to_delete = []
    for target in TARGETS:
        rows = [r[0].row for r in ws.iter_rows(min_row=2) if key(r) == target]
        if len(rows) != 1:
            raise ValueError(f"Esperava 1 linha para {target}, encontrou {len(rows)}. Abortando.")
        to_delete.append((rows[0], target))

    # De baixo para cima, para os índices das linhas restantes não mudarem.
    for row, target in sorted(to_delete, reverse=True):
        print(f"Removendo linha {row}: {target[0]} ({target[2]}, {target[1]})")
        ws.delete_rows(row)
    ws.auto_filter.ref = f"A1:{ws.cell(row=1, column=ws.max_column).column_letter}{ws.max_row}"

    total = ws.max_row - 1
    for srow in wb[SHEET_SUMMARY].iter_rows(min_row=2):
        if srow[0].value in ("Total de Programas", "Programas Presenciais STEM"):
            srow[1].value = total
        if srow[0].value == "Data de Compilação":
            srow[1].value = f"{srow[1].value} - Michigan State e Virginia Tech removidas em {datetime.now().strftime('%d/%m/%Y')}"

    wb.save(DB_PATH)
    print(f"Total de programas na base: {total}")


if __name__ == "__main__":
    main()
