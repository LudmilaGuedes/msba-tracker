# -*- coding: utf-8 -*-
"""
Remove da base os programas adicionados na varredura nacional
(etl_add_nationwide_sweep.py) que NAO sao presenciais, e corrige o formato
de outros dois. Pedido de set/2026: manter so programas STEM presenciais.
So as 35 linhas novas foram auditadas; as antigas ja tinham passado por
esse filtro.

REMOVIDOS:
- University of Tampa (Tampa, FL): a pagina oficial do M.S. in Business
  Analytics diz "All classes are offered online". Com visto F-1 so e
  permitida 1 disciplina online por semestre, entao o curso e inviavel.
- University of St. Thomas (Minneapolis, MN): o MS in Business Analytics e
  "a hybrid of on-campus and online classes designed for busy
  professionals", no ritmo do aluno (de 16 meses a 5 anos). Nao e um
  programa presencial full-time.

CORRIGIDOS:
- Loyola University Chicago: o "MS in Business Data Analytics" (MSBDA) nao
  aceita mais inscricoes (catalog.luc.edu); o programa atual e o "Business
  Analytics (MS)", com aulas presenciais, hibridas e online no Water Tower
  Campus. Nome e formato atualizados.
- Xavier University: o MSBA e hibrido, com ~70% presencial no campus de
  Cincinnati e ~30% online (xavier.edu). Mantido, com o formato explicito.

As outras 31 linhas novas tem opcao presencial confirmada.

Uso: python scripts/etl_remove_online_programs.py (a partir da raiz).
"""

import sys
from datetime import datetime

import openpyxl

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"
SHEET_SUMMARY = "Resumo"

TO_REMOVE = [
    ("University of Tampa", "FL", "Tampa"),
    ("University of St. Thomas", "MN", "Minneapolis"),
]

TO_UPDATE = {
    ("Loyola University Chicago", "IL", "Chicago"): {
        "Nome Oficial do Programa": "MS in Business Analytics (substituiu o MS in Business Data Analytics, que não aceita mais inscrições)",
        "Formato": "Presencial no Water Tower Campus (também há turmas híbridas e online)",
    },
    ("Xavier University", "OH", "Cincinnati"): {
        "Formato": "Híbrido: ~70% presencial no campus de Cincinnati e ~30% online",
    },
}


def key_of(row, col):
    return (row[col["Nome da Universidade"] - 1].value, row[col["Estado"] - 1].value, row[col["Cidade"] - 1].value)


def main():
    rotate_backup()
    print("Backup rotacionado (política fixa) em: data/msba_programs_database_backup.xlsx")

    wb = openpyxl.load_workbook(DB_PATH)
    ws = wb[SHEET_DATA]
    col = {c.value: i + 1 for i, c in enumerate(ws[1])}

    for key, updates in TO_UPDATE.items():
        rows = [r for r in ws.iter_rows(min_row=2) if key_of(r, col) == key]
        if len(rows) != 1:
            raise ValueError(f"{key}: esperava 1 linha, encontrou {len(rows)}")
        for header, value in updates.items():
            cell = rows[0][col[header] - 1]
            print(f"\n- {key[0]} / {header}:\n    Antes:  {cell.value!r}\n    Depois: {value!r}")
            cell.value = value

    # Remove de baixo para cima para não deslocar os índices ainda não processados.
    targets = sorted(
        (r[0].row for r in ws.iter_rows(min_row=2) if key_of(r, col) in TO_REMOVE), reverse=True
    )
    if len(targets) != len(TO_REMOVE):
        raise ValueError(f"Esperava {len(TO_REMOVE)} linhas para remover, encontrou {len(targets)}")
    for row_idx in targets:
        print(f"\nRemovendo linha {row_idx}: {ws.cell(row=row_idx, column=1).value}")
        ws.delete_rows(row_idx)

    ws.auto_filter.ref = f"A1:{ws.cell(row=1, column=ws.max_column).column_letter}{ws.max_row}"
    total = ws.max_row - 1
    for srow in wb[SHEET_SUMMARY].iter_rows(min_row=2):
        if srow[0].value in ("Total de Programas", "Programas Presenciais STEM"):
            srow[1].value = total
        if srow[0].value == "Data de Compilação":
            stamp = datetime.now().strftime("%d/%m/%Y")
            srow[1].value = f"{srow[1].value} - {len(TO_REMOVE)} programas online removidos em {stamp}"

    wb.save(DB_PATH)
    print(f"\nTotal de programas na base: {total}")


if __name__ == "__main__":
    main()
