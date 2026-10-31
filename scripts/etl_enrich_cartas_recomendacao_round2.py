# -*- coding: utf-8 -*-
"""
Cartas de Recomendação, rodada 2 (20/out/2026): preenche com número as 6
linhas que ainda estavam sem quantidade ("Verificar" ou "Até 2"), para que
toda a coluna comece por um número (a aba Estratégia lê esse número).

Pesquisado em páginas oficiais:
- UAH: checklist "All Business Graduate Programs - Except Executive MBA"
  (uah.edu/business/grad/degrees/admission-requirements) pede aplicação,
  transcripts, GMAT/GRE e resume; sem cartas -> 0.
- SCSU: página do programa (southernct.edu/academics/programs/business-
  analytics-ms) só exige bacharelado com GPA 3.0; nenhuma página oficial
  pede cartas -> 0.
- VT Global BA: a página de application guidelines saiu do ar (404); a lista
  de documentos registrada antes (statement, vídeo, experiência, inglês)
  não tem cartas -> 0.
- University of San Diego: página MSBA How to Apply lista "Letter(s) of
  recommendation" entre os itens, "up to two" -> 1 (até 2).
- Clark: página de requisitos dos mestrados exige "letters of
  recommendation" sem número (só o MAT diz 3); agregadores citam 2 -> 2
  (estimado).
- A&M-Texarkana: o catálogo do MS Business and Data Analytics (programa
  novo, Fall 2026) e a página de admissões de pós pedem cartas sem número;
  a página de admissões do College of Business exige 3 para MBA e MS
  Accounting -> 3 (estimado).

Uso: python scripts/etl_enrich_cartas_recomendacao_round2.py (a partir da raiz).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"
COLUMN = "Cartas de Recomendação"

# (Nome da Universidade, Estado, Cidade) -> novo valor
UPDATES = {
    ("UAH", "AL", "Huntsville"): (
        "0 (checklist oficial dos mestrados de business pede só aplicação, transcripts, "
        "GMAT/GRE e resume)"
    ),
    ("Southern Connecticut State University (SCSU)", "CT", "New Haven"): (
        "0 (nenhuma página oficial pede cartas)"
    ),
    ("VT Global BA", "VA", "Blacksburg"): (
        "0 (lista oficial de documentos sem cartas; página de aplicação fora do ar em out/2026)"
    ),
    ("University of San Diego", "CA", "San Diego"): "1 (até 2)",
    ("Clark University", "MA", "Worcester"): (
        "2 (estimado: site oficial exige 'letters' sem número; agregadores citam 2)"
    ),
    ("Texas A&M University-Texarkana (A&M-Texarkana)", "TX", "Texarkana"): (
        "3 (estimado: o College of Business exige 3 para MBA e MS Accounting; o catálogo "
        "do MS não informa número)"
    ),
}


def main():
    rotate_backup()
    print("Backup rotacionado (política fixa) em: data/msba_programs_database_backup.xlsx")

    wb = openpyxl.load_workbook(DB_PATH)
    ws = wb[SHEET_DATA]
    headers = [cell.value for cell in ws[1]]
    col = {h: i + 1 for i, h in enumerate(headers)}

    rows_by_key = {}
    for row in ws.iter_rows(min_row=2):
        key = (
            row[col["Nome da Universidade"] - 1].value,
            row[col["Estado"] - 1].value,
            row[col["Cidade"] - 1].value,
        )
        rows_by_key.setdefault(key, []).append(row[0].row)

    for key in UPDATES:
        if len(rows_by_key.get(key, [])) != 1:
            raise ValueError(f"Esperava 1 linha para {key}, achei {len(rows_by_key.get(key, []))}. Abortando.")

    no_fill = PatternFill(fill_type=None)
    changes = []
    for key, new_value in UPDATES.items():
        cell = ws.cell(row=rows_by_key[key][0], column=col[COLUMN])
        changes.append((key[0], cell.value, new_value))
        cell.value = new_value
        cell.fill = no_fill

    wb.save(DB_PATH)

    print(f"\n{COLUMN} (Antes -> Depois):")
    for name, old, new in changes:
        print(f"  - {name}:")
        print(f"      Antes:  {old!r}")
        print(f"      Depois: {new!r}")


if __name__ == "__main__":
    main()
