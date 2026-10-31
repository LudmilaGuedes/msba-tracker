# -*- coding: utf-8 -*-
"""
Data enrichment (empregabilidade): Texas Tech (Rawls), Lubbock, TX --
MSDS, versao PRESENCIAL (on-campus).

Fontes oficiais do Rawls Career Management Center (set/2026):
- depts.ttu.edu/rawlsbusiness/about/cmc/placement/ (AY 2024-25)
- .../placement/documents/AY2023-24CareerOutcomes-web.pdf
- .../placement/documents/AY2022-23CareerOutcomes-web.pdf
Nos Specialty Masters, "employed" significa emprego conseguido em ate 180
dias apos a formatura.

MS Data Science (on-campus):
  AY 2022-23: 33 formados, 23 respostas, 87,0% empregados, media US$76.750
  AY 2023-24: 20 formados, 20 respostas, 52,6% empregados, media US$73.000
  AY 2024-25: 72,7% empregados, salario "Insufficient Data"
MS Data Science (online), para comparacao e NAO usado na linha: 83,9-100%
empregados, US$86-114k. A turma online e de profissionais que ja
trabalham, entao os numeros nao sao comparaveis com a presencial.

Preenchido:
- Taxa Empregabilidade %: 72,7 (AY 2024-25, o dado mais recente)
- Salario Medio Pos USD: 73.000 (AY 2023-24, o mais recente com salario)
- Tamanho Turma: 20 (formados on-campus em AY 2023-24)
Principais empregadores: continua "Verificar". A lista que circula (KPMG,
Deloitte, GE, State Farm...) vem de noticia/terceiros, sem separar online
de presencial.

Uso: python scripts/etl_enrich_texas_tech_empregabilidade.py (a partir da raiz).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "Texas Tech",
    "Estado": "TX",
    "Cidade": "Lubbock",
}

UPDATES = {
    "Tamanho Turma": 20,
    "Taxa Empregabilidade %": 72.7,
    "Salário Médio Pós USD": 73000,
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
        raise ValueError("Nenhuma linha encontrada para Texas Tech (Lubbock, TX). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Texas Tech, Lubbock, TX)")

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

    print(f"\n{len(UPDATES)} campos atualizados na linha {row_idx} (Texas Tech):\n")
    for header in UPDATES:
        print(f"  - {header}:")
        print(f"      Antes:  {before[header]!r}")
        print(f"      Depois: {after[header]!r}")


if __name__ == "__main__":
    main()
