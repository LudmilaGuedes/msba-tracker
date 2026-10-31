# -*- coding: utf-8 -*-
"""
Data enrichment (dificuldade de entrada): Texas Tech (Rawls), Lubbock, TX
-- MSDS.

Nao ha taxa de aceitacao publicada; o Poets&Quants e o site do programa
nao trazem perfil de turma. Indicios usados para classificar:
- GMAT/GRE obrigatorio, mas sem nota minima (FAQ oficial);
- sem GPA minimo publicado (FAQ oficial);
- a pagina do programa nao pede cartas de recomendacao;
- GRE medio citado em snippet de busca: Verbal 150 / Quant 152, sem ano e
  sem fonte primaria verificada. E um patamar baixo, perto do percentil
  40-50;
- turma presencial pequena (20-33 formados por ano).
O FAQ diz que o processo e "highly competitive", mas isso nao vem com
nenhum dado. Classificado como "Acessível", e o GMAT/GRE obrigatorio e a
principal barreira pratica.

Uso: python scripts/etl_enrich_texas_tech_dificuldade.py (a partir da raiz).
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
    "GMAT Médio": (
        "Não divulgado (só GRE, citado em busca: Verbal 150 / Quant 152; ano e fonte "
        "primária não confirmados)"
    ),
    "Dificuldade Entrada": "Acessível",
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
