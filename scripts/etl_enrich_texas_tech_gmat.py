# -*- coding: utf-8 -*-
"""
Data enrichment (GMAT/GRE e GPA): Texas Tech (Rawls), Lubbock, TX --
Master of Science in Data Science (MSDS).

Resolve o conflito registrado em etl_add_texas_tech_msds.py. Os snippets
de busca vieram atribuidos ao FAQ oficial do MSDS
(depts.ttu.edu/rawlsbusiness/graduate/ms/datascience/faq/; o fetch direto
continua bloqueado com HTTP 403):
- "A GRE/GMAT submission is required; however, there is not a designated
  score requirement for acceptance into the program."
- A dispensa sai com 5+ anos de experiencia relevante e "possivelmente" com
  3 anos. O pedido e feito por e-mail a Cy Cawthron, com o ID da aplicacao
  e o resume.
- GPA: nao ha minimo publicado. O Rawls busca "above average grades, strong
  professional development, high test scores and a background in
  quantitative and technical work".

Uso: python scripts/etl_enrich_texas_tech_gmat.py (a partir da raiz).
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
    "GMAT/GRE Obrigatório": (
        "Obrigatório (GMAT ou GRE), sem nota mínima; dispensa com 5+ anos de experiência "
        "relevante (possível com 3 anos) -- pedido por e-mail a Cy Cawthron com o ID da "
        "aplicação e o resume"
    ),
    "GPA Mínimo": (
        "Sem mínimo publicado -- o Rawls busca notas acima da média, testes fortes e "
        "perfil quantitativo/técnico"
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
