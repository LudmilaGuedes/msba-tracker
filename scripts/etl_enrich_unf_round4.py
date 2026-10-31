"""
Data enrichment (round 4): corrige 'Testes Inglês Aceitos' da linha
'University of North Florida (UNF)' no msba_programs_database.xlsx.

Pergunta do usuário: "aceita duolingo english test?" -- a planilha só
listava TOEFL/IELTS (vindo da página específica de critérios do MSBA,
que não menciona Duolingo/PTE). A página oficial university-wide de
proficiência em inglês para pós-graduação
(https://www.unf.edu/admissions/english-language-proficiency.html,
acessada via WebFetch em 2026-09-08) confirma que Duolingo E PTE
também são aceitos:

- TOEFL 79
- IELTS 6.5
- Duolingo 110
- PTE Academic 53

Também documentadas alternativas que exigem aprovação do Graduate
Program Director (não automáticas): GRE Verbal >= 152, ou diploma de
instituição com ensino em inglês, ou diploma de bacharelado/mestrado
de instituição dos EUA credenciada, ou certificado do Pre-University
English Language Program da própria UNF (média mínima 3.0).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

VERIFICAR = "Verificar"

MATCH = {
    "Nome da Universidade": "University of North Florida (UNF)",
    "Estado": "FL",
    "Cidade": "Jacksonville",
}

UPDATES = {
    "Testes Inglês Aceitos": (
        "TOEFL 79 (iBT) / IELTS 6.5 / Duolingo 110 / PTE Academic 53 "
        "(CORRIGIDO: página específica do MSBA só citava TOEFL/IELTS, "
        "mas a página university-wide de proficiência em inglês "
        "confirma Duolingo e PTE também aceitos). Alternativas que "
        "exigem aprovação do Graduate Program Director (não "
        "automáticas): GRE Verbal >= 152, diploma de instituição com "
        "ensino em inglês, ou diploma de bacharelado/mestrado dos "
        "EUA."
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
        raise ValueError("Nenhuma linha encontrada para UNF. Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (UNF)")

    no_fill = PatternFill(fill_type=None)
    verificar_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    before = {}
    after = {}
    for header, new_value in UPDATES.items():
        col_idx = header_to_col[header]
        cell = ws.cell(row=row_idx, column=col_idx)
        before[header] = cell.value
        cell.value = new_value
        cell.fill = verificar_fill if new_value == VERIFICAR else no_fill
        after[header] = new_value

    wb.save(DB_PATH)

    print("\nCampos atualizados (Como Estava -> Como Ficou):")
    for header in UPDATES:
        print(f"  - {header}:")
        print(f"      Antes: {before[header]!r}")
        print(f"      Depois: {after[header]!r}")


if __name__ == "__main__":
    main()
