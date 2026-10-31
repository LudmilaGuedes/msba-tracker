"""
Data enrichment (RODADA 4): atualiza APENAS o campo 'Deadline Aplicação'
da linha 'Florida Institute of Technology (Florida Tech)' (Melbourne, FL).

Pedido da usuária (2026-10-03): para esta escola de admissão contínua, o
campo deve começar pela data de ABERTURA das aplicações, tratada como o
"Round 1" -- 1 jul 2027 para o Fall 2028. O prazo final (~15 jun 2028,
duas semanas antes do prazo do I-20 de 1 jul 2028) fica depois.

Fonte: https://www.fit.edu/graduate-admissions/ ("Application Opens:
July 1st the year prior to enrollment"; "I-20 Deadline: July 1st of
enrollment year" para Fall e "November 1st" para Spring).
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
    "Nome da Universidade": "Florida Institute of Technology (Florida Tech)",
    "Estado": "FL",
    "Cidade": "Melbourne",
}

UPDATES = {
    "Deadline Aplicação": (
        "Round 1: 1 jul 2027 (abertura das aplicações para o Fall 2028; "
        "admissão contínua, análise assim que a aplicação fica completa); "
        "prazo final: ~15 jun 2028 (2 semanas antes do prazo do I-20, 1 jul "
        "2028); Spring: prazo do I-20 em 1 nov"
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
        raise ValueError("Nenhuma linha encontrada para Florida Tech (Melbourne, FL). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Florida Tech, Melbourne, FL)")

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
