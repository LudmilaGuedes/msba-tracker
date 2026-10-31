"""
Data enrichment (RODADA 5): atualiza APENAS a linha 'Florida State'
(Herbert Wertheim College of Business, Tallahassee, FL) no
msba_programs_database.xlsx.

Rodada focada nas lacunas que ficaram 'Verificar' na rodada 4.

FONTE: https://bulletin.fsu.edu/matriculation/g-admissions (Graduate
Bulletin, acessado em 2026-10-17): "minimum 3.0 (on a 4.0 scale) grade
point average (GPA) in all work attempted while registered as an
upper-division undergraduate student" -- vale para todos os mestrados,
inclusive diplomas internacionais comparáveis.

AINDA SEM FONTE (continuam 'Verificar'): Taxa Empregabilidade, Salário
Médio Pós, Principais Empregadores (a Wertheim não publica relatório de
colocação do MS-BA; a única menção é um formando da 1ª turma contratado
pela Naval Air Systems Command -- caso isolado, não usado), Ranking US
News/QS (os rankings encontrados são do MBA online, não do MS-BA
presencial), Tamanho Turma, GPA Médio, GMAT Médio, Taxa Aceitação e
Entrevista (não mencionada na página do programa).
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
    "Nome da Universidade": "Florida State",
    "Estado": "FL",
    "Cidade": "Tallahassee",
}

UPDATES = {
    "GPA Mínimo": 3.0,  # GPA da fase upper-division da graduação (regra da FSU Graduate School)
    "Entrevista": "Não mencionada na página do programa",
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
        raise ValueError("Nenhuma linha encontrada para Florida State (Tallahassee, FL). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Florida State, Tallahassee, FL)")

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
