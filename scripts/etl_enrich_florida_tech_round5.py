"""
Data enrichment (RODADA 5): atualiza APENAS a linha 'Florida Institute of
Technology (Florida Tech)' (Nathan M. Bisk College of Business, Melbourne,
FL) no msba_programs_database.xlsx.

Programa: Master of Science in Business Analytics, presencial (campus de
Melbourne), STEM.

FONTES OFICIAIS (acessadas em 2026-10-19):
  - https://www.fit.edu/programs/business-analytics-ms/
  - https://www.fit.edu/financialaid/graduate-students/cost-of-attendance/
  - https://www.fit.edu/housing/cost-of-housing/
  - https://www.fit.edu/housing/applications-and-forms/

CONFIRMADO (sem alteração): 30 créditos, US$ 820/crédito (privada, igual
para todos) -> US$ 24.600; fees US$ 2.260 (rodada 3); seguro US$ 1.999;
sem GRE/GMAT, cartas ou redação; Duolingo 105; taxa US$ 60; Deadline no
formato pedido pela usuária na rodada 4 (abertura das aplicações como
Round 1) -- não alterado.

CORREÇÕES:
- Aluguel Privativo/mês: era 'Verificar'. Studio em Southgate Village
  (opção privativa mais barata do campus) = US$ 5.180 por semestre em
  2026-27; 2 semestres / 9 meses do ano letivo = ~US$ 1.151/mês. Moradia
  para pós não é garantida, e as opções fora do campus são limitadas
  (página de housing).

NÃO ALTERADOS (continuam 'Verificar'): Duração -- a página do programa
presencial não informa (só a versão online diz "as few as 12 months");
Custo Vida Total (depende da duração); empregabilidade, salário,
empregadores (a página só cita rankings institucionais); rankings do
programa; tamanho da turma; GPA/GMAT médios; taxa de aceitação;
entrevista.
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
    "Aluguel Privativo/mês USD": 1151,  # studio Southgate Village US$ 5.180/semestre x 2 / 9 meses (2026-27)
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
