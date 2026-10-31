"""
Data enrichment (RODADA 2): atualiza APENAS a linha 'Rutgers Business
School (Newark)' (Newark, NJ) no msba_programs_database.xlsx.

Esta linha já passou por uma rodada de enrichment (script
etl_enrich_rutgers_newark.py), restando 16/54 campos 'Verificar'.
Usuário confirmou via AskUserQuestion (ago/2026) para prosseguir com
uma Rodada 2 focada nesses campos remanescentes.

'Início' preenchido: 'Fall/Spring' -- a página oficial de FAQ do MITA
confirma que o programa aceita admissão tanto no Fall quanto no
Spring (consistente com os deadlines domésticos e internacionais já
presentes na planilha, que já diferenciavam os dois ciclos).

CUIDADO APLICADO: mesmo após tentativas adicionais nesta rodada
(incluindo a página de FAQ revisitada e uma tentativa de página de
admissions/employment report, esta última retornando 404), NENHUM
outro campo pôde ser preenchido com confiança -- nenhuma fonte oficial
consultada traz tamanho de turma, GPA mínimo/médio, taxa de aceitação,
política de entrevista, taxa de empregabilidade, salário médio, ou
rankings (US News/QS) específicos do MITA. Esses campos permanecem
'Verificar'.

Fontes principais (pesquisadas via WebFetch, ago/2026):
- https://www.business.rutgers.edu/masters-information-technology-analytics/frequently-asked-questions
  (confirma ciclos de admissão Fall e Spring; reafirma médias de
  GMAT/GRE/TOEFL já presentes da rodada 1)
- https://www.business.rutgers.edu/masters-information-technology-analytics/admissions
  (confirma apenas o pré-requisito de bacharelado + cálculo/
  estatística básica, já presente; sem dados novos de GPA, entrevista,
  ou taxa de aceitação)
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
    "Nome da Universidade": "Rutgers Business School (Newark)",
    "Estado": "NJ",
    "Cidade": "Newark",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Início": "Fall/Spring",  # oficial, ambos os ciclos confirmados
}


def main():
    rotate_backup()
    print(f"Backup rotacionado (política fixa) em: data/msba_programs_database_backup.xlsx")

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
        raise ValueError("Nenhuma linha encontrada para Rutgers Business School (Newark) (Newark, NJ). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Rutgers Business School (Newark), Newark, NJ)")

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
