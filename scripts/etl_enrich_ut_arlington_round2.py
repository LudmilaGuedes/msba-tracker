"""
Data enrichment (RODADA 2): atualiza APENAS a linha 'UT Arlington'
(College of Business, Arlington, TX) no msba_programs_database.xlsx.

Esta linha já passou por uma rodada de enrichment (script
etl_enrich_ut_arlington.py), restando ~26/54 campos 'Verificar'.
Usuário confirmou via AskUserQuestion (ago/2026) para prosseguir com
uma Rodada 2 focada nesses campos remanescentes.

'Principais Empregadores' preenchido com a lista oficial encontrada na
página do programa: Aflac, Amazon, Alcon, Deloitte, Toyota.

CUIDADO APLICADO: a página do programa cita "No. 18 Data analytics
program in North America", mas SEM citar a fonte/publicação dessa
classificação -- diferente de rankings claramente atribuídos a US
News ou QS. Por não haver atribuição de fonte, NÃO foi usado para
preencher 'Ranking US News' ou 'Ranking QS' (mesmo critério de cautela
já aplicado a outras universidades desta sessão com rankings de fontes
não identificadas ou não-tradicionais).

CUIDADO APLICADO: uma página alternativa do mesmo departamento
menciona "17-month program", divergindo do valor já presente em
'Duração (meses)' (16, não 'Verificar'). Como a diferença é pequena e
pode refletir apenas uma forma diferente de contar o mesmo programa
(ex.: incluindo ou não um módulo inicial), o valor existente foi
MANTIDO sem alteração.

CUIDADO APLICADO (dados NÃO encontrados mesmo após tentativas
adicionais nesta rodada, permanecem 'Verificar'): 'Custo/Crédito'
(ambos), 'Tuition Total (Residente)', 'Fees Adicionais', 'Tamanho
Turma', 'GPA Médio Admitidos', 'GMAT Médio', 'Taxa Aceitação %',
'Taxa Empregabilidade %', 'Salário Médio Pós USD', 'Matérias
Pré-Requisito', 'Documentos Necessários', 'Taxa Aplicação USD',
'Rounds Aplicação', 'Entrevista', 'TA/RA Disponível'/'Processo TA/RA'/
'Graduate Assistantships', 'Bolsas Disponíveis'/'Tipos de Bolsa',
'Caminho Tuition Residente', e a maior parte dos campos mensais de
custo de vida em Arlington, TX. Múltiplas páginas oficiais (catálogo
de admissão, tuition estimator, página específica do departamento)
foram consultadas nesta rodada sem sucesso em nenhum desses pontos.

Fontes principais (pesquisadas via WebFetch, ago/2026):
- https://www.uta.edu/academics/programs/business-analytics-ms
  (empregadores oficiais, ranking sem fonte atribuída — não usado)
- https://catalog.uta.edu/business/infosystems/graduate/ (sem dados
  de admissão específicos do MSBA nesta rodada)
- https://www.uta.edu/admissions/afford/tuition-estimator (sem dados
  de tuition acessíveis no conteúdo extraído)
- https://www.uta.edu/academics/schools-colleges/business/departments/information-systems-and-operations-management/master-of-science-business-analytics
  (menciona "17-month program", discrepância registrada, não usada)
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
    "Nome da Universidade": "UT Arlington",
    "Estado": "TX",
    "Cidade": "Arlington",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Principais Empregadores": "Aflac, Amazon, Alcon, Deloitte, Toyota",
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
        raise ValueError("Nenhuma linha encontrada para UT Arlington (Arlington, TX). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (UT Arlington, Arlington, TX)")

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
