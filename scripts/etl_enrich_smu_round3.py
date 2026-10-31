"""
Data enrichment (RODADA 3): atualiza APENAS a linha 'SMU' (Cox School
of Business, Dallas, TX) no msba_programs_database.xlsx.

Esta linha já passou por 2 rodadas de enrichment (etl_enrich_smu.py e
_round2.py). Usuário confirmou via AskUserQuestion (ago/2026) para
prosseguir com uma Rodada 3.

ACHADO PRINCIPAL: localizado e lido o PDF oficial "Employment
Statistics — MS Business Analytics Class of 2025" (Cox Career
Management Center, conforme padrão MBA CSEA Reporting Standards),
disponível para download na página do programa.

'Tamanho Turma' preenchido: 97 (Total Graduating Class 2024-25,
Tabela 1 do relatório oficial).

'Taxa Empregabilidade %' CORRIGIDO de 96 para 82 -- a Tabela 2
("Timing of First Job Offers") do relatório oficial mostra que, do
total de 93 formandos buscando emprego, 17 (18%) NÃO tinham oferta
reportada até 6 meses após a formatura -- ou seja, 82% TINHAM oferta
até esse marco. Esse número CONFIRMA, com uma fonte primária muito
mais detalhada, a discrepância já registrada na Rodada 2 (quando um
"82%" havia sido encontrado mas não usado por falta de confirmação
robusta o suficiente para sobrescrever o valor existente). Agora, com
o relatório oficial completo em mãos, o valor foi corrigido com
confiança alta.

'Salário Médio Pós USD' CONFIRMADO sem alteração: a Tabela 5 do
relatório mostra salário base médio de $79.822 (mediana $75.000, entre
51 respondentes de 77,3% de taxa de resposta) -- praticamente idêntico
ao valor já presente ($79.800), validando cruzadamente sem necessidade
de correção.

CUIDADO APLICADO: o relatório detalha compensação por função,
indústria, região e experiência (ex.: Business/Data Analytics $76.495
médio; Financial Services $82.631 médio; Southwest US $77.720 médio),
mas nenhuma dessas quebras corresponde a um único "Principais
Empregadores" nomeado -- o relatório categoriza por SETOR/FUNÇÃO, não
por empresa específica, então 'Principais Empregadores' permanece
'Verificar'.

CUIDADO APLICADO (dados NÃO encontrados mesmo após tentativas
adicionais nesta rodada, permanecem 'Verificar'): 'Mercado/mês',
'Seguro Saúde', 'Seguro Residencial/mês', 'Celular/mês', 'WiFi/mês',
'Transporte/mês', 'Outros Gastos/mês', 'Custo Vida Total (programa)',
'Quão Realista (Intl)', 'TA/RA Disponível', 'Processo TA/RA',
'Graduate Assistantships', 'Ranking QS', 'GPA Mínimo', 'GPA Médio
Admitidos', 'GMAT Médio', 'Matérias Pré-Requisito' (páginas de
curriculum e career management center consultadas nesta rodada não
trazem esses dados).

Fontes principais (pesquisadas via WebFetch, ago/2026):
- https://www.smu.edu/-/media/site/cox/centersandinstitutes/cmc/employment-reports/msba-employment-statistics_2025.pdf
  (relatório oficial completo Class of 2025: tamanho de turma,
  timing de ofertas de emprego, compensação detalhada por múltiplas
  dimensões — fonte primária de alta confiança)
- https://www.smu.edu/cox/business-degrees/masters-programs/ms-business-analytics/curriculum
  (sem dados de pré-requisitos, GPA, ou GMAT)
- https://www.smu.edu/cox/centers-institutes/career-management-center
  (sem empregadores nomeados, direciona ao PDF acima)
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
    "Nome da Universidade": "SMU",
    "Estado": "TX",
    "Cidade": "Dallas",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Tamanho Turma": 97,  # Total Graduating Class 2024-25, oficial
    "Taxa Empregabilidade %": 82,  # CORREÇÃO: oficial (100% - 18% sem oferta reportada em 6 meses), era 96
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
        raise ValueError("Nenhuma linha encontrada para SMU (Dallas, TX). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (SMU, Dallas, TX)")

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
