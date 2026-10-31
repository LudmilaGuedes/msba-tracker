"""
Data enrichment (RODADA 3): atualiza APENAS a linha 'UT Arlington'
(College of Business, Arlington, TX) no msba_programs_database.xlsx.

Esta linha já passou por 2 rodadas de enrichment (etl_enrich_ut_arlington.py
e _round2.py), restando 33/54 campos 'Verificar', incluindo tuition
(prioridade desta rodada 3, conforme confirmado pelo usuário via
AskUserQuestion, ago/2026).

TUITION: PERMANECE 'Verificar' -- localizada a página oficial de
"Cost of Attendance" 2026-27 da UTA (fao/average-cost), que traz uma
tabela ITEMIZADA de "Tuition and Fees": $11.044/ano (residente) /
$23.486/ano (não-residente). NÃO USADA para preencher 'Custo/Crédito'
ou 'Tuition Total' por DUAS razões documentadas: (1) a própria fonte
declara explicitamente que os valores são "average costs... to be used
for planning purposes only and are not reflective of actual tuition
and fees incurred by an individual student" -- ou seja, não é uma
tarifa por-crédito oficial do College of Business, é uma média
universitária geral; (2) calcular um total de programa exigiria
assumir tanto uma carga de créditos/ano quanto multiplicar por uma
duração de programa não totalmente confirmada, empilhando duas
suposições sobre um número já rotulado como não-representativo.
Preencher os campos de tuition com essa cadeia de suposições
violaria o princípio já estabelecido nesta sessão de nunca inventar
uma decomposição/estimativa sobre dados de baixa confiança.

CUSTO DE VIDA: da mesma tabela oficial (também rotulada como média
geral, não específica do College of Business), preenchidos apenas os
campos que NÃO exigem decompor um valor já preenchido de outra fonte:
'Transporte/mês USD' = $3.280/ano ÷ 12 = $273,33; 'Outros Gastos/mês
USD' = (Books & Supplies $500 + Personal/Misc $2.274) ÷ 12 = $231,17.
'Mercado/mês' NÃO foi preenchido: a fonte só dá "Housing and Meals"
como valor COMBINADO ($15.180/ano), e 'Aluguel Privativo/mês' já
estava preenchido com $1.034 de uma fonte DIFERENTE (rodada 1) --
subtrair um do outro misturaria duas fontes distintas numa proporção
inventada, violando o mesmo princípio de cautela.

CUIDADO APLICADO (dados NÃO encontrados mesmo após tentativas
adicionais nesta rodada, permanecem 'Verificar'): 'Custo/Crédito'
(ambos), 'Tuition Total (Residente)', 'Fees Adicionais', 'Mercado/
mês', 'Seguro Saúde', 'Seguro Residencial/mês', 'Utilities/mês',
'Celular/mês', 'WiFi/mês', 'Custo Vida Total (programa)', 'Tamanho
Turma', 'Bolsas Disponíveis'/'Tipos de Bolsa'/'Quão Realista (Intl)'/
'TA/RA Disponível'/'Processo TA/RA'/'Caminho Tuition Residente'/
'Graduate Assistantships', 'Ranking US News', 'Ranking QS', 'Taxa
Empregabilidade %', 'Salário Médio Pós USD', 'GPA Médio Admitidos',
'GMAT Médio', 'Taxa Aceitação %', 'Matérias Pré-Requisito', 'Taxa
Aplicação USD' (a página de admissões confirma apenas que existe uma
taxa não-reembolsável, sem valor), 'Rounds Aplicação', 'Documentos
Necessários' (apenas genéricos foram encontrados: transcripts, cartas/
statement/writing samples "variam por departamento"), 'Entrevista'.

Fontes principais (pesquisadas via WebFetch, ago/2026):
- https://www.uta.edu/administration/fao/average-cost (tabela oficial
  de custo de atendimento 2026-27, rotulada como média geral para
  fins de planejamento, não específica do College of Business)
- https://www.uta.edu/admissions/graduate/apply (documentos gerais,
  sem valor de taxa de aplicação, sem menção a entrevista)
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
    "Transporte/mês USD": 273.33,  # $3.280/ano (média geral UTA, não específica do College of Business) / 12
    "Outros Gastos/mês USD": 231.17,  # (Books $500 + Personal $2.274)/ano / 12, mesma fonte
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
