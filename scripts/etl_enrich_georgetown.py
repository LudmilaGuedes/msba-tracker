"""
Data enrichment: atualiza APENAS a linha da 'Georgetown' (McDonough
School of Business, Washington, DC) no msba_programs_database.xlsx.

Foco exclusivo: MS in Business Analytics, formato full-time presencial,
STEM, 30 créditos (18 disciplinas + capstone), 10 meses. Deep research
em: custo total real de tuition para internacionais, bolsas/
assistantships, requisitos de GPA/GMAT/GRE/inglês, e custo de vida em
Washington, DC.

TUITION (já estava correto, apenas complementado): $71.370 para a
turma de Fall 2026 (confirmado oficialmente). PREENCHIDO 'Custo/
Crédito' (estava 'Verificar' em ambos) dividindo por 30 créditos =
$2.379/crédito. Como é universidade PRIVADA, 'Tuition Total
(Residente)' preenchido com o mesmo valor do Intl.

Emprego/salário: dados PRÓPRIOS do MSBA (não misturados com o MBA),
extraídos do relatório oficial de emprego da Turma de Dezembro de 2023
(msb.georgetown.edu). 'Salário Médio Pós USD' preenchido com $123.268
(salário-base médio). 'Taxa Empregabilidade %' permanece 'Verificar'
DE PROPÓSITO: a única métrica de "emprego" divulgada pela própria
Georgetown é "76% mudaram de emprego em até 6 meses" -- uma métrica de
TROCA DE CARREIRA, não uma taxa de colocação/emprego no sentido
tradicional do campo do schema. Usá-la aqui misturaria conceitos
diferentes e poderia enganar a leitura do dado. 'Principais
Empregadores' preenchido com a lista oficial do mesmo relatório
(específica do MSBA).

'Ranking US News' e 'Ranking QS' permanecem 'Verificar': a única
menção de ranking encontrada foi "top 2 nacional pela Fortune" -- uma
publicadora diferente das duas que o schema pede especificamente: não
inventado um valor para US News/QS a partir de um ranking de outra
fonte.

'Seguro Saúde (programa)' preenchido com $4.450/ano -- taxa oficial
corrente do plano de seguro obrigatório para internacionais
(studenthealthinsurance.georgetown.edu). Demais campos de custo de vida
mensal (Aluguel, Mercado, Utilities, Transporte) preenchidos com
estimativas gerais de orçamento estudantil para Washington, DC (mesma
metodologia/fontes já usadas no enrichment da Johns Hopkins, também
localizada em DC, nesta mesma planilha) -- não são específicas de
Georgetown, mas refletem o mesmo mercado de custo de vida da cidade.

Fontes principais (pesquisadas via busca web, ago/2026):
- https://msb.georgetown.edu/msba/admissions-financing/ (tuition
  oficial, lista de bolsas com consideração automática: MSBA
  Distinguished Faculty Scholars, Innovation in Analytics and AI,
  Analytics for Social Impact, Global Citizen, Community in Diversity,
  Georgetown Undergrad/Alumni até 10%, Federal Employee até 10%)
- https://msb.georgetown.edu/msba/admissions-financing/application-
  process/ (documentos, taxa de aplicação $175, sem entrevista
  obrigatória, sem pré-requisito formal de estatística/programação)
- https://msb.georgetown.edu/msba/master-of-science-in-business-
  analytics-faq/ (GPA/GMAT/GRE sem mínimo divulgado, revisão holística)
- https://msb.georgetown.edu/news-story/career/georgetown-msba-
  graduates-exceed-career-goals-in-2024-employment-report/ (relatório
  oficial de emprego, Turma Dez/2023: salário-base médio $123.268,
  empregadores e indústrias específicos do MSBA)
- Busca web: deadlines Early Action 21 out 2026, Round 1 20 nov 2026,
  Round 2 20 jan 2027, Round 3 25 mar 2027 (corrigido o ano do Round 3,
  que uma fonte agregadora citou incorretamente como 2026)
- https://studenthealthinsurance.georgetown.edu (seguro saúde
  obrigatório internacional, $4.450/ano)
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
    "Nome da Universidade": "Georgetown",
    "Estado": "DC",
    "Cidade": "Washington",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Intl) USD": 2379,  # $71.370 / 30 créditos
    "Custo/Crédito (Residente) USD": 2379,  # privada, mesma taxa do Intl
    "Tuition Total (Residente) USD": 71370,  # privada, mesma taxa do Intl
    "Aluguel Privativo/mês USD": 1150,  # estimativa geral de orçamento estudantil para Washington, DC
    "Mercado/mês USD": 320,
    "Utilities/mês USD": 175,
    "Transporte/mês USD": 136,  # passe de Metro
    "Seguro Saúde (programa) USD": 4450,  # oficial, plano obrigatório para internacionais
    "Tamanho Turma": 74,  # turma recente (McDonough news story)
    "Tipos de Bolsa": (
        "Consideração automática para todo candidato: MSBA "
        "Distinguished Faculty Scholars, Innovation in Analytics and "
        "AI, Analytics for Social Impact, Global Citizen, Community in "
        "Diversity Scholarship; até 10% para ex-alunos/graduados da "
        "Georgetown ou funcionários federais"
    ),
    "Taxa Aplicação USD": 175,  # isenção disponível para militares ativos/reservistas/veteranos
    "Deadline Aplicação": "Early Action: 21 out 2026; Round 1: 20 nov 2026; Round 2: 20 jan 2027; Round 3: 25 mar 2027",
    "Rounds Aplicação": "Early Action, Round 1, Round 2, Round 3 (ver Deadline Aplicação)",
    "Documentos Necessários": (
        "Application online, resume/CV, ensaio pessoal (máx. 500 "
        "palavras, 3 opções de tema), video essay (1-3 min), 1 carta "
        "de recomendação (acadêmica ou profissional), histórico "
        "acadêmico (não-oficial na candidatura), teste de inglês "
        "(internacionais, se aplicável)"
    ),
    "Matérias Pré-Requisito": "Nenhum pré-requisito formal de estatística/programação — programa aceita alunos sem experiência prévia em código",
    "Salário Médio Pós USD": 123268,  # salário-base médio, Turma de Dezembro de 2023 (relatório oficial de emprego do MSBA)
    "Principais Empregadores": (
        "Accenture, Amazon, American Express, Bank of America, Booz "
        "Allen Hamilton, Capital One, Deloitte, Federal Reserve "
        "District Banks, Guidehouse, IBM, JPMorgan Chase, Microsoft, "
        "Morgan Stanley, PwC, Raytheon, S&P Global, U.S. Department of "
        "Defense"
    ),
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
        raise ValueError("Nenhuma linha encontrada para Georgetown (Washington, DC). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Georgetown, Washington, DC)")

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
