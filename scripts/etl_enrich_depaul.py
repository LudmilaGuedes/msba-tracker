"""
Data enrichment: atualiza APENAS a linha da 'DePaul University'
(Kellstadt Graduate School of Business, Chicago, IL) no
msba_programs_database.xlsx.

Foco: MS in Business Analytics and Artificial Intelligence (MSBAI),
formato presencial full-time (12 meses; há também trilha part-time de
18-24 meses, não tocada aqui), 48 créditos com possibilidade de
waivers (já confirmado na planilha), STEM.

TUITION: CONFIRMADO -- a página oficial de custos do Kellstadt declara
explicitamente "$1.228 per credit hour for the 2025-2026 academic
year", validando cruzadamente o valor já presente na planilha
($1.228/crédito, $58.900 total para ~48 créditos) sem necessidade de
correção. Nota: a fonte rotula o ano como 2025-26, um ano anterior ao
foco usual desta pesquisa (2026-27) -- mantido por ser a única fonte
oficial encontrada, com essa ressalva de ano registrada.

'Testes Inglês Aceitos' CORRIGIDO: os valores anteriores ("TOEFL 80
(mín 20/seção) / IELTS 7.0 (mín 6.5)") divergiam da página oficial
ESPECÍFICA de admissão internacional do Kellstadt, que lista: TOEFL 80
iBT (todas as seções ≥17) ou nova escala 4.0 (nenhuma sub-nota <3.5);
IELTS 6.5 (não 7.0); Duolingo 115+; PTE 53. Como a nova fonte é uma
página dedicada exclusivamente a requisitos de admissão internacional
(mais específica e autoritativa para essa pergunta do que a fonte
anterior desconhecida), os valores foram CORRIGIDOS para os oficiais.

'Taxa Aplicação USD' preenchido: $75 (não-reembolsável, confirmado na
mesma página internacional). 'Documentos Necessários' preenchido:
aplicação online + taxa; essay sobre interesse no programa e
objetivos profissionais; resume; transcripts oficiais de todas as
instituições; cartas de recomendação OPCIONAIS (não obrigatórias).

'Deadline Aplicação' e 'Rounds Aplicação' REFINADOS: o programa usa
ciclo trimestral Fall/Winter/Spring (não Fall/Spring/Summer como o
valor anterior sugeria) -- Fall: 15 ago (doméstico e internacional,
segundo uma página; 1 jun para internacionais segundo outra página do
mesmo site -- INCONSISTÊNCIA INTERNA do próprio site da DePaul,
registrada aqui); Winter: 1 dez (doméstico) / 1 out (internacional);
Spring: 1 mar (doméstico) / 1 fev (internacional).

'Destaques Currículo' preenchido: 9 disciplinas obrigatórias de
business analytics + 3 eletivas; foco em Python, R, SQL, Tableau,
Generative AI, machine learning, modelagem preditiva; projetos de
consultoria capstone com parceiros reais (Reyes Holdings, Labelmasters,
Morningstar).

'Tipos de Bolsa' preenchido com CAUTELA: a única bolsa com valor
concreto encontrada é o "Double Demon Scholarship" (25% de desconto na
tuition), mas ela é EXCLUSIVA para ex-alunos de GRADUAÇÃO da própria
DePaul -- não aplicável a uma candidata externa/internacional sem
bacharelado na DePaul. Outras bolsas são mencionadas genericamente
("range of tuition support and financial aid options") sem
detalhamento de valores ou critérios.

CUIDADO APLICADO: a fonte cita "Ranked 8th best master's in business
analytics program by Intelligent.com (2024, 2025)" e "Named a Best
Value College by U.S. News & World Report (2025)" -- nenhuma das duas
corresponde ao formato numérico esperado pelos campos 'Ranking US
News' ou 'Ranking QS' desta planilha (a primeira é de uma fonte
não-tradicional; a segunda é uma classificação genérica de "melhor
custo-benefício", não um ranking de programa específico) -- por isso
NÃO foram usadas para preencher esses campos, que permanecem
'Verificar'.

CUIDADO APLICADO (dados NÃO encontrados, permanecem 'Verificar'):
'Tamanho Turma', 'GPA Mínimo', 'GPA Médio Admitidos', 'GMAT/GRE
Obrigatório' (nenhuma menção encontrada em nenhuma fonte oficial
consultada, apesar de múltiplas tentativas -- a nota anterior "essay
pode substituir" foi mantida por não ter sido contradita), 'GMAT
Médio', 'Taxa Aceitação %', 'Taxa Empregabilidade %', 'Salário Médio
Pós USD', 'Principais Empregadores', 'Matérias Pré-Requisito',
'Entrevista', 'Fees Adicionais USD', 'Seguro Saúde (programa) USD',
'Processo TA/RA', 'TA/RA Disponível', e todos os campos mensais de
custo de vida em Chicago, IL.

Fontes principais (pesquisadas via WebFetch, ago/2026):
- https://www.depaul.edu/academics/programs/business-analytics-ms
  (formato, duração, STEM, currículo, deadlines por trimestre,
  rankings não-tradicionais)
- .../admission-requirements (documentos, deadlines -- versão
  ligeiramente diferente da anterior, inconsistência registrada)
- https://business.depaul.edu/admission-aid/international-student-admission
  (taxa de aplicação $75, mínimos de inglês corretos e específicos
  para internacionais)
- https://business.depaul.edu/admission-aid/tuition-aid (tuition
  oficial 2025-26 confirmando $1.228/crédito, Double Demon
  Scholarship exclusiva para alumni)
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
    "Nome da Universidade": "DePaul University",
    "Estado": "IL",
    "Cidade": "Chicago",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Testes Inglês Aceitos": (
        "CORREÇÃO: TOEFL 80 iBT (todas seções ≥17) ou nova escala "
        "4.0 (nenhuma sub-nota <3.5); IELTS 6.5; Duolingo 115+; PTE 53"
    ),
    "Taxa Aplicação USD": 75,  # não-reembolsável, oficial
    "Documentos Necessários": (
        "Aplicação online + taxa; essay sobre interesse no programa "
        "e objetivos profissionais; resume; transcripts oficiais de "
        "todas as instituições; cartas de recomendação OPCIONAIS"
    ),
    "Deadline Aplicação": (
        "Ciclo Fall/Winter/Spring: Fall 15 ago (nota: 1 jun para "
        "internacionais segundo outra página do site — "
        "inconsistência interna); Winter: 1 dez (doméstico) / 1 out "
        "(internacional); Spring: 1 mar (doméstico) / 1 fev "
        "(internacional)"
    ),
    "Rounds Aplicação": "3 ciclos trimestrais (Fall/Winter/Spring), sem rounds numerados dentro de cada ciclo",
    "Destaques Currículo": (
        "9 disciplinas obrigatórias de business analytics + 3 "
        "eletivas; Python, R, SQL, Tableau, Generative AI, machine "
        "learning, modelagem preditiva; capstone com parceiros reais "
        "(Reyes Holdings, Labelmasters, Morningstar)"
    ),
    "Tipos de Bolsa": (
        "Double Demon Scholarship (25% desconto) — EXCLUSIVA para "
        "ex-alunos de graduação da própria DePaul, não aplicável a "
        "candidata externa/internacional; outras bolsas mencionadas "
        "genericamente sem detalhamento"
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
        raise ValueError("Nenhuma linha encontrada para DePaul University (Chicago, IL). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (DePaul University, Chicago, IL)")

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
