"""
Data enrichment: atualiza APENAS a linha da 'UW-Madison' (Wisconsin
School of Business, Madison, WI) no msba_programs_database.xlsx.

Foco: MS in Business Analytics, presencial, 30 créditos (já
confirmado), STEM (já confirmado). Programa começa com um bootcamp
online no verão, seguido de 2 semestres presenciais.

'Duração (meses)' CORRIGIDO de 12 para 10 -- a página oficial confirma
"10 months total" (bootcamp de verão online + 2 semestres
presenciais), não 12.

TUITION: os valores já presentes ($1.500/crédito, $45.000 total,
$1.660 de fees) foram CONFIRMADOS pela página oficial ($1.500/crédito,
"$45.000 total program cost") -- nenhuma correção necessária, apenas
validação cruzada.

'GPA Mínimo' preenchido: 3.0 cumulativo (GPAs internacionais devem ser
equivalentes a 3.0); auto-admissão disponível para bacharéis da
própria UW-Madison com GPA ≥3.4. 'GPA Médio Admitidos' preenchido:
3.49 (Class of 2026, oficial).

'GMAT/GRE Obrigatório' preenchido com o critério oficial completo:
EXIGIDO, salvo dispensa seletiva concedida a candidatos com forte
preparo quantitativo (graduação STEM/quantitativa; GPA ≥3.2 geral ou
notas B+ em disciplinas quantitativas; 2+ anos de experiência
analítica/financeira/técnica; certificações como CPA/CFA/FRM; cursos
prep de MBA; ou mestrado/doutorado prévio). 'GMAT Médio' preenchido
com os dados oficiais disponíveis (a página de Class Profile só
divulga médias de GRE, não de GMAT): GRE Verbal médio 154 (faixa 80%:
144-162), GRE Quantitativo médio 165 (faixa 80%: 160-170) -- deixado
claro no valor do campo que são estatísticas de GRE, não de GMAT.

'Testes Inglês Aceitos' refinado com os dois patamares oficiais de
TOEFL (92 antes de 21/jan/26; 5.0 no novo formato a partir dessa
data), IELTS 7.0, Duolingo 125, e dispensa para diploma de graduação
de 4 anos lecionado em inglês.

'Taxa Aplicação USD' preenchido: $75 (doméstico e internacional) + $6
de taxa de processamento adicional para internacionais. 'Documentos
Necessários' preenchido com a lista oficial (3 essays escritos de 600
palavras + 1 opcional; 3 video essays de 2-3 min + até 2 adicionais em
72h; resume; transcripts; GMAT/GRE ou pedido de dispensa aprovado;
teste de inglês para internacionais).

'Deadline Aplicação' e 'Rounds Aplicação' preenchidos/refinados com os
5 prazos oficiais para o ciclo de início em julho de 2027 (mais
granular que o valor anterior "May 1 (intl), Jun 1 (domestic)", que
não capturava as 3 rounds intermediárias): 4 nov 2026; 13 jan 2027; 3
mar 2027; 29 abr 2027 (final internacionais); 3 jun 2027 (final
domésticos).

'Tamanho Turma' preenchido: 62 alunos (Class of 2026, oficial), dos
quais 32 internacionais (~52% da turma) -- consistente com a
classificação já existente de 'Quão Realista (Intl)' = 'Alta'.

'Taxa Empregabilidade %' preenchido: 89% (Class of 2025, ofertas de
emprego em até 6 meses). 'Salário Médio Pós USD' preenchido: $75.629
(Class of 2025, oficial). 'Principais Empregadores' preenchido com uma
amostra da lista oficial (FedEx, EY, McKinsey & Company, Amazon,
Goldman Sachs, Capital One, Walt Disney Company). 'Ranking QS'
preenchido: '#7 public master's em business analytics (QS 2025)' --
dado apresentado diretamente na página oficial do programa.

'Matérias Pré-Requisito' preenchido: nenhum pré-requisito formal de
disciplinas específicas; o programa aceita candidatos de "qualquer
área" para elegibilidade de auto-admissão.

CUIDADO APLICADO (dados NÃO encontrados, permanecem 'Verificar'):
'Taxa Aceitação %' (a página de Class Profile explicitamente não
fornece esse dado), 'Ranking US News' (não encontrado, distinto do QS
que foi encontrado), 'Entrevista' (não mencionado em nenhuma fonte
oficial consultada), 'Caminho Tuition Residente' já estava preenchido
("Não (F-1)", não Verificar, não tocado), e todos os campos mensais de
custo de vida (Aluguel, Mercado, Seguro Saúde, Utilities, Celular,
WiFi, Transporte, Outros Gastos, Custo Vida Total).

Fontes principais (pesquisadas via WebFetch, ago/2026):
- https://business.wisc.edu/graduate/masters/business-analytics/
  (formato, duração 10 meses, currículo, tuition confirmada, salário
  médio, ranking QS, deadlines)
- .../admissions/ (GPA, critérios completos de dispensa de GMAT/GRE,
  mínimos de inglês em 2 patamares, taxa de aplicação, documentos)
- .../admissions/class-profile/ (GPA médio 3.49, GRE médio, tamanho de
  turma 62, 32 internacionais, distribuição de áreas de graduação)
- .../careers/ (taxa de emprego 89%, salário médio $75.629,
  empregadores — todos Class of 2025)
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
    "Nome da Universidade": "UW-Madison",
    "Estado": "WI",
    "Cidade": "Madison",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Duração (meses)": 10,  # CORREÇÃO: oficial (bootcamp verão online + 2 semestres presenciais), não 12
    "GPA Mínimo": "3.0 cumulativo (GPAs internacionais devem ser equivalentes); auto-admissão para bacharéis UW-Madison com GPA ≥3.4",
    "GPA Médio Admitidos": 3.49,  # oficial, Class of 2026
    "GMAT/GRE Obrigatório": (
        "Exigido, salvo dispensa seletiva para candidatos com forte "
        "preparo quantitativo: graduação STEM/quantitativa; GPA ≥3.2 "
        "geral ou B+ em quantitativas; 2+ anos de experiência "
        "analítica/financeira/técnica; certificações (CPA/CFA/FRM); "
        "cursos prep de MBA; ou mestrado/doutorado prévio"
    ),
    "GMAT Médio": "Não divulgado (só GRE): Verbal médio 154 (faixa 80%: 144-162); Quantitativo médio 165 (faixa 80%: 160-170)",
    "Testes Inglês Aceitos": (
        "TOEFL mín. 92 (antes de 21/jan/26) ou 5.0 (novo formato, a "
        "partir dessa data); IELTS mín. 7.0; Duolingo mín. 125; "
        "dispensado para diploma de graduação de 4 anos lecionado em "
        "inglês"
    ),
    "Taxa Aplicação USD": 75,  # + $6 de processamento adicional para internacionais
    "Documentos Necessários": (
        "3 essays escritos (600 palavras cada) + 1 opcional; 3 video "
        "essays (2-3 min) + até 2 adicionais em 72h; resume; "
        "transcripts (não-oficiais na aplicação, oficiais até 1 jul); "
        "GMAT/GRE ou pedido de dispensa aprovado; teste de inglês "
        "(internacionais)"
    ),
    "Deadline Aplicação": "Ciclo jul/2027: 4 nov 2026; 13 jan 2027; 3 mar 2027; 29 abr 2027 (final internacionais); 3 jun 2027 (final domésticos)",
    "Rounds Aplicação": "5 rounds: 4 nov, 13 jan, 3 mar, 29 abr (final intl), 3 jun (final doméstico)",
    "Tamanho Turma": 62,  # Class of 2026, oficial (32 internacionais, ~52%)
    "Taxa Empregabilidade %": 89,  # Class of 2025, oficial (ofertas em até 6 meses)
    "Salário Médio Pós USD": 75629,  # Class of 2025, oficial
    "Principais Empregadores": "FedEx, Ernst & Young (EY), McKinsey & Company, Amazon, Goldman Sachs, Capital One, Walt Disney Company",
    "Ranking QS": "#7 public master's em business analytics (QS 2025)",
    "Matérias Pré-Requisito": "Nenhum pré-requisito formal de disciplinas específicas; aceita candidatos de qualquer área para elegibilidade de auto-admissão",
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
        raise ValueError("Nenhuma linha encontrada para UW-Madison (Madison, WI). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (UW-Madison, Madison, WI)")

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
