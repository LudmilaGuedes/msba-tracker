"""
Data enrichment: atualiza APENAS a linha da 'Cal Poly SLO' (Orfalea
College of Business, San Luis Obispo, CA) no msba_programs_database.xlsx.

Foco: MS in Business Analytics, presencial, acelerado, 30 unidades
(18 core + 12 eletivas), STEM (já confirmado).

'Duração (meses)' CORRIGIDO de 12 para 10 -- a página oficial atual
confirma "10 months" para o formato acelerado full-time. 'Início'
CORRIGIDO de 'Fall' para 'Summer' -- a mesma página confirma "Start
Date: Summer 2026", não mais entrada no Fall.

'Total Créditos' preenchido: 30 (18 core + 12 eletivas, confirmado
oficialmente pela página de currículo). 'Custo/Crédito' preenchido por
DERIVAÇÃO do total Intl já presente e não alterado ($37.389 ÷ 30 =
$1.246,30) -- consistente com a nota já existente 'Caminho Tuition
Residente' = 'Mesma tuition intl/resident' (programa self-support,
tarifa flat). 'Tuition Total (Residente)' preenchido com o MESMO valor
do Intl ($37.389), pelo mesmo motivo de tarifa flat.

'GMAT/GRE Obrigatório' refinado: dispensado ATUALMENTE para todos os
candidatos, mas envio incentivado se o GPA de graduação for abaixo de
3.0 ou se o histórico não usar uma escala de 4 pontos (código
institucional 4038). 'GPA Mínimo' (2.5, já presente) CONFIRMADO
oficialmente, sem alteração necessária.

'Deadline Aplicação' REFINADO com dados mais completos e precisos da
página oficial de aplicação: ciclo abre 1 out 2026; deadline
internacional 1 abr 2027; deadline doméstico 1 jun 2027 -- o valor
anterior ("Priority Feb 1, Intl Apr 1") citava uma data de prioridade
de 1 fev que não foi corroborada nesta fonte; mantida apenas a
informação confirmada oficialmente.

'Documentos Necessários' preenchido com a lista oficial: statement of
purpose, resume, 2 cartas de recomendação, transcripts oficiais
(enviados diretamente pela instituição), GMAT/GRE opcional, exame de
proficiência em inglês (internacionais).

'Bolsas Disponíveis' e 'TA/RA Disponível' preenchidos com achado
IMPORTANTE (semi-negativo) da FAQ oficial: "Generally, only financial
aid in the form of loans is available for this program" -- ou seja,
NÃO há bolsa/assistantship institucional própria do MSBA confirmada;
apenas empréstimos, e fontes externas (bolsas privadas, benefícios de
veteranos) são aceitas mas não oferecidas pelo programa em si.

'Salário Médio Pós USD' JÁ ESTAVA CORRETO -- a página oficial de
Career Paths cita EXATAMENTE o mesmo valor já presente na planilha
($95.375, mediana, 2025 Graduate Status Report, 13 respondentes, 60
dias pós-formatura), validando cruzadamente o dado sem necessidade de
correção. 'Principais Empregadores' preenchido com uma amostra da
lista oficial expandida: Amazon, Apple, Google, Microsoft, Deloitte,
Accenture, EY, KPMG, Kaiser Permanente, SpaceX.

CUIDADO APLICADO (dados NÃO encontrados, permanecem 'Verificar'):
'Fees Adicionais USD', 'Tamanho Turma', 'GPA Médio Admitidos', 'GMAT
Médio', 'Taxa Aceitação %', 'Ranking US News', 'Ranking QS' (a página
de career paths confirma explicitamente ausência de dado de ranking),
'Taxa Empregabilidade %' (já presente, não Verificar, não tocado --
nenhuma fonte nova confirmou ou contradisse), 'Entrevista', 'Rounds
Aplicação' (não há rounds numerados, apenas 2 deadlines por status),
'Seguro Saúde (programa) USD' (já presente, não Verificar, não
tocado), e todos os campos mensais de custo de vida ainda vazios em
San Luis Obispo, CA.

Fontes principais (pesquisadas via WebFetch, ago/2026):
- https://orfalea.calpoly.edu/graduate-programs/ms-business-analytics/
  (formato, duração 10 meses, início Summer 2026, currículo geral)
- .../apply (GPA, GMAT/GRE dispensado, deadlines, taxa de aplicação
  $70 já confirmada, documentos, pré-requisitos já confirmados)
- .../curriculum (30 unidades confirmadas: 18 core + 12 eletivas)
- .../faqs (achado de que só há financiamento via empréstimos, sem
  bolsa/assistantship institucional própria)
- .../msba-career-paths (salário mediano $95.375 validado
  cruzadamente, empregadores expandidos, ausência de ranking)
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
    "Nome da Universidade": "Cal Poly SLO",
    "Estado": "CA",
    "Cidade": "San Luis Obispo",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Duração (meses)": 10,  # CORREÇÃO: oficial, formato acelerado, não 12
    "Início": "Summer",  # CORREÇÃO: oficial "Start Date: Summer 2026", não Fall
    "Total Créditos": 30,  # oficial: 18 core + 12 eletivas
    "Custo/Crédito (Intl) USD": 1246.30,  # 37.389 / 30 (total já presente, não alterado)
    "Custo/Crédito (Residente) USD": 1246.30,  # tarifa flat (self-support), igual ao Intl
    "Tuition Total (Residente) USD": 37389,  # tarifa flat, igual ao Intl
    "GMAT/GRE Obrigatório": (
        "Dispensado atualmente para todos os candidatos; envio "
        "incentivado se GPA <3.0 ou histórico sem escala de 4 pontos "
        "(código institucional 4038)"
    ),
    "Deadline Aplicação": "Ciclo abre 1 out 2026; internacional 1 abr 2027; doméstico 1 jun 2027",
    "Documentos Necessários": (
        "Statement of purpose; resume; 2 cartas de recomendação; "
        "transcripts oficiais (enviados diretamente pela "
        "instituição); GMAT/GRE opcional; exame de proficiência em "
        "inglês (internacionais)"
    ),
    "Bolsas Disponíveis": "Geralmente apenas empréstimos disponíveis para este programa (confirmado oficialmente); bolsas privadas/externas, benefícios de veteranos aceitos, mas sem bolsa institucional própria do MSBA",
    "TA/RA Disponível": "Não confirmado — FAQ oficial menciona apenas financiamento via empréstimos, sem menção a assistantships para este programa",
    "Principais Empregadores": "Amazon, Apple, Google, Microsoft, Deloitte, Accenture, EY, KPMG, Kaiser Permanente, SpaceX",
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
        raise ValueError("Nenhuma linha encontrada para Cal Poly SLO (San Luis Obispo, CA). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Cal Poly SLO, San Luis Obispo, CA)")

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
