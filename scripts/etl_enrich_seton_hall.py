"""
Data enrichment: atualiza APENAS a linha da 'Seton Hall University'
(Stillman School of Business, South Orange, NJ) no
msba_programs_database.xlsx.

Foco exclusivo: M.S. in Business Analytics, formato híbrido com
componente presencial obrigatório (aulas em horário noturno, 1-2x por
semana, no campus de South Orange; eletivas majoritariamente online),
STEM, 30 créditos (10 disciplinas), 18-24 meses. Deep research em:
custo total real de tuition para internacionais, bolsas/assistantships,
requisitos de GPA/GMAT/GRE/inglês, e custo de vida em South Orange, NJ.

VERIFICACAO DE FORMATO: a planilha ja indicava "On Ground/Hybrid" com
eletivas online -- CONFIRMADO oficialmente que o programa tem um
componente presencial genuino no campus (nao e 100% online), entao
continua qualificado como "presencial" para o foco desta pesquisa.
Nenhuma alteracao no campo 'Formato'.

TUITION: PREENCHIDO (estava 'Verificar' em ambos): $1.570/credito flat
(confirmado por 2 fontes independentes, mesma taxa para todos,
independente de residencia -- universidade PRIVADA). Total = 30 x
$1.570 = $47.100 (Intl e Residente).

CORRECAO DE CATEGORIZACAO: 'Fees Adicionais USD' continha o custo dos 2
cursos de pre-qualificacao ($475 cada) -- um custo CONDICIONAL (so para
quem nao tem bagagem em estatistica/Excel, dispensavel, e explicitamente
"sem impacto na decisao de admissao"), nao uma fee obrigatoria do
programa. Esse dado foi MOVIDO para 'Materias Pre-Requisito' (campo mais
adequado) com detalhes completos (dispensas possiveis, exame de
challenge $60), e 'Fees Adicionais USD' foi resetado para 'Verificar'
ate uma fee mandatoria real ser encontrada.

'Principais Empregadores' preenchido com a lista ESPECIFICA de
empregadores de egressos do MSBA (pagina oficial): Zoetis, Verizon,
Stryker, BakerTilly. Um salario anual de mercado ($131.410, media
regional de "management/business analysts" na regiao de NY, fonte tipo
BLS) NAO foi usado para 'Salario Medio Pos USD' por nao ser especifico
dos egressos do programa -- mesmo criterio de cautela ja aplicado a
outras universidades desta planilha.

Custo de vida: 'Aluguel Privativo/mes' preenchido com $850 (quarto
privativo com banheiro, South Orange NJ, fonte especifica encontrada).
'Transporte/mes' preenchido com $110 (passe mensal do NJ Transit; o
shuttle/onibus local e gratuito com ID de aluno, mas o transporte
regional para NYC/arredores nao e). 'Seguro Saude (programa)'
permanece 'Verificar' -- confirmado como obrigatorio para
internacionais, mas nenhuma fonte trouxe o valor em dolares.

Fontes principais (pesquisadas via busca web, ago/2026):
- https://www.shu.edu/academics/master-of-science-in-business-
  analytics.html (formato hibrido confirmado, duracao 18-24 meses,
  curriculo, deadlines de prioridade por semestre, empregadores)
- https://www.shu.edu/business/m-s-in-business-analytics-
  prequalification-courses.html (cursos de pre-qualificacao, custos,
  dispensas)
- Busca web cruzada: tuition $1.570/credito flat (2 fontes)
- Busca web: aluguel de quarto privativo em South Orange, NJ; NJ
  Transit ($100-120/mes) vs onibus local gratuito com ID
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
    "Nome da Universidade": "Seton Hall University",
    "Estado": "NJ",
    "Cidade": "South Orange",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Intl) USD": 1570,  # oficial, flat
    "Custo/Crédito (Residente) USD": 1570,  # privada, mesma taxa
    "Tuition Total (Intl) USD": 47100,  # 30 créditos x $1.570
    "Tuition Total (Residente) USD": 47100,  # privada, mesma taxa
    "Fees Adicionais USD": VERIFICAR,  # CORREÇÃO: valor anterior era custo condicional de pré-requisito, não fee mandatória (movido para Matérias Pré-Requisito)
    "Matérias Pré-Requisito": (
        "2 cursos de pré-qualificação (BMBA 8019 Statistics + BMBA "
        "8020 IT Management) exigidos para quem não tem bagagem em "
        "estatística/Excel, $475 cada ($950 total); sem impacto na "
        "decisão de admissão; dispensável com B+ em estatística de "
        "graduação ou exame de challenge ($60/exame)"
    ),
    "Aluguel Privativo/mês USD": 850,  # quarto privativo com banheiro, South Orange NJ
    "Transporte/mês USD": 110,  # passe mensal NJ Transit (shuttle/ônibus local é gratuito com ID de aluno)
    "GPA Mínimo": 3.0,
    "Experiência Profissional": "Não obrigatória, mas experiência prévia em analytics é vista favoravelmente",
    "Testes Inglês Aceitos": "IELTS/TOEFL/Duolingo aceitos (scores mínimos não divulgados publicamente)",
    "Deadline Aplicação": "Fall priority: 31 mai; Spring priority: 31 out; Summer priority: 31 mar (admissão contínua/rolling)",
    "TA/RA Disponível": "Sim (mencionado junto às bolsas de mérito; candidatura até o priority deadline)",
    "Processo TA/RA": "Candidatura até o priority deadline (mesmo processo de bolsas de mérito, sem formulário separado detalhado)",
    "Graduate Assistantships": "Disponíveis, aplicar até o priority deadline (processo específico não detalhado publicamente)",
    "Quão Realista (Intl)": "Moderada — bolsas de mérito podem reduzir o custo em até 50%; GA mencionado mas processo detalhado não divulgado",
    "Principais Empregadores": "Zoetis, Verizon, Stryker, BakerTilly",
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
        raise ValueError("Nenhuma linha encontrada para Seton Hall University (South Orange, NJ). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Seton Hall University, South Orange, NJ)")

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
