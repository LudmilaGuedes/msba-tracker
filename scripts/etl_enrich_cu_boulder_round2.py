"""
Data enrichment (round 2): atualiza a linha da 'CU Boulder' (Leeds
School of Business, MSBA) no msba_programs_database.xlsx, a pedido do
usuário: confirmar tuition, bolsas, taxa de aplicação,
empregabilidade/salário e requisitos de aplicação.

FONTES (oficiais, colorado.edu/business, via WebFetch/WebSearch em
2026-09-08):
- https://www.colorado.edu/business/ms-programs/masters-program-business-analytics/msba-msma-and-mssca-estimated-cost-attendance
- https://www.colorado.edu/business/ms-programs/masters-programs-admissions
- https://www.colorado.edu/business/scholars (bolsas -- só cobre
  undergrad, não MS)
- https://www.colorado.edu/business/current-students/current-students/office-career-strategy/industry-coaching/placement-statistics
  (outcomes -- só cobre undergrad, não MS)

CONFIRMADO SEM MUDANÇA (o valor já na planilha bateu exatamente com a
fonte oficial):
- Custo/Crédito Intl ($1.753,12) e Residente ($1.261,58): confirmado
  batendo com Tuition & Fees oficiais / 33 créditos.
- Tuition Total Intl ($57.853) + Fees Adicionais ($1.424) = $59.277,
  que é EXATAMENTE o "International: Tuition & Fees $59.277" da
  página oficial de custo de atendimento (coorte Summer
  2026/Fall 2026/Spring 2027, baseada nas tarifas 2025-2026).
- Custo Vida Total ($31.818): bate exato com "Living Expenses
  (12-month): $31.818" da mesma página oficial.
- GPA Mínimo (2,75), GMAT/GRE (opcional), Testes de Inglês (TOEFL
  80/IELTS 6.5/Duolingo 115/PTE 58/Cambridge 180), Documentos
  Necessários (resume, 1 carta de recomendação, essay até 750
  palavras, transcripts não-oficiais, GMAT/GRE opcional), Deadlines/
  Rounds (5 out, 30 nov, 1 fev, 3 fev final intl, 10 mai final
  doméstico) e Entrevista (não exigida p/ MS): TODOS confirmados
  batendo exatamente com a página oficial de admissões -- nenhuma
  mudança necessária, dado já estava correto e atualizado.

CORRIGIDO/ESCLARECIDO -- TAXA DE APLICAÇÃO: a planilha já tinha $0, e
isso está correto NESTE MOMENTO, mas o motivo importa: a taxa NOMINAL
é $70 (não é uma escola que nunca cobra, como Illinois Tech/Seattle
U/Baylor/SMU) -- ela está "currently waived" (isenta atualmente, cita
textual da própria página oficial), uma decisão institucional em
vigor agora, não uma política permanente. Documentado como nota em
'Rounds Aplicação' para não dar a falsa impressão de que é sempre
gratuita.

PREENCHIDO (estava 'Verificar'):
- Bolsas Disponíveis: Sim.
- Tipos de Bolsa: bolsa de mérito, concedida no momento da aceitação
  (linguagem da página não confirma se é 100% automática ou requer
  alguma indicação no formulário, mas não há processo de aplicação de
  bolsa SEPARADO). Internacionais são orientados a aplicar no Round 1
  ou 2 para melhor chance de bolsa -- fundos são limitados e
  priorizados para quem aplica cedo. Alunos matriculados podem
  concorrer a bolsas adicionais a cada semestre. NENHUM valor ou
  faixa em dólares é publicado (diferente da Fordham). Dado de
  contexto (não específico do MSBA): Leeds School inteiro concedeu
  $2.963.707 em bolsas para 634 alunos de pós-graduação no ano letivo
  2024-25 (~$4.674/aluno em média, mas é uma média agregada de TODOS
  os programas de mestrado, não só MSBA).
- Quão Realista (Intl): Moderada -- fundos limitados, priorizados por
  ordem de aplicação (rounds 1-2), sem valor garantido.

NÃO CONFIRMADO NEM CONTESTADO -- 'Taxa Empregabilidade %' (90) e
'Salário Médio Pós USD' ($83.885) permanecem como estavam: a única
página oficial de outcomes que encontrei
(.../placement-statistics) é EXCLUSIVAMENTE de graduação (undergrad),
por curso (Accounting, Business Analytics, Finance, etc.), e não
cobre os programas de MS. Não encontrei o relatório que originou os
90%/$83.885/lista de empregadores já na planilha -- pode ser de uma
fonte legítima de uma pesquisa anterior que não foi re-localizada
nesta rodada, mas também não posso confirmá-la agora. Reportado ao
usuário como incerto, sem alterar o valor.
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
    "Nome da Universidade": "CU Boulder",
    "Estado": "CO",
    "Cidade": "Boulder",
}

UPDATES = {
    "Bolsas Disponíveis": "Sim",
    "Tipos de Bolsa": (
        "Bolsa de mérito concedida no momento da aceitação (sem "
        "aplicação de bolsa separada identificada). Internacionais: "
        "aplicar no Round 1 ou 2 dá melhor chance -- fundos limitados, "
        "priorizados por ordem de aplicação. Alunos matriculados podem "
        "concorrer a bolsas adicionais a cada semestre. Nenhum valor "
        "ou faixa em dólares publicado. Contexto: Leeds School inteiro "
        "concedeu $2.963.707 em bolsas p/ 634 alunos de pós-graduação "
        "(2024-25) -- média agregada de TODOS os mestrados, não só "
        "MSBA."
    ),
    "Quão Realista (Intl)": "Moderada -- fundos limitados, priorizados por ordem de aplicação (rounds 1-2), sem valor garantido",
    "Rounds Aplicação": (
        "5 prazos: 5 out, 30 nov, 1 fev, 3 fev (final intl), 10 mai "
        "(final doméstico). Taxa de aplicação: nominal $70, mas "
        "ATUALMENTE isenta ('currently waived', confirmado na página "
        "oficial) -- não é política permanente, confirmar vigência "
        "antes de aplicar."
    ),
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
        raise ValueError("Nenhuma linha encontrada para CU Boulder. Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (CU Boulder)")

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
