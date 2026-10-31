"""
Data enrichment (RODADA 3): atualiza APENAS a linha 'UCF (Professional)'
(Professional MS in Management, Business Analytics track, Orlando, FL) no
msba_programs_database.xlsx. A linha 'UCF' (MS in Business Analytics
tradicional) NÃO é tocada -- o match exige o nome exato.

FONTES OFICIAIS (acessadas em 2026-10-19):
  - https://search.cm.ucf.edu/api/v1/programs/942/ (catálogo da pós da
    UCF; + /deadlines/ e /outcomes/)
  - https://studentaccounts.ucf.edu/tf-profmsmgmt-ba/

ACHADOS CRÍTICOS:
- Catálogo: "This track has been suspended effective Fall 2023" -- a
  trilha não aceita alunos novos.
- Catálogo: "considered a part-time program" (2 noites por semana,
  Downtown Orlando). Aluno F-1 precisa de matrícula full-time, então,
  mesmo se a trilha reabrisse, o formato não serve para visto de
  estudante.

CORREÇÕES:
- Duração: 20 meses ("A 20-month program") -- deixa de ser estimativa.
- Formato: part-time noturno, cohort, Downtown Orlando.
- Pré-requisitos: OBRIGATÓRIOS ("Minimum of college-level
  microeconomics, statistics or quantitative methods and calculus
  courses required") + teste diagnóstico na admissão. A planilha dizia
  "recomendado" (isso vale para o MS-BA tradicional, não para esta trilha).
- Empregabilidade / salário: feed oficial de outcomes desta trilha; o
  dado mais recente é 2019-20: 71% empregados full-time, ganho médio
  anual US$ 83.036 (amostra pequena, varia muito ano a ano:
  48.940 a 108.012).
- Fees Adicionais: a tabela 2026-27 chama os US$ 30.000 de "Total Annual
  Tuition and Fees" mas diz que incluem "all credit hours" -- ambíguo
  para um programa de 20 meses. Tuition Total mantida em 30.000, com a
  ressalva no texto.
- Deadline / Rounds / Dificuldade / Quão Realista: refletem a suspensão.
  O texto do deadline não cita mês, para o classificador de rounds não
  colocar a trilha em nenhum round.

NÃO ALTERADOS: STEM (sem fonte), bolsas, rankings, empregadores, inglês.
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
    "Nome da Universidade": "UCF (Professional)",
    "Estado": "FL",
    "Cidade": "Orlando",
}

UPDATES = {
    "Formato": (
        "Presencial, PART-TIME (cohort noturno, 2 noites por semana, "
        "Downtown Orlando) -- part-time não permite visto F-1"
    ),
    "Duração (meses)": 20,
    "Fees Adicionais USD": (
        "Incluído na tuition (créditos, apostilas, laptop, orientação, "
        "estacionamento). Ressalva: a tabela 2026-27 chama os US$ 30.000 de "
        "\"Total Annual Tuition and Fees\" -- não fica claro se é por ano "
        "ou pelo programa todo (20 meses)"
    ),
    "Quão Realista (Intl)": (
        "Não aplicável -- trilha com admissão suspensa desde Fall 2023 e "
        "part-time (incompatível com F-1)"
    ),
    "Taxa Empregabilidade %": 71,  # employed_full_time 2019-20, feed oficial de outcomes da trilha
    "Salário Médio Pós USD": 83036,  # avg_annual_earnings 2019-20, mesmo feed
    "Matérias Pré-Requisito": (
        "Obrigatórias: microeconomia, estatística ou métodos quantitativos "
        "e cálculo (nível universitário); teste diagnóstico exigido na "
        "admissão"
    ),
    "Deadline Aplicação": (
        "Verificar -- trilha com admissão suspensa desde Fall 2023 "
        "(catálogo da pós da UCF); o feed de prazos ainda mostra só um prazo "
        "doméstico de Fall, nenhum internacional"
    ),
    "Rounds Aplicação": (
        "Não há: a trilha está suspensa desde Fall 2023 (\"This track has "
        "been suspended effective Fall 2023\", catálogo da pós da UCF)"
    ),
    "Dificuldade Entrada": (
        "Inviável no momento -- admissão suspensa desde Fall 2023; e, por "
        "ser part-time, não serve para visto F-1"
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
        raise ValueError("Nenhuma linha encontrada para UCF (Professional) (Orlando, FL). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (UCF (Professional), Orlando, FL)")

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
