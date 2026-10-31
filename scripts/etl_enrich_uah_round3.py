# -*- coding: utf-8 -*-
"""
Data enrichment (round 3): UAH (University of Alabama in Huntsville),
Huntsville, AL -- MS in Business Analytics, presencial, STEM.

O pedido da rodada 2 (etl_enrich_uah_round2.py) foi repetido. Tratei
como pedido para tentar de novo os campos que ainda estavam "Verificar",
por fontes que nao tinham sido usadas.

Fontes novas nesta rodada (set/2026):
- uah.edu/admissions/graduate/financial-aid/costs (Cost of Attendance
  oficial 2026-27)
- uah.edu/heoa/student-achievement
- uah.edu/health-services/students/student-health-insurance-plan e
  uah.myahpcare.com
- uah.edu/business/grad/degrees/analytics

ACHADOS:

1) CUSTO DE VIDA PREENCHIDO a partir do Cost of Attendance oficial de
   2026-27 para nao-residente. A tuition do orcamento e US$28.840, igual a
   2 x US$14.420, o que confirma que o orcamento cobre 2 semestres (~9
   meses) e bate com os valores de tuition da rodada 2. Valores:
   Food & Housing 12.952; Books 2.936; Transportation (nao-residente)
   2.610; Miscellaneous 3.328.
   Metodologia igual a de etl_enrich_cincinnati_round3.py:
     Transporte/mes     = 2.610 / 9 = 290,00
     Outros Gastos/mes  = 3.328 / 9 = 369,78
     Custo Vida Total   = (12.952 + 2.936 + 2.610 + 3.328) / 9 x 12
                        = 29.101,33
   Diferenca em relacao a Cincinnati: o MSBA da UAH dura 12 meses, e nao
   9, entao o total foi pro-rateado para 12 meses. O seguro saude nao
   entra nesse total, porque tem campo proprio.
   Aluguel/mes e Mercado/mes nao foram separados, porque a fonte junta
   "Food & Housing" (mesmo criterio de Cincinnati). Mercado ja tinha
   US$300 e nao foi alterado.

2) SEGURO SAUDE: continua "Verificar". Internacionais sao inscritos
   automaticamente no plano BCBS da UAH, mas os valores so aparecem dentro
   do portal de adesao (uah.myahpcare.com), que nao foi possivel ler.

3) EMPREGABILIDADE/SALARIO: continua "Verificar". Os unicos dados
   publicados agregam a universidade inteira: 78% empregados ou na
   pos-graduacao em 1 ano (todos os niveis), media de US$65.592 (todos os
   niveis, 2025) e ~US$76.934 para todos os mestrados, este ultimo de
   fonte terciaria. Nenhum e especifico do MSBA, entao nao foram usados,
   para evitar contaminacao entre programas.

4) DESTAQUES CURRICULO enriquecido com o texto da pagina oficial do
   programa (R/Python, desenho de experimentos, estatistica e ML) e com o
   credenciamento AACSB.

Continua "Verificar": Tamanho Turma, Aluguel/mes, Seguro Saude,
Seguro Residencial, Utilities, Celular, WiFi, Rankings, Taxa
Empregabilidade, Salario Medio, Taxa Aceitacao, Entrevista.

Uso: python scripts/etl_enrich_uah_round3.py (a partir da raiz).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "UAH",
    "Estado": "AL",
    "Cidade": "Huntsville",
}

UPDATES = {
    "Transporte/mês USD": 290.00,
    "Outros Gastos/mês USD": 369.78,
    "Custo Vida Total (programa) USD": 29101.33,
    "Destaques Currículo": (
        "R/Python, desenho de experimentos, métodos estatísticos e machine learning; "
        "College of Business com acreditação AACSB; aulas noturnas e opções online "
        "(desenhado também para profissionais que trabalham)"
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
        raise ValueError("Nenhuma linha encontrada para UAH (Huntsville, AL). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (UAH, Huntsville, AL)")

    no_fill = PatternFill(fill_type=None)

    before = {}
    after = {}
    for header, new_value in UPDATES.items():
        col_idx = header_to_col[header]
        cell = ws.cell(row=row_idx, column=col_idx)
        before[header] = cell.value
        cell.value = new_value
        cell.fill = no_fill
        after[header] = new_value

    wb.save(DB_PATH)

    print(f"\n{len(UPDATES)} campos atualizados na linha {row_idx} (UAH):\n")
    for header in UPDATES:
        print(f"  - {header}:")
        print(f"      Antes:  {before[header]!r}")
        print(f"      Depois: {after[header]!r}")


if __name__ == "__main__":
    main()
