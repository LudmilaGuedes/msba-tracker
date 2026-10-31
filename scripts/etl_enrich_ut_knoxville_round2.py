# -*- coding: utf-8 -*-
"""
Data enrichment (round 2): UT Knoxville (Haslam College of Business),
Knoxville, TN -- MS in Business Analytics & AI (MSBAi). Rodada anterior:
etl_enrich_ut_knoxville.py.

Fontes (set/2026):
- haslam.utk.edu/business-analytics-statistics/msbai/ (10 meses, inicio
  no verao, lock-step de 36 creditos em 3 semestres; tuition para entrada
  em 2025: US$26.606 residente / US$54.371 nao-residente; >70% dos
  admitidos recebem algum funding; GA por merito com waiver de tuition +
  stipend, consideracao automatica; GRE/GMAT ou Canvas Exam de US$9;
  entrevista de 30 min; 2 recomendacoes; prazos para internacionais;
  ~90% aceitam oferta em ate 6 meses; salarios medianos por cargo)
- gradschool.utk.edu (taxa US$70 intl / US$60 domestico; Duolingo 120,
  TOEFL 80, IELTS 6.5; validade de 2 anos)

ALTERACOES:
- Taxa de aplicacao: Verificar -> 70.
- Ingles: Duolingo 120 confirmado na Graduate School (a pagina do programa
  so cita TOEFL/IELTS).
- Prazos no formato de rounds (regra do CLAUDE.md), com os prazos de
  internacionais: submissao 15/out, completa 15/dez (prioridade), final
  15/fev; admissao rolling desde agosto.
- Destaques: curriculo com IA (generative AI, deep learning, NLP,
  reinforcement learning), R/SQL/Python, eletivas.
- Principais Empregadores: salarios medianos por cargo da pagina oficial,
  mais a lista de empregadores de 2019 (antiga) com ressalva.
- GMAT/GRE: redacao mais clara (o Canvas Exam substitui).
O relatorio de outcomes 2024-25 da Haslam e so da graduacao e nao foi
usado. Taxa de empregabilidade (90) e salario (87.563) mantidos.
Uso: python scripts/etl_enrich_ut_knoxville_round2.py (a partir da raiz).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {'Nome da Universidade': 'UT Knoxville', 'Estado': 'TN', 'Cidade': 'Knoxville'}

UPDATES = {'Destaques Currículo': 'MS in Business Analytics & AI: lock-step de 36 créditos em 3 semestres '
                        '(início no verão); generative AI, deep learning, NLP e reinforcement '
                        'learning, além de R, SQL e Python; eletivas em supply chain analytics, '
                        'forecasting, causal inference e A/B testing',
 'Principais Empregadores': 'Salários medianos por cargo (página oficial): Data Analyst US$84.731, '
                            'BI Analyst US$97.963, Data Scientist US$108.020, Operations Analyst '
                            'US$83.640. Empregadores citados em 2019 (lista antiga): 3M, Dell, '
                            'Eastman Chemical, JPMorgan Chase, Microsoft, Pilot Flying J, Boeing, '
                            'Delta',
 'GMAT/GRE Obrigatório': 'Obrigatório (GRE ou GMAT), com dispensa possível; alternativa: Canvas '
                         'Exam próprio da Haslam (US$9); a avaliação quantitativa é obrigatória',
 'Testes Inglês Aceitos': 'TOEFL iBT 80 / IELTS 6.5 / Duolingo 120 (mínimos da Graduate School; '
                          'válidos por 2 anos)',
 'Deadline Aplicação': 'Round 1: 15 out (submissão, internacionais); Round 2: 15 dez (aplicação '
                       'completa, prioridade); Round 3: 15 fev (final para internacionais); '
                       'admissão rolling desde agosto até preencher a turma',
 'Taxa Aplicação USD': 70,
 'Rounds Aplicação': 'Rolling (aberto desde agosto até preencher a turma); internacionais: '
                     'submissão até 15/out, aplicação completa até 15/dez (prioridade) ou 15/fev '
                     '(final); taxa de US$70 (intl)',
 'Documentos Necessários': 'Resume; respostas às essay questions; histórico acadêmico oficial; 2 '
                           'cartas de recomendação (de preferência 1 acadêmica de área '
                           'quantitativa + 1 profissional); GRE/GMAT ou Canvas Exam (US$9); '
                           'TOEFL/IELTS/Duolingo; entrevista de 30 min',
 'Entrevista': 'Obrigatória: ~30 minutos, virtual ou presencial'}


def main():
    rotate_backup()
    print("Backup rotacionado (política fixa) em: data/msba_programs_database_backup.xlsx")

    wb = openpyxl.load_workbook(DB_PATH)
    ws = wb[SHEET_DATA]
    headers = [cell.value for cell in ws[1]]
    col = {h: i + 1 for i, h in enumerate(headers)}

    unknown = [h for h in UPDATES if h not in col]
    if unknown:
        raise ValueError(f"Colunas inexistentes no schema: {unknown}")

    rows = [
        r[0].row for r in ws.iter_rows(min_row=2)
        if all(r[col[k] - 1].value == v for k, v in MATCH.items())
    ]
    if len(rows) != 1:
        raise ValueError(f"Esperava 1 linha para {MATCH}, encontrou {len(rows)}. Abortando.")
    row_idx = rows[0]

    changed = []
    for header, new_value in UPDATES.items():
        cell = ws.cell(row=row_idx, column=col[header])
        if cell.value == new_value:
            continue
        changed.append((header, cell.value, new_value))
        cell.value = new_value
        cell.fill = PatternFill(fill_type=None)

    wb.save(DB_PATH)
    print(f"\n{len(changed)} campos atualizados na linha {row_idx} ({MATCH['Nome da Universidade']}):")
    for header, before, after in changed:
        print(f"  - {header}:")
        print(f"      Antes:  {before!r}")
        print(f"      Depois: {after!r}")


if __name__ == "__main__":
    main()
