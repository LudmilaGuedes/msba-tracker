# -*- coding: utf-8 -*-
"""
Data enrichment: Xavier University (Williams College of Business),
Cincinnati, OH -- Master of Science in Business Analytics. A linha veio da
varredura nacional (etl_add_nationwide_sweep.py).

Fontes (set/2026):
- xavier.edu/master-science-business-analytics/index (hibrido ~70%
  presencial / ~30% online; 12 meses full-time; 30 creditos; inicio no
  Fall; aplicacao recomendada ate 1/ago, sem prazo formal; GMAT/GRE
  opcional; 2 recomendacoes; concentracoes em Customer e Healthcare
  Services; empregadores)
- xavier.edu/tuition-fees/graduate-current (2026-27: US$966/credito; Student
  ID US$15 e GSA US$15 por semestre)
- xavier.edu/international-admission/applications-and-requirements/graduate
  (TOEFL iBT 79 / 4.5 na nova escala; IELTS 6.5; avaliacao de credenciais)
- xavier.edu/williams/business-graduate-scholarships (bolsas de 25% de
  desconto na tuition por semestre; Loyal Musketeer 25%) e busca (GA com
  ate 24 creditos de tuition + salario por hora)

*** DUOLINGO NAO CONFIRMADO: o valor anterior "Duolingo 105+" veio de um
snippet de busca. As paginas oficiais de admissao internacional e de pos
citam so TOEFL/IELTS. Registrado com ressalva; se a Xavier nao aceitar
Duolingo, a linha deixa de atender ao criterio da varredura. ***

Taxa de aplicacao: nenhuma pagina oficial informa; fica "Verificar".
Uso: python scripts/etl_enrich_xavier.py (a partir da raiz).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {'Nome da Universidade': 'Xavier University', 'Estado': 'OH', 'Cidade': 'Cincinnati'}

UPDATES = {'Duração (meses)': 12,
 'Início': 'Fall',
 'Destaques Currículo': 'Data Wrangling, Data Mining & Predictive Analytics, Data Visualization & '
                        'Storytelling, Ethics in Data Analytics, Database for Analytics, Analytics '
                        'Practicum ou estágio; concentrações em Customer ou Healthcare Services',
 'Total Créditos': 30,
 'Tuition Total (Intl) USD': 28980,
 'Tuition Total (Residente) USD': 28980,
 'Fees Adicionais USD': 'Student ID Fee de US$15 + Graduate Student Association Fee de US$15 por '
                        'semestre (2026-27)',
 'Bolsas Disponíveis': 'Sim',
 'Tipos de Bolsa': 'Bolsas da Williams College de 25% de desconto na tuition por semestre '
                   '(Graduate Business Ignite, Breaking Barriers, JAB in Business); Loyal '
                   'Musketeer 25% (ex-alunos Xavier); Graduate Assistantship com até 24 créditos '
                   'de tuition + salário por hora',
 'Quão Realista (Intl)': 'Moderada -- os descontos de 25% e o GA existem, mas a página não diz se '
                         'internacionais são elegíveis; o GA cobriria até 24 dos 30 créditos',
 'TA/RA Disponível': 'Sim (Graduate Assistantship: até 24 créditos de tuition + salário por hora)',
 'Graduate Assistantships': 'Até 24 créditos de tuition + salário por hora (valor não divulgado)',
 'Principais Empregadores': "Procter & Gamble, Cincinnati Children's, GE Aerospace",
 'GMAT/GRE Obrigatório': 'Opcional',
 'Testes Inglês Aceitos': 'TOEFL iBT 79 (ou 4.5 na nova escala) / IELTS 6.5 (página oficial de '
                          'admissão internacional); Duolingo 105 citado em busca, mas ausente das '
                          'páginas oficiais -- confirmar',
 'Deadline Aplicação': 'Sem prazo formal; recomendado aplicar até 1 ago (início no Fall)',
 'Rounds Aplicação': 'Sem rounds; admissão contínua (rolling)',
 'Documentos Necessários': 'Aplicação online; personal statement; resume; transcripts oficiais + '
                           'avaliação de credenciais estrangeiras; 2 cartas de recomendação; '
                           'GMAT/GRE opcional; teste de inglês'}


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
