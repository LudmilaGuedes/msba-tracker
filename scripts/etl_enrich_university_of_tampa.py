# -*- coding: utf-8 -*-
"""
Data enrichment: University of Tampa (Sykes College of Business), Tampa, FL
-- M.S. in Business Analytics, STEM. A linha foi adicionada na varredura
nacional (etl_add_nationwide_sweep.py).

Fontes (set/2026):
- ut.edu/graduate-degrees/graduate-business/ms-in-business-analytics
- ut.smartcatalogiq.com (catalogo do MSBA; graduate-tuition 2026-27)
- ut.edu/graduate-degrees/graduate-international-admissions
- ut.edu/admissions/financial-aid/financial-aid-for-graduates

*** CONFLITO DE FORMATO: a pagina do programa diz "All classes are offered
online". Ja a pagina de admissao internacional lista o MSBA como elegivel
para F-1 e diz que alunos F-1 so podem fazer 1 disciplina online por
semestre. Se todas as disciplinas forem online, na pratica o curso fica
inviavel com visto F-1. Registrado em Formato para confirmar com
gradinfo@ut.edu antes de aplicar. ***

TUITION: US$789 por credito em 2026-27 (antes US$758) + Student Service
Fee de US$50/semestre. Sao 32 creditos obrigatorios, mais 8 de
pre-requisitos se nao forem dispensados. Total: 32 x 789 = US$25.248
(US$31.560 com os pre-requisitos).

BOLSAS: Graduate Assistantship por convite e competitivo: waiver de ate 12
creditos por semestre (ate 36 creditos com renovacao) + stipend de US$1.500
por semestre; 20h/semana no campus (seg-sex, 8h-17h); GPA >= 3,25. A
pagina nao diz se vale para internacionais.

GMAT/GRE obrigatorio. A dispensa e so para diploma de business de
instituicao AACSB; experiencia profissional nao conta.

EMPREGABILIDADE: continua "Verificar". Os dados achados sao do MBA (35%
empregados em 3 meses, US$79.375) ou do PayScale (Sykes em geral).
Uso: python scripts/etl_enrich_university_of_tampa.py (a partir da raiz).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {'Nome da Universidade': 'University of Tampa', 'Estado': 'FL', 'Cidade': 'Tampa'}

UPDATES = {'Formato': "VERIFICAR -- conflito: a página do programa diz 'All classes are offered online', mas "
            'a admissão internacional lista o MSBA como elegível para F-1 (que só pode fazer 1 '
            'disciplina online por semestre); confirmar com gradinfo@ut.edu se há turmas '
            'presenciais suficientes',
 'Duração (meses)': 12,
 'Início': 'Fall, Spring ou Summer',
 'Destaques Currículo': 'Econometrics, Analytics in Finance, Data Mining for Business, Machine '
                        'Learning for Business Analytics, Big Data for Analytics, Marketing '
                        'Analytics, Decision Modeling, Advanced Forecasting & Predictive '
                        'Analytics; mínimo de 3 semestres',
 'Total Créditos': 32,
 'Custo/Crédito (Intl) USD': 789,
 'Custo/Crédito (Residente) USD': 789,
 'Tuition Total (Intl) USD': 25248,
 'Tuition Total (Residente) USD': 25248,
 'Fees Adicionais USD': 'Student Service Fee de US$50/semestre (2026-27); mais 8 créditos de '
                        'pré-requisitos (~US$6.312) se não forem dispensados',
 'Bolsas Disponíveis': 'Sim',
 'Tipos de Bolsa': 'Graduate Assistantship (por convite, competitivo): waiver de até 12 créditos '
                   'por semestre (até 36 créditos com renovação) + stipend de US$1.500/semestre',
 'Quão Realista (Intl)': 'Moderada -- o GA é por convite e competitivo, e a página não diz se vale '
                         'para internacionais; ele exige presença no campus em horário comercial, '
                         'o que conflita com um curso só online',
 'TA/RA Disponível': 'Sim (Graduate Assistantship por convite)',
 'Processo TA/RA': 'O Office of Graduate and Continuing Studies convida candidatos academicamente '
                   'qualificados; 20h/semana no campus (seg-sex, 8h-17h), 8-14 créditos por '
                   'semestre e GPA ≥ 3.25',
 'Graduate Assistantships': 'GA com waiver de até 12 créditos/semestre + US$1.500/semestre; '
                            'disponível para novos alunos full-time de pós em business',
 'GMAT/GRE Obrigatório': 'Obrigatório (GMAT ou GRE); dispensa só para diploma de business de '
                         'instituição AACSB, analisada por comitê (experiência profissional não '
                         'conta)',
 'Matérias Pré-Requisito': 'Foundations of Finance, Database Fundamentals e Statistics for '
                           'Decision Makers (8 créditos); dispensáveis com nota B ou melhor nos '
                           'últimos 7 anos em escola AACSB',
 'Testes Inglês Aceitos': 'IELTS 6.5 / TOEFL 90 / Duolingo 120',
 'Deadline Aplicação': 'Fall: fim de maio (recomendado para internacionais); Spring: outubro; '
                       'admissão por semestre',
 'Taxa Aplicação USD': 40,
 'Rounds Aplicação': 'Sem rounds; admissão contínua por semestre (Fall, Spring ou Summer)',
 'Documentos Necessários': 'Transcripts oficiais; GMAT/GRE (salvo dispensa); personal statement de '
                           '500-1.000 palavras; resume; 2 formulários de referência profissional; '
                           'extrato bancário carimbado (até 6 meses) para o I-20'}


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
