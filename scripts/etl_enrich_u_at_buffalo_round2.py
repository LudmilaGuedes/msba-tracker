# -*- coding: utf-8 -*-
"""
Data enrichment (round 2): University at Buffalo (SUNY), School of
Management -- MS in Business Analytics and Applied AI. Rodada anterior:
etl_enrich_u_at_buffalo.py.

Fontes (set/2026):
- management.buffalo.edu/.../business-analytics-applied-ai.html
  (full-time, 10 meses, STEM)
- .../business-analytics-applied-ai/admissions.html (turma de 2027
  comeca em 6/jul/2027; taxa de US$100, isenta so para militares e
  veteranos dos EUA; test-optional; resume + transcripts nao oficiais)
- .../admissions/ms-faq.html (TOEFL 95 preferido, Duolingo 120, IELTS 6.5,
  PTE 55; "No direct School of Management aid is available for the MS
  programs"; assistantships e empregos no campus via Handshake)
- .../information-systems/funding-scholarships-fellowships.html (tarifas
  de MS da School of Management para o FALL 2026: nao-residente
  US$13.255/semestre, custo total estimado US$39.556,53; residente
  US$5.655/semestre, US$19.910,53; prioridade para assistantships ate
  15/jan). A pagina e do MS in Information Systems, mas as tarifas sao as
  mesmas para os MS da School of Management; a pagina do MSBA ainda mostra
  as tarifas de 2024 (US$36.178,22).
- management.buffalo.edu/career-resource-center/.../ms-salary/business-analytics.html
  (turma de 2025: 52 formados, 33 procurando emprego, 58% empregados na
  formatura/3 meses, 94% em 6 meses; media US$87.143, mediana US$80.000,
  faixa US$65-135k)

ALTERACOES:
- Inicio: Fall -> julho (verao).
- Tuition atualizada para as tarifas do Fall 2026; custo por credito =
  total / 30.
- Bolsas: sem bolsa direta da SOM para MS. Assistantships sao competitivas
  (20h/semana, stipend + tuition scholarship). GOP e CyberCorps sao para
  cidadaos ou residentes. Realismo rebaixado de "Moderada" para "Baixa".
- Empregabilidade 94% (6 meses, sobre 33 que procuravam emprego numa
  turma de 52) e salario medio US$87.143.
Uso: python scripts/etl_enrich_u_at_buffalo_round2.py (a partir da raiz).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {'Nome da Universidade': 'U at Buffalo SUNY', 'Estado': 'NY', 'Cidade': 'Buffalo'}

UPDATES = {'Início': 'Julho (verão) -- a turma de 2027 começa em 6/jul/2027',
 'Tamanho Turma': 52,
 'Destaques Currículo': 'Full-time de 10 meses; STEM (24 meses extras de OPT); foco em analytics, '
                        'IA aplicada, visualização de dados e storytelling',
 'Custo/Crédito (Intl) USD': 1318.55,
 'Custo/Crédito (Residente) USD': 663.68,
 'Tuition Total (Intl) USD': 39556.53,
 'Tuition Total (Residente) USD': 19910.53,
 'Fees Adicionais USD': 'Incluídas no custo total estimado (tarifas de MS da School of Management '
                        'para o Fall 2026: US$13.255/semestre não-residente, US$5.655 residente)',
 'Bolsas Disponíveis': 'Limitadas',
 'Tipos de Bolsa': 'Não há bolsa direta da School of Management para os MS; assistantships '
                   '(GA/TA/RA) competitivas pelo Handshake (20h/semana, stipend + tuition '
                   'scholarship). GOP e CyberCorps (até US$34k) são para cidadãos/residentes dos '
                   'EUA',
 'Quão Realista (Intl)': 'Baixa -- a SOM não dá bolsa própria para o MS; só sobram assistantships '
                         'competitivas e empregos no campus',
 'Taxa Empregabilidade %': 94,
 'Salário Médio Pós USD': 87143,
 'Principais Empregadores': 'M&T Bank, BNY Mellon, Chubb, Comcast, Centene Corp., BlueCross '
                            'BlueShield (entre mais de 50 listados); turma de 2025: mediana '
                            'US$80k, faixa US$65-135k',
 'Testes Inglês Aceitos': 'TOEFL iBT 95 (preferido pelo comitê) / IELTS 6.5 / PTE 55 / Duolingo '
                          '120 (mínimo da universidade)',
 'Rounds Aplicação': 'Rolling; prioridade para bolsas, fellowships e assistantships até 15/jan; '
                     'taxa de US$100 (isenta só para militares e veteranos dos EUA)',
 'Documentos Necessários': 'Resume em PDF com datas; transcripts não oficiais (PDF) de todas as '
                           'instituições; GMAT/GRE opcionais (test-optional); teste de inglês; '
                           'cartas e entrevista não mencionadas na página de admissão'}


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
