# -*- coding: utf-8 -*-
"""
Data enrichment (round 3): Texas A&M (Mays Business School), Houston, TX --
MS in Analytics. Rodadas anteriores: etl_enrich_texas_am.py e
etl_enrich_texas_am_round2.py.

Fontes (set/2026):
- mays.tamu.edu/ms-analytics/admissions/ (datas do Fall 2027, formato,
  custo de US$70.000, waiver de GRE/GMAT)
- mays.tamu.edu/graduate/masters/ms-programs/analytics/
- Texas A&M ISSS (via busca): "I-20s cannot be issued for distance or
  online programs"
- Busca sobre as regras de ingles da Graduate School da Texas A&M (as
  paginas grad.tamu.edu dao 404 ao fetch)
- Help center da Liaison/TAMU: taxa de aplicacao de US$114 para
  internacionais (US$89 domesticos)

*** PONTOS CRITICOS ***
1) ELEGIBILIDADE F-1 NAO CONFIRMADA: as aulas sao as tercas e quintas a
   noite, e o aluno pode assistir presencialmente no CityCentre Houston ou
   ao vivo online. A pagina do programa nao fala de alunos internacionais,
   e a Texas A&M nao emite I-20 para programas online/a distancia.
   Registrado em Formato para confirmar.
2) DUOLINGO PROVAVELMENTE NAO ACEITO: o valor anterior "Duolingo (min.
   100)" nao foi confirmado. A Graduate School aceitou o DET so de forma
   temporaria num ciclo antigo; os testes atuais citados sao TOEFL 80,
   IELTS 6.0 e GRE Verbal 146. O texto nao cita mais o DET como aceito, e o
   dashboard passa a mostrar o alerta. Precisa ser confirmado com o
   programa.

OUTRAS MUDANCAS:
- GMAT/GRE: "waivers are available", ou seja, e exigido salvo dispensa
  (antes estava "Opcional").
- Fees Adicionais: o custo de US$70.000 e "all-inclusive" (tuition, fees,
  software, livros e orientacao), cobrado como US$17.500 por semestre em 4
  semestres, igual para residente e nao-residente.
- Prazos no formato de rounds (regra do CLAUDE.md), ciclo Fall 2027.
- Taxa de aplicacao: US$114 (internacionais).
- Documentos: participar de uma information session e obrigatorio.
Uso: python scripts/etl_enrich_texas_am_round3.py (a partir da raiz).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {'Nome da Universidade': 'Texas A&M', 'Estado': 'TX', 'Cidade': 'Houston'}

UPDATES = {'Formato': 'Aulas às terças e quintas à noite, presenciais no CityCentre Houston ou ao vivo '
            'online (cohort, 4 semestres). VERIFICAR elegibilidade F-1: a página do programa não '
            'fala de internacionais e a Texas A&M não emite I-20 para programas online/a distância',
 'Fees Adicionais USD': 'Incluídas no total de US$70.000 (tuition, fees, software, livros e '
                        'orientação), cobrado como US$17.500 por semestre em 4 semestres',
 'GMAT/GRE Obrigatório': "Obrigatório salvo dispensa -- 'GRE/GMAT waivers are available'; "
                         'conversar com a equipe do programa sobre o processo de waiver',
 'Testes Inglês Aceitos': 'TOEFL 80 / IELTS 6.0 / GRE Verbal 146 (mínimos da Graduate School da '
                          'Texas A&M). O DET só foi aceito de forma temporária num ciclo antigo e '
                          'não aparece entre os testes atuais -- confirmar com o programa',
 'Deadline Aplicação': 'Round 1: 3 dez 2026; Round 2: 1 abr 2027 (consideração para bolsa); Round '
                       '3: 25 jun 2027 (final) -- ciclo Fall 2027; aplicações abrem em 1 set 2026',
 'Taxa Aplicação USD': 114,
 'Rounds Aplicação': '3 rounds (3 dez, 1 abr com consideração para bolsa, 25 jun); taxa de US$114 '
                     'para internacionais (US$89 domésticos)',
 'Documentos Necessários': 'Participar de uma information session (obrigatório); GPA; resume; '
                           'essay; transcript não oficial; comprovação dos pré-requisitos; '
                           'GRE/GMAT (salvo dispensa); teste de inglês; entrevista avaliativa'}


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
