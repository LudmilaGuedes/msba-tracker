# -*- coding: utf-8 -*-
"""
Data enrichment (tuition): Texas A&M University-Texarkana, Texarkana, TX --
MS in Business and Data Analytics (MSBDA).

Fonte: tamut.edu/business-office/files/tuition-fee-schedules-for-va.pdf,
"Guaranteed Tuition & Mandatory Fee Schedule for Two-Year Graduate Cohort --
Entering in Academic Year 2025-2026". E a tabela oficial mais recente
encontrada; a de 2026-27 nao esta publicada em texto e a calculadora do
site nao carrega os valores.

Com 9 creditos por semestre (full-time), a tuition + mandatory fees e:
  residente     US$3.017,99 -> 3.017,99 / 9 = US$335,33 por credito
  nao-residente US$7.701,07 -> 7.701,07 / 9 = US$855,67 por credito
Os valores ja incluem statutory + designated tuition e as fees
obrigatorias. Programa de 30 creditos (estimativa proporcional):
  residente     30 x 335,33 = US$10.060
  nao-residente 30 x 855,67 = US$25.670
A differential tuition de pos (US$46,43/credito) vale so para as
concentracoes de Energy Leadership, IT, SCM e Accounting; o MSBDA nao
aparece na lista.

Com a International Eagle Scholarship (isencao da tarifa de fora do estado
+ US$2.000/ano), o internacional pagaria a coluna de residente (~US$10k no
curso), menos a bolsa.
Uso: python scripts/etl_enrich_texas_am_texarkana_tuition.py (a partir da raiz).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {'Nome da Universidade': 'Texas A&M University-Texarkana (A&M-Texarkana)',
 'Estado': 'TX',
 'Cidade': 'Texarkana'}

UPDATES = {'Custo/Crédito (Intl) USD': 855.67,
 'Custo/Crédito (Residente) USD': 335.33,
 'Tuition Total (Intl) USD': 25670,
 'Tuition Total (Residente) USD': 10060,
 'Fees Adicionais USD': 'Já incluídas nos valores de tuition (tabela oficial de tuition + '
                        'mandatory fees da turma de pós de 2025-26, com valor garantido por 2 '
                        'anos)',
 'Caminho Tuition Residente': 'Sim -- a International Eagle Scholarship inclui isenção da tuition '
                              'de fora do estado (primeiros 100) e o internacional passa a pagar '
                              'como residente (~US$335 vs ~US$856 por crédito com fees; ~US$10k vs '
                              '~US$25,7k no curso); há também o Competitive Scholarship Waiver do '
                              'Texas'}


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
