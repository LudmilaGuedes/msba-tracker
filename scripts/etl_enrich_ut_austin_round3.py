# -*- coding: utf-8 -*-
"""
Data enrichment (rodada 3, 20/out/2026): atualiza APENAS a linha da 'UT Austin'
(McCombs School of Business, Austin, TX) no msba_programs_database.xlsx.

Foco: MS Business Analytics and AI (MSBAi), presencial em Austin, STEM.
A rodada 2 (etl_enrich_ut_austin_round2.py, out/2026) ja tinha reconferido a
linha inteira; esta rodada refaz a pesquisa nas paginas oficiais (que mudaram
de URL: agora em /graduate/specialized-masters/...) e corrige o que mudou.

CONFIRMADO SEM ALTERACAO:
- Tuition: US$55.600 residente / US$59.700 nao-residente (= internacional),
  por semestre 10.600/21.500/21.500 e 11.900/22.900/22.900, mais pre-matricula
  de US$2.000; sem tarifa por credito publicada (Custo/Credito fica Verificar).
- Seguro e ISSS fora da tuition; bolsas de merito limitadas, automaticas, sem
  valores publicados e sem dizer se internacionais sao elegiveis; sem TA/RA;
  emprestimos para internacionais (alguns sem co-signer americano).
- Rounds 2026-27: R1 13/out/2026, R2 12/jan/2027 (final para internacionais),
  R3 30/mar/2027, R4 4/mai/2027; taxa US$90 intl / US$65 domestico.
- GMAT/GRE opcional (competitivo >= 60o percentil); TOEFL 100 / IELTS 7.0
  esperados pelo programa; minimos da Graduate School TOEFL 79 / IELTS 6.5 /
  DET 115; mestrado nao isenta.
- 3 ensaios (2 obrigatorios + 1 opcional, 500 palavras), 2 cartas (>= 1
  academica), employment history, transcripts, Kira (30 min) + entrevista
  com selecionados.
- Class of 2027: 64 alunos, GPA medio 3.55, GMAT medio 655.
- Empregabilidade: relatorio Class of 2025 (publicado mar/2026) continua o
  mais recente -- 81,4% aceitaram oferta ate 6 meses apos a formatura
  (internacionais: 83,3%); salario base medio US$104.798 (mediana 105.000;
  internacionais: media 117.734 / mediana 120.000).

CORRIGIDO:
- Ranking QS: 12 -> 15. QS Business Masters Rankings 2027 (Business
  Analytics), divulgado em 17/set/2026: #15 no mundo, #9 nos EUA.

MANTIDO COMO Verificar: Custo/Credito (Intl e Residente); Taxa de Aceitacao %
(nao divulgada).

Fontes (out/2026):
- mccombs.utexas.edu/graduate/specialized-masters/ms-business-analytics-and-artificial-intelligence/on-campus/admissions/tuition-financial-aid/
- .../on-campus/admissions/application-process/
- .../on-campus/admissions/faq
- mccombs.utexas.edu/wp-content/uploads/2026/03/25-MSBA-Employment-Report.pdf
- graduate.utexas.edu/admissions/apply/international
- news.mccombs.utexas.edu/rankings/mccombs-m-s-programs-take-high-rankings/ (QS 2027)

Uso: python scripts/etl_enrich_ut_austin_round3.py (a partir da raiz).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = ("UT Austin", "TX", "Austin")

UPDATES = {
    "Ranking QS": 15,
}


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

    target_rows = [
        row[0].row
        for row in ws.iter_rows(min_row=2)
        if (
            row[col["Nome da Universidade"] - 1].value,
            row[col["Estado"] - 1].value,
            row[col["Cidade"] - 1].value,
        ) == MATCH
    ]
    if len(target_rows) != 1:
        raise ValueError(f"Esperava 1 linha para {MATCH}, achei {len(target_rows)}. Abortando.")
    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (UT Austin, Austin, TX)")

    no_fill = PatternFill(fill_type=None)
    changes = []
    for header, new_value in UPDATES.items():
        cell = ws.cell(row=row_idx, column=col[header])
        changes.append((header, cell.value, new_value))
        cell.value = new_value
        cell.fill = no_fill

    wb.save(DB_PATH)

    print("\nCampos atualizados (Antes -> Depois):")
    for header, old, new in changes:
        print(f"  - {header}:")
        print(f"      Antes:  {old!r}")
        print(f"      Depois: {new!r}")


if __name__ == "__main__":
    main()
