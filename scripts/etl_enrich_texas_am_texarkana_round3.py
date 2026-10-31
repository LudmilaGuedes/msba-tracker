# -*- coding: utf-8 -*-
"""
Data enrichment (round 3): Texas A&M University-Texarkana, Texarkana, TX --
Master of Science in Business and Data Analytics (MSBDA), STEM, presencial.
Rodadas anteriores: etl_enrich_texas_am_texarkana.py (e a adicao inicial).

Fontes (set/2026):
- tamut.edu/academic-programs/business-and-data-analytics/
- tamut.edu/enrollment-services/scholarships/
- tamut.edu/admissions/graduate/
- catalog.tamut.edu (graduate e international admissions)

ACHADOS:
1) *** BOLSA PARA INTERNACIONAIS: "International Eagle Scholarship" --
   "$2000 Yearly plus an out of state tuition waiver" para estudantes
   internacionais de pos-graduacao com GPA >= 3.0 e 9 creditos por
   semestre; renovavel por 2 anos; o waiver vale para os primeiros 100
   contemplados. Tambem ha a "Eagle New Grad Scholarship" (US$1.500/ano com
   9h por semestre, GPA >= 3.0) e o formulario de Competitive Scholarship
   Waiver para quem paga tarifa de nao-residente. As ofertas chegam por
   e-mail depois da admissao. ***
2) INICIO: internacionais so comecam no Fall ou no Spring (nunca no
   Summer).
3) PRAZOS: admissao rolling, que abre em 1/abr (Fall) e 1/out (Spring); os
   documentos de internacionais devem chegar 45 dias antes do 1o dia de
   aula (~10/jul para o Fall 2026, data calculada na rodada anterior).
4) DOCUMENTOS: acrescentada a avaliacao de credenciais estrangeiras e a
   comprovacao financeira de 1 ano.
5) GMAT/GRE: a Graduate School diz que "some programs" exigem GRE/GMAT/MAT;
   nada publicado para o MSBDA, entao fica "Verificar".
6) TUITION: continua "Verificar". A calculadora oficial nao mostra os
   valores (template sem dados). Agregadores citam ~US$299/credito
   (residente) e ~US$834 (nao-residente), sem ano confirmado. Registrado
   so como texto em Caminho Tuition Residente.
7) TAXA DE APLICACAO: mantida em US$50 (confirmada na rodada anterior para
   a pos). Os US$30 que aparecem em busca sao da graduacao.

O programa comecou no Fall 2026, entao ainda nao ha formados nem dados de
empregabilidade ou salario.
Uso: python scripts/etl_enrich_texas_am_texarkana_round3.py (a partir da raiz).
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

UPDATES = {'Início': 'Fall ou Spring (internacionais não podem começar no Summer); programa lançado no Fall '
           '2026',
 'Bolsas Disponíveis': 'Sim',
 'Tipos de Bolsa': 'International Eagle Scholarship: US$2.000/ano + isenção da tuition de fora do '
                   'estado (paga como residente), para internacionais de pós com GPA ≥ 3.0 e 9 '
                   'créditos por semestre; renovável por 2 anos; isenção limitada aos 100 '
                   'primeiros. Eagle New Grad Scholarship: US$1.500/ano (9h/semestre, GPA ≥ 3.0). '
                   'A oferta chega por e-mail após a admissão',
 'Quão Realista (Intl)': 'Alta -- a International Eagle Scholarship exige só GPA ≥ 3.0 e 9 '
                         'créditos por semestre; a isenção da tuition de fora do estado é limitada '
                         'aos 100 primeiros, então vale aplicar cedo',
 'Caminho Tuition Residente': 'Sim -- a International Eagle Scholarship inclui isenção da tuition '
                              'de fora do estado (primeiros 100); há também o Competitive '
                              'Scholarship Waiver do Texas (formulário na página de bolsas). '
                              'Tarifas citadas por terceiros: ~US$299/crédito residente vs ~US$834 '
                              'não-residente (ano não confirmado)',
 'GMAT/GRE Obrigatório': "Verificar -- a Graduate School diz que 'some programs' exigem "
                         'GRE/GMAT/MAT; nada publicado especificamente para o MSBDA',
 'Deadline Aplicação': 'Internacional (documentos): ~10/jul para o Fall (45 dias antes do 1º dia '
                       'de aula; data calculada); aplicações rolling abrem em 1/abr (Fall) e 1/out '
                       '(Spring); prazo próprio do programa não publicado',
 'Rounds Aplicação': 'Sem rounds; admissão rolling (Fall abre em 1/abr, Spring em 1/out); '
                     'internacionais só começam no Fall ou no Spring',
 'Documentos Necessários': 'Resume/CV (1-2 páginas); carta de intenção (2-4 páginas); cartas de '
                           'recomendação; transcripts oficiais de todas as instituições; avaliação '
                           'de credenciais estrangeiras (Foreign Credential Evaluation); '
                           "comprovação financeira ou Sponsor's Statement para 1 ano; teste de "
                           'inglês'}


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
