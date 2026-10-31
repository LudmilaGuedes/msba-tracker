# -*- coding: utf-8 -*-
"""
Data enrichment (round 5): SMU (Cox School of Business), Dallas, TX -- MS in
Business Analytics. Rodadas anteriores: etl_enrich_smu.py ate round4 e
etl_fix_tuition_audit_usc_smu.py.

Fontes (set/2026):
- smu.edu/cox/.../ms-business-analytics/admissions (full-time 2026-27:
  tuition US$50.000 + university/program fees US$10.872 = US$61.502;
  custo total estimado do programa US$97.377)
- smu.edu/cox/.../ms-business-analytics (presencial, 9-21 meses, 33
  creditos, trilhas)
- smu.edu/cox/ms-degrees/faqs (TOEFL/IELTS/Duolingo aceitos e opcionais;
  sem taxa de aplicacao; entrevista por convite; o MSBA inclui um estagio
  no Spring, com CPT)
- smu.edu/-/media/site/cox/.../msba-employment-report_2025.pdf (turma de
  2025: 97 alunos; 82% com outcome e 80% com emprego full-time em 6 meses;
  media US$79,8k; 27% com bonus de assinatura, media US$12,3k; 81% ficaram
  no Southwest)

ALTERACOES:
- Fees Adicionais: 2.772 -> 10.872 (university + program fees de 2026-27).
- Custo/credito: 50.000 / 33 = 1.515,15 (o valor anterior de 1.570 nao
  batia com o total).
- Custo Vida Total: 35.875 = 97.377 (custo total estimado oficial) -
  61.502 (tuition + fees). Pode incluir seguro saude e livros; a pagina nao
  detalha.
- Principais Empregadores: amostra do relatorio de 2025. ATENCAO: a lista
  cobre TODOS os MS da Cox, nao so o MSBA.
- Destaques, Tipos de Bolsa, Rounds, Entrevista e Documentos: detalhados
  com as fontes acima.
Uso: python scripts/etl_enrich_smu_round5.py (a partir da raiz).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {'Nome da Universidade': 'SMU', 'Estado': 'TX', 'Cidade': 'Dallas'}

UPDATES = {'Destaques Currículo': '33 créditos; trilhas de business process analytics ou customer analytics; '
                        'Python, R, SQL, Tableau e Excel; estágio no Spring incluído no currículo '
                        '(com CPT para internacionais); 9 a 21 meses (full-time ou part-time)',
 'Custo/Crédito (Intl) USD': 1515.15,
 'Custo/Crédito (Residente) USD': 1515.15,
 'Fees Adicionais USD': 10872,
 'Custo Vida Total (programa) USD': 35875,
 'Tipos de Bolsa': 'Bolsas de mérito (consideração automática, sem aplicação separada; baseadas no '
                   'desempenho acadêmico e no GMAT/GRE; os Rounds 1 e 2 têm prioridade para bolsas '
                   'e fellowships); bolsa garantida no 4+1 para graduados da própria Cox',
 'Principais Empregadores': 'Amostra de empresas que contrataram formados de TODOS os MS da Cox '
                            '(turma de 2025, não só o MSBA): Amazon, American Airlines, Bank of '
                            'America, Capital One, Deloitte, EY, KPMG, PwC, Goldman Sachs, '
                            'JPMorgan, Fidelity, Samsung, Siemens, Toyota, Southwest Airlines, '
                            'CBRE, JLL, KKR; 81% ficaram no Southwest (Texas)',
 'Rounds Aplicação': '5 rounds (Rounds 1 e 2 prioritários para bolsas e fellowships); rolling '
                     'depois de 25/mai; sem taxa de aplicação',
 'Documentos Necessários': 'Resume; essays; transcripts não oficiais (oficiais só se for '
                           'admitido); 2 referências (só contato, sem carta); GMAT/GRE e teste de '
                           'inglês opcionais; entrevista por convite',
 'Entrevista': 'Por convite do comitê de admissão (via Zoom ou presencial no campus)'}


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
