# -*- coding: utf-8 -*-
"""
Data enrichment (round 2): U of Utah (David Eccles School of Business), Salt
Lake City, UT -- MS in Business Analytics. Rodada anterior:
etl_enrich_utah.py.

Fontes (set/2026). As paginas eccles.utah.edu vem truncadas no fetch; os
dados vieram de snippets de busca dessas paginas e de PDFs oficiais:
- fbs.admin.utah.edu/download/income/Graduate/BusSMP.pdf (tabela 2025-26
  dos Specialized Masters de business: 10 creditos = US$13.000 com TODAS as
  fees, ou seja, US$1.300/credito inclusivo)
- eccles.utah.edu/.../master-of-science-in-business-analytics (academics,
  finances, admissions, careers/employment-success), via snippets
- isss.utah.edu (regra F-1: no maximo 3 creditos online contam para a
  matricula full-time)

*** PONTO CRITICO -- FORMATO: as disciplinas sao a noite, uma vez por
semana; muitas existem presencial e online, e algumas SO online. O perfil
da turma mostra 80% ja empregados no inicio do programa (turma
profissional). Para F-1, a matricula precisa ser majoritariamente
presencial, entao isso fica registrado para confirmar com MSBA@utah.edu. ***

ALTERACOES:
- Duracao 16 -> 12: full-time conclui em 3 semestres (12 meses).
- Fees: ja incluidas nos US$1.300/credito (confirmado na tabela oficial).
- Caminho Tuition Residente: tarifa unica para residentes e
  nao-residentes.
- Perfil/resultados (pagina oficial, via snippet): turma de 51 alunos, GPA
  medio 3,53, 95% de colocacao, salario medio US$96.959. ATENCAO: 80% ja
  trabalhavam ao entrar, entao os numeros refletem em parte profissionais
  experientes.
- Bolsas: admissions scholarships de US$2.500 a US$18.000, conforme
  residencia e desempenho; exige aplicar ate o scholarship deadline.
- TA/RA: o dado anterior (stipend de ~US$28,5k) nao foi confirmado para o
  MSBA (parece ser de assistantships de pesquisa ou PhD) -- nota
  acrescentada.
Uso: python scripts/etl_enrich_utah_round2.py (a partir da raiz).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {'Nome da Universidade': 'U of Utah', 'Estado': 'UT', 'Cidade': 'Salt Lake City'}

UPDATES = {'Formato': 'Presencial/online: aulas à noite, 1x por semana por disciplina; muitas disciplinas '
            'têm versão presencial e online, e algumas são só online; turma majoritariamente '
            'profissional (80% já empregados no início). VERIFICAR viabilidade F-1: só 3 créditos '
            'online por semestre contam para a matrícula full-time',
 'Duração (meses)': 12,
 'Tamanho Turma': 51,
 'Destaques Currículo': 'Currículo lockstep de 30 créditos (todos fazem as mesmas disciplinas na '
                        'mesma ordem); full-time em 3 semestres ou part-time; #4 no TFE Times 2025 '
                        '(melhor programa público); fast track para ex-alunos da Eccles com GPA ≥ '
                        '3.4',
 'Fees Adicionais USD': 'Incluídas nos US$1.300/crédito (tabela oficial 2025-26: 10 créditos = '
                        'US$13.000 com todas as fees)',
 'Tipos de Bolsa': 'Admissions scholarships de US$2.500 a US$18.000, conforme residência e '
                   'desempenho acadêmico; é preciso aplicar até o scholarship deadline',
 'Quão Realista (Intl)': 'Moderada -- há bolsas de admissão com faixa ampla (US$2,5-18k), mas o '
                         'valor depende de residência e desempenho, e exige aplicar até o prazo de '
                         'bolsa',
 'TA/RA Disponível': 'Verificar -- o dado anterior (stipend de ~US$28,5k + tuition) parece ser de '
                     'assistantships de pesquisa ou PhD e não foi confirmado para o MSBA',
 'Caminho Tuition Residente': 'Não aplicável -- tarifa única de US$1.300/crédito para residentes e '
                              'não-residentes',
 'Taxa Empregabilidade %': 95,
 'Salário Médio Pós USD': 96959,
 'Principais Empregadores': 'University of Utah Health, Swire Coca-Cola, Zions Bancorporation, '
                            'Layton Construction, DigiCert, CHG Healthcare, bioMérieux, Goldman '
                            'Sachs, qBotica (lista oficial; formados empregados em UT, NY, CA, TX, '
                            'WA, DC, IL e Austrália)',
 'GPA Médio Admitidos': 3.53,
 'Testes Inglês Aceitos': 'TOEFL iBT 90 / IELTS Academic 7.0 / Duolingo 110 (mínimo da '
                          'universidade; agregadores citam 115 para o MSBA -- confirmar)'}


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
