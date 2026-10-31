# -*- coding: utf-8 -*-
"""
Data enrichment (round 2): NC State (Institute for Advanced Analytics),
Raleigh, NC -- MS in Analytics (MSA), presencial, STEM, 10 meses.

Fontes (set/2026):
- analytics.ncsu.edu/?page_id=4202 (tuition do ciclo Summer II '26/Fall
  '26/Spring '27: nao-residente 59.679,06, residente 30.642,68; saude e
  custo de vida por fora). Re-confirmado, sem mudanca.
- analytics.ncsu.edu/admissions/deadlines-and-dates/ (turma de 2028)
- analytics.ncsu.edu/?page_id=13892 (internacionais: taxa de US$95, DUAS
  recomendacoes com a terceira opcional, personal statement, resume,
  transcripts)
- catalog.ncsu.edu/graduate/graduate-handbook/international-student-admissions/
  (minimos de ingles com sub-notas)
- analytics.ncsu.edu/?page_id=1948 (FAQ: sem RA/TA para alunos do MSA;
  auxilio parcial baseado em necessidade; GRE/GMAT nao considerados;
  entrevista + triagem tecnica)
- analytics.ncsu.edu/2026/05/20/celebrating-our-msa-26-grads/ (turma de
  2026: 61 formados; 70,5% com oferta aceita ate 18/mai; mediana de US$105k)
- Relatorio de emprego MSA2025, via snippet: 94 formados; 67% empregados
  na formatura e 90% em 90 dias; media US$92.295; 167 empregadores
  entrevistaram e 49 contrataram.

ACHADOS:
1) PRAZOS ATUALIZADOS para a turma de 2028 (inicio em 24/jun/2027):
   prioritario 1/nov/2026; final internacional 15/fev/2027; final
   domestico 1/abr/2027. O valor antigo "Dec 1 / Feb 15" era de um ciclo
   anterior.
2) DOCUMENTOS CORRIGIDOS: sao 2 recomendacoes (a terceira e opcional), e
   nao 3.
3) BOLSAS: o FAQ confirma que alunos do MSA NAO sao elegiveis a RA/TA. O
   Instituto oferece auxilio PARCIAL baseado em necessidade. Os campos de
   TA/RA passaram de "Limitado/busca ativa" para "Nao".
4) ENTREVISTA detalhada: ~30 min com 3-4 pessoas, mais uma triagem tecnica
   remota de 30 min (12 questoes: coding em Python/R/SAS e estatistica).
5) TAMANHO DA TURMA: 94 era a turma de 2025; a de 2026 formou 61.
6) Empregabilidade (90%) e salario (US$92.295) sao da turma de 2025 em 90
   dias e foram mantidos. A turma de 2026 ainda nao tem o numero de 90
   dias; na formatura estava em 70,5%, com mediana de US$105k.

Uso: python scripts/etl_enrich_nc_state_round2.py (a partir da raiz).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "NC State",
    "Estado": "NC",
    "Cidade": "Raleigh",
}

UPDATES = {
    "Início": "Final de junho (24/jun/2027 para a turma de 2028)",
    "Tamanho Turma": 61,
    "Tipos de Bolsa": (
        "Auxílio financeiro PARCIAL do Instituto, baseado em necessidade; fellowships "
        "pontuais da Graduate School (ex.: Diversity Enhancement Grants, Warren-Wolfpack "
        "Club Graduate Fellowships). Sem assistantships: alunos do MSA não são elegíveis "
        "a RA/TA"
    ),
    "Quão Realista (Intl)": (
        "Baixa -- sem assistantships e sem bolsa de mérito estruturada; a maioria se "
        "financia com recursos próprios ou empréstimo; existe só auxílio parcial por "
        "necessidade, e muitas fellowships exigem cidadania americana"
    ),
    "TA/RA Disponível": "Não -- o FAQ oficial diz que alunos do MSA não são elegíveis a RA/TA",
    "Processo TA/RA": "Não se aplica (MSA não é elegível a RA/TA)",
    "Graduate Assistantships": "Não disponível para alunos do MSA (FAQ oficial)",
    "Principais Empregadores": (
        "167 empregadores entrevistaram e 49 contrataram pelo menos 1 aluno (turma de "
        "2025); os nomes não são listados nas fontes públicas consultadas"
    ),
    "GMAT/GRE Obrigatório": "Não aceito nem considerado -- o programa não avalia GRE/GMAT (não é exigido)",
    "Matérias Pré-Requisito": (
        "Pelo menos 1, idealmente 2 semestres de métodos estatísticos em nível "
        "universitário (incluindo regressão); experiência com programação em pelo menos "
        "1 linguagem -- os pré-requisitos podem ser concluídos durante a aplicação"
    ),
    "Testes Inglês Aceitos": (
        "TOEFL iBT 80 (até 21/jan/2026; ≥18 em cada seção) ou 4 (nova escala; ≥4 em cada "
        "seção) / IELTS 6.5 (≥6.5 em cada banda) / Duolingo 110 (≥110 em Literacy, "
        "Comprehension e Conversation)"
    ),
    "Deadline Aplicação": (
        "Round 1: 1 nov 2026 (prioritário); Round 2: 15 fev 2027 (final internacional); "
        "Round 3: 1 abr 2027 (final doméstico) -- turma de 2028, início em 24/jun/2027"
    ),
    "Rounds Aplicação": (
        "Prioritário (entrevistas e triagem técnica em novembro; decisão em 4/dez; "
        "aceite até 15/jan) e regular (rolling até preencher as vagas; turma fechada até "
        "o início de maio)"
    ),
    "Documentos Necessários": (
        "Transcripts não oficiais de cada instituição; personal statement (1-2 páginas); "
        "resume; 2 cartas de recomendação (3ª opcional); teste de inglês"
    ),
    "Entrevista": (
        "Sim, por convite: ~30 min com 3-4 docentes/staff/ex-alunos, mais uma triagem "
        "técnica remota de 30 min (12 questões: coding em Python, R ou SAS e estatística)"
    ),
}


def main():
    rotate_backup()
    print("Backup rotacionado (política fixa) em: data/msba_programs_database_backup.xlsx")

    wb = openpyxl.load_workbook(DB_PATH)
    ws = wb[SHEET_DATA]

    headers = [cell.value for cell in ws[1]]
    header_to_col = {h: i + 1 for i, h in enumerate(headers)}

    unknown_cols = [h for h in UPDATES if h not in header_to_col]
    if unknown_cols:
        raise ValueError(f"Colunas inexistentes no schema: {unknown_cols}")

    col_uni = header_to_col["Nome da Universidade"]
    col_estado = header_to_col["Estado"]
    col_cidade = header_to_col["Cidade"]

    target_rows = []
    for row in ws.iter_rows(min_row=2):
        if (
            row[col_uni - 1].value == MATCH["Nome da Universidade"]
            and row[col_estado - 1].value == MATCH["Estado"]
            and row[col_cidade - 1].value == MATCH["Cidade"]
        ):
            target_rows.append(row[0].row)

    if len(target_rows) == 0:
        raise ValueError("Nenhuma linha encontrada para NC State (Raleigh, NC). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (NC State, Raleigh, NC)")

    no_fill = PatternFill(fill_type=None)

    before = {}
    after = {}
    for header, new_value in UPDATES.items():
        col_idx = header_to_col[header]
        cell = ws.cell(row=row_idx, column=col_idx)
        before[header] = cell.value
        cell.value = new_value
        cell.fill = no_fill
        after[header] = new_value

    wb.save(DB_PATH)

    print(f"\n{len(UPDATES)} campos atualizados na linha {row_idx} (NC State):\n")
    for header in UPDATES:
        print(f"  - {header}:")
        print(f"      Antes:  {before[header]!r}")
        print(f"      Depois: {after[header]!r}")


if __name__ == "__main__":
    main()
