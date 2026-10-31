# -*- coding: utf-8 -*-
"""
Data enrichment (round 2): atualiza APENAS a linha 'TCU' (Texas Christian
University, Neeley School of Business, Fort Worth, TX) no
msba_programs_database.xlsx.

CONTEXTO: rodada 1 (scripts/etl_enrich_tcu.py) ja confirmou bem tuition
($2.289/credito, identico p/ residente e intl -- universidade PRIVADA --
total $71.663 p/ 30 creditos), Fees Adicionais $3.000, duracao 11 meses,
inicio Summer, GMAT/GRE opcional, pre-requisitos, taxa de aplicacao $100,
documentos basicos. Esta rodada foca nos campos que ainda estavam
"Verificar".

ATUALIZADO com fonte oficial (neeley.tcu.edu):
- Bolsas Disponíveis / Tipos de Bolsa / Quão Realista (Intl): estavam
  "Verificar". Confirmado em neeley.tcu.edu/admissions/graduate/
  scholarships-financial-aid.php: existem Assistantships (posicoes
  remuneradas de apoio a docentes/departamentos), bolsas de merito
  (selecao competitiva), bolsas para veteranos (GI Bill/Yellow Ribbon --
  NAO aplicavel a internacional) e Community Partners Scholarships. SEM
  candidatura separada -- avaliacao automatica junto com o material de
  admissao; candidatos internacionais sao EXPLICITAMENTE considerados
  automaticamente. Nenhum valor especifico do MSBA foi divulgado.
- Salário Médio Pós USD: estava "Verificar". CONFIRMADO oficialmente na
  propria pagina do programa (neeley.tcu.edu/degrees-programs/graduate/
  business-analytics.php): $81.000 ("average total compensation for
  graduates").
- Tamanho Turma: estava "Verificar". CONFIRMADO: 22 estudantes
  matriculados (neeley.tcu.edu/about/facts-figures.php, secao Enrollment).
- Testes Inglês Aceitos: estava so a lista de testes sem notas ("TOEFL/
  IELTS/PTE/Duolingo (obrigatorio para quem nao e de pais de lingua
  oficial inglesa)"). CONFIRMADO com notas minimas exatas
  (neeley.tcu.edu/admissions/graduate/admission-requirements.php): TOEFL
  iBT min. 100 / IELTS min. 7.0 / PTE min. 68 / Duolingo min. 130.
- Entrevista: estava "Verificar". CONFIRMADO: Sim, candidatos selecionados
  sao convidados para entrevista ("by invitation").
- Documentos Necessários: acrescentado "ensaios de aplicacao", confirmados
  oficialmente na mesma pagina de admission-requirements, que nao estava
  mencionado na rodada 1.
- GPA Mínimo: estava "Verificar". A pagina oficial afirma explicitamente
  que NAO existe um GPA minimo unico que garanta admissao (avaliacao
  holistica) -- isso e uma informacao confirmada (ausencia de minimo), nao
  um dado desconhecido, entao foi registrado em vez de deixado "Verificar".

CORRIGIDO -- Deadline Aplicação / Rounds Aplicação: a rodada 1 tinha
registrado apenas "Prioritaria (consideracao p/ assistantship): 15 out;
Final (Summer/Fall): 1 fev". A pagina oficial de admission-requirements.php
(fonte mais detalhada e completa) lista uma estrutura de admissao por
ROUNDS que nao bate com isso: 5 rounds -- 1/out, 15/jan, 1/mar, 1/abr,
1/mai -- mais admissao rolling apos 1/mai. Substituido pela estrutura mais
completa e detalhada; a estrutura antiga pode ter sido uma versao
simplificada ou de um ciclo anterior.

NAO ALTERADO por falta de fonte oficial confiavel (mantido "Verificar"):
- TA/RA Disponível, Processo TA/RA, Graduate Assistantships: a mesma
  pagina de financial aid confirma que "assistantships" existem como
  categoria de ajuda financeira do Neeley em geral, mas NAO ha detalhe de
  cobertura (paga tuition? so stipend?) nem processo especifico para o
  MSBA -- informacao generica demais para preencher com confianca.
- Ranking US News / Ranking QS: nenhum ranking especifico do MS Business
  Analytics encontrado (a pagina de Facts & Figures do Neeley so cobre
  BBA/MBA/Accounting/Entrepreneurship/Supply Chain).
- Taxa Empregabilidade %: nenhum placement rate real do MSBA encontrado.
  ATENCAO: a pagina oficial lista "86% Entered immediately after
  undergrad" (Turma 2023) -- isso e composicao da turma (% que entrou
  direto da graduacao), NAO taxa de empregabilidade pos-formatura -- nao
  foi usado para nao confundir os dois conceitos.
- Principais Empregadores: nenhuma lista nomeada especifica do MSBA
  encontrada (so descricao generica de funcoes: data analyst, business
  analyst, consultant em finance/consulting/healthcare/tech).
- GPA Médio Admitidos, GMAT Médio, Taxa Aceitação %: nao encontrados.
- Fees Adicionais ($3.000): nao reconfirmado diretamente nesta rodada
  (achado indireto de "books and materials estimated $2.500-$3.000" e
  consistente mas nao e fonte oficial direta para este campo) -- mantido
  sem alteracao.

Fontes principais (pesquisadas via WebSearch/WebFetch, set/2026):
- https://neeley.tcu.edu/admissions/graduate/scholarships-financial-aid.php
  (oficial -- tipos de bolsa/assistantship, avaliacao automatica p/ internacionais)
- https://neeley.tcu.edu/degrees-programs/graduate/business-analytics.php
  (oficial -- pagina do programa: salario medio $81k, composicao de turma)
- https://neeley.tcu.edu/admissions/graduate/admission-requirements.php
  (oficial -- testes de ingles com notas, GMAT/GRE opcional, documentos,
  5 rounds de aplicacao, entrevista por convite, ausencia de GPA minimo)
- https://neeley.tcu.edu/about/facts-figures.php
  (oficial -- tamanho de turma: 22 estudantes matriculados)
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "TCU",
    "Estado": "TX",
    "Cidade": "Fort Worth",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Bolsas Disponíveis": "Sim",
    "Tipos de Bolsa": (
        "Assistantships (posições remuneradas de apoio a docentes/"
        "departamentos), bolsas de mérito (seleção competitiva), "
        "Community Partners Scholarships, e bolsas para veteranos "
        "(GI Bill/Yellow Ribbon -- não aplicável a internacional). "
        "Avaliação automática junto com o material de admissão, sem "
        "candidatura separada. Nenhum valor específico do MSBA divulgado "
        "publicamente"
    ),
    "Quão Realista (Intl)": (
        "Moderada -- candidatos internacionais são explicitamente "
        "considerados automaticamente para bolsas junto com o material de "
        "admissão (sem etapa extra), mas nenhum valor ou taxa de sucesso "
        "específica do MSBA foi divulgado publicamente"
    ),
    "GPA Mínimo": "Não há mínimo único divulgado -- Neeley afirma explicitamente que a admissão é holística, sem um GPA de corte",
    "Tamanho Turma": 22,
    "Testes Inglês Aceitos": "TOEFL iBT mín. 100 / IELTS mín. 7.0 / PTE mín. 68 / Duolingo mín. 130",
    "Deadline Aplicação": "5 rounds: 1/out, 15/jan, 1/mar, 1/abr, 1/mai; admissão rolling após 1/mai",
    "Rounds Aplicação": "5 rounds (1/out a 1/mai) + rolling após 1/mai -- corrige/detalha estrutura anterior de 2 deadlines",
    "Salário Médio Pós USD": 81000,
    "Entrevista": "Sim -- candidatos selecionados são convidados por convite (by invitation)",
    "Documentos Necessários": "Transcripts oficiais, resume, 3 referências (contato, não carta de recomendação formal), ensaios de aplicação",
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
        raise ValueError("Nenhuma linha encontrada para TCU (Fort Worth, TX). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (TCU, Fort Worth, TX)")

    no_fill = PatternFill(fill_type=None)

    before = {}
    after = {}
    for header, new_value in UPDATES.items():
        col_idx = header_to_col[header]
        cell = ws.cell(row=row_idx, column=col_idx)
        before[header] = cell.value
        cell.value = new_value
        cell.fill = no_fill  # nenhum novo valor aqui e "Verificar"
        after[header] = new_value

    wb.save(DB_PATH)

    print("\nCampos atualizados (Como Estava -> Como Ficou):")
    for header in UPDATES:
        print(f"  - {header}:")
        print(f"      Antes: {before[header]!r}")
        print(f"      Depois: {after[header]!r}")


if __name__ == "__main__":
    main()
