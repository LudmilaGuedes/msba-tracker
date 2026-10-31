# -*- coding: utf-8 -*-
"""
Data enrichment (round 3): auditoria + pesquisa aprofundada focada em UMA
unica linha -- University of Washington (Foster School of Business),
Seattle, WA -- MS in Business Analytics (MSBA), presencial, STEM.

Pedido da usuaria: atuar como engenheiro de dados fazendo Data Enrichment
de uma unica instituicao, cobrindo tuition, bolsas, taxa de aplicacao,
requisitos completos de aplicacao, empregabilidade/salario/empregadores e
duracao/formato. Substituir dados incorretos, preencher lacunas, manter
"Verificar" para o que nao for encontrado.

Fontes oficiais (.edu, todas verificadas por WebFetch direto em set/2026):
- Tuition: foster.uw.edu/academics/degree-programs/master-science-business-analytics/program-fees-financing-msba/
- Curriculo/creditos/formato: foster.uw.edu/academics/degree-programs/master-science-business-analytics/curriculum-msba
- Admissoes (GPA, GMAT/GRE, testes de ingles, documentos, deadlines):
  foster.uw.edu/academics/degree-programs/master-science-business-analytics/admissions-msba/
- Class Profile oficial (tamanho de turma, GPA/GMAT/GRE medios):
  foster.uw.edu/wp-content/uploads/2025/10/Class-of-2026-MSBA-Class-Profile.pdf
- Resultados de carreira (taxa de oferta, salario, bonus, empregadores):
  foster.uw.edu/academics/degree-programs/master-science-business-analytics/career-services-msba/

ACHADOS PRINCIPAIS:

1) TUITION ATUALIZADA (ano letivo 2026-27, valor mudou desde a ultima
   pesquisa): Residente USD 49.763 (era 48.172); Internacional/nao-
   residente USD 62.955 (era 61.364). Custo/credito recalculado com o
   total de creditos correto (ver item 2): residente ~USD 1.131/credito,
   internacional ~USD 1.431/credito (o programa fatura por trimestre, nao
   estritamente por credito -- valor por credito e uma aproximacao,
   mesma metodologia ja usada no registro anterior).

2) TOTAL DE CREDITOS CORRIGIDO: pagina oficial do curriculo diz 44
   creditos (estava registrado como 46 -- numero incorreto/desatualizado).

3) TAMANHO DE TURMA PREENCHIDO (estava "Verificar"): 99 alunos, Class of
   2026 (fonte: Class Profile oficial em PDF).

4) GMAT MEDIO CORRIGIDO COM RESSALVA IMPORTANTE: Class Profile oficial
   mostra "615 Average GMAT FE" -- isso e GMAT FOCUS EDITION (escala
   205-805), NAO o GMAT classico (escala 200-800) que o valor anterior
   (655) parecia representar. Sao escalas diferentes, 615 no Focus
   Edition NAO significa uma nota pior que 655 no GMAT classico -- sao
   numeros no mesmo padrao numerico mas escalas distintas. Tambem
   confirmado: 70% dos alunos da Class of 2026 NAO enviaram GMAT/GRE
   (teste e opcional, dado ja registrado em "GMAT/GRE Obrigatorio").

5) *** CORRECAO IMPORTANTE DE DADO PROVAVELMENTE ERRADO ***
   "Salario Medio Pos USD" estava registrado como 140.354 e "Taxa
   Empregabilidade %" como 91. Pesquisa direta na pagina oficial de
   carreira do MSBA (career-services-msba) mostra, para a turma mais
   recente com dado publicado (Class of 2024): salario medio inicial
   USD 106.000 + bonus de contratacao medio USD 13.000; 92% dos
   formandos receberam oferta de emprego em ate 6 meses. O numero
   140.354 NAO aparece em nenhuma pagina oficial do MSBA encontrada
   nesta pesquisa -- ele e quase identico ao salario medio do MBA
   Full-Time Class of 2025 da mesma escola (USD 144.411, fonte: Foster
   MBA employment report), o que sugere fortemente que o dado anterior
   veio de uma contaminacao MBA->MSBA em pesquisa anterior (um MSBA de
   1 ano tipicamente nao supera o salario de um MBA completo). Corrigido
   para o dado oficial e diretamente fonteado do MSBA: 106.000 (salario
   base; o bonus de 13.000 nao tem campo proprio no schema, mencionado
   aqui e no resumo ao usuario). Taxa de empregabilidade corrigida para
   92 (metrica oficial: % com oferta em ate 6 meses -- metrica um pouco
   diferente de "empregado em 3 meses", mas e o dado oficial mais
   granular e mais recente encontrado).

6) EMPREGADORES ATUALIZADOS: pagina oficial de carreira lista Microsoft,
   T-Mobile, Meta, Amazon, Google como empregadores de destaque dos
   formados do MSBA (substitui a lista anterior "Brooks Running,
   T-Mobile, Google", que nao foi confirmada na fonte oficial atual).

7) DEADLINES DE APLICACAO com anos confirmados: Round 1 = 20/out/2026
   (23h59 PDT); Round 2 = 5/jan/2027 (23h59 PDT) -- confirma o padrao
   ja registrado ("20 out; 5 jan"), so adiciona os anos.

8) DOCUMENTOS NECESSARIOS enriquecidos com detalhe da pagina oficial de
   admissoes: Statement of Purpose obrigatorio + personal statement
   opcional (ate 300 palavras cada), self-review de qualificacoes, alem
   dos itens ja registrados (resume, 2 cartas, entrevista em video).

9) GPA MINIMO enriquecido com o detalhe oficial "ultimos 90 creditos
   trimestre / 60 creditos semestre" (mantendo o texto ja existente
   sobre admissao abaixo do minimo com qualificacoes fortes).

NAO ALTERADO (ja confirmado correto ou fora do escopo pedido):
- GPA Medio Admitidos (3.5) -- bate exatamente com o Class Profile
  oficial, nao mudou.
- Testes de Ingles Aceitos (TOEFL 92/4.5, IELTS 7.0, Duolingo 120-125)
  -- confirmado identico na pagina oficial de admissoes.
- Bolsas Disponiveis / Tipos de Bolsa / Quao Realista (Intl) -- pagina
  oficial de tuition/financing confirma exatamente o que ja estava
  registrado ("pool limitado, considerado automaticamente na
  candidatura"); nenhuma bolsa por merito/assistantship/fellowship
  adicional encontrada.
- Taxa de Aplicacao USD (90) -- confirmado, sem isencao divulgada em
  nenhuma pagina oficial encontrada.
- Formato e Duracao (12 meses, presencial evening/weekend, full-time)
  -- confirmados, sem mudanca.
- Ranking US News e Taxa de Aceitacao % -- fora do escopo pedido pela
  usuaria nesta rodada; permanecem "Verificar" (nao encontrados com
  confianca em pesquisas anteriores nem nesta).

Uso: python etl_enrich_university_of_washington_round3.py (executar a
partir da raiz do repositorio).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "University of Washington",
    "Estado": "WA",
    "Cidade": "Seattle",
}

UPDATES = {
    "Total Créditos": 44,
    "Custo/Crédito (Intl) USD": 1431,
    "Custo/Crédito (Residente) USD": 1131,
    "Tuition Total (Intl) USD": 62955,
    "Tuition Total (Residente) USD": 49763,
    "Tamanho Turma": 99,
    "GMAT Médio": 615,
    "GPA Mínimo": (
        "3.0 (últimos 90 créditos trimestre / 60 créditos semestre; "
        "admite abaixo com qualificações fortes)"
    ),
    "Taxa Empregabilidade %": 92,
    "Salário Médio Pós USD": 106000,
    "Principais Empregadores": "Microsoft, T-Mobile, Meta, Amazon, Google",
    "Deadline Aplicação": "Round 1: 20/out/2026; Round 2: 5/jan/2027",
    "Rounds Aplicação": (
        "Round 1: 20/out/2026 (23h59 PDT); Round 2: 5/jan/2027 (23h59 "
        "PDT); início único em junho (verão)"
    ),
    "Documentos Necessários": (
        "Perfil de aplicação online; Statement of Purpose (obrigatório) + "
        "personal statement opcional (até 300 palavras cada); resume; 2 "
        "cartas de recomendação (preferência profissionais); transcripts; "
        "self-review de qualificações; teste de inglês (se aplicável); "
        "entrevista em vídeo curta após submissão"
    ),
    "Dificuldade Entrada": (
        "Moderada -- GPA médio 3.5, GMAT FE médio 615 (GMAT Focus "
        "Edition, escala diferente do GMAT clássico), GRE médio ~317 "
        "(Class of 2026; 70% dos alunos não enviaram GMAT/GRE -- teste é "
        "opcional); taxa de aceitação não divulgada publicamente"
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
        raise ValueError("Nenhuma linha encontrada para University of Washington (Seattle, WA). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (University of Washington, Seattle, WA)")

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

    print(f"\n{len(UPDATES)} campos atualizados na linha {row_idx} (University of Washington):\n")
    for header in UPDATES:
        print(f"  - {header}:")
        print(f"      Antes:  {before[header]!r}")
        print(f"      Depois: {after[header]!r}")


if __name__ == "__main__":
    main()
