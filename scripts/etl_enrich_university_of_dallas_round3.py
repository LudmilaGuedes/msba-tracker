# -*- coding: utf-8 -*-
"""
Data enrichment (round 3): auditoria + pesquisa aprofundada focada em UMA
unica linha -- University of Dallas (Satish & Yasmin Gupta College of
Business), Irving, TX -- Master of Science in Business Analytics,
presencial, STEM.

Pedido da usuaria: Data Enrichment de uma unica instituicao, cobrindo
tuition, bolsas, taxa de aplicacao, requisitos completos de aplicacao,
empregabilidade/salario/empregadores e duracao/formato. Substituir dados
incorretos, preencher lacunas, manter "Verificar" para o que nao for
encontrado.

*** ATENCAO CRITICA DE NOME (ja documentada no CLAUDE.md do projeto):
"University of Dallas" (privada, catolica, Irving TX, Gupta College of
Business) e "University of Texas at Dallas" / "UT Dallas" (publica,
Richardson TX, Jindal School of Management) sao DUAS instituicoes
DIFERENTES. A maioria dos resultados de busca sobre "business analytics
dallas" retorna dados da UT Dallas (Jindal), que tem um programa MSBA
muito mais divulgado publicamente. Toda informacao usada neste script foi
verificada como vindo especificamente de paginas *.udallas.edu (nunca
utdallas.edu), incluindo casos em que a propria busca alertou sobre essa
confusao antes de eu descartar o resultado. ***

Fontes oficiais (todas em udallas.edu, verificadas por WebFetch direto em
set/2026):
- udallas.edu/academics/programs/business-analytics/index.php
- udallas.edu/academics/programs/business-analytics/program.php
- udallas.edu/academics/programs/business-analytics/more-infornation.php
- udallas.edu/admissions-aid/graduate/index.php
- udallas.edu/admissions-aid/graduate/college-of-business/ms-mba-admission-requirements.php

ACHADOS:

1) MATERIAS PRE-REQUISITO: o registro anterior ja tinha a informacao
   correta, mas como uma INFERENCIA por ausencia ("nao confirmacao
   textual direta"). Nesta pesquisa, a pagina oficial de requisitos de
   admissao (ms-mba-admission-requirements.php) CONFIRMA textualmente:
   pre-requisitos de curso especificos so existem para MS Accounting, MS
   Finance e MS Financial Planning -- MS Business Analytics nao consta
   nessa lista. Atualizado de "inferencia" para "confirmado", removendo
   a ressalva.

2) TESTES DE INGLES: re-confirmados EXATAMENTE iguais aos ja registrados
   (TOEFL 80 / IELTS 6.5 / Duolingo 110 / MET 55 / PTE 67 / iTEP 3.9,
   direto da pagina oficial de requisitos) -- sem mudanca. Nota: um
   agregador terceiro (mastersportal.com) mostrava TOEFL 100 e Duolingo
   120 para este mesmo programa -- valores INCORRETOS (parecem ser
   templates genericos do agregador, nao dados reais do programa) --
   descartados e NAO usados.

3) DEADLINE DE SPRING: pagina oficial de admissoes de pos-graduacao
   confirma "MBA/MS: December 1, 2025 (Spring 2026)" como prazo -- bate
   exatamente com o "1º dez (Spring)" ja registrado. Sem mudanca.

4) NAO ENCONTRADO, mesmo apos pesquisa extensa em multiplas paginas
   oficiais (nenhum dado fabricado, permanece "Verificar"):
   - Duracao (meses): a pagina oficial fala em "30 horas-credito / 11-13
     cursos" e diz que profissionais em tempo parcial (2 materias/
     semestre) levam "1-2 anos" -- mas NAO ha declaracao oficial da
     duracao para um aluno em tempo integral (o caso do publico-alvo,
     visto F-1). Um agregador (mastersportal) alegava "12 meses", mas
     como esse mesmo agregador errou os dados de teste de ingles para
     este programa (ver item 2), esse numero NAO foi considerado
     confiavel o suficiente para uso.
   - Tamanho de Turma, GPA Minimo, GPA Medio Admitidos, GMAT Medio, Taxa
     de Aceitacao %, Taxa de Empregabilidade %, Salario Medio Pos,
     Principais Empregadores, Entrevista, Dificuldade de Entrada: NENHUMA
     pagina oficial do Gupta College publica esses dados especificamente
     para o MS in Business Analytics (nem para o college como um todo,
     em nivel de outcomes report). Uma pagina de terceiro (Princeton
     Review) tinha estatisticas de "Gupta College of Business", mas os
     numeros (aceitacao 75%, tuition anual $15.150, $560/credito, media
     de 8 anos de experiencia e 34 anos de idade) sao claramente de um
     programa DIFERENTE (provavelmente MBA part-time/executivo) --
     tuition nao bate nem de longe com os $1.300/credito ja confirmados
     para o MSBA -- entao NAO foram usados para evitar contaminacao de
     dados entre programas da mesma escola.

CONCLUSAO: este e um programa de escola pequena/menos divulgada
publicamente -- ao contrario de outras faculdades ja auditadas nesta
sessao (UW, Rochester), a Gupta College of Business nao publica relatorio
de outcomes/class profile para o MSBA. Recomenda-se contato direto com a
Graduate Admissions Office ((972) 721-5004) para os dados que permanecem
"Verificar".

Uso: python etl_enrich_university_of_dallas_round3.py (executar a partir
da raiz do repositorio).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "University of Dallas",
    "Estado": "TX",
    "Cidade": "Irving",
}

UPDATES = {
    "Matérias Pré-Requisito": (
        "Nenhum pré-requisito formal de disciplinas -- confirmado na "
        "página oficial de requisitos de admissão do Gupta College: "
        "pré-requisitos de curso específicos só existem para MS "
        "Accounting, MS Finance e MS Financial Planning; MS Business "
        "Analytics não é listado entre eles"
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
        raise ValueError("Nenhuma linha encontrada para University of Dallas (Irving, TX). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (University of Dallas, Irving, TX)")

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

    print(f"\n{len(UPDATES)} campo(s) atualizado(s) na linha {row_idx} (University of Dallas):\n")
    for header in UPDATES:
        print(f"  - {header}:")
        print(f"      Antes:  {before[header]!r}")
        print(f"      Depois: {after[header]!r}")


if __name__ == "__main__":
    main()
