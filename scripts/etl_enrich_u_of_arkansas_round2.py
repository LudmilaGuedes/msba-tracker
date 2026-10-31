# -*- coding: utf-8 -*-
"""
Data enrichment (round 2): U of Arkansas (Walton College of Business),
Fayetteville, AR -- Master of Applied Business Analytics (MABA), trilha
FULL-TIME, presencial, STEM.

Fontes (verificadas por WebFetch/WebSearch em set/2026):
- walton.uark.edu/graduate-programs/business-analytics-masters-degree/
- walton.uark.edu/graduate-programs/admissions/tuition-cost.php (2026-27)
- walton.uark.edu/graduate-programs/apply.php
- catalog.uark.edu/graduatecatalog/admissions/ (catalogo 2026-27)
- international-students.uark.edu/admissions/english-language-proficiency.php
- waltoncareers.uark.edu/employment-outcomes/

ACHADOS:

1) *** INICIO CORRIGIDO: o full-time NAO comeca no Fall. A pagina oficial
   diz "The full-time program begins in mid-May", e a grade segue Summer I
   (ERP Fundamentals) -> Fall -> Spring -> Summer II (Practicum). O site
   chama a coorte de "Full-Time (Fall 2027)", com deadline intl de 15/fev,
   mas o inicio real das aulas e em meados de maio. ***

2) TESTES DE INGLES PREENCHIDOS (antes: "minimos nao divulgados"). O
   catalogo 2026-27 e a pagina de proficiencia concordam: TOEFL iBT 4.0
   (nova escala), IELTS Academic 6.5, Duolingo 110, PTE-A 58, iTEP 3.9,
   C1 Advanced 176, todos validos por 2 anos. Um snippet de busca antigo
   citava TOEFL 79 / IELTS 6.0 / DET 105, mas as duas fontes oficiais
   atuais prevalecem. Apos a admissao ha ainda uma avaliacao de redacao
   em ingles, que pode exigir cursos ELAC no 1o ano.

3) GPA MINIMO PREENCHIDO: 3.0 nos ultimos 60 creditos da graduacao, regra
   geral da Graduate School no catalogo 2026-27.

4) DOCUMENTOS: a pagina oficial fala em "Three letters of recommendation"
   (antes estava registrado "3 referencias profissionais"), alem de GMAT/
   GRE e teste de ingles.

5) TUITION re-confirmado para 2026-27, sem mudanca: US$741,88 / US$1.974,78
   por credito; total intl oficial US$62.583,40.

6) BOLSAS re-confirmadas como inexistentes: "No other Walton master's degree
   provides any direct funding" (so o MBA full-time tem assistantship).

7) EMPREGABILIDADE: continua "Verificar". Os unicos numeros publicados sao
   do MBA full-time (89% empregados em 3 meses, US$82.185, 2024-25) e nao
   foram atribuidos ao MABA. Um agregador cita US$70-85k de salario
   esperado, mas sem fonte oficial, entao nao foi usado.

8) ENTREVISTA: a pagina oficial de admissao nao menciona entrevista.
   TAXA DE APLICACAO: US$75, confirmada; nenhuma isencao mencionada.

Uso: python scripts/etl_enrich_u_of_arkansas_round2.py (a partir da raiz).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "U of Arkansas",
    "Estado": "AR",
    "Cidade": "Fayetteville",
}

UPDATES = {
    "Formato": (
        "Presencial (full-time); existe também a trilha Professional MABA "
        "part-time híbrida (1 sábado/mês, 5 semestres) -- não é o foco desta linha"
    ),
    "Início": (
        "Meados de maio (Summer) -- sequência Summer I → Fall → Spring → Summer II "
        "(practicum); o site rotula a coorte como 'Full-Time (Fall <ano>)'"
    ),
    "Destaques Currículo": (
        "ERP Fundamentals (Summer I), Data Analytics Fundamentals, Data Management "
        "Systems, decision support e business intelligence; practicum multi-semestre "
        "com dados reais de empresas; laços corporativos (Walmart, Tyson); opção 4+1"
    ),
    "GPA Mínimo": "3.0 nos últimos 60 créditos da graduação (regra da Graduate School)",
    "GMAT/GRE Obrigatório": (
        "OBRIGATÓRIO para candidatos só com histórico internacional (sem direito a "
        "dispensa); domésticos podem dispensar com GPA ≥3.2 de instituição dos EUA, "
        "3+ anos de experiência profissional, ou credenciais excepcionais (JD/MD/PhD)"
    ),
    "Testes Inglês Aceitos": (
        "TOEFL iBT 4.0 (nova escala) / IELTS Academic 6.5 / Duolingo 110 / PTE-A 58 / "
        "iTEP 3.9 / C1 Advanced 176 -- válidos por 2 anos; avaliação de redação após a "
        "admissão pode exigir cursos ELAC no 1º ano"
    ),
    "Deadline Aplicação": (
        "Full-time coorte 'Fall 2027': internacional 15/fev, doméstico 1/jul (aulas "
        "começam em meados de maio); coorte Fall 2026 já encerrada"
    ),
    "Rounds Aplicação": (
        "Sem rounds -- prazo único por coorte (ver Deadline Aplicação); nenhuma isenção "
        "da taxa de aplicação mencionada na página oficial"
    ),
    "Documentos Necessários": (
        "Online application; transcripts oficiais; GMAT/GRE (obrigatório para "
        "internacionais); Statement of Purpose (máx. 2 páginas); resume; 3 cartas de "
        "recomendação via portal; teste de inglês"
    ),
    "Entrevista": "Não mencionada na página oficial de admissão",
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
        raise ValueError("Nenhuma linha encontrada para U of Arkansas (Fayetteville, AR). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (U of Arkansas, Fayetteville, AR)")

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

    print(f"\n{len(UPDATES)} campos atualizados na linha {row_idx} (U of Arkansas):\n")
    for header in UPDATES:
        print(f"  - {header}:")
        print(f"      Antes:  {before[header]!r}")
        print(f"      Depois: {after[header]!r}")


if __name__ == "__main__":
    main()
