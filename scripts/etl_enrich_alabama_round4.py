# -*- coding: utf-8 -*-
"""
Data enrichment (round 4): U of Alabama (Manderson/Culverhouse),
Tuscaloosa, AL -- MS in Business Analytics, presencial, STEM.

Fontes (verificadas por WebFetch/WebSearch em set/2026):
- studentaccounts.ua.edu/tuition-rates/ (Fall 2026/Spring 2027)
- manderson.culverhouse.ua.edu/programs/specialized-masters/business-analytics/
- manderson.culverhouse.ua.edu/business-administration-mba/tuition-financial-aid/
- manderson.culverhouse.ua.edu/business-administration-mba/faq/
- manderson.culverhouse.ua.edu/business-administration-mba/career-outcomes/
- catalog.ua.edu/graduate/about/academic-policies/application/
- international.ua.edu/isss/future-new-students/scholarships-aid/

ACHADOS:

1) TUITION CORRIGIDO: os valores anteriores (US$1.139,07 / US$406 por
   credito; US$34.172 / US$12.180 no total) batiam com as tarifas de
   2025-26. A UA cobra uma TARIFA FIXA por semestre para 9-15 horas; em
   Fall 2026/Spring 2027 sao US$17.962 (nao residente) e US$5.959
   (residente) por semestre. O MSBA tem 15 creditos no Fall + 15 no Spring,
   entao sao 2 semestres na tarifa fixa: US$35.924 intl / US$11.918
   residente. O custo por credito e o equivalente efetivo (total / 30).
   Obs.: US$1.840/h e a tarifa de overload/part-time, nao a base. A taxa
   de US$308 por curso de Business segue vigente em 2026-27 (Fees
   Adicionais = 10 x 308 = US$3.080, sem mudanca).

2) BOLSAS: a pagina do MSBA nao cita funding. O que existe: (a) bolsas de
   merito da Manderson, com consideracao automatica para admitidos e
   preferencia para quem envia GMAT/GRE ("cutoffs geralmente no top 10%
   do pool"); (b) GA da Manderson: metade da tuition + stipend de
   ~US$7.800 por 9 meses, 10h/semana, "tipicamente para os melhores
   alunos". Esses valores vem das paginas do MBA, e nao esta explicito se
   valem para o MSBA. (c) GTA/GRA de outros departamentos, via Graduate
   School. O realismo foi rebaixado de "Alta" para "Moderada", porque
   nada e garantido para o MSBA e a disputa e pelo top do pool.

3) TAXA DE APLICACAO: US$80 intl (US$65 domestico), confirmada no
   catalogo oficial. Free Application Week de 6-10/out/2025 (valida para
   Fall 2026): para internacionais, exige GPA >= 3.5 no ultimo diploma.
   Tambem ha isencao para veteranos/militares dos EUA, McNair e
   ex-alunos da UA. Fonte: noticias e agregadores; o catalogo so cita a
   isencao para graduandos da UA.

4) EMPREGABILIDADE: continua "Verificar". Os unicos numeros publicados
   (90,4% em 90 dias, turma de 2023) sao do MBA, e nao foram atribuidos ao
   MSBA para evitar contaminacao entre programas. Salario de US$82k
   (base, 2022, pagina do MSBA) foi mantido.

Re-confirmados sem mudanca: duracao de 9 meses (Fall + Spring, coorte
de Fall), presencial, 30 creditos, GPA 3.3, GMAT/GRE percentil 60,
TOEFL 90/IELTS 7.0/Duolingo 110/PTE 59 e a inconsistencia de deadline
(15/mar vs 15/abr).

Uso: python scripts/etl_enrich_alabama_round4.py (a partir da raiz).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "U of Alabama",
    "Estado": "AL",
    "Cidade": "Tuscaloosa",
}

UPDATES = {
    "Custo/Crédito (Intl) USD": 1197.47,
    "Custo/Crédito (Residente) USD": 397.27,
    "Tuition Total (Intl) USD": 35924,
    "Tuition Total (Residente) USD": 11918,
    "Tipos de Bolsa": (
        "Bolsa de mérito Manderson (consideração automática na admissão; "
        "preferência a quem envia GMAT/GRE; cutoff ~top 10% do pool); GA Manderson "
        "(metade da tuition + stipend ~US$7.800/9 meses, 10h/semana -- valores da "
        "página do MBA, aplicabilidade ao MSBA não explícita); GTA/GRA de outros "
        "departamentos via Graduate School (stipend + tuition + seguro)"
    ),
    "Quão Realista (Intl)": (
        "Moderada -- não há funding garantido/publicado para o MSBA; bolsa de mérito "
        "e GA são competitivos (top do pool) e favorecem quem envia GMAT/GRE; GTA/GRA "
        "exige buscar vaga ativamente em outros departamentos"
    ),
    "Rounds Aplicação": (
        "Sem rounds numerados -- admissão contínua (rolling); prioridade de bolsa de "
        "mérito e assistantship para quem aplica até 15/mar (ver Deadline Aplicação). "
        "Isenção da taxa (US$80 intl): 'Free Application Week' anual da Graduate School "
        "(edição de 6-10/out/2025 valeu para Fall 2026; internacionais precisam de GPA "
        "≥ 3.5 no último diploma); também isentos veteranos/militares dos EUA, "
        "McNair Scholars e ex-alunos da UA -- confirmar datas no site da Graduate School"
    ),
    "Documentos Necessários": (
        "Transcript oficial; evidência de Intro Stats (B+) e Calculus (B+); resume; "
        "statement of purpose; 3 cartas de recomendação; GMAT/GRE (salvo dispensa por "
        "pós-graduação); teste de inglês (internacionais)"
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
        raise ValueError("Nenhuma linha encontrada para U of Alabama (Tuscaloosa, AL). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (U of Alabama, Tuscaloosa, AL)")

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

    print(f"\n{len(UPDATES)} campos atualizados na linha {row_idx} (U of Alabama):\n")
    for header in UPDATES:
        print(f"  - {header}:")
        print(f"      Antes:  {before[header]!r}")
        print(f"      Depois: {after[header]!r}")


if __name__ == "__main__":
    main()
