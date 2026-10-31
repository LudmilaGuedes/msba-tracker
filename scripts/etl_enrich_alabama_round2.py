# -*- coding: utf-8 -*-
"""
Data enrichment (round 2): auditoria + pesquisa aprofundada focada em UMA
unica linha -- University of Alabama (Manderson Graduate School of
Business / Culverhouse College), Tuscaloosa, AL -- MS in Business
Analytics, presencial, STEM.

Pedido da usuaria: Data Enrichment de uma unica instituicao, cobrindo
tuition, bolsas, taxa de aplicacao, requisitos completos de aplicacao,
empregabilidade/salario/empregadores e duracao/formato. Substituir dados
incorretos, preencher lacunas, manter "Verificar" para o que nao for
encontrado.

Fontes oficiais (verificadas por WebFetch/WebSearch em set/2026):
- manderson.culverhouse.ua.edu/programs/specialized-masters/business-analytics/
- manderson.culverhouse.ua.edu/business-administration-mba/faq/
- manderson.culverhouse.ua.edu/career-services/
- catalog.ua.edu/graduate/business/information-systems-statistics-management-science/business-analytics-msba/
- studentaccounts.ua.edu/tuition-rates/

ACHADOS:

1) *** INCONSISTENCIA REAL NO PROPRIO SITE OFICIAL: a pagina oficial do
   programa (business-analytics/) tem DUAS informacoes de deadline que
   se contradizem -- a secao de "informacoes do programa" diz "Application
   Deadline: April 15", mas a secao de FAQ da MESMA pagina diz "The
   application deadline is March 15". Nao ha como saber qual e a
   correta so pela leitura do site (nao e um caso de round1 vs round2 --
   sao afirmacoes conflitantes sobre O prazo). Registrado ambas as datas
   com a ressalva explicita da inconsistencia, seguindo o mesmo padrao ja
   usado em outras linhas desta base quando o proprio site oficial se
   contradiz (ex.: DePaul). ***

2) TUITION: PESQUISADO EXTENSIVAMENTE, MAS NAO ALTERADO -- a pagina
   oficial do programa diz textualmente que nao lista valores exatos
   ("Typical graduate tuition rates apply"). A pagina oficial de tarifas
   (studentaccounts.ua.edu) lista USD 490/hora residente e USD 1.840/hora
   nao-residente para Fall 2026/Spring 2027 -- MAS uma busca cruzada
   revelou que o valor de US$1.840/hora e especificamente a tarifa de
   "overload" (sobrecarga, acima de 15 horas), nao a tarifa-base
   equivalente de tempo integral -- ou seja, NAO da pra multiplicar esse
   numero por 30 creditos com confianca. Agregadores terceiros (Mastersportal,
   Yocket/Collegedunia) davam 3 valores diferentes e mutuamente
   contraditorios (US$34.542, US$39.840, US$31.460), um deles inclusive
   com um erro obvio de calculo (tratando o programa de 9 meses como se
   fosse 2 anos). Diante desse nivel de contradicao em TODAS as fontes
   novas, e como o valor ja registrado (Custo/Credito x Total Creditos =
   Total Tuition, exatamente consistente internamente) nao foi
   contradito por nenhuma fonte oficial direta e inequivoca, optei por
   NAO alterar -- risco de piorar o dado em vez de melhorar. Recomenda-se
   confirmar diretamente com studentaccounts@ua.edu.

3) FEES ADICIONAIS PREENCHIDO (estava "Verificar"): calculado a partir
   de 2 fontes oficiais diretas cruzadas -- o catalogo oficial confirma
   que o curriculo tem 5 cursos no Fall + 4 no Spring + 1 eletiva = 10
   cursos de 3 creditos (30 creditos totais), e a pagina oficial de
   tarifas confirma "fee de tecnologia/infraestrutura de US$308 POR
   CURSO" para cursos de Commerce/Business Administration (Fall
   2026/Spring 2027). 10 cursos x US$308 = US$3.080. Registrado como
   valor calculado (nao um numero unico publicado diretamente), com a
   metodologia documentada aqui.

4) DEADLINE/PRAZO DE BOLSA E ISENCAO DE TAXA: confirmado que a UA nao
   tem sistema de rounds -- e admissao rolling, com recomendacao
   explicita de aplicar ate 15/mar para prioridade de bolsa merito e
   assistantship. Tambem confirmado (fonte oficial de admissoes da UA,
   fora do Manderson): a universidade promove uma "Free Application Week"
   anual com isencao da taxa de aplicacao para elegiveis -- vale
   verificar as datas atualizadas a cada ano.

5) PRINCIPAIS EMPREGADORES EXPANDIDO: a pagina oficial do programa lista
   uma lista bem mais completa que a ja registrada (que so tinha 4
   nomes). Adicionados: Adobe, Bank of America, Ford, ExxonMobil,
   Chick-Fil-A (mantidos os 4 ja confirmados: P&G, Walmart, NBA, ESPN,
   ja que nao foram contraditos, so nao re-confirmados individualmente --
   a pagina cita "e mais de 15 outros" alem dos nomeados).

6) CURRICULO DETALHADO: catalogo oficial lista os cursos exatos (BAN
   500 Business Intelligence & Data Visualization, MIS 501/502, OM
   500/501, ST 541/542, BAN 501 Predictive Modeling, BAN 580 Capstone) --
   adicionado a "Destaques Curriculo".

7) SALARIO MEDIO: o valor ja registrado (US$82.000, 2022) CONTINUA sendo
   o unico numero encontrado especificamente atribuido a pagina do MS in
   Business Analytics (mesmo que datado). Uma cifra mais recente (Class
   of 2025: US$73.582 salario medio + US$4.478 bonus) foi encontrada na
   pagina de Career Services do Manderson, mas e uma media AGREGADA de
   TODOS os "Specialized Masters" (Accountancy, Applied Statistics,
   Business Analytics, Economics, Finance, Management, MIS, Marketing,
   Operations Management e Tax Accounting juntos) -- nao e especifica do
   MSBA. Para evitar o mesmo erro de contaminacao entre programas ja
   identificado e corrigido nesta sessao (caso da University of
   Washington, onde dado de MBA tinha contaminado a linha do MSBA),
   NAO foi usada para substituir o valor especifico ja registrado.
   Mencionada apenas no resumo ao usuario, nao gravada na planilha.

NAO ENCONTRADO (permanece "Verificar", nao inventado): Tamanho de Turma,
GPA Medio Admitidos, GMAT Medio, Taxa de Aceitacao %, Taxa de
Empregabilidade % (so achei salario, nao a taxa), Entrevista de
admissao, Seguro Saude do programa, Utilities/mes, Transporte/mes, Custo
Vida Total. Nenhuma pagina oficial do Manderson especifica do MSBA
publica esses numeros.

Uso: python etl_enrich_alabama_round2.py (executar a partir da raiz do
repositorio).
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
    "Deadline Aplicação": (
        "15/mar (FAQ oficial do programa) OU 15/abr (seção de admissões "
        "da mesma página oficial -- inconsistência interna do próprio "
        "site); recomendação oficial é aplicar até 15/mar para "
        "prioridade de bolsa/assistantship"
    ),
    "Rounds Aplicação": (
        "Sem rounds numerados -- admissão contínua (rolling); prioridade "
        "de bolsa de mérito e assistantship para quem aplica até 15/mar "
        "(ver Deadline Aplicação). A UA promove anualmente uma 'Free "
        "Application Week' com isenção da taxa de aplicação para "
        "elegíveis -- datas variam a cada ano, confirmar no site da "
        "Graduate School"
    ),
    "Fees Adicionais USD": 3080,
    "Principais Empregadores": (
        "P&G, Walmart, NBA, ESPN, Adobe, Bank of America, Ford, "
        "ExxonMobil, Chick-Fil-A (e mais de 15 outros, por página "
        "oficial do programa)"
    ),
    "Destaques Currículo": (
        "2 semestres, recent grads; currículo: Fall (BAN 500 Business "
        "Intelligence & Data Visualization, MIS 501, MIS 502, OM 500, ST "
        "541), Spring (BAN 501 Predictive Modeling, OM 501, ST 542, BAN "
        "580 Capstone) + 1 eletiva"
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
