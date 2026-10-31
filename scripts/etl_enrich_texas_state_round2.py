# -*- coding: utf-8 -*-
"""
Data enrichment (round 2): atualiza APENAS a linha 'Texas State' (McCoy
College of Business, San Marcos, TX) no msba_programs_database.xlsx.

CONTEXTO: rodada 1 (scripts/etl_enrich_texas_state.py) ja confirmou muito
bem: tuition, Fees Adicionais, Seguro Saude, bolsas/TA-RA (com a mesma
politica formal de waiver de tuition nao-residente ja vista em outras
universidades publicas do Texas), Salario Medio Pos ($98.333), Principais
Empregadores, testes de ingles detalhadissimos, deadlines, taxa de
aplicacao, documentos. Esta rodada focou nos campos que ainda estavam
"Verificar".

ATUALIZADO com fonte oficial (mccoy.txst.edu, mycatalog.txstate.edu):
- GPA Mínimo: estava "Verificar". O padrao institucional GERAL da Texas
  State para mestrados e 2,75, mas as paginas do MSDAIS especificamente so
  exigem "a competitive overall GPA or a competitive GPA in the last 60
  hours", sem um corte proprio divulgado -- registrado refletindo essa
  nuance (minimo institucional confirmado, mas sem corte especifico do
  programa).
- Matérias Pré-Requisito: estava "Verificar". ESCLARECIDO: NAO ha
  pre-requisito formal de disciplinas para ADMISSAO no MSDAIS. As
  mencoes a "pre-requisitos" que aparecem no catalogo (ex.: ISAN 5357 e
  ANLY 5336 com nota C ou melhor) sao pre-requisitos INTERNOS ao
  curriculo -- exigidos para cursar disciplinas avancadas (ex.: Machine
  Learning) depois de ja admitido no programa -- nao sao requisito de
  entrada. Essa distincao e relevante e foi registrada em vez de deixar
  "Verificar" (que sugeriria simplesmente "nao encontrado").
- GMAT/GRE Obrigatório: mantido "obrigatorio em geral, com possibilidade
  de waiver", mas o CRITERIO da dispensa foi detalhado com uma ressalva
  de CONFLITO entre fontes: a pagina oficial especifica do MSDAIS
  (mccoy.txst.edu/graduate/admissions/gmat-gre/waiver/msdais.html) diz
  que a dispensa e concedida com base em diploma de pos-graduacao
  PREVIAMENTE OBTIDO ("based on a previously earned advanced degree").
  Um resultado de busca nao verificado por fetch direto cita um criterio
  DIFERENTE (GPA dos ultimos 60 creditos >= 3,5) -- nao foi possivel
  reconciliar as duas fontes nesta rodada, entao ambas foram registradas
  com a ressalva de conflito, priorizando a fonte oficial verificada por
  fetch direto como mais confiavel.

NAO ALTERADO por falta de fonte oficial confiavel (mantido "Verificar"):
- Tamanho Turma, GPA Médio Admitidos, GMAT Médio, Taxa Aceitação %,
  Entrevista, Ranking US News, Ranking QS: nenhuma fonte oficial
  encontrada em nenhuma das paginas consultadas (mccoy.txst.edu,
  gradcollege.txst.edu, mycatalog.txstate.edu).
- Taxa Empregabilidade %: dois achados considerados e DESCARTADOS por
  risco de mistura de conceitos/programas: (1) "92,68% empregados em ate
  3 meses" e um numero do McCoy College INTEIRO (nao especifico do
  MSDAIS), de fonte terceira (niche.com), nao oficial; (2) um achado
  especifico do MSDAIS sobre "aumento salarial medio de $29.000 para
  quem continua no mesmo empregador" e uma metrica DIFERENTE (aumento de
  quem ja trabalhava, nao taxa de emprego de recem-formado) -- nenhum dos
  dois foi usado.

Fontes principais (pesquisadas via WebSearch/WebFetch, set/2026):
- https://mccoy.txst.edu/graduate/admissions/gmat-gre/waiver/msdais.html
  (oficial -- critério de dispensa de GMAT/GRE: diploma de pós-graduação prévio)
- https://mycatalog.txstate.edu/graduate/mccoy-business-administration/computer-information-systems-quantitative-methods/data-analytics-info-systems-ms/
  (oficial -- catálogo: GPA "competitivo" exigido, pré-requisitos internos ao currículo)
- https://mccoy.txst.edu/graduate/data-analytics.html (oficial -- página do programa)
- https://gradcollege.txst.edu/programs/msdais.html (oficial -- GPA institucional geral 2,75)
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "Texas State",
    "Estado": "TX",
    "Cidade": "San Marcos",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "GPA Mínimo": (
        "2,75 (mínimo institucional geral da Texas State para "
        "mestrados); o MSDAIS especificamente não divulga um corte "
        "próprio, exigindo apenas \"a competitive overall GPA or a "
        "competitive GPA in the last 60 hours\""
    ),
    "Matérias Pré-Requisito": (
        "Nenhum pré-requisito formal de disciplinas para ADMISSÃO. "
        "Menções a \"pré-requisitos\" no catálogo (ex.: ISAN 5357 e ANLY "
        "5336 com nota C ou melhor) são internas ao currículo -- exigidas "
        "para cursar disciplinas avançadas (ex.: Machine Learning) depois "
        "de já admitido, não para entrar no programa"
    ),
    "GMAT/GRE Obrigatório": (
        "Obrigatório em geral (GMAT ou GRE, pontuação competitiva), com "
        "possibilidade de dispensa. CRITÉRIO COM CONFLITO ENTRE FONTES: "
        "a página oficial específica do MSDAIS diz que a dispensa é "
        "concedida com base em diploma de pós-graduação previamente "
        "obtido (\"based on a previously earned advanced degree\"); uma "
        "fonte de busca não verificada por fetch direto cita um critério "
        "diferente (GPA dos últimos 60 créditos ≥ 3,5) -- não "
        "reconciliado nesta rodada, priorize a primeira (fonte oficial "
        "confirmada) até nova verificação"
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
        raise ValueError("Nenhuma linha encontrada para Texas State (San Marcos, TX). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Texas State, San Marcos, TX)")

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
