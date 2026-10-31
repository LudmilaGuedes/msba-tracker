# -*- coding: utf-8 -*-
"""
Data enrichment (round 2): atualiza APENAS a linha 'U of Houston' (Bauer
College of Business, Houston, TX) no msba_programs_database.xlsx.

CONTEXTO: rodada 1 (scripts/etl_enrich_u_of_houston.py) ja confirmou bem
tuition ($1.680,56/credito intl, $1.041,67/credito residente; totais
$60.500 intl / $37.500 residente; 36 creditos), um caminho de tuition
"residente" para internacionais sob "certas condicoes" ($42.500), bolsas
(Bauer Excellence Scholarship ~$5.000), GMAT/GRE obrigatorio salvo
dispensa, testes de ingles com notas, taxa de aplicacao $80, deadlines,
documentos. Esta rodada focou nos campos que ainda estavam "Verificar".

ATUALIZADO com fonte oficial (bauer.uh.edu/ms-business-analytics/):
- GPA Médio Admitidos: estava "Verificar". CONFIRMADO: 3,62 (turma que
  entra Fall 2025) -- a mesma pagina oficial do programa tambem traz
  idade media 26 anos, 0,3 anos de experiencia profissional media, e
  composicao 90% masculino / 10% feminino (dados extras sem coluna
  dedicada no schema, registrados aqui apenas como contexto).
- Entrevista: estava "Verificar". CONFIRMADO: Sim, por convite -- texto
  oficial: "applicants selected to continue the process will be invited
  to interview" (nao e universal, so para pre-selecionados).
- Experiência Profissional: mantido "Não obrigatória" mas ENRIQUECIDO com
  nuance oficial: a mesma pagina afirma "post-bachelor's professional
  work experience strongly preferred" -- ou seja, nao e exigencia formal,
  mas e fortemente preferida na pratica, o que bate com a media de so 0,3
  anos de experiencia da turma admitida (indicando que a maioria entra
  vinda direto da graduacao, apesar da preferencia declarada).

RECONFIRMADO sem alteracao: GMAT/GRE obrigatorio salvo dispensa ("GRE/GMAT
required, waivers available upon request") -- bate exatamente com o que
ja estava cadastrado. Deadlines de aplicacao tambem reconfirmados sem
mudanca.

NAO RECONFIRMADO -- Caminho Tuition Residente ($42.500 sob "certas
condicoes" para internacionais): esta rodada tentou reconfirmar a
condicao exata mas so achou mencao generica (nao especifica do MSBA) de
que algumas bolsas do Bauer vem com "in-state tuition waiver if needed"
numa pagina de OUTRO programa (MS Accountancy) -- nao e confirmacao
direta da condicao ja cadastrada para o MSBA. Como o valor original nao
foi contradito, apenas nao pode ser reconfirmado com uma fonte nova, foi
MANTIDO sem alteracao (nao ha base para dizer que esta errado).

NAO ALTERADO por falta de fonte oficial confiavel (mantido "Verificar"):
- GPA Mínimo, GMAT Médio, Taxa Aceitação %, Tamanho Turma, Rounds
  Aplicação (alem dos 2 deadlines ja confirmados): a propria pagina
  oficial do programa NAO divulga esses numeros (confirmado que a
  ausencia e por nao-divulgacao, nao por falha de busca).
- Fees Adicionais USD: a pagina oficial de custo
  (bauer.uh.edu/ms-business-analytics/cost.php) retornou erro 404 nas
  duas tentativas desta rodada.
- TA/RA Disponível, Processo TA/RA, Graduate Assistantships: nenhuma
  fonte oficial especifica do Bauer/MSBA encontrada.
- Ranking US News, Ranking QS: nao encontrado nesta rodada.
- Taxa Empregabilidade %, Salário Médio Pós USD, Principais Empregadores:
  o Rockwell Career Center do Bauer so publica estatisticas de emprego
  para MBA e para graduacao -- CONFIRMADO que nao ha linha especifica do
  MSBA nessas estatisticas. Numeros encontrados em buscas (89% placement
  MBA, $70.162 salario medio de GRADUACAO, lista de empregadores da
  Class of 2025 de graduacao) sao de programas DIFERENTES -- usa-los
  aqui misturaria dados de programas distintos, o que o projeto proibe.
  Mantidos "Verificar".

Fontes principais (pesquisadas via WebSearch/WebFetch, set/2026):
- https://www.bauer.uh.edu/ms-business-analytics/
  (oficial -- GPA médio 3,62, entrevista por convite, experiência
  profissional preferida mas não exigida, reconfirma GMAT/GRE e deadlines)
- https://careercenter.bauer.uh.edu/graduate-employment-statistics/
  (oficial -- confirmado que NÃO há dado específico do MSBA, só MBA/undergrad)
- https://www.bauer.uh.edu/ms-business-analytics/cost.php (erro 404, não acessível)
- https://www.bauer.uh.edu/ms-accountancy/apply/international.php
  (oficial, mas de OUTRO programa -- não usado para reconfirmar tuition residente do MSBA)
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "U of Houston",
    "Estado": "TX",
    "Cidade": "Houston",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "GPA Médio Admitidos": 3.62,
    "Entrevista": "Sim, por convite (não universal -- só candidatos pré-selecionados são convidados)",
    "Experiência Profissional": (
        "Não obrigatória, mas \"strongly preferred\" (fortemente "
        "preferida) segundo a página oficial; na prática, a turma "
        "admitida tem em média apenas 0,3 anos de experiência "
        "profissional, sugerindo que a maioria entra direto da graduação"
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
        raise ValueError("Nenhuma linha encontrada para U of Houston (Houston, TX). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (U of Houston, Houston, TX)")

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
