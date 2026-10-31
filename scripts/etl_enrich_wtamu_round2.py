"""
Data enrichment (round 2): atualiza APENAS a linha 'West Texas A&M
University (WTAMU)' (Paul and Virginia Engler College of Business,
Canyon, TX) no msba_programs_database.xlsx.

Pedido da usuaria: reconfirmar tuition para o perfil dela, confirmar
bolsas/descontos (se existem de verdade e quao realistas sao), e buscar
rankings/empregabilidade/salario medio inicial.

=========================================================================
1) TUITION -- JA CORRIGIDA EM AUDITORIA ANTERIOR DESTA SESSAO
=========================================================================
Reconfirmado na fonte oficial: $980/crédito internacional (Residente
$490, Out-of-state doméstico $920). Bate com o que já estava corrigido
($32.340 total intl). Sem alteração.

=========================================================================
2) ACHADO IMPORTANTE -- MESMA REGRA "BOLSA-DESTRAVA-TARIFA" DO TEXAS,
   MAS COM PRAZO DE VALIDADE
=========================================================================
Fonte oficial (wtamu.edu/.../international-resident-tuition.html),
citação exata: "International students who are awarded a university
scholarship in the amount of $1,000 or more per academic year may be
eligible to pay lower tuition rates." CONFIRMADO que se aplica tanto a
graduação quanto a MESTRADO ("Estudantes de pós-graduação (Masters)").
Processo: candidatar-se à bolsa DEPOIS de admitida, ~48h para os
sistemas sincronizarem.

ALERTA DE PRAZO (achado crítico, citação exata da própria página):
"Please note that after the 2025-2026 academic year, not all students
who receive a $1,000 or more scholarship may qualify for the lower
tuition rates." Ou seja: a regra pode ficar mais restrita a partir do
ciclo 2026-2027 -- exatamente o ciclo em que a usuária provavelmente se
candidataria. Isso NÃO significa que a regra vai acabar, mas que ela
está sinalizada para mudar; vale confirmar o estado atual direto com o
programa antes de contar com isso no planejamento financeiro.

=========================================================================
3) BOLSAS GERAIS DA GRADUATE SCHOOL -- EXIGEM FAFSA (PROVÁVEL BARREIRA
   PARA INTERNACIONAIS)
=========================================================================
Fonte oficial (wtamu.edu/.../financial-information/scholarships.html):
elegibilidade exige "Demonstrate need with submission of the FAFSA".
FAFSA é ajuda federal americana, tipicamente restrita a cidadãos/
residentes permanentes/não-cidadãos elegíveis -- NA PRÁTICA, isso
exclui a maioria dos candidatos internacionais em visto F-1 desse pool
específico de bolsas (mesmo padrão de barreira já visto com o A2O da
UCF nesta sessão). O caminho realista para a usuária continua sendo a
regra do item 2 (bolsa de $1.000+ -> tarifa reduzida), não as bolsas
gerais da Graduate School.
As "Merit Scholarships" (President's/Provost's/Honor Scholar, $1.500-
$2.500 + $12.000 de waiver) encontradas em busca NÃO foram usadas --
nomenclatura e valores são típicos de bolsa de CALOURO (undergraduate
freshman), sem confirmação de que se aplicam a mestrado.

=========================================================================
4) RANKING -- ENCONTRADO, ESPECÍFICO DO PROGRAMA (COM RESSALVA)
=========================================================================
Fonte oficial (wtamu.edu/.../rankings-and-recognition.html): o MSCISBA
(o programa desta linha) é ranqueado pelo U.S. News & World Report como
"#18 Best Online Graduate Information Technology Program" (2026) --
trajetória: #16 (2020-22) -> #16 (2023) -> #14 (2024) -> #22 (2025) ->
#18 (2026). RESSALVA: é um ranking da categoria "Online" e "Information
Technology" (não "Business Analytics" nem especificamente a modalidade
presencial) -- mas é o MESMO programa/grau (MSCISBA), só ranqueado pela
publicação nessa categoria. Sem ranking QS encontrado.

=========================================================================
5) EMPREGABILIDADE/SALÁRIO -- NÃO USADOS (RISCO DE CONTAMINAÇÃO)
=========================================================================
Encontrados "66,70% empregados" e "salário médio $69.316" -- mas o
primeiro não especifica de qual programa (pode ser do college como um
todo), e o segundo é EXPLICITAMENTE do MBA, não do MSCISBA. Um resultado
de busca sobre "graduate assistantship stipend" também misturou dados de
grad.tamu.edu -- que é da Texas A&M University (College Station), uma
universidade DIFERENTE e maior do sistema A&M, não a West Texas A&M
(WTAMU/Canyon) desta linha. Nenhum dos dois foi usado, mesmo critério de
sempre nesta sessão (evitar contaminação entre programas/instituições).
'Taxa Empregabilidade %', 'Salário Médio Pós USD' e 'Principais
Empregadores' permanecem "Verificar".
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "West Texas A&M University (WTAMU)",
    "Estado": "TX",
    "Cidade": "Canyon",
}

UPDATES = {
    "Caminho Tuition Residente": (
        "Sim, com alerta de prazo -- internacionais que recebem bolsa "
        "universitária da WTAMU de $1.000+/ano podem pagar tarifa "
        "reduzida, aplicável também a mestrado (candidatar-se à bolsa "
        "APÓS a admissão, ~48h de sincronização). RESSALVA OFICIAL: "
        "\"after the 2025-2026 academic year, not all students who "
        "receive a $1,000 or more scholarship may qualify\" -- regra "
        "pode ficar mais restrita a partir de 2026-27; confirmar status "
        "atual direto com o programa"
    ),
    "Quão Realista (Intl)": (
        "Moderada -- bolsas gerais da Graduate School exigem FAFSA "
        "(ajuda federal americana, na prática inacessível a "
        "internacionais F-1); mas a regra de bolsa de $1.000+ -> tarifa "
        "reduzida É real e aplicável a mestrado internacional -- "
        "caminho mais realista aqui é focar nela, não nas bolsas gerais"
    ),
    "Ranking US News": (
        "#18 (2026) -- \"Best Online Graduate Information Technology "
        "Program\", mesmo grau (MSCISBA); trajetória: #16 (2020-22), "
        "#16 (2023), #14 (2024), #22 (2025), #18 (2026)"
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
        raise ValueError("Nenhuma linha encontrada para WTAMU (Canyon, TX). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (WTAMU, Canyon, TX)")

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

    print("\nCampos atualizados (Como Estava -> Como Ficou):")
    for header in UPDATES:
        print(f"  - {header}:")
        print(f"      Antes: {before[header]!r}")
        print(f"      Depois: {after[header]!r}")


if __name__ == "__main__":
    main()
