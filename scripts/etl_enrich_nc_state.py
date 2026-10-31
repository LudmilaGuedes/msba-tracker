"""
Data enrichment: atualiza APENAS a linha da 'NC State' (Institute for
Advanced Analytics, Raleigh, NC) no msba_programs_database.xlsx.

Foco exclusivo: MS in Analytics (MSA), presencial, STEM, programa intensivo
de 10 meses (Summer II + Fall + Spring). Deep research em: custo total real
de tuition para internacionais, bolsas/assistantships, requisitos de GPA/
GMAT/GRE/inglês, e custo de vida local.

ACHADO DA AUDITORIA: 'Total Créditos' estava em 32; a fonte oficial (página
de tuition do Instituto, com o calendário exato Summer II '26 / Fall '26 /
Spring '27 -- ou seja, o PRÓPRIO ciclo letivo atual) confirma 34 créditos
(6 no Summer II + 14 no Fall + 14 no Spring). CORRIGIDO. Com isso, também
foi possível atualizar 'Tuition Total (Intl) USD': a planilha tinha $57.890
(valor de um ciclo letivo anterior, ainda muito próximo); a tabela oficial
para o ciclo 2026-27 (o mesmo em que a coleta foi feita) dá $59.679,06 para
não-residentes/internacionais -- CORRIGIDO para o valor mais atual. A tabela
já inclui as fees obrigatórias embutidas por semestre, por isso 'Fees
Adicionais USD' foi MANTIDO "Verificar": somá-las de novo separadamente
contaria a mesma taxa duas vezes no custo total (mesmo critério já aplicado
em Georgia Tech e NC State's própria vizinha, UGA).

'Custo/Crédito (Intl/Residente)' MANTIDOS "Verificar": a fonte oficial
descreve explicitamente a MSA como "packaged as a flat tuition" para os 34
créditos, sem uma tarifa por crédito real -- calcular um "por crédito"
dividindo o total inventaria uma granularidade que a própria universidade
não usa.

Custo de vida: 'Transporte/mês' foi preenchido com a estimativa oficial de
Cost of Attendance da NC State ($2.000/ano ÷ 10 meses do programa). Como
'Utilities/mês' continua sem uma fonte isolada confiável (a COA oficial
embute utilities dentro de uma categoria agregada de "Personal Expenses",
que não pode ser decomposta com confiança), 'Custo Vida Total (programa)
USD' permanece "Verificar" -- falta 1 dos 9 componentes para uma soma
completa.

Fontes principais (pesquisadas manualmente, fora deste script, ago/2026):
- https://analytics.ncsu.edu/?page_id=4202 (tabela oficial de tuition e
  fees, ciclo Summer II '26/Fall '26/Spring '27 -- o ciclo corrente)
- Busca por admissions: GRE/GMAT não exigidos nem considerados (já
  cadastrado corretamente); pré-requisitos de estatística/regressão e
  programação; documentos: transcript não-oficial, personal statement
  (1-2 páginas), 3 cartas de recomendação
- Busca por funding: internacionais não elegíveis a ajuda federal/estadual,
  maioria financia via empréstimo; bolsas pontuais existem (Graduate School
  Diversity Enhancement Grants, Warren-Wolfpack Club Graduate Fellowships)
  mas sem bolsa integral garantida; assistantships não são um caminho
  estruturado dentro do MSA, exigem busca em outras unidades da
  universidade
- studentservices.ncsu.edu (COA oficial 2025-26): transporte $2.000/ano
  para graduate students full-time
- GPA médio de 3.60 encontrado é da "Class of 2017" -- descartado por ser
  velho demais (quase 10 anos) para representar o programa hoje
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

VERIFICAR = "Verificar"

MATCH = {
    "Nome da Universidade": "NC State",
    "Estado": "NC",
    "Cidade": "Raleigh",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Total Créditos": 34,  # CORRECAO: planilha tinha 32; oficial: 6 (Summer II) + 14 (Fall) + 14 (Spring) = 34
    "Tuition Total (Intl) USD": 59679,  # CORRECAO: planilha tinha 57.890 (ciclo anterior); oficial 2026-27 = $59.679,06
    "Tuition Total (Residente) USD": 30643,  # oficial 2026-27: $30.642,68
    "Transporte/mês USD": 200,  # COA oficial NC State: $2.000/ano ÷ 10 meses do programa
    "Matérias Pré-Requisito": (
        "Pelo menos 1, idealmente 2 semestres de métodos estatísticos em "
        "nível universitário (incluindo regressão); experiência com "
        "programação em pelo menos 1 linguagem de computador"
    ),
    "Documentos Necessários": (
        "Transcript não-oficial de cada instituição cursada, personal "
        "statement (1-2 páginas), 3 cartas de recomendação"
    ),
    "Bolsas Disponíveis": "Sim (limitadas)",
    "Tipos de Bolsa": (
        "Bolsas menores/fellowships pontuais (ex.: Graduate School "
        "Diversity Enhancement Grants, Warren-Wolfpack Club Graduate "
        "Fellowships); sem bolsa integral garantida"
    ),
    "Quão Realista (Intl)": (
        "Baixa -- internacionais não são elegíveis para ajuda federal/"
        "estadual; maioria financia via empréstimo educacional; prêmios "
        "pontuais existem mas são raros"
    ),
    "TA/RA Disponível": (
        "Limitado -- não é uma via estruturada dentro do MSA; vagas de "
        "GA podem existir em outros departamentos/unidades da "
        "universidade, exigindo busca ativa fora do Instituto"
    ),
    "Processo TA/RA": (
        "Sem processo formal centralizado no MSA; contato direto com o "
        "Instituto ou busca de vagas em outras unidades da universidade"
    ),
    "Graduate Assistantships": (
        "Não é o caminho típico de financiamento do MSA (programa "
        "intensivo de 10 meses); vagas fora do Instituto podem existir, "
        "mas exigem busca ativa"
    ),
}


def main():
    rotate_backup()
    print(f"Backup rotacionado (política fixa) em: data/msba_programs_database_backup.xlsx")

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
