"""
Data enrichment (RODADA 2): atualiza APENAS a linha 'Purdue' (Mitch
Daniels School of Business, West Lafayette, IN) no
msba_programs_database.xlsx.

Esta linha já passou por uma rodada de enrichment muito completa
(script etl_enrich_purdue.py), restando 8/54 campos 'Verificar'.
Usuário confirmou via AskUserQuestion (ago/2026) para prosseguir com
uma Rodada 2.

'Ranking US News' preenchido: "#9 Business Analytics (U.S. News and
World Report, 2023)" -- ranking específico da especialidade de
Business Analytics (o mais diretamente comparável aos rankings de
outras universidades já registrados nesta planilha), com nota de que o
ano-fonte é 2023 (mais antigo que o ideal). A mesma página também cita
"#7 Most Innovative Schools (U.S. News, 2025)", mas essa é uma
categoria diferente (inovação institucional geral, não a
especialidade de Business Analytics) -- não usada para não misturar
métricas diferentes no mesmo campo.

'Ranking QS' preenchido: "#24 MS in Business Analytics (QS Global,
2026)" -- dado oficial e atual, diretamente da página do programa.

'Matérias Pré-Requisito' preenchido: nenhum pré-requisito formal de
admissão confirmado; algumas ELETIVAS específicas dentro do currículo
exigem cursos-base prévios ou proficiência em Python (não é uma
barreira de entrada, é um requisito interno de matrícula em certas
disciplinas).

CUIDADO APLICADO: a página do Class Profile confirma GPA médio de 3.4
(consistente, arredondamento do valor já presente e mais preciso,
3.41 -- não alterado por já estar correto e mais detalhado) e "92%
Employment Rate Within 6 months" (bate exatamente com o valor já
presente em 'Taxa Empregabilidade %', validado sem necessidade de
correção).

CUIDADO APLICADO (dados NÃO encontrados mesmo após tentativas
adicionais nesta rodada, permanecem 'Verificar'): 'Tamanho Turma'
(o PDF de perfil oficial e a página de Class Profile trazem
demografia detalhada -- experiência, gênero, origem geográfica,
diversidade de undergrad majors -- mas NENHUM número absoluto de
turma), 'Seguro Saúde (programa) USD', 'Custo Vida Total (programa)
USD', 'GMAT Médio' (nenhuma fonte oficial consultada divulga essa
média, apenas o preferencial "GMAT ≥600" já presente em outro campo),
'Taxa Aceitação %'.

Fontes principais (pesquisadas via WebFetch, ago/2026):
- https://business.purdue.edu/masters/programs/ms-business-analytics-and-information-management/profile-msbaim.pdf
  (PDF oficial de perfil de turma: GPA 3.41 confirmado, demografia
  detalhada, sem tamanho de turma nem GMAT)
- https://business.purdue.edu/masters/programs/ms-business-analytics-and-information-management/
  (rankings QS #24 e US News #9/#7, 92% empregabilidade confirmado)
- .../class-profile.php (GPA 3.4, experiência 1-3 anos, gênero,
  8 estados + 12 países, 15 majors — sem tamanho de turma/GMAT/
  taxa de aceitação)
- .../curriculum.php (ausência de pré-requisito formal de admissão,
  pré-requisitos apenas para eletivas específicas)
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
    "Nome da Universidade": "Purdue",
    "Estado": "IN",
    "Cidade": "West Lafayette",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Ranking US News": "#9 Business Analytics (U.S. News and World Report, 2023)",
    "Ranking QS": "#24 MS in Business Analytics (QS Global, 2026)",
    "Matérias Pré-Requisito": (
        "Nenhum pré-requisito formal para admissão; algumas eletivas "
        "específicas do currículo exigem cursos-base prévios ou "
        "proficiência em Python"
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
        raise ValueError("Nenhuma linha encontrada para Purdue (West Lafayette, IN). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Purdue, West Lafayette, IN)")

    no_fill = PatternFill(fill_type=None)
    verificar_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    before = {}
    after = {}
    for header, new_value in UPDATES.items():
        col_idx = header_to_col[header]
        cell = ws.cell(row=row_idx, column=col_idx)
        before[header] = cell.value
        cell.value = new_value
        cell.fill = verificar_fill if new_value == VERIFICAR else no_fill
        after[header] = new_value

    wb.save(DB_PATH)

    print("\nCampos atualizados (Como Estava -> Como Ficou):")
    for header in UPDATES:
        print(f"  - {header}:")
        print(f"      Antes: {before[header]!r}")
        print(f"      Depois: {after[header]!r}")


if __name__ == "__main__":
    main()
