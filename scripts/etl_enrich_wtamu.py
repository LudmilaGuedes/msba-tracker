"""
Data enrichment: atualiza APENAS a linha 'West Texas A&M University
(WTAMU)' (Paul and Virginia Engler College of Business, Canyon, TX) no
msba_programs_database.xlsx.

Foco exclusivo: MS in Computer Information Systems and Business Analytics
(MSCISBA), presencial, STEM.

Fontes principais (pesquisadas manualmente, fora deste script):
- https://www.wtamu.edu/academics/college-business/academics/graduate-programs/programs/mscisba-program.html
  (estrutura de 4 semestres/ano, formato hibrido/presencial)
- busca por proficiencia em ingles (TOEFL 79 / IELTS 6.0 ou ESLI Level III,
  padrao geral da WTAMU)
- busca por assistantships e bolsas da WTAMU (Graduate Assistantships por
  colegio/departamento; prioridade de bolsas em 1 fev; taxa de inscricao
  internacional $90, sem dispensa)
- busca por custo de vida em Canyon, TX (aluguel medio de 1-quarto ~$725-
  781/mes, 2026)

Nao encontrado mesmo apos pesquisa profunda (mantido "Verificar"): duracao
exata em meses (a estrutura de 33-36 horas em 4 semestres/ano nao permite
calcular um numero confiavel sem inventar), tamanho de turma, GPA medio,
GMAT medio, taxa de empregabilidade, salario medio, ranking e principais
empregadores especificos do MSCISBA.
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.stdout.reconfigure(encoding="utf-8")

from db_retention import rotate_backup  # noqa: E402

DB_PATH = "data/msba_programs_database.xlsx"
SHEET_DATA = "MSBA Programs"

VERIFICAR = "Verificar"

MATCH = {
    "Nome da Universidade": "West Texas A&M University (WTAMU)",
    "Estado": "TX",
    "Cidade": "Canyon",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Testes Inglês Aceitos": "TOEFL 79 (iBT) / IELTS 6.0 (ou conclusão do ESLI Level III)",
    "Taxa Aplicação USD": 90,  # taxa internacional, sem dispensa disponível
    "Bolsas Disponíveis": "Sim",
    "Tipos de Bolsa": "Bolsas gerais da Graduate School; prioridade em 1 fev para o próximo ano acadêmico",
    "Deadline Aplicação": (
        "Prioridade de bolsas: 1 fev (próximo ano acadêmico); admissão em "
        "4 semestres/ano (Fall, Spring, Summer I e II)"
    ),
    "TA/RA Disponível": "Possível (Graduate Assistantships por colégio/departamento, cobertura parcial + salário)",
    "Graduate Assistantships": (
        "Disponíveis por colégio/departamento; contatar o Engler College "
        "of Business (COBGrad@wtamu.edu) para vagas específicas"
    ),
    "Documentos Necessários": (
        "Application, comprovação de proficiência em inglês "
        "(internacionais), pagamento da taxa de inscrição, cumprimento de "
        "todos os requisitos do programa"
    ),
    "Aluguel Privativo/mês USD": 725,  # 1-quarto médio em Canyon, TX, 2026
}


def main():
    rotate_backup()
    print("Backup da versão anterior salvo em: data/msba_programs_database_backup.xlsx")

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
        raise ValueError("Nenhuma linha encontrada para West Texas A&M University (Canyon, TX). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (West Texas A&M University (WTAMU), Canyon, TX)")

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
