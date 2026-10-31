"""
Data enrichment: atualiza APENAS a linha 'Cal Poly Pomona' (Singelyn
Graduate School, Pomona, CA) no msba_programs_database.xlsx.

ATENCAO: distinta da linha 'Cal Poly SLO' (San Luis Obispo, CA), uma
universidade diferente do mesmo sistema Cal Poly. O matching exige Nome +
Estado + Cidade.

Foco exclusivo: MS in Business Analytics, presencial, STEM.

CORRECAO PRINCIPAL: a planilha tinha 'Tuition Total (Intl)': $20.380, valor
muito proximo da tarifa de RESIDENTE calculada a partir das tarifas
oficiais por unidade. A CSU cobra em faixas (nao linear): full-time (7+
unidades) e um valor fixo por termo, mas a tarifa part-time (0-6 unidades)
e linear e serve de base confiavel para calcular um "custo por unidade"
generico: residente $413,33/unidade (2.480/6), nao-residente +$471/unidade
de sobretaxa = $884,33/unidade. Confirmado tambem que a 'Professional
Program Fee' (+$321/unidade) NAO se aplica ao MS Business Analytics (so
MBA, MS Accountancy e MS Information Security).

Fontes principais (pesquisadas manualmente, fora deste script):
- https://www.cpp.edu/student-accounting/tuition/graduate-tuition-and-fees.shtml
  (tarifas oficiais 2026-27 por unidade, residente e nao-residente)
- https://www.cpp.edu/cba/masters-business-analytics.shtml (pre-requisitos,
  documentos, GPA minimo ja confirmado, sem GRE/GMAT)
- busca por custo de vida em Pomona, CA (aluguel medio de 1-quarto ~$2.175/
  mes, 2026 -- regiao cara do sul da California)

Nao encontrado mesmo apos pesquisa profunda (mantido "Verificar"): taxa de
empregabilidade especifica do MSBA (um resultado isolado citava "26,70%
empregados na formatura", mas essa metrica parece medir emprego NO MOMENTO
da formatura, nao apos um periodo de busca -- diferente da convencao usada
no resto desta planilha -- entao NAO foi usada, para nao subestimar
artificialmente o programa), salario medio, ranking, tamanho de turma exato.
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
    "Nome da Universidade": "Cal Poly Pomona",
    "Estado": "CA",
    "Cidade": "Pomona",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Intl) USD": 884.33,  # 413,33 (residente) + 471 (sobretaxa não-residente)
    "Custo/Crédito (Residente) USD": 413.33,  # 2.480 / 6 unidades (faixa part-time, linear)
    "Tuition Total (Intl) USD": 30067,  # CORRECAO: 34 x 884,33 (valor antigo parecia usar tarifa de residente)
    "Tuition Total (Residente) USD": 14053,  # 34 x 413,33
    "Matérias Pré-Requisito": (
        "Álgebra universitária, estatística e programação, nota C ou "
        "superior"
    ),  # CORRECAO: planilha tinha "Calculus + 2 Stats"; fonte oficial do CBA especifica algebra+estatistica+programacao
    "Testes Inglês Aceitos": "TOEFL/IELTS; CPP English Language Institute (CPELI) nível 6 também aceito",
    "Documentos Necessários": "2 cartas de recomendação, resume, certificados; sem GRE/GMAT exigido",
    "Bolsas Disponíveis": (
        "Sim (bolsas gerais via Financial Aid Office; nenhum programa "
        "específico do MSBA identificado)"
    ),
    "Aluguel Privativo/mês USD": 2175,  # 1-quarto médio em Pomona, CA, 2026
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
        raise ValueError("Nenhuma linha encontrada para Cal Poly Pomona (Pomona, CA). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Cal Poly Pomona, Pomona, CA)")

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
