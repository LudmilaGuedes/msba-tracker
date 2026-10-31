# -*- coding: utf-8 -*-
"""
CORRECAO: a rodada anterior (etl_enrich_batch_pesquisar_mais_round1.py)
sobrescreveu por engano 2 campos da Wayne State University que JA TINHAM
dados bons (nao eram "Verificar"), perdendo informacao no processo -- o
script tratou incorretamente esses campos como lacunas vazias. Este
script MESCLA a informacao antiga (preservada) com a nova encontrada
nesta sessao, em vez de substituir.

- Matérias Pré-Requisito: o valor antigo mencionava cursos de nivelamento
  no verao e possibilidade de admissao condicional -- informacao perdida
  na sobrescrita. Restaurado + mesclado com o detalhe novo (linguagens de
  programacao especificas aceitas).
- Tipos de Bolsa: o valor antigo mencionava o portal WSU
  ScholarshipUniverse (~200 alunos de business contemplados/ano, funding
  need-based) -- tambem perdido. Restaurado + mesclado com o AGRADE
  Scholarship (achado novo, mais especifico sobre waiver de tuition
  in-state).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "Wayne State University",
}

UPDATES = {
    "Matérias Pré-Requisito": (
        "Competência básica em probabilidade, estatística e programação "
        "(C, C++, Java, Python, R ou FORTRAN); cursos de nivelamento "
        "disponíveis no verão para candidatos com lacunas, podendo "
        "receber admissão condicional"
    ),
    "Tipos de Bolsa": (
        "Bolsas nomeadas via portal WSU ScholarshipUniverse (~200 alunos "
        "de business contemplados/ano; maioria do funding need-based, "
        "FAFSA, aplicável a domésticos); AGRADE Scholarship -- cobre até "
        "4 créditos na tarifa in-state por 2 semestres consecutivos "
        "(Fall+Winter ou Winter+Fall)"
    ),
}


def main():
    rotate_backup()
    print("Backup rotacionado (política fixa) em: data/msba_programs_database_backup.xlsx")

    wb = openpyxl.load_workbook(DB_PATH)
    ws = wb[SHEET_DATA]

    headers = [cell.value for cell in ws[1]]
    header_to_col = {h: i + 1 for i, h in enumerate(headers)}

    col_uni = header_to_col["Nome da Universidade"]

    target_rows = [
        row[0].row for row in ws.iter_rows(min_row=2)
        if row[col_uni - 1].value == MATCH["Nome da Universidade"]
    ]
    if len(target_rows) == 0:
        raise ValueError("Nenhuma linha encontrada para Wayne State University. Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Wayne State University)")

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

    print("\nCampos corrigidos (Antes -> Depois):")
    for header in UPDATES:
        print(f"  - {header}:")
        print(f"      Antes: {before[header]!r}")
        print(f"      Depois: {after[header]!r}")


if __name__ == "__main__":
    main()
