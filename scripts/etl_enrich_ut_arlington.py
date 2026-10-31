"""
Data enrichment: atualiza APENAS a linha 'UT Arlington' (College of
Business, Arlington, TX) no msba_programs_database.xlsx.

Foco exclusivo: MS in Business Analytics, presencial (aulas noturnas em
Arlington), STEM.

Tuition NAO foi alterada: a pesquisa encontrou uma tarifa geral de pos-
graduacao nao-residente de $1.357/credito, mas isso nao bate com o valor
ja existente na planilha (Tuition Total Intl $32.055 / 33 creditos =
~$971/credito) -- uma diferenca de ~40% grande demais para assumir que uma
das duas esta certa sem confirmacao mais solida. Prefer-se nao sobrescrever
um dado que nao foi provado incorreto com um numero de fonte generica
potencialmente desatualizada ou de outro programa.

Fontes principais (pesquisadas manualmente, fora deste script):
- https://www.uta.edu/academics/programs/business-analytics-ms (duracao de 16
  meses, 33 creditos -- ja confirmado na planilha --, formato noturno,
  curriculo)
- GPA minimo de 3.0 (ultimas 60 horas), ja existente na planilha, confirmado
  por fonte independente sem alteracao
- busca por custo de vida em Arlington, TX (aluguel medio de 1-quarto
  ~$1.034/mes, 2026)

Nao encontrado mesmo apos pesquisa profunda (mantido "Verificar"): taxa de
empregabilidade, salario medio e ranking especificos do MSBA (so achei
medias genericas do setor e de outros programas da UTA, que NAO foram
usadas para nao inventar um dado especifico deste programa).
"""

import shutil
import sys
from datetime import datetime

import openpyxl
from openpyxl.styles import PatternFill

sys.stdout.reconfigure(encoding="utf-8")

DB_PATH = "data/msba_programs_database.xlsx"
SHEET_DATA = "MSBA Programs"

VERIFICAR = "Verificar"

MATCH = {
    "Nome da Universidade": "UT Arlington",
    "Estado": "TX",
    "Cidade": "Arlington",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Duração (meses)": 16,
    "Aluguel Privativo/mês USD": 1034,  # mercado de aluguel de Arlington, TX, 2026
}


def main():
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    backup_path = f"data/msba_programs_database.backup-{timestamp}.xlsx"
    shutil.copy2(DB_PATH, backup_path)
    print(f"Backup criado em: {backup_path}")

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
        raise ValueError("Nenhuma linha encontrada para UT Arlington (Arlington, TX). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (UT Arlington, Arlington, TX)")

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
