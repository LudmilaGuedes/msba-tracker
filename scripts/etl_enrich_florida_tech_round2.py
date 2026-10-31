"""
Data enrichment (round 2): atualiza APENAS o campo 'Fees Adicionais USD' da
linha 'Florida Institute of Technology (Florida Tech)' (Nathan M. Bisk
College of Business, Melbourne, FL) no msba_programs_database.xlsx.

Foco exclusivo: Master of Science in Business Analytics, presencial (campus
de Melbourne, distinto da versão 100% online da mesma escola), STEM.

Contexto: o round 1 (etl_enrich_florida_tech.py) nunca foi de fato aplicado
ao arquivo vigente -- os valores atuais de tuition/crédito ($820) e taxa de
aplicação ($60) já estavam corretos e foram reconfirmados nesta rodada via
https://www.fit.edu/student-financial-services/tuition-and-fees/ (tabela
2026-27: "College of Business (MBA, Masters): $820/crédito", distinto da
tarifa de $1.390 usada por engenharia/aeronáutica/psicologia -- o round 1
parece ter aplicado a tarifa errada por engano, mas isso nunca chegou a
sujar o arquivo vigente).

Fonte desta atualização: mesma página oficial de tuition-and-fees, seção
"Mandatory Graduate Fees":
  - Technology Fee: $160/semestre (tempo integral, campus principal)
  - Activity Fee: $240/semestre (tempo integral, campus Melbourne)
  - Orientation Fee: $150 (única vez, primeiro semestre)

DECISÃO: só a Orientation Fee ($150) foi gravada em 'Fees Adicionais USD'.
Technology e Activity Fee são recorrentes POR SEMESTRE, e a duração exata
do programa presencial em semestres (campo 'Duração (meses)') continua
'Verificar' apesar de busca extensa -- não há base confiável para multiplicar
essas taxas sem inventar uma suposição sobre quantos semestres o programa
dura. Quando a duração for confirmada, este campo deve ser revisto para
somar Technology + Activity Fee proporcionalmente.
"""

import sys

import openpyxl

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "Florida Institute of Technology (Florida Tech)",
    "Estado": "FL",
    "Cidade": "Melbourne",
}

UPDATES = {
    "Fees Adicionais USD": 150,  # Orientation Fee (única vez) -- ver nota acima sobre Technology/Activity Fee
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
        raise ValueError("Nenhuma linha encontrada para Florida Institute of Technology. Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por segurança.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Florida Institute of Technology (Florida Tech), Melbourne, FL)")

    before = {}
    after = {}
    for header, new_value in UPDATES.items():
        col_idx = header_to_col[header]
        cell = ws.cell(row=row_idx, column=col_idx)
        before[header] = cell.value
        cell.value = new_value
        after[header] = new_value

    wb.save(DB_PATH)

    print("\nCampos atualizados (Como Estava -> Como Ficou):")
    for header in UPDATES:
        print(f"  - {header}:")
        print(f"      Antes: {before[header]!r}")
        print(f"      Depois: {after[header]!r}")


if __name__ == "__main__":
    main()
