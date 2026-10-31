"""
Anonimiza o texto de 'Deadline Aplicação': troca "Para a usuária (...)" por
"Para ingresso (...)" nas linhas que traziam a data equivalente para o ciclo
alvo. Só muda o rótulo; a data e o início do texto ficam iguais, então a
classificação Round 1/2/3 (primeiro mês do campo) não se altera.

Rodar a partir da raiz do repositório.
"""

import shutil

import openpyxl

from db_retention import rotate_backup, DB_PATH

PUBLIC_COPY = "msba-dashboard/public/data/msba_programs_database.xlsx"
COLUMN = "Deadline Aplicação"

rotate_backup()
wb = openpyxl.load_workbook(DB_PATH)
ws = wb["MSBA Programs"]
headers = [c.value for c in ws[1]]
col = headers.index(COLUMN) + 1

changed = 0
for row in range(2, ws.max_row + 1):
    cell = ws.cell(row=row, column=col)
    if isinstance(cell.value, str) and "Para a usuária" in cell.value:
        cell.value = cell.value.replace("Para a usuária", "Para ingresso")
        changed += 1
        print(f"  {ws.cell(row=row, column=1).value}")

wb.save(DB_PATH)
shutil.copy2(DB_PATH, PUBLIC_COPY)
print(f"{changed} célula(s) alterada(s); cópia atualizada em {PUBLIC_COPY}")
