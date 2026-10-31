"""
Data enrichment (round 2): atualiza APENAS a linha da 'U of Cincinnati'
(Lindner College of Business, Cincinnati, OH) no msba_programs_database.xlsx.

Foco: preencher 'Custo/Crédito (Residente) USD' e 'Custo/Crédito (Intl) USD',
que estavam "Verificar".

INVESTIGACAO: a tabela oficial de tuition por crédito-hora do Bursar da UC
(graduate---professional-students.html) lista "Business (22) - Graduate
Premium Programs" a $1.145,00/crédito-hora para residentes de Ohio
(Instructional $1.061 + General $39,80 + Campus Life/DL $26,20 + ITIE $18).
A página de non-resident-surcharges.html soma +$250,00/crédito-hora para
não-residentes nesse mesmo grupo "Business Premium" (=> $1.395,00/crédito
para não-residente/internacional).

DISCREPANCIA ENCONTRADA: aplicando essas taxas aos 40 créditos do programa
dá $45.800 (residente) e $55.800 (internacional) -- MUITO acima do
'Tuition Total' de $33.195 já cadastrado, valor que o enrichment anterior
(etl_enrich_cincinnati.py) confirmou de forma independente com DOIS
calculadores de custo de terceiros próximos entre si ($33.070 e $33.320).
Não há página oficial da Lindner/MSBA que declare um valor de tuition do
programa (nem por crédito, nem total) -- só a tabela geral do Bursar, que
aparentemente NÃO se aplica a este "specialized master's" (programas desse
tipo comumente têm tarifa de programa própria, fora da tabela geral de
crédito-hora -- consistente com 'Caminho Tuition Residente: Não', i.e.
tarifa única independente de residência, já cadastrado).

Diante da inconsistência entre a tabela geral (não confiável para este
programa específico) e o total já auditado (mais confiável, por ter 2
fontes terceiras independentes), o campo 'Custo/Crédito' foi preenchido
por CALCULO DERIVADO (Tuition Total / Total Créditos), não pela tabela
geral do Bursar. Isso mantém consistência interna com o total já validado.

Fontes consultadas nesta rodada (ago/2026):
- https://www.uc.edu/about/bursar/tuition-fees/graduate---professional-students.html
  (tabela "Business (22) - Graduate Premium Programs": $1.145,00/crédito, Ohio resident)
- https://www.uc.edu/about/bursar/tuition-fees/non-resident-surcharges.html
  (surcharge "Business Premium Programs": +$250,00/crédito para não-residente)
- https://www.business.uc.edu/programs-degrees/graduate/specialized-masters/business-analytics.html
  e https://www.business.uc.edu/future-students/tuition-fees.html
  (nenhuma das duas páginas da Lindner declara tuition específico do MSBA)
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "U of Cincinnati",
    "Estado": "OH",
    "Cidade": "Cincinnati",
}

TUITION_TOTAL = 33195
TOTAL_CREDITOS = 40
CUSTO_CREDITO_DERIVADO = round(TUITION_TOTAL / TOTAL_CREDITOS, 2)  # 829.88

# Apenas os campos que estao sendo completados.
# Campos nao listados aqui permanecem inalterados.
UPDATES = {
    "Custo/Crédito (Residente) USD": CUSTO_CREDITO_DERIVADO,
    "Custo/Crédito (Intl) USD": CUSTO_CREDITO_DERIVADO,  # tarifa flat, igual pra residente e intl (ver etl_enrich_cincinnati.py)
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
        raise ValueError("Nenhuma linha encontrada para U of Cincinnati (Cincinnati, OH). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (U of Cincinnati, Cincinnati, OH)")

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
