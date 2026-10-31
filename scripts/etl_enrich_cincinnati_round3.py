"""
Data enrichment (round 3): atualiza APENAS a linha da 'U of Cincinnati'
(Lindner College of Business, Cincinnati, OH) no msba_programs_database.xlsx.

Foco: custo de vida (não-tuition) e precisao do seguro saude, usando fontes
oficiais adicionais indicadas pela usuaria.

FONTES CONSULTADAS NESTA RODADA (ago/2026):

1. https://www.uc.edu/about/financial-aid/starting/costs/non-tuition-costs/non-tuition-costs-2027.html
   Tabela oficial "Graduate or Professional Student" 2026-27 (ano academico
   de 9 meses -- BATE EXATO com a duracao do programa, 9 meses):
     - Housing and Food:              $17.526  (nao separado em aluguel vs mercado)
     - Books and Course Materials:    $1.300
     - Transportation (Ohio/metro/reciprocity): $2.038
     - Transportation (non-resident): $3.038   <- usado aqui (maioria dos MSBA e intl/non-resident)
     - Miscellaneous:                 $4.892
   Pagina NAO diferencia on-campus/off-campus/com-familia para graduate
   (so undergrad tem essa divisao) -- e uma unica tabela para todos os
   grad/professional.

2. https://www.uc.edu/about/bursar/tuition-fees/program-fees.html
   Nao lista nenhuma taxa de programa para "Business Analytics"/"MSBA"/
   Lindner (so achou "Business (22) Undergraduate" $42/credito ou $500/
   semestre, e programas de Taxation grad a $100/credito -- nenhum dos
   dois se aplica ao MSBA). 'Fees Adicionais USD' permanece "Verificar":
   ausencia de mencao nao e evidencia de que a taxa seja $0.

3. https://www.uc.edu/about/bursar/tuition-fees/special-course-fees-.html
   Nao lista nenhum curso com prefixo BANA (Business Analytics). Sem
   impacto em campos desta linha.

4. https://www.uc.edu/about/bursar/tuition-fees/student-health-insurance-.html
   Confirma com mais precisao o UC Student Health Insurance (Aetna)
   2026-27: Fall $1.107,90 + Spring $1.571,17 = $2.679,07 (o valor
   cadastrado, 2679, ja era o arredondamento correto; atualizado para a
   cifra exata).

CALCULOS DERIVADOS (a partir da tabela oficial #1, ano de 9 meses = duracao
exata do programa, entao nao ha necessidade de pro-ratear por duracao):
  - Custo Vida Total (programa) = Housing+Food + Books + Transportation(non-resident) + Misc
                                 = 17.526 + 1.300 + 3.038 + 4.892 = 26.756
    (Seguro Saude NAO entra aqui -- ja e campo proprio separado, para nao
    duplicar contagem.)
  - Transporte/mes = 3.038 / 9 = 337,56 (arredondado)
  - Outros Gastos/mes = 4.892 / 9 = 543,56 (arredondado; corresponde a "Miscellaneous")

NAO PREENCHIDOS (mantidos "Verificar"): Aluguel Privativo/mes e Mercado/mes
-- a fonte oficial junta os dois em "Housing and Food" sem separar; separar
exigiria estimar uma proporcao arbitraria, o que nao e uma cifra fonteada
(mesmo criterio ja aplicado a outras linhas do dataset -- UC Davis, UMN,
Utah -- ver etl_enrich_cincinnati.py). Seguro Residencial, Utilities,
Celular e WiFi tambem nao aparecem nesta tabela de COA (tipicamente ja
embutidos em "Housing and Food" ou "Miscellaneous" nas tabelas de
financial aid dos EUA) -- permanecem "Verificar".
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

HOUSING_FOOD = 17526
BOOKS = 1300
TRANSPORT_NON_RESIDENT = 3038
MISC = 4892

CUSTO_VIDA_TOTAL = HOUSING_FOOD + BOOKS + TRANSPORT_NON_RESIDENT + MISC  # 26756
TRANSPORTE_MES = round(TRANSPORT_NON_RESIDENT / 9, 2)  # 337.56
OUTROS_GASTOS_MES = round(MISC / 9, 2)  # 543.56

# Apenas os campos que estao sendo completados/corrigidos.
UPDATES = {
    "Custo Vida Total (programa) USD": CUSTO_VIDA_TOTAL,
    "Transporte/mês USD": TRANSPORTE_MES,
    "Outros Gastos/mês USD": OUTROS_GASTOS_MES,
    "Seguro Saúde (programa) USD": 2679.07,  # precisao: 1107.90 (Fall) + 1571.17 (Spring)
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
