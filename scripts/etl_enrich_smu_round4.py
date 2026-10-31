"""
Data enrichment (RODADA 4): atualiza APENAS a linha 'SMU' (Cox School
of Business, Dallas, TX) no msba_programs_database.xlsx.

Usuário forneceu diretamente o link oficial do Bursar da SMU com a
tabela de tuition e fees 2026-27, apontando que os preços já presentes
na planilha (de rodadas anteriores) divergiam do valor oficial atual.
Verificado e CONFIRMADA a discrepância -- corrigido nesta rodada.

TUITION: CORRIGIDO. A tabela oficial do Bursar (jay2026-2027) para o
M.S. in Business Analytics mostra uma estrutura de cobrança que NÃO
existia nas rodadas anteriores desta linha: "1–11 hrs: $1.570 por
crédito" (tarifa part-time) vs. "12-18 hrs: $25.000 flat rate por
semestre" (tarifa full-time). O valor anterior de 'Custo/Crédito'
($1.392, derivado do total antigo ÷ 33 créditos) e 'Tuition Total'
($45.936) NÃO batem com nenhuma das duas tarifas oficiais atuais.

CORRIGIDO 'Custo/Crédito' (Intl e Residente, mesmo valor -- SMU é
privada, tuition flat sem distinção de residência) para $1.570,
usando a tarifa oficial por crédito divulgada.

CORRIGIDO 'Tuition Total' (Intl e Residente) para $50.000 --
ESTIMATIVA baseada na tarifa flat full-time ($25.000/semestre) × 2
semestres, assumindo o ritmo full-time padrão mais comum para
completar os 33 créditos do programa (a página oficial NÃO especifica
quantos semestres um aluno full-time tipicamente permanece nessa
faixa de cobrança, então esse número de semestres é uma suposição
razoável, documentada como tal, não uma confirmação direta).

CORRIGIDO 'Fees Adicionais USD' para $2.772 -- "Program Fee"
full-time oficial ($1.386/semestre) × 2 semestres. RESSALVA
IMPORTANTE: a página também menciona que "Mandatory Fees" adicionais
se aplicam a este programa, mas SEM especificar o valor na seção
consultada (referenciada como estando "listada separadamente na
página de mandatory fees da SMU") -- ou seja, o total real pode ser
AINDA MAIOR do que os $52.772 (tuition + program fee) calculados
aqui. Documentado para transparência; não foi possível quantificar
essas fees adicionais nesta rodada.

Fonte fornecida pelo usuário e verificada diretamente:
https://www.smu.edu/businessfinance/officeoffinanceandplanning/bursar/tuition_and_fees/graduate/jay2026-2027
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
    "Nome da Universidade": "SMU",
    "Estado": "TX",
    "Cidade": "Dallas",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Intl) USD": 1570,  # CORREÇÃO: oficial 2026-27 (era 1.392, derivado de total desatualizado)
    "Custo/Crédito (Residente) USD": 1570,  # privada, tarifa flat (igual ao Intl)
    "Tuition Total (Intl) USD": 50000,  # CORREÇÃO: $25.000/semestre (flat, full-time) x 2 semestres, estimativa (era 45.936)
    "Tuition Total (Residente) USD": 50000,  # privada, tarifa flat (igual ao Intl)
    "Fees Adicionais USD": 2772,  # CORREÇÃO: $1.386/semestre (Program Fee full-time) x 2 (era 9.700); "Mandatory Fees" adicionais não quantificadas
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
        raise ValueError("Nenhuma linha encontrada para SMU (Dallas, TX). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (SMU, Dallas, TX)")

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
