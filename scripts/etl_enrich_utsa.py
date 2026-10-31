"""
Data enrichment: atualiza APENAS a linha 'UTSA' (Carlos Alvarez College of
Business, San Antonio, TX) no msba_programs_database.xlsx.

Foco exclusivo: MS in Business Analytics, presencial/hibrido, STEM.

Tuition NAO foi alterada: a unica tarifa clara que a pesquisa encontrou
($775/credito) e explicitamente do formato 100% ONLINE ("no additional
fees, no out-of-state tuition"), o que sugere ser uma estrutura de preco
diferente do programa presencial que esta linha documenta. Como nao foi
possivel confirmar a tarifa correta para a trilha presencial/hibrida sem
arriscar misturar dados do programa online, os campos de tuition
permaneceram "Verificar".

Emprego/salario tambem NAO foram preenchidos: os unicos numeros encontrados
vinham de uma fonte intitulada "UTSA (Alvarez) Employment Report | MBA
Guidance", que aparenta ser especifica do programa de MBA, nao do MSBA --
por isso nao foram usados, para nao contaminar esta linha com dados de um
programa diferente.

Fontes principais (pesquisadas manualmente, fora deste script):
- busca por duracao (11 ou 16 meses full-time cohort, ou part-time flexivel)
- busca por admissao (GMAT/GRE nao obrigatorio, bolsa competitiva de $1.000
  para nao-residentes, GTA/GRA competitivos)
- busca por deadlines gerais da pos-graduacao UTSA (Fall 1 mai, Spring 1 nov)
- busca por custo de vida em San Antonio, TX (aluguel medio de apartamento
  estudantil ~$1.072-$1.320/mes, 2026)
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
    "Nome da Universidade": "UTSA",
    "Estado": "TX",
    "Cidade": "San Antonio",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Duração (meses)": "11 (full-time) ou 16 (cohort estendido); também há opção part-time flexível",
    "GMAT/GRE Obrigatório": "Não obrigatório",
    "Bolsas Disponíveis": "Sim",
    "Tipos de Bolsa": (
        "Graduate Competitive Scholarship de $1.000 para alunos "
        "não-residentes (admitidos ou já matriculados)"
    ),
    "TA/RA Disponível": "Sim (GTA/GRA competitivos, conforme necessidade do departamento)",
    "Graduate Assistantships": (
        "GTA/GRA disponíveis em base competitiva; posições de correção de "
        "provas também disponíveis"
    ),
    "Deadline Aplicação": "Fall: 1 mai; Spring: 1 nov",
    "Aluguel Privativo/mês USD": 1150,  # mercado de aluguel de San Antonio, TX, 2026
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
        raise ValueError("Nenhuma linha encontrada para UTSA (San Antonio, TX). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (UTSA, San Antonio, TX)")

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
