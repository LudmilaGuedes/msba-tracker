"""
Data enrichment: atualiza APENAS a linha 'TCU' (M.J. Neeley School of
Business, Fort Worth, TX) no msba_programs_database.xlsx.

Foco exclusivo: MS in Business Analytics, presencial, STEM.

CORRECAO PRINCIPAL: a tuition ja existente ($1.794/credito, $56.100 total)
estava desatualizada. A pagina oficial de tuition para a coorte 2026-2027
mostra uma tarifa por semestre diferente (Summer 2026: $2.210/hora; Fall
2026/Spring 2027: $2.315/hora), com custo total de ~$71.663 (ja incluindo
$1.000/semestre de taxas adicionais, 3 semestres = $3.000).

Fontes principais (pesquisadas manualmente, fora deste script):
- https://neeley.tcu.edu/admissions/graduate/tuition-fees.php (tuition e
  creditos atualizados para a coorte 2026-2027)
- https://www.neeley.tcu.edu/Page/Requirements-and-Deadlines (deadlines,
  taxa de inscricao $100, GMAT/GRE opcional com dispensas, TOEFL/IELTS/PTE/
  Duolingo, pre-requisito de algebra/QBA, 3 referencias sem carta formal)

Nao encontrado mesmo apos pesquisa profunda (mantido "Verificar"): taxa de
empregabilidade, salario medio, tamanho de turma e ranking especificos do
MSBA (so encontrei dados do programa de graduacao/BBA da TCU, que NAO foram
usados por serem de um programa diferente).
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
    "Nome da Universidade": "TCU",
    "Estado": "TX",
    "Cidade": "Fort Worth",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Total Créditos": 30,
    "Custo/Crédito (Intl) USD": 2289,  # CORRECAO: media ponderada da tarifa 2026-27 (antes: 1794, desatualizado)
    "Custo/Crédito (Residente) USD": 2289,  # privada, mesma tarifa para todos
    "Tuition Total (Intl) USD": 71663,  # CORRECAO: valor oficial 2026-27, ja incluindo $3.000 de taxas
    "Tuition Total (Residente) USD": 71663,
    "Fees Adicionais USD": 3000,  # $1.000/semestre x 3 semestres, ja incluso no total acima
    "GMAT/GRE Obrigatório": (
        "Opcional para admissão geral (dispensa possível para CFA charter "
        "holders ou pós-graduação prévia em área STEM/quantitativa); pode "
        "ser exigido para consideração de bolsa"
    ),
    "Testes Inglês Aceitos": (
        "TOEFL/IELTS/PTE/Duolingo (obrigatório para quem não é de país de "
        "língua oficial inglesa)"
    ),
    "Matérias Pré-Requisito": "Álgebra universitária ou métodos quantitativos de negócios (QBA), nota C ou superior",
    "Documentos Necessários": "Transcripts oficiais, 3 referências (contato, não carta de recomendação)",
    "Deadline Aplicação": "Prioritária (consideração para assistantship): 15 out; Final (Summer/Fall): 1 fev",
    "Taxa Aplicação USD": 100,
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
        raise ValueError("Nenhuma linha encontrada para TCU (Fort Worth, TX). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (TCU, Fort Worth, TX)")

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
