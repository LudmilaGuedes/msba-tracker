"""
Data enrichment: atualiza APENAS a linha 'Baylor' (Hankamer School of
Business, Waco, TX) no msba_programs_database.xlsx.

Foco exclusivo: MS in Business Analytics (Residential), presencial, STEM.

Confirmado (nao alterado): a bolsa automatica de 65% de tuition para todos
os admitidos, que ja constava na planilha, foi verificada como precisa em
fonte independente.

Fontes principais (pesquisadas manualmente, fora deste script):
- busca por tuition (specialty master's rate: $2.651/credito, tarifa unica
  privada aplicavel a todos os alunos)
- https://hankamer.baylor.edu/msba/admissions/deadlines (datas exatas por rodada)
- https://hankamer.baylor.edu/msba/admissions (sem taxa de inscricao, documentos,
  pre-requisitos ja confirmados)
- busca por custo de vida em Waco (aluguel medio de 1-quarto ~$1.100/mes, 2026)

Nao encontrado mesmo apos pesquisa profunda (mantido "Verificar"): tamanho
de turma, taxa de empregabilidade, salario medio e ranking especificos do
MSBA (so encontrei taxa de colocacao geral da Baylor, 92%, que NAO foi usada
por nao ser especifica deste programa).
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
    "Nome da Universidade": "Baylor",
    "Estado": "TX",
    "Cidade": "Waco",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Intl) USD": 2651,  # privada, mesma tarifa para todos
    "Custo/Crédito (Residente) USD": 2651,
    "Tuition Total (Intl) USD": 98087,  # 37 x 2651 (valor de tabela, antes da bolsa de 65%)
    "Tuition Total (Residente) USD": 98087,
    "Taxa Aplicação USD": 0,  # sem taxa de inscrição
    "Deadline Aplicação": (
        "Fall: prioridade/bolsa 15 fev; admissão geral até 15 abr; final "
        "15 jun (EUA) / 15 abr (internacional)"
    ),
    "Documentos Necessários": (
        "Transcripts (não oficiais na inscrição, oficiais após admissão), "
        "resume, 3 cartas de recomendação; internacionais: TOEFL/IELTS/PTE"
    ),
    "Aluguel Privativo/mês USD": 1100,  # mercado de aluguel de Waco, 2026
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
        raise ValueError("Nenhuma linha encontrada para Baylor (Waco, TX). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Baylor, Waco, TX)")

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
