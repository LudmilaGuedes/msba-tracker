"""
Data enrichment: atualiza APENAS a linha 'U of Central Oklahoma' (College
of Business, Edmond, OK) no msba_programs_database.xlsx.

Foco exclusivo: MS in Business Analytics, presencial, STEM.

Tuition NAO foi alterada: um resultado isolado citava "$19.088 para o
primeiro ano de aluno internacional", mas o valor ja existente na planilha
($13.344 = 32 x $417/credito) e internamente consistente e mais
especificamente sourced -- o valor de $19.088 provavelmente inclui custo de
vida/outras taxas, nao apenas tuition, entao nao foi usado para substituir
um dado ja limpo e coerente.

Fontes principais (pesquisadas manualmente, fora deste script):
- busca por duracao (confirmado: 2 anos / "as little as two years")
- busca por admissao (GRE/GMAT opcional, GPA 3.0 ja confirmado, documentos:
  resume + personal statement de 500 palavras)
- busca por bolsas para internacionais na UCO (UCO International Achievement
  Award, Distinguished Graduate Scholar ~$3.000/ano, Graduate Merit Global
  Scholar ~$2.000/ano por ate 2 anos -- bolsas gerais da universidade, nao
  especificas do MSBA, mas abertas a qualquer programa de pos-graduacao)
- busca por custo de vida em Edmond, OK (aluguel medio ~$1.078/mes, 2026)

Nao encontrado mesmo apos pesquisa profunda (mantido "Verificar"): taxa de
empregabilidade, salario medio, ranking, tamanho de turma exato e
assistantships especificos do MSBA (o programa e descrito qualitativamente
como "turmas pequenas", mas sem numero divulgado).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.stdout.reconfigure(encoding="utf-8")

from db_retention import rotate_backup  # noqa: E402

DB_PATH = "data/msba_programs_database.xlsx"
SHEET_DATA = "MSBA Programs"

VERIFICAR = "Verificar"

MATCH = {
    "Nome da Universidade": "U of Central Oklahoma",
    "Estado": "OK",
    "Cidade": "Edmond",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Duração (meses)": 24,
    "GMAT/GRE Obrigatório": "Opcional (GRE/GMAT não obrigatório para admissão)",
    "Documentos Necessários": (
        "Application completa, resume, personal statement (500 palavras); "
        "internacionais: comprovantes de qualificação + teste de inglês"
    ),
    "Bolsas Disponíveis": "Sim",
    "Tipos de Bolsa": (
        "UCO International Achievement Award (mérito); Distinguished "
        "Graduate Scholar (~$3.000/ano); Graduate Merit Global Scholar "
        "(~$2.000/ano, até 2 anos)"
    ),
    "Quão Realista (Intl)": (
        "Moderada (aplicar cedo aumenta chances; ~$700 mil/ano distribuídos "
        "a internacionais na UCO como um todo)"
    ),
    "Aluguel Privativo/mês USD": 1078,  # mercado de aluguel de Edmond, OK, 2026
}


def main():
    rotate_backup()
    print("Backup da versão anterior salvo em: data/msba_programs_database_backup.xlsx")

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
        raise ValueError("Nenhuma linha encontrada para U of Central Oklahoma (Edmond, OK). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (U of Central Oklahoma, Edmond, OK)")

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
