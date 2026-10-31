"""
Data enrichment: atualiza APENAS a linha 'U of Houston' (C.T. Bauer College
of Business, Houston, TX) no msba_programs_database.xlsx.

Foco exclusivo: MS Business Analytics (on-campus), presencial, STEM.

Fontes principais (pesquisadas manualmente, fora deste script):
- https://www.bauer.uh.edu/ms-business-analytics/cost.php (tabela oficial de
  tuition on-campus com 4 faixas: residente TX, não-residente, internacional
  com tuition "residente" e internacional padrão)
- https://www.bauer.uh.edu/ms-business-analytics/faqs.php (documentos,
  deadlines, pré-requisitos)
- busca por admissão (TOEFL 100/IELTS 6.5, GRE/GMAT com dispensa condicional,
  bolsas via aplicação única, Bauer Excellence Scholarship)
- Custo de vida em Houston reaproveitado da auditoria anterior de Texas A&M
  (mesma cidade): aluguel médio de 1-quarto ~$1.199/mês, 2026

Nao encontrado mesmo apos pesquisa profunda (mantido "Verificar"): taxa de
empregabilidade, salario medio e ranking especificos do MSBA (so encontrei
dados do Full-Time MBA da Bauer, que NAO foram usados por serem de um
programa diferente).
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
    "Nome da Universidade": "U of Houston",
    "Estado": "TX",
    "Cidade": "Houston",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Duração (meses)": "18-24 (full-time ou part-time; 36 créditos)",
    "Custo/Crédito (Intl) USD": 1680.56,  # 60500 / 36 (tarifa padrão internacional)
    "Custo/Crédito (Residente) USD": 1041.67,  # 37500 / 36
    "Tuition Total (Intl) USD": 60500,
    "Tuition Total (Residente) USD": 37500,
    "Caminho Tuition Residente": (
        "Sim — internacionais podem pagar tarifa 'residente' de $42.500 "
        "(em vez de $60.500 padrão) sob certas condições"
    ),
    "GMAT/GRE Obrigatório": (
        "Obrigatório, salvo dispensa (para quem demonstra conhecimento "
        "quantitativo forte ou experiência profissional relevante)"
    ),
    "Testes Inglês Aceitos": "TOEFL 100 (iBT) / IELTS 6.5",
    "Matérias Pré-Requisito": (
        "Nenhum pré-requisito formal (familiaridade com R/Python é "
        "recomendada, não obrigatória)"
    ),
    "Documentos Necessários": (
        "Resume, goal statement, transcripts oficiais, GRE/GMAT (ou "
        "dispensa), 3 cartas de recomendação"
    ),
    "Deadline Aplicação": (
        "Intl: 1 out (Spring)/1 mai (Fall); Doméstico: 1 dez (Spring)/1 jul "
        "(Fall) — admissão contínua, aplicar cedo ajuda na bolsa"
    ),
    "Bolsas Disponíveis": "Sim",
    "Tipos de Bolsa": (
        "Aplicação única para bolsas (bauer.uh.edu/invest/scholarships); "
        "ex.: Bauer Excellence Scholarship (~$5.000, reduz tuition a 50% da "
        "taxa residente)"
    ),
    "Quão Realista (Intl)": "Moderada (aplicar cedo aumenta chances; bolsas não são garantidas)",
    "Aluguel Privativo/mês USD": 1199,  # mercado de aluguel de Houston, TX, 2026
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
        raise ValueError("Nenhuma linha encontrada para U of Houston (Houston, TX). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (U of Houston, Houston, TX)")

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
