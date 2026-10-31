"""
Data enrichment: atualiza APENAS a linha 'Florida Institute of Technology
(Florida Tech)' (Nathan M. Bisk College of Business, Melbourne, FL) no
msba_programs_database.xlsx.

Foco exclusivo: Master of Science in Business Analytics, presencial (campus
de Melbourne, distinto da versao 100% online da mesma escola), STEM.

Fontes principais (pesquisadas manualmente, fora deste script):
- https://www.fit.edu/student-financial-services/tuition-and-fees/ (tarifa por
  credito 2026-2027: $1.390, atualizando o valor antigo de $1.310 de 2024-25)
- busca por admissao MS Business Analytics (GPA 3.0, GMAT/GRE nao obrigatorio,
  sem taxa de inscricao, literacia computacional como pre-requisito)
- https://www.fit.edu/admission/applying/international/demonstrating-english-proficiency/
  (TOEFL/IELTS)
- https://www.fit.edu/provost/academic-guidelines/guideline-grad-student-assist-employment-contract/graduate-assistantship-application-process/
  (processo de GSA/DGRATS)
- https://www.fit.edu/financialaid/graduate-students/cost-of-attendance/ (orcamento
  oficial de custo de vida 2025-2026 para pos-graduandos de Melbourne)

Custo de vida: Aluguel e Mercado ficaram "Verificar" porque a Florida Tech
publica 'Housing & Food' como uma unica cifra combinada, sem separar as
duas categorias -- dividir isso exigiria inventar uma proporcao, o que nao
foi feito. Transporte e Outros Gastos foram derivados normalmente. Custo
Vida Total tambem ficou "Verificar": a duracao exata do programa presencial
(em meses) nao foi confirmada apesar de busca extensa, entao nao ha uma base
confiavel para multiplicar os custos mensais.
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
    "Nome da Universidade": "Florida Institute of Technology (Florida Tech)",
    "Estado": "FL",
    "Cidade": "Melbourne",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Intl) USD": 1390,  # CORRECAO: tarifa atualizada 2026-27 (era 1310, de 2024-25)
    "Custo/Crédito (Residente) USD": 1390,  # privada, mesma tarifa para todos
    "Tuition Total (Intl) USD": 41700,  # 30 x 1390
    "Tuition Total (Residente) USD": 41700,
    "Transporte/mês USD": 167,  # orçamento oficial Florida Tech 2025-26, pós-graduandos Melbourne (1500/9 meses)
    "Outros Gastos/mês USD": 367,  # pessoal (2100/9=233) + livros (1200/9=133)
    "GPA Mínimo": 3.0,
    "GMAT/GRE Obrigatório": "Não obrigatório",
    "Testes Inglês Aceitos": (
        "TOEFL iBT 79-80 (testes antes de 1/jan/2026) ou 4.0 (a partir de "
        "1/jan/2026, conforme fontes oficiais-adjacentes); IELTS 6.5"
    ),
    "Taxa Aplicação USD": 0,  # sem taxa de inscrição
    "Documentos Necessários": (
        "Bacharelado de instituição credenciada, histórico acadêmico "
        "oficial, resume; literacia computacional (curso prévio, exame de "
        "proficiência ou curso complementar)"
    ),
    "Rounds Aplicação": "Admissão contínua (rolling)",
    "TA/RA Disponível": (
        "Possível (contatar o departamento; GSA/DGRATS para quem tem ≥15h/"
        "semana em pesquisa/bolsa, aberto a internacionais)"
    ),
    "Graduate Assistantships": (
        "Prioridade de aplicação: 15 jan (Fall) / 15 set (Spring); "
        "elegibilidade varia por departamento; DGRATS para pesquisa/bolsa "
        "≥15h/semana"
    ),
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
        raise ValueError("Nenhuma linha encontrada para Florida Institute of Technology. Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Florida Institute of Technology (Florida Tech), Melbourne, FL)")

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
