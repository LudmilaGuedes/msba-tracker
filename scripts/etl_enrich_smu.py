"""
Data enrichment: atualiza APENAS a linha 'SMU' (Cox School of Business,
Dallas, TX) no msba_programs_database.xlsx.

Foco exclusivo: MS in Business Analytics, presencial, STEM.

Fontes principais (pesquisadas manualmente, fora deste script):
- EMPLOYMENT REPORT MS Business Analytics (Class of 2025): salário médio
  $79.800 (ja confirmado na planilha, sem alteracao) e mais de 96% dos
  formandos recentes receberam propostas em 6 meses (atualiza a Taxa de
  Empregabilidade antiga de 82%)
- https://www.smu.edu/cox/business-degrees/masters-programs/ms-business-analytics/admissions
  (documentos, GMAT/GRE e TOEFL/IELTS opcionais, taxa de aceitacao 46%, sem
  taxa de inscricao)
- https://www.smu.edu/cox/business-degrees/masters-programs/ms-business-analytics/admissions/deadlines
  (5 rounds de aplicacao com datas exatas)
- busca por custo de vida em Dallas, TX (studio ~$900-1.300/mes, usado como
  proxy de quarto privativo; utilities ~$225/mes, 2026)

Tuition NAO foi alterada: o valor ja existente ($45.936 + $9.700 de taxas)
bate exatamente com o confirmado para a coorte 2025-2026; nao foi possivel
confirmar um valor diferente e mais atual para 2026-2027, entao mantido.
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
    "Nome da Universidade": "SMU",
    "Estado": "TX",
    "Cidade": "Dallas",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Taxa Empregabilidade %": 96,  # atualização (planilha tinha 82; fonte mais recente confirma 96%+ em 6 meses)
    "Taxa Aceitação %": 46,
    "Taxa Aplicação USD": 0,  # sem taxa de inscrição
    "Testes Inglês Aceitos": (
        "TOEFL/IELTS (opcional, não obrigatório para internacionais; "
        "validade de 2 anos)"
    ),  # CORRECAO: antes implicava exigência
    "Documentos Necessários": (
        "Resume, ensaios, transcripts não oficiais, 2 referências (contato, "
        "não carta de recomendação), GMAT/GRE opcional (validade de 5 anos)"
    ),
    "Entrevista": "Faz parte do processo de avaliação (não listada como obrigatória formalmente)",
    "Deadline Aplicação": (
        "5 rounds: 13 out/2026, 12 jan/2027 (prioritários para bolsa), "
        "9 mar/2027, 20 abr/2027, 25 mai/2027; admissão contínua após isso"
    ),
    "Rounds Aplicação": "5 rounds (ver Deadline Aplicação)",
    "Bolsas Disponíveis": "Sim",
    "Tipos de Bolsa": (
        "Merit-based (número limitado, consideração automática, baseada em "
        "desempenho acadêmico e GMAT/GRE); bolsa garantida no programa 4+1 "
        "para graduados da própria Cox"
    ),
    "Aluguel Privativo/mês USD": 1100,  # studio em Dallas, TX, 2026 (proxy para quarto privativo)
    "Utilities/mês USD": 225,
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
        raise ValueError("Nenhuma linha encontrada para SMU (Dallas, TX). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (SMU, Dallas, TX)")

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
