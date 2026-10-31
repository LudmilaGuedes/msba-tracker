"""
Data enrichment: atualiza APENAS a linha 'UT Dallas' (Naveen Jindal School
of Management, Richardson, TX) no msba_programs_database.xlsx.

Foco exclusivo: MS in Business Analytics and AI, trilha presencial/Flex
(distinta das trilhas 'Cohort' de 36 creditos e 'Online Cohort', que sao
programas com estrutura diferente e NAO foram usados para nao misturar
dados de tracks distintos).

Os valores de tuition e Total Creditos (30 creditos, ja existentes na
planilha e internamente consistentes com os totais ja calculados) NAO foram
alterados: a nova pesquisa encontrou uma tarifa de $1.390... na verdade uma
estrutura de tuition em blocos por faixa de creditos (nao um valor fixo por
credito) especifica da trilha 'Cohort' (36 creditos), que e uma trilha
DIFERENTE da que ja estava documentada -- substituir arriscaria misturar
dados de duas ofertas distintas do mesmo programa.

Fontes principais (pesquisadas manualmente, fora deste script):
- https://sites.utdallas.edu/infosystems/ms-business-analytics/ (trilha Flex:
  24 meses padrao mas autoritmada, full-time ou part-time, entrada em
  Fall/Spring/Summer)
- https://jindal.utdallas.edu/admission-requirements/masters/ e
  https://jindal.utdallas.edu/ms-business-analytics-faq/ (GPA medio 3.3,
  GMAT/GRE opcional, pre-requisito de calculo, 3 cartas de recomendacao,
  taxa de inscricao $75)
- Busca por rankings: os unicos rankings encontrados (US News Online, Fortune
  Best Online) sao especificos do programa ONLINE, entao NAO foram usados
  para o campo Ranking US News desta linha (que e sobre a trilha presencial)
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
    "Nome da Universidade": "UT Dallas",
    "Estado": "TX",
    "Cidade": "Richardson",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Duração (meses)": "~24 (trilha Flex, autoritmada; full-time pode completar mais rápido)",
    "Início": "Fall, Spring, Summer",  # CORRECAO: Flex permite entrada nos 3 semestres, nao so Fall/Spring
    "GPA Médio Admitidos": 3.3,
    "Taxa Aplicação USD": 75,
    "Matérias Pré-Requisito": (
        "Conhecimento de cálculo; pode ser suprido cursando OPRE 6303 "
        "(Quantitative Foundations of Business) se não atendido previamente"
    ),
    "Documentos Necessários": (
        "3 cartas de recomendação; teste de proficiência em inglês com no "
        "máximo 2 anos da data da candidatura (para internacionais)"
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
        raise ValueError("Nenhuma linha encontrada para UT Dallas (Richardson, TX). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (UT Dallas, Richardson, TX)")

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
