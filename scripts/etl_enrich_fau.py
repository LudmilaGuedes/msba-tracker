"""
Data enrichment: atualiza APENAS a linha da 'Florida Atlantic University (FAU)'
(Boca Raton, FL) no msba_programs_database.xlsx, com base em pesquisa web
focada no programa 'Master of Science in Business Analytics (Traditional)'
(presencial, STEM, CIP 30.7102) do College of Business.

Nao confundir com o 'Professional/Online' MSBA da mesma escola (formato
online/hibrido para profissionais) -- o programa correto para este schema
(presencial, publico F-1) e o Traditional, ja era o que constava na planilha
antes desta atualizacao.

Fontes principais (pesquisadas manualmente, fora deste script):
- https://business.fau.edu/masters-phd/business-analytics/traditional/program-information/
- https://business.fau.edu/masters-phd/business-analytics/traditional/admissions/
- https://business.fau.edu/masters-phd/admissions-mastersphd/tuition-financial-aid/
- https://www.fau.edu/controllers-office/student-services/tuition-breakdown/ (tarifas por
  credito 2026-2027, oficiais)
- https://www.fau.edu/finaid/other/cost-of-attendance/ (orcamento oficial de custo de vida
  2026-2027 para pos-graduandos, usado como proxy para os campos de custo de vida do MSBA,
  que nao tem orcamento proprio publicado)
"""

import shutil
import sys
from datetime import datetime

sys.stdout.reconfigure(encoding="utf-8")

import openpyxl
from openpyxl.styles import PatternFill

DB_PATH = "data/msba_programs_database.xlsx"
SHEET_DATA = "MSBA Programs"

VERIFICAR = "Verificar"

MATCH = {
    "Nome da Universidade": "Florida Atlantic University (FAU)",
    "Estado": "FL",
    "Cidade": "Boca Raton",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Início": "Fall, Spring, Summer",
    "Destaques Currículo": (
        "10 cursos de 3 créditos (30 no total), sem pré-requisitos de negócios, "
        "sem opção de tese; foco prático (não teórico): Excel, Tableau, Power BI, "
        "SQL, Python, SPSS, SAS, Google Analytics; estatística aplicada, "
        "forecasting, modelagem preditiva, visualização de dados"
    ),
    "Custo/Crédito (Intl) USD": 1125.06,
    "Custo/Crédito (Residente) USD": 371.82,
    "Tuition Total (Intl) USD": 33752,  # 30 x 1125.06, tarifa oficial 2026-27 (ja inclui taxas)
    "Tuition Total (Residente) USD": 11155,  # 30 x 371.82
    "Aluguel Privativo/mês USD": 1355,  # orcamento oficial FAU p/ pos-grad, off-campus, 2026-27 (12198/9)
    "Mercado/mês USD": 611,  # 5498/9
    "Transporte/mês USD": 396,  # 3562/9
    "Outros Gastos/mês USD": 680,  # Personal (4756/9=528) + Books (1368/9=152), sem coluna propria p/ livros
    "Custo Vida Total (programa) USD": 36504,  # (1355+611+396+680) x 12 meses de programa
    "Bolsas Disponíveis": "Sim",
    "Tipos de Bolsa": (
        "Fellowships do Graduate College (anuais); bolsas via Office of Student "
        "Financial Aid; Graduate Teaching Assistantships (GTA) no departamento ITOM"
    ),
    "Quão Realista (Intl)": (
        "Moderada (GTAs são competitivas, por semestre, 20h/semana; waiver de "
        "tuition total é improvável e não garantido, geralmente parcial)"
    ),
    "TA/RA Disponível": "Sim (Graduate Teaching Assistantship - GTA, departamento ITOM)",
    "Processo TA/RA": (
        "Seleção competitiva por semestre; indicar interesse na aplicação; "
        "contato com o Graduate Advising Office do College of Business"
    ),
    "Graduate Assistantships": (
        "Disponíveis (stipend + possível waiver parcial de tuition); waiver "
        "total é raro e não garantido"
    ),
    "GPA Mínimo": 3.0,
    "GMAT/GRE Obrigatório": (
        "Obrigatório (GRE mínimo 295 ou GMAT mínimo 495); dispensa se GPA ≥ 3.3 "
        "ou pós-graduação prévia"
    ),  # CORRECAO: planilha tinha "Não", mas o teste e exigido salvo dispensa
    "Testes Inglês Aceitos": "TOEFL 79 (iBT) / IELTS 6.0",
    "Deadline Aplicação": (
        "Fall: dom 1 jul / intl 15 fev; Spring: dom 1 nov / intl 15 jul; "
        "Summer: dom 1 abr / intl 15 jan"
    ),
    "Taxa Aplicação USD": 30,
    "Documentos Necessários": (
        "Histórico acadêmico, GRE ou GMAT (salvo dispensa), comprovante de "
        "imunização, avaliação de transcript por curso (NACES ou FAU para "
        "diplomas internacionais), TOEFL/IELTS se aplicável"
    ),
    "Entrevista": "Não exigida",
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
        raise ValueError("Nenhuma linha encontrada para Florida Atlantic University (FAU). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Florida Atlantic University (FAU), Boca Raton, FL)")

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
