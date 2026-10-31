"""
Data enrichment: atualiza APENAS a linha 'UCF' (Barry S. Miller College of
Business, Orlando, FL) no msba_programs_database.xlsx.

ATENCAO: existe uma segunda linha, 'UCF (Professional)' (Professional MS in
Management/BA), que e um programa DIFERENTE e NAO deve ser tocada. O
matching abaixo exige Nome da Universidade == 'UCF' (exato) E Estado == 'FL'
E Cidade == 'Orlando'.

Foco exclusivo: Master of Science in Business Analytics, presencial, STEM,
12 meses / 3 termos.

Fontes principais (pesquisadas manualmente, fora deste script):
- https://business.ucf.edu/degree/ms-business-analytics/ e https://www.ucf.edu/degree/business-analytics-ms/
  (duracao, creditos, tuition por credito, GPA, GMAT/GRE nao obrigatorio, TOEFL/IELTS/Duolingo,
  documentos, deadlines, taxa de inscricao)
- https://graduate.ucf.edu/graduate-guide/assistantships (politica de GTA: 0,5 FTE,
  stipend + waiver de tuition + seguro saude)
- https://www.ucf.edu/financial-aid/cost/ (orcamento oficial de custo de atendimento
  2026-2027 para pos-graduandos, cenario off-campus, usado como proxy para os
  campos de custo de vida do MSBA, que nao tem orcamento proprio publicado)

Nao encontrado mesmo apos pesquisa profunda (mantido "Verificar"): taxa de
empregabilidade, salario medio pos-programa e ranking especificos do MSBA
(so encontrei dados do MS em Data Analytics, um programa diferente da UCF,
que NAO foi usado para nao contaminar esta linha com dados de outro curso).
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
    "Nome da Universidade": "UCF",
    "Estado": "FL",
    "Cidade": "Orlando",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Duração (meses)": 12,
    "Custo/Crédito (Residente) USD": 369.65,  # refinamento de precisão (planilha tinha 370, arredondado)
    "Custo/Crédito (Intl) USD": 1194.05,  # refinamento de precisão (planilha tinha 1194)
    "Tuition Total (Residente) USD": 11090,  # 30 x 369,65
    "Tuition Total (Intl) USD": 35822,  # 30 x 1194,05
    "Aluguel Privativo/mês USD": 1507,  # orçamento oficial UCF 2026-27, off-campus (13566/9 meses)
    "Mercado/mês USD": 686,  # 6170/9
    "Transporte/mês USD": 122,  # 1094/9
    "Outros Gastos/mês USD": 478,  # pessoal (3104/9=345) + livros (1200/9=133), sem coluna própria p/ livros
    "Custo Vida Total (programa) USD": 33516,  # (1507+686+122+478) x 12 meses de programa
    "GPA Mínimo": 3.0,
    "GMAT/GRE Obrigatório": "Não obrigatório",
    "Testes Inglês Aceitos": "TOEFL iBT 91 mínimo / IELTS 7.0 / Duolingo 125",
    "Matérias Pré-Requisito": (
        "Bacharelado em área quantitativa (negócios, economia, ciência da "
        "computação, engenharia, matemática, estatística ou ciências naturais)"
    ),
    "Documentos Necessários": (
        "Resume, goal statement (500-1.000 palavras), 3 cartas de "
        "recomendação; internacionais: avaliação de credencial via WES ou "
        "Joseph Silny and Associates"
    ),
    "Deadline Aplicação": "Somente admissão Fall; internacional: 15 jan; final (doméstico): 1 jul",
    "Taxa Aplicação USD": 31,
    "Processo TA/RA": (
        "GTA de 0,5 FTE (20h/semana), até 2 disciplinas; inclui stipend + "
        "waiver de tuition (9 créditos) + seguro saúde quando ≥ 0,5 FTE"
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
        raise ValueError("Nenhuma linha encontrada para UCF (Orlando, FL). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (UCF, Orlando, FL) -- NAO 'UCF (Professional)'")

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
