"""
Data enrichment: atualiza APENAS a linha da 'U of Louisville' (College of
Business, Louisville, KY) no msba_programs_database.xlsx.

Foco exclusivo: MS in Business Analytics (MSBA), presencial, STEM,
AACSB, 30 créditos, formato full-time acelerado de 13 meses (cohort-
based, 3 semestres). Deep research em: custo total real de tuition para
internacionais, bolsas/assistantships, requisitos de GPA/GMAT/GRE/
inglês, e custo de vida em Louisville, KY.

TUITION (já estava correto, apenas verificado e complementado): a página
oficial do programa confirma explicitamente "$34.500 bundled tuition and
fees – regardless of residency" -- ou seja, o valor já inclui TODAS as
taxas obrigatórias, sem distinção de residência. Isso confirma os
$34.500 já presentes tanto em 'Tuition Total (Intl)' quanto '(Residente)'
e explica por que 'Fees Adicionais USD' deve ser PREENCHIDO COM 0 (não
'Verificar') -- não existe uma fee adicional separada, ela já está
embutida no pacote. 'Custo/Crédito' (estava 'Verificar' em ambos)
PREENCHIDO dividindo $34.500 por 30 créditos = $1.150/crédito (igual
para residente e internacional).

Custo de vida: 'Transporte/mês' preenchido com $0 -- TARC (sistema de
transporte público de Louisville) é gratuito para alunos da UofL
mediante Cardinal Card/UofL ID (desde abr/2025, via app Token Transit).
'Seguro Saúde (programa)' preenchido com $3.290 -- soma oficial dos
períodos Fall ($1.645) + Spring/Summer ($1.645) do plano internacional
Anthem 2026-27 (cost sheet oficial da Academic Health Plans).
'Utilities/mês' permanece 'Verificar': as únicas fontes encontradas são
médias de domicílio inteiro (~$185-188/mês), sem uma forma confiável de
estimar a fração de um quarto individual -- e por isso 'Custo Vida Total
(programa) USD' também permanece 'Verificar'.

Fontes principais (pesquisadas via busca web, ago/2026):
- https://business.louisville.edu/academics-programs/graduate-programs/
  msba/ (tuition bundled oficial, pré-requisito de Statistics Prep,
  bolsas merit + GMAT 680+ para bolsa adicional, concentração em IA)
- https://catalog.louisville.edu/graduate/programs-study/master-science-
  business-analytics/ (documentos exigidos, formato full-time 13 meses/
  part-time 20-24 meses, estrutura curricular)
- Busca web: TARC gratuito para alunos UofL (news.louisville.edu,
  louisvillecardinal.com)
- https://myahpcare.com/wp-content/uploads/U-of-L-Anthem-2026-27-
  PC_INTL_Approved.pdf (cost sheet oficial do seguro saúde internacional
  2026-27)
- Busca web (agregador, confiança moderada -- não confirmado em página
  oficial direta): taxa de emprego 90% em 3 meses e salário médio
  $72.000 para o MSBA
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

VERIFICAR = "Verificar"

MATCH = {
    "Nome da Universidade": "U of Louisville",
    "Estado": "KY",
    "Cidade": "Louisville",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Intl) USD": 1150,  # $34.500 (bundled, oficial) / 30 créditos
    "Custo/Crédito (Residente) USD": 1150,  # mesmo valor — "regardless of residency" (oficial)
    "Fees Adicionais USD": 0,  # oficial: "$34.500 bundled tuition and fees" — já incluídas no Tuition Total, não há fee separada
    "Seguro Saúde (programa) USD": 3290,  # oficial 2026-27: Fall $1.645 + Spring/Summer $1.645 (Anthem International)
    "Transporte/mês USD": 0,  # TARC gratuito para alunos UofL (Cardinal Card / app Token Transit)
    "Matérias Pré-Requisito": (
        "Ao menos 1 curso de estatística ou cálculo na graduação, ou "
        "competência quantitativa demonstrada; obrigatório completar o "
        "curso online gratuito e self-paced 'Statistics Prep' (~5 "
        "semanas) antes do início do programa"
    ),
    "Documentos Necessários": (
        "Application com taxa, histórico acadêmico oficial, resume, "
        "personal statement, 2 cartas de referência"
    ),
    "Processo TA/RA": "N/A — sem TA/RA dedicado ao programa",
    "Taxa Empregabilidade %": 90,  # fonte agregadora (não confirmada em página oficial direta)
    "Salário Médio Pós USD": 72000,  # fonte agregadora (não confirmada em página oficial direta)
}


def main():
    rotate_backup()
    print(f"Backup rotacionado (política fixa) em: data/msba_programs_database_backup.xlsx")

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
        raise ValueError("Nenhuma linha encontrada para U of Louisville (Louisville, KY). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (U of Louisville, Louisville, KY)")

    no_fill = PatternFill(fill_type=None)
    verificar_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    before = {}
    after = {}
    for header, new_value in UPDATES.items():
        col_idx = header_to_col[header]
        cell = ws.cell(row=row_idx, column=col_idx)
        before[header] = cell.value
        cell.value = new_value
        cell.fill = verificar_fill if new_value == VERIFICAR else no_fill
        after[header] = new_value

    wb.save(DB_PATH)

    print("\nCampos atualizados (Como Estava -> Como Ficou):")
    for header in UPDATES:
        print(f"  - {header}:")
        print(f"      Antes: {before[header]!r}")
        print(f"      Depois: {after[header]!r}")


if __name__ == "__main__":
    main()
