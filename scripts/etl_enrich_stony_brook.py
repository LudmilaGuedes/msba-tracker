"""
Data enrichment: atualiza APENAS a linha da 'Stony Brook SUNY' (College
of Business, Long Island, NY) no msba_programs_database.xlsx.

Foco exclusivo: MS in Business Analytics & Intelligence (MSBAI),
presencial (formato híbrido parcial), STEM, 30 créditos (10
disciplinas). Deep research em: custo total real de tuition para
internacionais, bolsas/assistantships, requisitos de GPA/GMAT/GRE/
inglês, e custo de vida em Long Island, NY.

NOTA IMPORTANTE: o MSBAI é um programa NOVO, com lançamento anunciado
para Fall 2026 -- ou seja, esta é a PRIMEIRA turma do programa. Isso
explica a ausência quase total de dados de outcomes proprios
(ranking, taxa de emprego, salario, empregadores, tamanho de turma):
o programa ainda nao formou nenhum aluno. Nenhum desses campos foi
preenchido com dados de outros programas do College of Business para
evitar atribuir estatisticas de um programa diferente a este.

TUITION (já estava correto, apenas verificado e complementado): a
página oficial de FAQs do College of Business confirma exatamente os
valores já presentes na planilha -- $471/crédito (residente) / $1.097/
crédito (não-residente/internacional). PREENCHIDOS os campos
'(Residente)' que estavam 'Verificar': $471/crédito e $14.130 total (30
créditos).

'Seguro Saúde (programa)' preenchido com $2.403 -- soma oficial e
CORRENTE (ciclo Fall 2026 + Spring/Summer 2027) do International
Student Health Insurance Plan (SHIP) de Stony Brook: $1.006,73 (Fall
2026) + $1.396,28 (Spring/Summer 2027).

Bolsas/Assistantships: a Graduate Tuition Scholarship (GTS) permite que
estudantes com stipend recebam cobertura parcial ou total de tuition à
taxa de RESIDENTE -- preenchido em 'Graduate Assistantships' e 'Caminho
Tuition Residente'. Assistantships gerais existem na universidade, mas
a disponibilidade especifica no College of Business não é garantida
(a própria fonte oficial recomenda contato direto) -- 'TA/RA
Disponível' preenchido refletindo essa incerteza, sem superestimar.

Fontes principais (pesquisadas via busca web, ago/2026):
- https://www.stonybrook.edu/business/graduate/faqs.html (tuition
  oficial confirmando os valores já presentes, GTS, deadlines de
  aplicação por semestre, documentos exigidos, GMAT/GRE opcional
  recomendado se GPA<3.2)
- Busca web: programa lançando em Fall 2026 (STEM, 30 créditos/10
  disciplinas, calculus MAT 123 ou equivalente com nota C+ como
  pré-requisito -- já confirmado e consistente com o que já estava na
  planilha)
- https://www.stonybrook.edu/commcms/studentaffairs/shs/insurance/
  international.php (SHIP internacional, tarifas oficiais correntes
  Fall 2026/Spring-Summer 2027)
- Busca web: aluguel de quarto privativo em Stony Brook, NY (faixa
  $1.200-2.000, região com custo de vida ~71% acima da média nacional)
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
    "Nome da Universidade": "Stony Brook SUNY",
    "Estado": "NY",
    "Cidade": "Long Island",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Residente) USD": 471,  # oficial (College of Business FAQ), confirma o valor já presente do Intl
    "Tuition Total (Residente) USD": 14130,  # 30 créditos x $471
    "Seguro Saúde (programa) USD": 2403,  # oficial, corrente: SHIP internacional Fall 2026 + Spring/Summer 2027
    "Aluguel Privativo/mês USD": 1400,  # quarto privativo, Stony Brook/Long Island (faixa $1.200-2.000 encontrada)
    "GPA Mínimo": 3.0,
    "Deadline Aplicação": (
        "Fall 2026: Doméstico 15 jul, Internacional 1º mai; Spring "
        "2026: Doméstico 15 jan, Internacional 1º dez; Summer 2026: "
        "Doméstico 15 mar"
    ),
    "Rounds Aplicação": "Admissão contínua (rolling); sem rounds numerados",
    "Documentos Necessários": (
        "Diploma de bacharelado, resume, personal statement, TOEFL/"
        "IELTS (maioria dos internacionais), GMAT/GRE opcional "
        "(recomendado se GPA < 3.2)"
    ),
    "TA/RA Disponível": (
        "Incerto/limitado — assistantships existem na universidade em "
        "geral, mas disponibilidade específica no College of Business "
        "não é garantida (fonte oficial recomenda contato direto)"
    ),
    "Graduate Assistantships": (
        "Graduate Tuition Scholarship (GTS): estudantes com stipend "
        "podem receber cobertura parcial ou total de tuition à taxa "
        "de residente"
    ),
    "Caminho Tuition Residente": "Via GTS (Graduate Tuition Scholarship) para quem recebe stipend — cobertura à taxa de residente",
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
        raise ValueError("Nenhuma linha encontrada para Stony Brook SUNY (Long Island, NY). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Stony Brook SUNY, Long Island, NY)")

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
