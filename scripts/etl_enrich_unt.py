"""
Data enrichment: atualiza APENAS a linha da 'University of North Texas
(UNT)' (G. Brint Ryan College of Business, Denton, TX) no
msba_programs_database.xlsx.

Foco: Master of Science in Business Analytics, trilha ON-CAMPUS
(distinta da versão 100% online de 33 créditos da mesma escola --
ambas administradas pelo Dept. de Information Technology and Decision
Sciences/ITDS), STEM (até 12 meses de CPT + 36 meses de OPT).

'Total Créditos' CORRIGIDO de 36 para 33 -- a página oficial atual do
programa confirma "33 credit hours (18 core + 15 electives)", não 36.
Duração oficial "18-24 months" é consistente com o valor já presente
(24, extremo superior da faixa), não alterado.

TUITION: PREENCHIDO (estava 'Verificar' em todos os campos). A página
oficial cita um custo APROXIMADO assumindo ritmo de 9 créditos/
semestre para completar o programa: In-State $19.628 / Out-of-State
$37.488 -- tratado como o custo TOTAL aproximado do programa completo
(33 créditos), já que a página rotula os valores como referentes à
conclusão nesse ritmo, não como um valor por semestre isolado.
PREENCHIDO 'Tuition Total (Residente)' = $19.628; '(Intl)' = $37.488
(usando a tarifa out-of-state). 'Custo/Crédito' derivado dividindo
pelos 33 créditos: $594,79 (residente) / $1.136,00 (internacional).
RESSALVA: a própria fonte rotula esses valores como "approximate", não
uma tarifa fechada oficial -- documentado como estimativa razoável, não
certeza absoluta.

'Salário Médio Pós USD' preenchido: $86.000 -- salário médio oficial e
específico do programa MSBA, medido em até 90 dias pós-formatura.
'Principais Empregadores' preenchido com a lista oficial: Texas
Instruments, Toyota, PepsiCo, IBM, Hitachi.

'Taxa Aplicação USD' preenchido: $75 (não-reembolsável, taxa geral da
Graduate School, aplicável a todos os programas incluindo o MSBA).

CUIDADO APLICADO: a página do programa não fornece uma taxa de
empregabilidade em percentual (apenas o salário médio), então 'Taxa
Empregabilidade %' permanece 'Verificar' -- não foi inferida a partir
do dado de salário para evitar inventar uma estatística não fornecida.

CUIDADO APLICADO (dados NÃO encontrados, permanecem 'Verificar'):
'Fees Adicionais USD' (o valor de tuition da fonte já pode incluir
fees, não especificado com clareza suficiente para separar),
'Tamanho Turma', 'GPA Médio Admitidos', 'GMAT Médio', 'Taxa Aceitação
%', 'Experiência Profissional', 'Matérias Pré-Requisito', 'Rounds
Aplicação' (a planilha já tinha 3 deadlines por período, sem indicação
de rounds numerados adicionais), 'Entrevista', 'Dificuldade Entrada',
'Ranking US News', 'Ranking QS', 'Bolsas Disponíveis'/'Tipos de
Bolsa'/'TA/RA Disponível'/'Processo TA/RA'/'Graduate Assistantships'/
'Caminho Tuition Residente' (a página oficial do programa não detalha
nenhum desses tópicos, apenas direciona para contato direto), e a
maior parte dos campos mensais de custo de vida (a planilha já tinha
Transporte, Outros Gastos e Custo Vida Total preenchidos de uma fonte
anterior a esta sessão, mantidos sem alteração).

Fontes principais (pesquisadas via WebFetch, ago/2026):
- https://cob.unt.edu/masters/programs/ms-business-analytics.html
  (créditos corretos, duração, currículo, tuition aproximada por
  residência, salário médio oficial $86k, empregadores)
- https://www.unt.edu/admissions/graduate/steps-to-apply.html (taxa
  de aplicação $75, documentos gerais da Graduate School)
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
    "Nome da Universidade": "University of North Texas (UNT)",
    "Estado": "TX",
    "Cidade": "Denton",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Total Créditos": 33,  # CORREÇÃO: oficial (18 core + 15 eletivas), não 36
    "Custo/Crédito (Residente) USD": 594.79,  # 19.628 / 33, estimativa oficial
    "Custo/Crédito (Intl) USD": 1136.00,  # 37.488 / 33, estimativa oficial
    "Tuition Total (Residente) USD": 19628,  # estimativa oficial (ritmo de 9 créditos/semestre)
    "Tuition Total (Intl) USD": 37488,  # estimativa oficial (ritmo de 9 créditos/semestre)
    "Salário Médio Pós USD": 86000,  # oficial, específico do MSBA, até 90 dias pós-formatura
    "Principais Empregadores": "Texas Instruments, Toyota, PepsiCo, IBM, Hitachi",
    "Taxa Aplicação USD": 75,  # não-reembolsável, taxa geral da Graduate School
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
        raise ValueError("Nenhuma linha encontrada para University of North Texas (UNT) (Denton, TX). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (University of North Texas (UNT), Denton, TX)")

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
