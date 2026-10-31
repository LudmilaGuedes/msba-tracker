"""
Data enrichment (RODADA 2): atualiza APENAS a linha 'Oakland
University' (School of Business Administration, Rochester Hills, MI)
no msba_programs_database.xlsx.

Esta linha já passou por uma rodada de enrichment (script
etl_enrich_oakland.py), restando 30/54 campos 'Verificar'. Usuário
confirmou via AskUserQuestion (ago/2026) para prosseguir com uma
Rodada 2 focada nesses campos remanescentes.

'Destaques Currículo' preenchido: estrutura em 4 módulos --
Preparatory and Foundation Courses, Core Courses, Elective Courses, e
Capstone Project.

'Matérias Pré-Requisito' preenchido: cursos preparatórios/de
nivelamento exigidos para quem não tem background em information
systems e/ou estatística; os créditos desses cursos NÃO contam para o
diploma.

'Documentos Necessários' preenchido: aplicação online; CV/resume; 1
contato de referência (não uma carta completa, apenas o contato para
referência); transcripts oficiais; internacionais: avaliação de
transcript + scores de TOEFL (se aplicável).

'Processo TA/RA' preenchido: aplicações de Graduate Assistantship para
o semestre de Fall devem chegar até 1 de abril.

'Dificuldade Entrada' preenchido: Moderada -- processo seletivo que
considera histórico acadêmico/bolsa, capacidade de comunicação e
experiência profissional (a página oficial descreve o processo como
"selective" sem taxa de aceitação numérica).

CUIDADO APLICADO: mesmo após tentativas adicionais nesta rodada 2, os
seguintes campos permanecem 'Verificar' por falta de fonte acessível
ou dado não divulgado: 'Início' (não especificado explicitamente),
'Tamanho Turma', 'GPA Médio Admitidos', 'GMAT Médio', 'Taxa Aceitação
%' (a página confirma apenas que o processo é "seletivo", sem número),
'Experiência Profissional' (valorizada na admissão, mas não
"obrigatória" nem quantificada), 'Entrevista', 'Rounds Aplicação',
'Ranking US News', 'Ranking QS', 'Principais Empregadores', 'Fees
Adicionais USD', a maior parte dos campos mensais de custo de vida
restantes, 'Bolsas Disponíveis'/'Tipos de Bolsa'/'Quão Realista
(Intl)'/'TA/RA Disponível'/'Graduate Assistantships'/'Caminho Tuition
Residente' (além do detalhe de deadline de GA já preenchido acima, não
foram encontrados valores ou processos adicionais).

Fontes principais (pesquisadas via WebFetch, ago/2026):
- https://www.oakland.edu/business/graduate-executive-programs/master-of-science-in-business-analytics/
  (estrutura curricular em 4 módulos, pré-requisitos de nivelamento,
  documentos exigidos, deadline de GA para Fall, avaliação qualitativa
  de dificuldade de admissão)
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
    "Nome da Universidade": "Oakland University",
    "Estado": "MI",
    "Cidade": "Rochester (MI)",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Destaques Currículo": "4 módulos: Preparatory and Foundation Courses, Core Courses, Elective Courses, Capstone Project",
    "Matérias Pré-Requisito": (
        "Cursos preparatórios/de nivelamento exigidos para quem não "
        "tem background em information systems e/ou estatística; "
        "créditos não contam para o diploma"
    ),
    "Documentos Necessários": (
        "Aplicação online; CV/resume; 1 contato de referência; "
        "transcripts oficiais; internacionais: avaliação de "
        "transcript + TOEFL (se aplicável)"
    ),
    "Processo TA/RA": "Aplicação de Graduate Assistantship para Fall deve chegar até 1 abr",
    "Dificuldade Entrada": "Moderada — processo seletivo considera histórico acadêmico/bolsa, comunicação e experiência profissional (sem taxa de aceitação numérica divulgada)",
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
        raise ValueError("Nenhuma linha encontrada para Oakland University (Rochester (MI), MI). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Oakland University, Rochester (MI), MI)")

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
