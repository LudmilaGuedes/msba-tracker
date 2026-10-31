"""
Data enrichment: atualiza APENAS a linha da 'U of Alabama' (Manderson
Graduate School of Business / Culverhouse College of Business,
Tuscaloosa, AL) no msba_programs_database.xlsx.

NOTA DE METODOLOGIA: orçamento de busca web da sessão esgotado
(200/200 WebSearch); pesquisa feita via WebFetch, navegando
diretamente pelo site oficial (manderson.culverhouse.ua.edu,
graduate.ua.edu). Várias tentativas de páginas de seguro saúde
internacional e tuition detalhada da universidade retornaram erro de
DNS/404.

Foco: MS in Business Analytics, presencial, cohort-based, entrada
apenas no Fall, 1 ano acadêmico (Fall + Spring), 30 créditos, STEM
(já confirmado na planilha).

'Duração (meses)' preenchido com 9 -- CORREÇÃO/PRECISÃO: a página
oficial confirma "1 academic year (Fall and Spring terms)", ou seja,
2 semestres (~agosto a maio), não um ano civil completo de 12 meses.
'Total Créditos' confirmado oficialmente em 30.

'Custo/Crédito' PREENCHIDO por DERIVAÇÃO ARITMÉTICA dos totais de
tuition já confirmados e presentes na planilha (não vindos de
agregador, já eram valores "limpos"): com 30 créditos agora
confirmados oficialmente, 'Custo/Crédito (Intl)' = $34.172 ÷ 30 =
$1.139,07 e 'Custo/Crédito (Residente)' = $12.180 ÷ 30 = $406,00. Os
totais de tuition em si NÃO foram alterados (já estavam corretos).

'GPA Mínimo' preenchido: 3.3 (com exceções avaliadas caso a caso).
'GMAT/GRE Obrigatório' refinado com o critério oficial: GMAT (versão
antiga ou Focus Edition) ou GRE, percentil 60 ou superior; dispensa
disponível para candidatos que já possuem pós-graduação. 'Testes
Inglês Aceitos' preenchido com os mínimos oficiais: TOEFL 90 / IELTS
7.0 / Duolingo 110 / PTE 59. 'Deadline Aplicação' preenchido: 15 abr.
'Documentos Necessários' preenchido com a lista oficial (transcript,
evidência de Intro Stats B+ e Calculus B+ -- já presente na planilha
como pré-requisito --, resume, statement of purpose, 3 cartas de
recomendação).

'Processo TA/RA' PREENCHIDO com detalhe oficial que CONFIRMA e refina
a nota já existente em 'Caminho Tuition Residente' ("Spring GA →
summer in-state"): assistantships (GTA/GRA/GA) cobrem stipend
(mínimo definido por departamento) + tuition (até 15 créditos ou custo
real, o que for menor) + taxas de tecnologia/curso + seguro saúde
individual; candidatura via contato direto com o Graduate Program
Director (GTA/GRA) ou UA Student Job Site (GA); GTAs internacionais
com função de ensino devem completar o programa ITAP; alunos
segurando assistantship no semestre da primavera QUALIFICAM para
tarifa in-state durante interim/verão -- confirma oficialmente a nota
já presente na planilha.

'Salário Médio Pós USD' preenchido com $82.000 -- salário médio base
específico do MS Business Analytics (coorte 2022, também há bônus
médio de $8.000 não incluído no campo por não haver campo próprio no
schema). RESSALVA: dado de 2022, portanto potencialmente desatualizado
(~4 anos) -- documentado como tal, mas usado por ser explicitamente
específico deste programa (não de MBA ou outro curso do Culverhouse).

CUIDADO APLICADO (dados NÃO encontrados, permanecem 'Verificar'):
'Seguro Saúde (programa) USD' (página de seguro internacional retornou
erro de DNS), 'Tamanho Turma', 'GPA Médio Admitidos', 'GMAT Médio',
'Taxa Aceitação %', 'Ranking US News', 'Ranking QS', 'Taxa
Empregabilidade %', 'Entrevista', 'Rounds Aplicação' (deadline único
de 15 abr, sem menção a rounds), 'Fees Adicionais USD', e os campos
mensais de custo de vida ainda vazios (Utilities, Transporte, Custo
Vida Total).

Fontes principais (pesquisadas via WebFetch, ago/2026):
- https://manderson.culverhouse.ua.edu/programs/specialized-masters/business-analytics/
  (formato, duração, créditos, currículo, GPA/GMAT/GRE/TOEFL/IELTS/
  Duolingo/PTE, deadline, documentos, salário médio 2022, empregadores)
- https://graduate.ua.edu/applicants/funding/ (detalhamento oficial de
  GTA/GRA/GA: stipend, cobertura de tuition, seguro saúde, processo de
  candidatura, ITAP para GTAs internacionais, elegibilidade a tarifa
  in-state via assistantship de primavera)
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
    "Nome da Universidade": "U of Alabama",
    "Estado": "AL",
    "Cidade": "Tuscaloosa",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Duração (meses)": 9,  # oficial: 1 ano acadêmico = Fall + Spring (~ago a mai), não 12 meses corridos
    "Total Créditos": 30,  # oficial
    "Custo/Crédito (Intl) USD": 1139.07,  # 34.172 / 30 (total já confirmado, agora com créditos oficiais)
    "Custo/Crédito (Residente) USD": 406.00,  # 12.180 / 30
    "GPA Mínimo": "3.3 (exceções avaliadas caso a caso)",
    "GMAT/GRE Obrigatório": (
        "GMAT (versão antiga ou Focus Edition) ou GRE, percentil 60 "
        "ou superior; dispensa disponível para quem já possui "
        "pós-graduação"
    ),
    "Testes Inglês Aceitos": "TOEFL mín. 90 / IELTS mín. 7.0 / Duolingo mín. 110 / PTE mín. 59",
    "Deadline Aplicação": "15 abr",
    "Documentos Necessários": (
        "Transcript oficial; evidência de Intro Stats (B+) e Calculus "
        "(B+); resume; statement of purpose; 3 cartas de recomendação"
    ),
    "Processo TA/RA": (
        "GTA/GRA: candidatura via contato direto com o Graduate "
        "Program Director; GA: vagas no UA Student Job Site. Pacote "
        "inclui stipend (mínimo por departamento) + tuition (até 15 "
        "créditos ou custo real, o menor) + taxas de tecnologia/curso "
        "+ seguro saúde individual. GTAs internacionais com função de "
        "ensino devem completar o programa ITAP. Assistantship no "
        "semestre da primavera qualifica para tarifa in-state no "
        "interim/verão"
    ),
    "Salário Médio Pós USD": 82000,  # coorte 2022, específico do MS Business Analytics (dado potencialmente desatualizado)
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
        raise ValueError("Nenhuma linha encontrada para U of Alabama (Tuscaloosa, AL). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (U of Alabama, Tuscaloosa, AL)")

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
