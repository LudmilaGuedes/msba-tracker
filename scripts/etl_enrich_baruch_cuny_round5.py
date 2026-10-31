"""
Data enrichment (round 5): atualiza APENAS a linha da 'Baruch CUNY'
(Zicklin School of Business, New York, NY) no msba_programs_database.xlsx.

Objetivo desta rodada: auditoria completa (custos, bolsas, requisitos,
duração/formato) do MS in AI & Business Analytics (antigo MS Business
Analytics), presencial, STEM.

FONTES OFICIAIS (acessadas em 2026-10-03 via download direto do HTML --
zicklin.baruch.cuny.edu segue com erro de certificado SSL na ferramenta
de fetch padrão, mesma limitação dos rounds anteriores):
  - Tuition & Fees:
    https://zicklin.baruch.cuny.edu/academic-programs/graduate/grad-admissions/financing-your-education/tuition-fees/
  - Scholarships & Financial Aid:
    https://zicklin.baruch.cuny.edu/academic-programs/graduate/grad-admissions/financing-your-education/scholarships-financial-aid/
  - Application Deadlines:
    https://zicklin.baruch.cuny.edu/academic-programs/graduate/grad-admissions/application-process/application-deadlines/
  - International Applicants:
    https://zicklin.baruch.cuny.edu/academic-programs/graduate/grad-admissions/application-process/international-applicants/
  - Graduate Advisement Hub (MS Business Analytics/AI & Business Analytics
    + New Student Information):
    https://zicklin.baruch.cuny.edu/current/graduate-advisement-hub/academic-programs-curriculum/ms-programs/ms-business-analytics/

ACHADOS / CORREÇÕES:
- Fees: a linha tinha US$ 539, que corresponde só às taxas por semestre
  (Technology $125 full-time + Activities $39,60 + Consolidated Services
  $15 = $179,60 x 3 semestres). Faltava a Academic Excellence Fee de
  US$ 110 POR CRÉDITO para programas MS (34,5 x 110 = $3.795). Total
  corrigido: $4.333,80 ~ $4.334. Confere com a estimativa oficial da
  Zicklin (MS 31,5-42 créditos, não-residente: $30.900-$41.200).
- Tuition por crédito (MS presencial): $470 residente NY (1-11 créditos;
  12+ créditos = flat $5.545/semestre) / $855 não-residente -- confere
  com a linha, sem alteração.
- Duração: Advisement Hub informa "Average Time to Completion: 3-4
  semesters"; a página New Student Information mostra o plano mínimo de
  3 semestres full-time (1º: STA 9708/CIS 9340/CIS 9650; 2º: CIS 9660;
  último: capstone BUS 9430). Entrada no Fall -> formatura em dezembro do
  ano seguinte = ~16 meses.
- Pré-requisitos: nenhum para admissão; STA 9708 Managerial Statistics
  (3 cr.) é dispensada com 6 créditos de estatística na graduação, média
  >= B-, cursados nos últimos 5 anos.
- Bolsas: nenhuma bolsa nomeada para o MSBA/MSAIBA (as listadas são p/
  MBA, MS Accountancy/Taxation, Real Estate, Marketing). Internacionais
  são considerados para bolsas de mérito; aplicar em rounds iniciais dá
  "consideration for scholarships". Baruch College Fund scholarships:
  número limitado. Empréstimos federais/Work-Study via FAFSA (não
  disponíveis para internacionais).
- Graduate Assistantships: vagas de meio período no campus, só para
  alunos já matriculados (não-Executive); pagas por hora, sem waiver.
- Documentos: internacionais precisam de avaliação de diploma
  course-by-course de agência NACES (WES, ECE, The Evaluation Company)
  -- faltava na linha.
- Formato: presencial (full-time ou part-time); existe um Online MS in AI
  & Business Analytics separado.

NÃO ALTERADOS (sem fonte oficial nova): GPA Mínimo, GPA Médio, Taxa
Aceitação, GMAT Médio, Rankings, Tamanho Turma, Seguro Saúde (CUNY não
exige plano para F-1 -- não há custo de programa), Custo Vida Total.
Total Créditos mantido em 34,5: o Advisement Hub mostra "Credit Range:
33-37", que parece ser a faixa do currículo antigo; o round anterior
usou a faixa do currículo novo (34,5-38,5). Empregabilidade/salário
mantidos (relatório GCMC 2023-2024; não há relatório mais recente
publicado).
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
    "Nome da Universidade": "Baruch CUNY",
    "Estado": "NY",
    "Cidade": "New York",
}

UPDATES = {
    "Formato": (
        "Presencial (full-time ou part-time, Manhattan); existe também um "
        "Online MS in AI & Business Analytics separado -- não é o foco "
        "desta linha"
    ),
    "Duração (meses)": 16,  # 3 semestres full-time (Fall -> dez do ano seguinte); média oficial 3-4 semestres
    "Fees Adicionais USD": 4334,  # Academic Excellence $110/cr x 34,5 + ($125 + $39,60 + $15) x 3 semestres
    "Tipos de Bolsa": (
        "Mérito na admissão: internacionais são considerados "
        "automaticamente (sem aplicação separada; valores não divulgados) "
        "e aplicar nos rounds iniciais aumenta a chance. Nenhuma bolsa "
        "nomeada específica do MSBA/MSAIBA (as listadas são p/ MBA, MS "
        "Accountancy/Taxation, Real Estate e Marketing). Baruch College "
        "Fund scholarships: número limitado, aplicação interna. Graduate "
        "Assistantship paga por hora, sem waiver. Empréstimos federais e "
        "Work-Study (FAFSA) não se aplicam a internacionais."
    ),
    "Quão Realista (Intl)": (
        "Baixa — existe mérito, mas é limitado e de valor não divulgado, "
        "sem nada específico do programa; GA só paga salário por hora. "
        "Planejar o custo integral (a tuition já é das mais baixas: "
        "~US$ 29,5k + ~US$ 4,3k de fees)"
    ),
    "Processo TA/RA": (
        "Sem processo na aplicação: as Graduate Assistantships são vagas "
        "de meio período no campus (ex.: apoio a professores em aulas e "
        "provas), abertas só a alunos já matriculados (não-Executive); "
        "candidatura depois de começar o curso, pagas por hora, sem "
        "waiver de tuition"
    ),
    "Matérias Pré-Requisito": (
        "Nenhuma para admissão. No programa, STA 9708 Managerial "
        "Statistics (3 cr.) é obrigatória, mas é dispensada com 6 créditos "
        "de estatística na graduação, média >= B-, cursados nos últimos "
        "5 anos"
    ),
    "Testes Inglês Aceitos": (
        "TOEFL iBT mín. recomendado 100 (escala antiga) / 5 (escala nova) "
        "/ Duolingo DET 120 / IELTS Academic 7 / PTE Academic 70 — "
        "dispensado com diploma dos EUA ou de país de língua inglesa "
        "listado, ou histórico de 1º-12º ano em escola dos EUA. "
        "TOEFL/IELTS válidos por 2 anos."
    ),
    "Documentos Necessários": (
        "Transcripts não-oficiais de todas as faculdades cursadas "
        "(oficiais só após admissão), avaliação de diploma "
        "course-by-course de agência NACES (WES Basic Course-by-Course, "
        "ECE ou The Evaluation Company) para diploma estrangeiro, "
        "CV/currículo, 2 cartas de recomendação obrigatórias (2 "
        "profissionais; recém-formados: 1 profissional + 1 acadêmica), 2 "
        "ensaios obrigatórios (500 e 250 palavras) + 1 opcional, "
        "comprovante de proficiência em inglês (se aplicável). Após a "
        "admissão, para o I-20: comprovação de recursos para o 1º ano "
        "inteiro"
    ),
    "Deadline Aplicação": (
        "Round 1: 1 mar 2027 (decisão 10 abr); Round 2: 1 abr 2027 "
        "(decisão 10 mai); Round 3: 1 mai 2027 (decisão 10 jun); Round 4: "
        "1 jun 2027 (decisão 10 jul) -- Fall 2027. Para ingresso (Fall "
        "2028), o equivalente seria ~1 mar 2028"
    ),
    "Rounds Aplicação": (
        "4 rounds para Fall: 1 mar, 1 abr, 1 mai, 1 jun (decisões no dia "
        "10 do mês seguinte); Spring: 1 out, 1 nov, 1 dez; documentos "
        "completos até o round para receber decisão na data. Rounds "
        "iniciais dão prioridade para bolsas e escolha de disciplinas"
    ),
    "Entrevista": "Não mencionada nas páginas oficiais de admissão",
}


def main():
    rotate_backup()
    print("Backup rotacionado (política fixa) em: data/msba_programs_database_backup.xlsx")

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
        raise ValueError("Nenhuma linha encontrada para Baruch CUNY (New York, NY). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Baruch CUNY, New York, NY)")

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
