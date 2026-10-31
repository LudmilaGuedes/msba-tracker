"""
Data enrichment (RODADA 3): atualiza APENAS a linha 'Florida Institute of
Technology (Florida Tech)' (Nathan M. Bisk College of Business, Melbourne,
FL) no msba_programs_database.xlsx.

Programa: Master of Science in Business Analytics, presencial (campus de
Melbourne), STEM.

FONTES OFICIAIS (acessadas em 2026-10-03):
  - https://www.fit.edu/student-financial-services/tuition-and-fees/
    (tuition 2026-27, Other Fees, tabela do seguro saúde)
  - https://www.fit.edu/programs/business-analytics-ms/
  - https://www.fit.edu/graduate-admissions/
    (requisitos gerais, tabela de requisitos por programa, prazos de I-20)
  - https://www.fit.edu/admission/scholarships--aid/funding-graduate-studies/
  - https://www.fit.edu/admission/admitted-students/international-student-next-steps/demonstrating-english-proficiency/
  - https://www.fit.edu/student-financial-services/student-health-insurance-program/

ACHADOS / CORREÇÕES:
- Tuition: "College of Business (MBA, Masters)" = US$ 820/crédito em
  2026-27 (confirmado de novo; privada, mesma tarifa para todos).
- Fees Adicionais: o valor anterior (US$ 480) era só 1 ano de Activity
  Fee. Fees obrigatórias de pós full-time em Melbourne:
    Activity Fee US$ 240/semestre + Technology Fee US$ 160/semestre
    Panther Ready Book Charge US$ 27/crédito (inscrição automática) x 30 = US$ 810
    New Orientation Fee US$ 150 (uma vez) + Graduation Fee US$ 100
  Supondo 3 semestres (30 créditos a 9-12 créditos/semestre):
    3 x 400 + 810 + 150 + 100 = US$ 2.260.
- Seguro Saúde: US$ 1.999 (Fall US$ 1.007 + Spring US$ 992, Summer
  incluído) para quem começa no Fall 2026 -- confirmado. Obrigatório e sem
  opção de dispensa para internacionais.
- Bolsas: CORREÇÃO. O texto anterior dizia que internacionais são
  considerados automaticamente para bolsa de mérito (fonte EducationUSA),
  mas as bolsas institucionais automáticas da Florida Tech são de
  graduação. Na pós, a página oficial diz só que "Exceptionally qualified
  graduate students may be selected by their academic department to
  receive a scholarship or fellowship" e lista bolsas nomeadas, quase
  todas de engenharia/ciências/psicologia; a única de negócios é a
  "Barbara A. and William G. Roy Fellowship in Management Studies".
  Assistantships (salário e/ou waiver) existem para "well-qualified
  students" do campus de Melbourne, a critério do departamento.
- GA para internacionais: além dos requisitos do departamento, exige nota
  mínima de inglês específica e aprovação no International Teaching
  Assistant Speaking Assessment.
- Documentos (tabela oficial por programa): só Transcripts + Résumé/CV.
  Internacionais: tradução para inglês e, se preciso, avaliação
  course-by-course NACES (parceiras: The Evaluation Company e Atlas). Para
  o I-20: depósito de US$ 300, comprovação de inglês e documentos
  financeiros.
- Inglês: o corte do TOEFL mudou em 1 jan 2026 (o texto anterior não
  tinha a data); ELS 112 também é aceito.
- Prazos: admissão contínua; prazo do I-20 para internacionais é 1 jul
  (Fall) e 1 nov (Spring), com aplicação pelo menos 2 semanas antes;
  aplicações abrem 1 jul do ano anterior. Aplicar tarde pode prejudicar a
  consideração para bolsa/assistantship.
- Caminho Tuition Residente: não se aplica (privada, tarifa única).

NÃO ALTERADOS: Duração (meses) continua 'Verificar' -- nenhuma fonte
oficial informa a duração do MSBA presencial e o catálogo
(catalog.fit.edu) não abriu; dependendo da carga (9 ou 12 créditos por
semestre) fica entre ~12 e ~20 meses. Empregabilidade, salário,
empregadores, rankings e tamanho da turma: não publicados para o
programa.
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
    "Nome da Universidade": "Florida Institute of Technology (Florida Tech)",
    "Estado": "FL",
    "Cidade": "Melbourne",
}

UPDATES = {
    "Formato": (
        "Presencial (campus de Melbourne); existe também uma versão online "
        "separada (Florida Tech Online) -- não é o foco desta linha"
    ),
    "Início": (
        "Fall ou Spring para internacionais (admissão contínua; o MSBA não "
        "está na lista de programas só-Fall)"
    ),
    "Fees Adicionais USD": 2260,  # (240 + 160) x 3 semestres + 27 x 30 cr. + 150 + 100
    "Tipos de Bolsa": (
        "Não há bolsa automática de mérito na pós (as automáticas da "
        "Florida Tech são de graduação). Bolsas/fellowships de pós são "
        "escolhidas pelo departamento para alunos excepcionais; a única "
        "nomeada de negócios é a Barbara A. and William G. Roy Fellowship "
        "in Management Studies (valor não divulgado). Graduate "
        "Assistantships (salário e/ou waiver de tuition) para alunos "
        "'well-qualified' do campus de Melbourne, a critério do "
        "departamento. Sem bolsa need-based nem empréstimo federal para "
        "internacionais"
    ),
    "Quão Realista (Intl)": (
        "Baixa -- nada é automático nem tem valor divulgado; fellowship "
        "depende de indicação do departamento e GA de vaga no próprio "
        "departamento, com exigência extra de inglês. Planejar o custo "
        "integral (~US$ 24,6k de tuition + ~US$ 2,3k de fees + US$ 2k de "
        "seguro)"
    ),
    "TA/RA Disponível": (
        "Possível -- Graduate Assistantships (salário e/ou waiver de "
        "tuition) para mestrado no campus de Melbourne, seleção pelo "
        "departamento, sem valores divulgados para Business Analytics. "
        "Internacionais precisam de nota mínima de inglês específica para "
        "GA e passar no International Teaching Assistant Speaking "
        "Assessment. DGRATS é exclusivo de doutorado"
    ),
    "Caminho Tuition Residente": (
        "Não se aplica -- universidade privada, tarifa única de US$ 820/"
        "crédito para residentes e internacionais"
    ),
    "Experiência Profissional": "Não obrigatória (não consta na tabela oficial de requisitos do programa)",
    "Matérias Pré-Requisito": (
        "Nenhuma formal; o programa aceita bacharéis em business analytics, "
        "administração, TI ou áreas afins, e outras formações podem ser "
        "consideradas. Dependendo do histórico, o departamento pode exigir "
        "cursos de nivelamento em áreas básicas de negócios"
    ),
    "Testes Inglês Aceitos": (
        "TOEFL iBT 80 (provas até 31 dez 2025) ou 4.0 (a partir de 1 jan "
        "2026) / TOEFL Essentials 9 / IELTS 6.5 / PTE Academic 58 / "
        "Cambridge Advanced 180 / Duolingo 105 / ELS nível 112; nota de no "
        "máximo 2 anos, enviada pela agência; comprovação exigida antes da "
        "emissão do I-20"
    ),
    "Deadline Aplicação": (
        "1 jul (Fall, prazo do I-20 para internacionais; enviar a aplicação "
        "pelo menos 2 semanas antes, ~15 jun); Spring: 1 nov. Admissão "
        "contínua; aplicações abrem 1 jul do ano anterior. Para ingresso "
        "(Fall 2028): ~15 jun 2028"
    ),
    "Rounds Aplicação": (
        "Sem rounds -- admissão contínua (rolling) com decisões o ano todo; "
        "aplicar tarde pode prejudicar a consideração para bolsa e "
        "assistantship"
    ),
    "Documentos Necessários": (
        "Aplicação online (taxa US$ 60, não reembolsável); transcripts "
        "oficiais de cada instituição com 12+ créditos (não oficiais "
        "servem para a análise preliminar); tradução para inglês e, se "
        "necessário, avaliação course-by-course NACES (parceiras: The "
        "Evaluation Company e Atlas); Résumé/CV. NÃO exige carta de "
        "recomendação, statement of objectives nem GRE/GMAT (tabela oficial "
        "por programa). Para o I-20: depósito de US$ 300, comprovante de "
        "inglês e documentos financeiros"
    ),
    "Dificuldade Entrada": "Acessível (GPA 3.0, sem GRE/GMAT, cartas ou redação)",
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
        raise ValueError("Nenhuma linha encontrada para Florida Tech (Melbourne, FL). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Florida Tech, Melbourne, FL)")

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
