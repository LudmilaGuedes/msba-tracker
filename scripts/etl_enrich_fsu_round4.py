"""
Data enrichment (RODADA 4): atualiza APENAS a linha 'Florida State'
(Herbert Wertheim College of Business, Tallahassee, FL) no
msba_programs_database.xlsx.

Programa: MS in Business Analytics (MS-BA), presencial, STEM.

FONTES OFICIAIS (acessadas em 2026-10-17):
  - https://wertheim.fsu.edu/msba
  - https://wertheim.fsu.edu/graduate/program-costs
  - https://wertheim.fsu.edu/students/scholarships/graduate
  - https://tuition.fsu.edu/tuition-and-fees-0
  - https://studentinsurance.fsu.edu/purchase-insurance
  - https://gradschool.fsu.edu/admissions/international-admissions

CONFIRMADO (sem alteração): US$ 1.278,04/crédito não-residente e
US$ 479,32 residente, já com as fees por crédito ("tuition plus fees");
33 créditos -> US$ 42.175 / US$ 15.818; lock-step de 3 semestres
(Summer/Fall/Spring), entrada só no Summer; GMAT/GRE opcional; prazos 15
jan (prioridade, também para assistantship) e 1 mar (final); taxa US$ 30;
2 cartas (3 preferível); avaliação course-by-course NACES obrigatória;
Duolingo 125, IELTS 7, TOEFL 5.0 (nova escala), PTE 66, C1 Advanced 180,
Michigan 55.

CORREÇÕES:
- Duração: a página diz que o programa começa "late June" e termina
  "early May" -> ~11 meses (era 12).
- Fees Adicionais: a página avisa que o valor por crédito não inclui
  taxas fixas por semestre; a tabela da FSU traz a Student Facility Use
  Fee de US$ 20/semestre -> 3 x 20 = US$ 60 (a taxa do FSUCard não tem
  valor publicado na tabela).
- Seguro Saúde: obrigatório (plano da FSU ou cobertura equivalente). O
  plano anual vai de 15 ago a 14 ago (US$ 2.797); como o MS-BA começa no
  fim de junho, é preciso também o plano de verão (US$ 728) -> US$ 3.525.
- Assistantships: CORREÇÃO. O texto anterior dizia "stipend ~US$
  29.000/ano + waiver + seguro" e "10h/semana". As 10 h/semana estão na
  seção do Department of Accounting, e o valor de ~US$ 29 mil não aparece
  nas páginas do MS-BA. A página do programa diz só que há "competitive
  assistantships that provide financial assistance", com prioridade para
  quem aplica até 15 jan; a página de custos diz que podem "waive up to
  the full tuition cost". Valores e carga horária: não divulgados.
- Bolsas: as específicas da área (AgileThought MIS Scholarship e Detlev
  H. Smaltz Scholarship for MIS) são do departamento de Business
  Analytics, Information Systems and Supply Chain; entre as bolsas gerais
  há a Danny Persaud International Student Scholarship. Valores não
  publicados.
- Taxa de aplicação: a pós da FSU não dá isenção ("unable to waive the
  application fee payment for graduate applicants").
- Documentos: proof of degree separado se o histórico não mostrar o
  diploma; para o I-20, comprovação financeira do 1º ano.

NÃO ALTERADOS: GPA Mínimo (sem valor publicado para o MS-BA),
empregabilidade, salário, empregadores, rankings, tamanho da turma,
taxa de aceitação, entrevista.
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
    "Nome da Universidade": "Florida State",
    "Estado": "FL",
    "Cidade": "Tallahassee",
}

UPDATES = {
    "Formato": "Presencial (full-time, lock-step de 3 semestres)",
    "Duração (meses)": 11,  # fim de junho -> início de maio
    "Fees Adicionais USD": 60,  # Student Facility Use Fee US$ 20/semestre x 3
    "Seguro Saúde (programa) USD": 3525,  # anual US$ 2.797 + verão US$ 728
    "Tipos de Bolsa": (
        "Bolsas da área (AgileThought MIS Scholarship e Detlev H. Smaltz "
        "Scholarship for MIS) e bolsas gerais da Wertheim para mestrados, "
        "incluindo a Danny Persaud International Student Scholarship -- "
        "valores não publicados. Assistantships competitivas que podem "
        "isentar até a tuition inteira; prioridade para quem aplica até 15 "
        "jan"
    ),
    "Quão Realista (Intl)": (
        "Baixa a moderada -- existem assistantships e bolsas abertas a "
        "internacionais, mas sem valores nem número de vagas publicados; "
        "para concorrer é preciso aplicar até 15 jan. Sem assistantship, a "
        "tuition (~US$ 42k) já é das mais baixas entre programas fortes"
    ),
    "TA/RA Disponível": (
        "Possível -- assistantships competitivas para o MS-BA, que podem "
        "isentar até a tuition inteira; valor do stipend e carga horária "
        "não divulgados; não garantidas"
    ),
    "Processo TA/RA": (
        "Sem aplicação separada publicada: quem completa a aplicação até o "
        "prazo prioritário de 15 jan é considerado para as assistantships "
        "do MS-BA. Bolsas da Wertheim pelo portal fsu.academicworks.com"
    ),
    "Graduate Assistantships": (
        "Competitivas, podem isentar até a tuition inteira; stipend e carga "
        "horária não publicados para o MS-BA (o valor de ~US$ 29k/ano e as "
        "10 h/semana citados antes não se referem a este programa)"
    ),
    "Deadline Aplicação": (
        "Round 1: 15 jan 2027 (prioridade, também para assistantship); "
        "Round 2: 1 mar 2027 (prazo final, só se houver vaga) -- início no "
        "Summer 2027 (fim de junho). Para ingresso (Summer 2028): ~15 jan "
        "2028"
    ),
    "Documentos Necessários": (
        "Applicant statement; resume/CV com datas e cargos (tempo integral "
        "ou parcial); 2 cartas de recomendação (3 preferível) de "
        "empregadores ou professores; 1 transcript oficial de cada "
        "instituição (internacionais: avaliação course-by-course NACES "
        "obrigatória antes da análise, e prova do diploma se o histórico não "
        "mostrar); GMAT/GRE opcional (códigos PN8K567 / 5219); teste de "
        "inglês; taxa US$ 30 (sem isenção na pós). Para o I-20: comprovação "
        "financeira do 1º ano"
    ),
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
        raise ValueError("Nenhuma linha encontrada para Florida State (Tallahassee, FL). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Florida State, Tallahassee, FL)")

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
