"""
Data enrichment (round 2): atualiza APENAS a linha da 'Florida State'
(Herbert Wertheim College of Business, Tallahassee, FL) no
msba_programs_database.xlsx.

Foco exclusivo: Master of Science in Business Analytics (MS-BA), presencial,
STEM. Deep research em: custo total real de tuition para internacionais,
oportunidades de bolsas/assistantships, requisitos de GPA/GMAT/GRE/inglês, e
custo de vida local -- a pedido explicito da usuaria, atuando como engenheiro
de dados fazendo enrichment de uma unica instituicao.

CORRECAO CRITICA -- TUITION NAO-RESIDENTE DESATUALIZADA:
'Custo/Crédito (Intl) USD' e 'Tuition Total (Intl) USD' estavam calculados
com a taxa antiga ($1.173,86/credito, de um enrichment anterior). A tabela
oficial atual (2026-27) mudou para $1.278,04/credito -- CONFIRMADO por DUAS
fontes independentes da propria FSU:
  - https://tuition.fsu.edu/tuition-and-fees-0 (tabela geral de graduate
    tuition: $479,32 resident + $760,69 out-of-state fee + $38,03 out-of-state
    financial aid fee = $1.278,04 out-of-state)
  - https://wertheim.fsu.edu/graduate/program-costs ("A Smart Choice in
    Value" -- pagina de custos de programas de mestrado da propria Lindner..
    Wertheim, cita MS-BA especificamente: "$479.32 in state" / "$1,278.04
    out of state")
Corrigido: 33 creditos x $1.278,04 = $42.175,32 (arredondado p/ $42.175).
Tuition residente ($479,32/credito, $15.818 total) permanece inalterado --
CONFIRMADO sem mudanca pelas mesmas 2 fontes.

BOLSAS E ASSISTANTSHIPS (novo, antes generico/incompleto):
- https://wertheim.fsu.edu/students/scholarships/graduate: MS-BA tem 2
  bolsas proprias -- "AgileThought MIS Scholarship" e "Detlev H. Smaltz
  Scholarship for MIS". Collegewide (40+ bolsas p/ qualquer mestrado de
  negocios) inclui a "Danny Persaud International Student Scholarship"
  (especifica p/ internacionais). Elegibilidade: aluno degree-seeking
  matriculado em >= 6 creditos; concorrer exige se inscrever no programa
  ate o deadline prioritario de 15/jan; avaliacao ocorre na primavera p/
  pagamento no outono+primavera seguintes. Candidatura via
  fsu.academicworks.com.
- https://wertheim.fsu.edu/graduate/program-costs: confirma que
  assistantships competitivos "podem isentar ate o custo total de tuition"
  (nao apenas parcial).

CUSTO DE VIDA (fonte oficial nova -- antes so tinha aluguel/mercado
"oficial-adjacente" aproximados):
https://tuition.fsu.edu/cost-attendance/cost-estimates-fall-2026-spring-2027
Tabela oficial "Cost Estimates Fall 2026 - Spring 2027" p/ graduate, main
campus (cobre 2 semestres = 9 meses; NAO inclui o semestre de Summer, que e
o 3o semestre do programa lock-step de 12 meses da FSU -- o worksheet
especifico de Summer 2026 existe mas e um PDF nao extraivel de forma
confiavel, entao os 3 meses de verao ficam sem breakdown por categoria):
  - Housing off-campus: $8.840 / 9 meses = $982,22/mes -> usado em
    'Aluguel Privativo/mês' (troca de on-campus $936 para off-campus,
    mais fiel ao conceito de "aluguel PRIVATIVO" no mercado, nao dormitorio)
  - Food: $5.740 / 9 meses = $637,78/mes -> 'Mercado/mês' (era $622,
    baseado em estimativa antiga de $5.600)
  - Transportation (out-of-state/intl): $3.590 / 9 meses = $398,89/mes ->
    'Transporte/mês' (usado o valor non-resident, coerente com o foco em
    internacionais desta rodada; resident seria $2.320/9=$257,78)
  - Personal: $4.736 / 9 meses = $526,22/mes -> 'Outros Gastos/mês'
  - Books/Materials: $1.380 (custo unico, sem campo dedicado no schema --
    nao alocado a nenhuma coluna)
'Custo Vida Total (programa)' PERMANECE "Verificar": somar so 9 dos 12
meses do programa (falta o breakdown do Summer) daria um total
sistematicamente subestimado -- mesmo criterio conservador do enrichment
anterior (nao calcular total parcial/incompleto).

SEGURO SAUDE (novo, antes "Verificar"):
https://studentinsurance.fsu.edu/purchase-insurance -- FSU Student Health
Insurance Plan (Aetna), taxa internacional 2026-27: Annual (15/ago-14/ago)
$2.797. Esse periodo anual bate bem com a duracao de 12 meses do programa
(Summer+Fall+Spring), entao foi usado o valor anual cheio em vez de somar
semestres separados.

REQUISITOS DE ADMISSAO (GPA/GMAT/ingles): sem mudanca -- o enrichment
anterior ja tinha a politica oficial de dispensa automatica do GMAT/GRE
(GPA geral >= 3.4 e upper-division >= 3.2) e os scores minimos de ingles,
ambos confirmados como corretos nesta rodada (mesma pagina oficial). Buscas
adicionais so encontraram uma recomendacao GENERICA do "College of
Business" (GMAT 550 / GPA 3.0 upper-division) vinda de agregadores
(princetonreview/campusreel, misturado com dados do MBA) -- NAO e
especifico do MS-BA e nao tem fonte primaria direta, entao 'GPA Mínimo'
permanece "Verificar" em vez de usar um numero de segunda mao. Pelo mesmo
motivo, 'Taxa Aceitação %', 'GPA Médio Admitidos', 'GMAT Médio', 'Tamanho
Turma', 'Ranking US News', 'Ranking QS', 'Taxa Empregabilidade %', 'Salário
Médio Pós USD' e 'Principais Empregadores' seguem "Verificar" -- nenhuma
fonte publica (oficial ou agregador confiavel) com esses numeros
especificos ao MS-BA foi encontrada apesar de multiplas buscas.

CAMINHO TUITION RESIDENTE (novo, antes "Verificar"):
https://admissions.fsu.edu/residency + busca sobre F-1 e residencia na
Florida: alunos com visto F-1 NAO sao elegiveis a tuition de residente
(FL Regulation 7.005) -- so haveria caminho apos mudanca de status
migratorio para residente permanente (I-485/green card), o que nao e uma
via pratica durante um programa de 12 meses. Preenchido como "Não".

RODADAS DE APLICACAO (novo, antes "Verificar"): FSU nao usa rounds
formais -- apenas 2 prazos (ja cadastrados em 'Deadline Aplicação'):
prioritario 15/jan (tambem vale p/ concorrer a assistantship/bolsa) e
final 1/mar.

PROCESSO TA/RA (novo, antes "Verificar"): GA/RA/TA exigem 10h/semana de
trabalho; combinam stipend em dinheiro + tuition waiver (pode chegar a
isentar o total); concorrencia e avaliada automaticamente ao se inscrever
no programa ate o deadline prioritario de 15/jan (mesmo processo que as
bolsas collegewide).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "Florida State",
    "Estado": "FL",
    "Cidade": "Tallahassee",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Intl) USD": 1278.04,  # CORRECAO: taxa 2026-27 subiu de 1173.86 p/ 1278.04
    "Tuition Total (Intl) USD": 42175,  # CORRECAO: 33 x 1278.04 = 42175.32
    "Aluguel Privativo/mês USD": 982.22,  # 8840 / 9 (housing off-campus, fonte oficial FSU 2026-27)
    "Mercado/mês USD": 637.78,  # 5740 / 9 (food, fonte oficial FSU 2026-27; era estimativa de 622)
    "Transporte/mês USD": 398.89,  # 3590 / 9 (transportation non-resident/intl, fonte oficial)
    "Outros Gastos/mês USD": 526.22,  # 4736 / 9 (personal expenses, fonte oficial)
    "Seguro Saúde (programa) USD": 2797,  # FSU SHIP (Aetna) internacional, anual 15/ago-14/ago
    "Tipos de Bolsa": (
        "Bolsas específicas do MS-BA: AgileThought MIS Scholarship e "
        "Detlev H. Smaltz Scholarship for MIS. Collegewide (40+ bolsas "
        "para qualquer mestrado de negócios da Wertheim), incluindo a "
        "Danny Persaud International Student Scholarship. Graduate "
        "Assistantships competitivos, que podem isentar até o custo "
        "total de tuition"
    ),
    "Processo TA/RA": (
        "GA/RA/TA exigem 10h/semana de trabalho; combinam stipend em "
        "dinheiro + tuition waiver (pode isentar até o total); "
        "concorrência avaliada automaticamente ao se inscrever no "
        "programa até o deadline prioritário de 15/jan, mesmo processo "
        "das bolsas collegewide (candidatura via fsu.academicworks.com)"
    ),
    "Caminho Tuition Residente": (
        "Não -- alunos com visto F-1 não são elegíveis a tuition de "
        "residente (FL Regulation 7.005); só haveria via após mudança "
        "de status para residente permanente (I-485/green card), não "
        "prática durante um programa de 12 meses"
    ),
    "Rounds Aplicação": (
        "Sem rounds formais -- 2 prazos: prioritário 15/jan (também "
        "vale para concorrer a assistantship/bolsa) e final 1/mar"
    ),
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
        raise ValueError("Nenhuma linha encontrada para Florida State (Tallahassee, FL). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Florida State, Tallahassee, FL)")

    no_fill = PatternFill(fill_type=None)

    before = {}
    after = {}
    for header, new_value in UPDATES.items():
        col_idx = header_to_col[header]
        cell = ws.cell(row=row_idx, column=col_idx)
        before[header] = cell.value
        cell.value = new_value
        cell.fill = no_fill
        after[header] = new_value

    wb.save(DB_PATH)

    print("\nCampos atualizados (Como Estava -> Como Ficou):")
    for header in UPDATES:
        print(f"  - {header}:")
        print(f"      Antes: {before[header]!r}")
        print(f"      Depois: {after[header]!r}")


if __name__ == "__main__":
    main()
