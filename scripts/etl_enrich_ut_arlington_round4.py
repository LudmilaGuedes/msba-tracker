# -*- coding: utf-8 -*-
"""
Data enrichment (RODADA 4): atualiza APENAS a linha 'UT Arlington'
(College of Business, Arlington, TX) no msba_programs_database.xlsx.

CONTEXTO: 3 rodadas anteriores ja tentaram (com bastante cautela, evitando
inventar decomposicoes/estimativas) preencher tuition e a maioria dos
demais campos, mas a maior parte permanece "Verificar". Esta rodada 4
tentou uma abordagem diferente para tuition (tabela do Student Accounts/
Bursar em vez da pagina generica de "Cost of Attendance" ja descartada na
rodada 3) e trouxe 2 achados relevantes.

TUITION: CONTINUA "Verificar". A pagina de "About Tuition" do Student
Accounts da UTA redireciona para uma "Tuition Estimator" interativa (nao
uma tabela estatica) e PDFs de "Tuition Charts" nao acessiveis via fetch
direto. Os componentes fixos encontrados (Statutory Tuition $50/credito,
Graduate Program Enhancement Fee $6/credito, Computer/Technology Fee
$37,90/credito, Athletics Fee $10/credito) NAO incluem a "Designated
Tuition" -- o componente principal e variavel por programa -- entao somar
so as pecas fixas encontradas daria um total incompleto e enganoso.
Mantido "Verificar" por disciplina (mesmo principio ja aplicado nas 3
rodadas anteriores: nao inventar decomposicao sobre dado incompleto).

CORRIGIDO -- Duração (meses): estava 16. A pagina OFICIAL do proprio
programa (uta.edu/academics/schools-colleges/business/departments/
information-systems-and-operations-management/master-of-science-business-
analytics) descreve explicitamente um "17-month program" -- fonte mais
direta que a origem do valor anterior (nao documentada). Corrigido para 17.

ATUALIZADO -- achado importante sobre bolsas (reduz o otimismo que a
ausencia de dado poderia sugerir): CONFIRMADO oficialmente
(uta.edu/administration/fao/apply-for-aid) que o "Scholarships Office"
afirma textualmente: "the Scholarships Office does not offer scholarships
for incoming graduate students" -- ou seja, NAO HA bolsa de entrada para
novos alunos de pos-graduacao. O portal de bolsas da UTA ("Mav
ScholarShop") so fica acessivel DEPOIS de ja estar matriculado, com pelo
menos 18 creditos concluidos e GPA >= 3,0 -- ou seja, exige bancar boa
parte do programa (33 creditos totais) sem ajuda antes de poder concorrer
a qualquer bolsa. Fellowships/assistantships sao mencionados de forma
GENERICA como disponiveis "em departamentos pelo campus" (fonte:
gradfaq da business school), sem detalhe de processo, cobertura ou
elegibilidade de internacional especifico ao MSBA -- entao os campos TA/RA
Disponivel, Processo TA/RA, Graduate Assistantships e Caminho Tuition
Residente permanecem "Verificar" (nao ha confirmacao suficiente para
preencher com confianca), mas 'Bolsas Disponíveis', 'Tipos de Bolsa' e
'Quão Realista (Intl)' foram atualizados para refletir esse achado
especifico sobre bolsa de ENTRADA, que e o que mais importa para uma
candidata aplicando de fora.

RECONFIRMADO sem alteracao: Principais Empregadores ("Aflac, Amazon,
Alcon, Deloitte, Toyota") -- a mesma pagina oficial do programa lista
exatamente esses nomes, agora com fonte documentada (nao estava
documentada nas rodadas anteriores).

NAO USADO -- Salário Médio Pós USD: a pagina oficial cita uma faixa "DFW
Data Scientist Salary Range: $79k-$130k" -- isso e faixa de MERCADO
REGIONAL (Dallas-Fort Worth) para a funcao "Data Scientist" em geral, NAO
um dado de outcome real dos formados do programa. Mapear isso para este
campo misturaria conceitos (mesmo erro ja evitado em rodadas de UTSA/Texas
A&M) -- mantido "Verificar".

NAO ALTERADO por falta de fonte oficial confiavel (mantido "Verificar"):
Custo/Crédito (ambos), Tuition Total (Residente), Fees Adicionais, Taxa
Aplicação USD (so confirmado que existe, sem valor), TA/RA Disponível,
Processo TA/RA, Caminho Tuition Residente, Graduate Assistantships,
Ranking US News, Ranking QS, Taxa Empregabilidade %, Tamanho Turma, GPA
Médio Admitidos, GMAT Médio, Taxa Aceitação %, Matérias Pré-Requisito,
Testes Inglês Aceitos (notas minimas especificas do MSBA nao encontradas),
Rounds Aplicação, Documentos Necessários, Entrevista.

Fontes principais (pesquisadas via WebSearch/WebFetch, set/2026):
- https://www.uta.edu/academics/schools-colleges/business/departments/information-systems-and-operations-management/master-of-science-business-analytics
  (oficial -- pagina do programa: duracao 17 meses, empregadores, faixa
  salarial regional de mercado)
- https://www.uta.edu/administration/fao/apply-for-aid
  (oficial -- bolsas de entrada NAO oferecidas para novos pos-graduandos;
  Mav ScholarShop so apos matricula com 18+ creditos e GPA >=3,0)
- https://www.uta.edu/academics/schools-colleges/business/graduate-business-services/gradfaq
  (oficial -- mencao generica a fellowships/assistantships "em
  departamentos pelo campus", sem detalhe especifico do MSBA)
- https://www.uta.edu/business-affairs/student-accounts/about-tuition e
  /mandatory-fees, catalog.uta.edu/academicregulations/tuition_fees/
  (oficial -- componentes fixos de tuition/fees por credito, mas SEM a
  Designated Tuition variavel por programa; nao usado para nao criar total
  incompleto)
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "UT Arlington",
    "Estado": "TX",
    "Cidade": "Arlington",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Duração (meses)": 17,
    "Bolsas Disponíveis": "Não para novos admitidos -- só após matrícula",
    "Tipos de Bolsa": (
        "\"Mav ScholarShop\" -- confirmado oficialmente que o "
        "Scholarships Office NÃO oferece bolsas para novos alunos de "
        "pós-graduação (\"does not offer scholarships for incoming "
        "graduate students\"); esse portal só fica acessível DEPOIS de "
        "já estar matriculado, com pelo menos 18 créditos concluídos e "
        "GPA ≥ 3,0. Fellowships/assistantships são mencionados de forma "
        "genérica como disponíveis \"em departamentos pelo campus\", sem "
        "detalhe específico do MSBA"
    ),
    "Quão Realista (Intl)": (
        "Baixa no momento da admissão -- não há bolsa de entrada para "
        "novos alunos de pós-graduação (confirmado oficialmente); "
        "qualquer bolsa via Mav ScholarShop exige já ter cursado pelo "
        "menos 18 dos 33 créditos do programa com GPA ≥3,0 antes de "
        "poder concorrer, ou seja, é preciso bancar boa parte do "
        "programa sem ajuda primeiro"
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
        raise ValueError("Nenhuma linha encontrada para UT Arlington (Arlington, TX). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (UT Arlington, Arlington, TX)")

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
