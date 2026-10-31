"""
Data enrichment: atualiza APENAS a linha da 'Wayne State University'
(Mike Ilitch School of Business + College of Engineering, Detroit, MI)
no msba_programs_database.xlsx.

NOTA DE METODOLOGIA: orçamento de busca web da sessão esgotado
(200/200 WebSearch); toda a pesquisa deste enrichment foi feita via
WebFetch, navegando diretamente pelo site oficial da Wayne State
(ilitchbusiness.wayne.edu, bulletins.wayne.edu, gradschool.wayne.edu).
Várias tentativas de localizar páginas de seguro saúde, taxa de
aplicação e custo de vida em Detroit retornaram 404 ou domínio
inexistente -- esses campos permanecem 'Verificar'.

Foco: M.S. in Data Science and Business Analytics (MSDSBA), programa
interdisciplinar entre o Mike Ilitch School of Business e o College of
Engineering, presencial (com opção de cursos online), STEM ("fast-
growing STEM field", conforme catálogo oficial), 30 créditos, 1 ano
full-time. Nome do programa já inclui "Business Analytics"
explicitamente -- sem ambiguidade de escopo como em casos anteriores
desta sessão (Rutgers/Pitt).

TUITION: CORRIGIDO. Os valores anteriores ('Custo/Crédito (Intl)' =
1937, '(Residente)' = 966, 'Tuition Total (Intl)' = '~29.100
(agregador — Verificar)') vinham de um agregador de baixa confiança,
já sinalizado como tal na própria célula. A tabela oficial de tuition
da Wayne State (Fall 2026) para o Mike Ilitch School of Business
(masters) lista: $1.005,13/crédito (residente MI) / $2.014,90/crédito
(fora do estado/internacional) -- CORRIGIDO para esses valores oficiais
(divergem significativamente da estimativa anterior, especialmente no
caso Intl: quase o dobro do valor do agregador). 'Tuition Total' (30
créditos): $30.153,90 (residente) / $60.447,00 (internacional). NOTA:
como o programa é interdisciplinar, 1 das 9 disciplinas "Major" pode
ser cursada no College of Engineering (CSC/IE) e teoricamente cobrada
por uma tarifa diferente da School of Business -- não confirmado nem
quantificado; a tarifa da Business School foi usada como base por ser
a escola-sede do programa (majoria das disciplinas: DSB/DSA/DSE).

'Fees Adicionais USD' preenchido com $3.297,63 -- ESTIMATIVA de
confiança moderada: Registration Fee oficial $402,91/semestre + Student
Service Fee oficial $69,63/crédito, mas o número exato de semestres
necessários para completar os 30 créditos em regime full-time de 1 ano
NÃO foi confirmado oficialmente (assumidos 3 semestres, ~10 créditos
cada, dado que o currículo inclui um Applied Analytics Practicum de 6
créditos que pode rodar no verão) -- documentado como estimativa, não
valor oficial fechado.

'GMAT/GRE Obrigatório' preenchido com nota de ausência: a página
oficial do bulletin de admissão (que detalha GPA mínimo e
pré-requisitos) NÃO menciona GMAT, GRE, TOEFL ou IELTS como exigência
de teste -- forte indício de que não são exigidos para admissão em si
(distinto do requisito de proficiência em inglês, que É exigido
separadamente e tem página própria). 'Testes Inglês Aceitos'
preenchido com os mínimos oficiais da Wayne State (universais, não
específicos deste programa): TOEFL iBT mín. 79 (antes de 21/jan/26) ou
4.5 (novo formato, a partir de 21/jan/26); IELTS mín. 6.5 (Indicator
NÃO aceito); Duolingo mín. 110; validade de 2 anos.

'Matérias Pré-Requisito' preenchido com o requisito oficial:
competência básica em probabilidade, estatística e programação; cursos
de nivelamento disponíveis no verão para candidatos com lacunas,
podendo receber admissão condicional.

'Deadline Aplicação' preenchido com os deadlines GERAIS da Graduate
School (Fall: doméstico 1 jun / internacional 1 mai) -- a própria
página oficial avisa que programas individuais podem ter deadlines
PRÓPRIOS que substituem esses, e não foi encontrado um deadline
específico do MSDSBA -- documentado como ressalva explícita no valor
do campo.

'Bolsas Disponíveis' e 'Tipos de Bolsa' preenchidos com CAUTELA: a
página de aid do Ilitch School confirma acesso a "$350 milhões" em
ajuda financeira universitária (need+merit) e ~200 bolsas nomeadas via
portal ScholarshipUniverse (~200 alunos de business contemplados/ano),
mas SEM valores, critérios ou confirmação específica para
internacionais/pós-graduação -- fica claro que existe alguma
disponibilidade, mas os detalhes não foram encontrados.

'TA/RA Disponível' preenchido com nota hedged: assistantships existem
em nível universitário/departamental na Wayne State (Graduate School
confirma a modalidade como opção de funding geral), mas nenhuma fonte
confirma disponibilidade específica para o MSDSBA -- recomendação
oficial é contatar o departamento diretamente.

CUIDADO APLICADO (dados NÃO encontrados, permanecem 'Verificar'):
'Taxa Aplicação USD', 'Documentos Necessários', 'Entrevista', 'Seguro
Saúde (programa) USD', TODOS os campos mensais de custo de vida
(aluguel, mercado, utilities, celular, wifi, transporte, outros
gastos, seguro residencial, custo vida total), 'Taxa Empregabilidade
%', 'Salário Médio Pós USD', 'Principais Empregadores', Rankings
(US News/QS), 'GPA Médio Admitidos', 'Taxa Aceitação %', 'Tamanho
Turma', 'Início', 'Experiência Profissional', 'Processo TA/RA',
'Graduate Assistantships', 'Caminho Tuition Residente', 'Quão
Realista (Intl)' e 'Dificuldade Entrada' -- não encontrados nas
páginas oficiais acessíveis via WebFetch nesta rodada (várias
tentativas de páginas de seguro saúde internacional e custo de vida em
Detroit retornaram erro 404/DNS).

Fontes principais (pesquisadas via WebFetch, ago/2026):
- https://bulletins.wayne.edu/graduate/school-business/programs/data-science-business-analytics-ms/
  (créditos, GPA mínimo, pré-requisitos, estrutura curricular oficial)
- https://ilitchbusiness.wayne.edu/tisa/msdsba (formato, duração,
  concentrações, link ao bulletin)
- https://wayne.edu/registrar/tuition/fee-charts (tuition e fees
  oficiais Fall 2026, Mike Ilitch School of Business)
- https://gradschool.wayne.edu/admissions/english-proficiency
  (mínimos oficiais de TOEFL/IELTS/Duolingo)
- https://gradschool.wayne.edu/admissions/deadlines (deadlines gerais
  da Graduate School, com ressalva de que programas têm prazos
  próprios)
- https://ilitchbusiness.wayne.edu/admissions/scholarships-aid e
  https://gradschool.wayne.edu/students/funding (bolsas gerais,
  assistantships gerais, sem detalhamento específico do MSDSBA)
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
    "Nome da Universidade": "Wayne State University",
    "Estado": "MI",
    "Cidade": "Detroit",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Intl) USD": 2014.90,  # CORREÇÃO: oficial Fall 2026, Mike Ilitch School of Business (era estimativa de agregador)
    "Custo/Crédito (Residente) USD": 1005.13,  # CORREÇÃO: oficial Fall 2026 (era estimativa de agregador)
    "Tuition Total (Intl) USD": 60447.00,  # CORREÇÃO: 30 créditos x $2.014,90 (era string vaga de agregador)
    "Tuition Total (Residente) USD": 30153.90,  # 30 créditos x $1.005,13, oficial
    "Fees Adicionais USD": 3297.63,  # estimativa: 3 semestres x $402,91 (Registration) + 30 créditos x $69,63 (Student Service)
    "GMAT/GRE Obrigatório": "Não mencionado como exigência na página oficial de admissão (que detalha GPA e pré-requisitos) — indício de que não é exigido, mas sem confirmação explícita",
    "Testes Inglês Aceitos": (
        "TOEFL iBT mín. 79 (exames antes de 21/jan/26) ou 4.5 (novo "
        "formato, a partir de 21/jan/26); IELTS mín. 6.5 (\"IELTS "
        "Indicator\" não aceito); Duolingo mín. 110; validade de 2 anos"
    ),
    "Matérias Pré-Requisito": (
        "Competência básica em probabilidade, estatística e "
        "programação; cursos de nivelamento disponíveis no verão para "
        "candidatos com lacunas, podendo receber admissão condicional"
    ),
    "Deadline Aplicação": "Deadlines gerais da Graduate School — Fall: doméstico 1 jun / internacional 1 mai (programa pode ter prazo próprio, não encontrado especificamente para o MSDSBA)",
    "Bolsas Disponíveis": "Acesso geral a ajuda financeira universitária (~$350 milhões, need+merit) e ~100 bolsas nomeadas via portal ScholarshipUniverse; sem valores/critérios específicos confirmados para internacionais/pós-graduação",
    "Tipos de Bolsa": "Bolsas nomeadas via portal WSU ScholarshipUniverse (~200 alunos de business contemplados/ano); maioria do funding need-based (FAFSA, aplicável a domésticos)",
    "TA/RA Disponível": "Não confirmado especificamente para o MSDSBA — assistantships existem em nível universitário/departamental na Wayne State, mas disponibilidade para este programa não foi confirmada nem negada nas fontes oficiais consultadas",
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
        raise ValueError("Nenhuma linha encontrada para Wayne State University (Detroit, MI). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Wayne State University, Detroit, MI)")

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
