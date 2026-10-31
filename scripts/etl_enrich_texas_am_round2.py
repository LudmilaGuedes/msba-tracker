# -*- coding: utf-8 -*-
"""
Data enrichment (round 2): atualiza APENAS a linha 'Texas A&M' (Mays
Business School, cidade cadastrada como "Houston" -- o programa MS in
Analytics NAO e no campus de College Station, ver rodada 1) no
msba_programs_database.xlsx.

CONTEXTO: rodada 1 (scripts/etl_enrich_texas_am.py) ja descobriu que este
programa e ministrado no CityCentre HOUSTON, formato hibrido noturno
(tercas/quintas, com opcao de livestream), voltado a profissionais que ja
trabalham -- NAO e um programa full-time diurno tradicional. Isso segue
valendo e e reforcado nesta rodada com um novo achado (ver abaixo).

CORRECAO -- Duração (meses): estava "~20 (5 semestres em formato
lock-step...)" na rodada 1, baseado em fonte indireta. A pagina OFICIAL do
programa (mays.tamu.edu/graduate/masters/ms-programs/analytics/) diz
explicitamente: **17 meses**, estrutura lock-step de **4 semestres
consecutivos** (fall, spring, summer, fall seguinte) -- nao 5 semestres/20
meses. Corrigido para a fonte mais direta (pagina oficial do proprio
programa). Tuition ($70.000, identico p/ residente e intl) confere com essa
estrutura: cobrada em 4 parcelas de $17.500/semestre.

CORRECAO -- Caminho Tuition Residente: estava "Sim ($1k+ competitive =
waiver)", sugerindo que uma bolsa converteria a tarifa para residente. Essa
logica NAO SE APLICA a este programa: a tuition e a MESMA ($70.000) para
residente E internacional -- nao existe uma tarifa diferenciada de
residente para se converter via bolsa. O valor anterior parece ter sido
copiado por engano do padrao de outro programa (ex.: UT Dallas, onde essa
logica de fato existe). Corrigido para refletir que o campo nao se aplica
aqui.

ATUALIZADO -- Documentos Necessários: confirmado oficialmente
(mays.tamu.edu/ms-analytics/admissions) que a lista de documentos inclui
GPA, resume, ENSAIO, transcript nao-oficial na inscricao, e comprovacao de
pre-requisitos -- alem da entrevista avaliativa ja cadastrada. "Ensaio" e
"pre-requisitos" nao estavam no texto anterior; acrescentados.

ATUALIZADO -- Destaques Currículo: corrigida a mencao interna a "5
semestres" para "4 semestres / 17 meses" (consistente com a correcao acima)
e ACRESCENTADO um achado relevante desta rodada: a pagina oficial divulga
metricas de carreira BEM DIFERENTES do que os campos "Taxa Empregabilidade
%"/"Salário Médio Pós USD" pressupoem -- 96% dos participantes recebem
aumento salarial durante/apos o curso, 92% assumem novo cargo, 51% relatam
aumento de $15.000+, valor medio de projeto de capstone de $127,8 milhoes
(2025). Essas metricas refletem o publico-alvo real do programa
(profissionais JA EMPREGADOS buscando progressao de carreira, nao
recem-formados entrando no mercado pela primeira vez) -- por isso NAO foram
usadas para preencher "Taxa Empregabilidade %" ou "Salário Médio Pós USD"
(que continuam "Verificar"), o que seria comparar coisas diferentes. Essa
natureza do programa reforca a ressalva ja registrada na rodada 1 sobre
compatibilidade com o perfil de aluno F-1 em tempo integral.

NAO ALTERADO por falta de fonte oficial confiavel (mantido como estava,
incluindo "Verificar" onde ja era o caso):
- Taxa Aplicação USD: um resultado de busca nao verificado por fetch direto
  citava "$90", mas sem confirmacao de que e especifico do MS Analytics
  (pode ser taxa generica da Mays) -- nao confiavel o suficiente para
  substituir "Verificar".
- Testes Inglês Aceitos ("TOEFL/IELTS/Duolingo (mín. 100)"): NAO alterado.
  Esta rodada encontrou numeros de TOEFL/IELTS numa pagina oficial, mas
  essa pagina e do programa IRMAO "MS Supply Chain and Analytics" -- um
  programa DIFERENTE. Usar esses numeros aqui misturaria dados de dois
  programas distintos, o que o projeto proibe. O valor existente nao foi
  nem confirmado nem contradito para o MS Analytics especificamente.
- Tipos de Bolsa, TA/RA Disponível, Processo TA/RA, Graduate Assistantships:
  a pagina oficial so menciona "financial aid available (federal, state,
  institutional, private sources)" de forma generica, sem nomear bolsas,
  TA/RA ou valores -- nao encontrado nada que confirme ou contradiga os
  valores ja cadastrados (que tambem carecem de fonte). Mantidos sem
  alteracao.
- GPA Mínimo/Médio, GMAT Médio, Taxa Aceitação %, Tamanho Turma, Matérias
  Pré-Requisito (conteudo especifico), Principais Empregadores, Ranking US
  News/QS: nenhuma fonte oficial encontrada. Mantidos "Verificar".
- Deadlines/Rounds Aplicação: reconfirmados sem alteracao (ciclo Fall 2027:
  abertura 1/set/2026, R1 3/dez/2026, R2-bolsa 1/abr/2027, Final
  25/jun/2027).
- Trilha full-time em College Station: reconfirmado que NAO existe uma
  oferta full-time diurna tradicional deste "MS Analytics" -- o "MBA + MS
  Analytics" e um dual-degree diferente, ja descartado na rodada 1.

Fontes principais (pesquisadas via WebSearch/WebFetch, set/2026):
- https://mays.tamu.edu/graduate/masters/ms-programs/analytics/
  (oficial -- duracao 17 meses/4 semestres, tuition $70k em 4 parcelas de
  $17.500, metricas de carreira/aumento salarial, financial aid generico)
- https://mays.tamu.edu/ms-analytics/admissions
  (oficial -- documentos exigidos: GPA, resume, ensaio, transcript,
  pre-requisitos, entrevista)
- https://mays.tamu.edu/graduate/masters/ms-programs/supply-chain-and-analytics/international-student-information/
  (oficial, mas de OUTRO programa -- TOEFL/IELTS encontrados aqui NAO usados
  nesta linha para evitar mistura de dados entre programas)
- https://mays.tamu.edu/graduate/mba/mba-ms-analytics/
  (oficial -- confirma que o dual-degree MBA+MS Analytics e diferente,
  nao usado)
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "Texas A&M",
    "Estado": "TX",
    "Cidade": "Houston",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Duração (meses)": (
        "17 (4 semestres consecutivos lock-step: fall, spring, summer, "
        "fall seguinte; não pode ser acelerado nem estendido) -- CORRIGE "
        "valor anterior de \"~20 meses/5 semestres\" com fonte oficial "
        "mais direta"
    ),
    "Caminho Tuition Residente": (
        "Não aplicável -- a tuition é a MESMA ($70.000) para residente e "
        "internacional; não existe tarifa diferenciada de residente para "
        "se converter via bolsa neste programa (diferente de outros "
        "programas da própria UT/A&M, ex. UT Dallas, onde esse caminho "
        "existe de fato)"
    ),
    "Documentos Necessários": (
        "GPA, resume, ensaio, transcript não-oficial na inscrição, "
        "comprovação de pré-requisitos, entrevista avaliativa (confirmado "
        "oficialmente em mays.tamu.edu/ms-analytics/admissions)"
    ),
    "Destaques Currículo": (
        "ATENÇÃO: formato híbrido -- aulas presenciais no CityCentre "
        "Houston às terças/quintas à noite, com opção de assistir ao vivo "
        "remotamente; NÃO é ministrado no campus de College Station "
        "apesar do nome da universidade. 4 semestres em formato lock-step "
        "(17 meses; não pode ser acelerado nem estendido). Voltado a "
        "profissionais que já trabalham e recém-formados. Ranqueado entre "
        "os top 10-16 nacionais pelo US News na categoria de programas "
        "online de negócios (não há ranking QS confirmado). Métricas de "
        "carreira divulgadas pelo programa (96% recebem aumento salarial "
        "durante/após o curso, 92% assumem novo cargo, 51% relatam "
        "aumento de $15.000+, valor médio de projeto de capstone de "
        "$127,8 milhões em 2025) refletem participantes JÁ EMPREGADOS "
        "buscando progressão de carreira -- NÃO equivalem a taxa de "
        "empregabilidade/salário pós-formatura tradicional de "
        "recém-formados, por isso não usadas para preencher esses campos"
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
        raise ValueError("Nenhuma linha encontrada para Texas A&M (Houston, TX). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Texas A&M, Houston, TX)")

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
