"""
Data enrichment: atualiza APENAS a linha 'U of South Carolina' (Darla
Moore School of Business, Columbia, SC) no msba_programs_database.xlsx.

Pedido da usuaria: confirmar se a tuition esta certa "para o meu perfil"
(internacional), confirmar bolsas/descontos, e buscar rankings/
empregabilidade que estavam "Verificar".

=========================================================================
1) TUITION -- RECONFIRMADA, SEM ALTERACAO
=========================================================================
Fonte oficial: sc.edu/.../masters_business_analytics/tuition_fees.php
  "$1,100 per credit hour (mesmo para residentes e não-residentes de SC)"
  "The Moore School does not offer an out-of-state rate for the MSBA
   program"
Ou seja: NAO EXISTE sobretaxa de internacional/fora do estado neste
programa especifico -- $1.100/credito e $33.000 total valem
IGUALMENTE para a usuaria (F-1) e para um aluno residente da Carolina
do Sul. Ja estava correto na planilha; apenas reconfirmado.

'Fees Adicionais' ($1.692) tambem RECONFIRMADO, agora com o detalhamento
oficial que faltava (a planilha ja tinha o numero certo, mas sem a conta
por tras): Matriculation fee $80 (unica) + International Student fee
$112 (unica) + 3 semestres (Fall/Spring/Summer, programa de 12 meses) x
(Student Health Center $100 + Technology $200 + International Student
fee $200) = $192 + 3x$500 = $192 + $1.500 = $1.692,00 -- bate EXATO.

=========================================================================
2) BOLSAS/DESCONTOS -- PARCIALMENTE CONFIRMADO, 1 DADO DESCARTADO
=========================================================================
CONFIRMADO na pagina oficial do programa: "consider each candidate for
merit-based financial awards in the form of a fellowship (amounts vary)
or graduate assistantship upon admission" -- bate com o que ja estava
cadastrado ("Fellowships de mérito (valores variáveis) + Graduate
Assistantships no Moore School Data Lab").

DADO ENCONTRADO MAS DESCARTADO (risco de contaminacao entre programas):
uma busca retornou "Fellowships provide a waiver so students only pay
in-state tuition and a scholarship of $250 per semester". Esse texto
NAO PODE se referir ao MSBA -- o proprio MSBA ja confirmou que NAO TEM
tarifa de fora do estado a ser "reduzida para in-state" (item 1 acima).
Uma bolsa que "reduz para a tarifa in-state" so faz sentido em um
programa que TEM diferenciacao de tarifa -- por isso essa frase quase
certamente veio de uma pagina de outro mestrado da Moore School (ex.:
Master of Accountancy, que aparecia misturado nos mesmos resultados de
busca). NAO foi usada para nao contaminar esta linha com dado de outro
programa (mesmo criterio ja aplicado a UGA/USC/UCF nesta sessao).

Prazo de candidatura a GA: a planilha ja tinha "Accept GA by Apr 15".
Uma busca sugeriu "April 1", mas essa informacao veio como paráfrase de
busca (sem aspas/citacao direta), enquanto a pagina de FAQ oficial do
programa (fetch direto) NAO menciona nenhum prazo especifico. Como
nenhuma das duas fontes e mais confiavel que a outra, e a que ja estava
na planilha tem uma origem de pesquisa anterior mais especifica, NAO foi
alterada.
Elegibilidade de internacionais para GA/fellowship: a pagina de FAQ
oficial nao confirma nem nega -- apenas recomenda contato com
"International Students Services" a parte. Sem novidade, mantido
'Quão Realista (Intl)' = "Moderada".

=========================================================================
2b) DIFICULDADE DE ENTRADA -- resposta a pergunta de acompanhamento
=========================================================================
Taxa de aceitação especifica do MSBA NAO e divulgada publicamente (nem
oficial nem terceiros -- so ha dados de aceitacao do MBA, outro
programa, descartados). Confirmado via busca: turma media de 23 alunos,
GPA mediano 3.80 (ja batia com o que estava cadastrado), 20 areas/
industrias de formacao representadas na turma. Pagina oficial de FAQ:
"the MSBA program does not admit/deny solely based on GRE/GMAT scores"
-- avaliacao holistica (rigor academico, base quantitativa, cartas de
recomendacao, personal statement, experiencia profissional, entrevista).
GRE esperado (nao eliminatorio): minimo 305. Adicionado a 'GMAT/GRE
Obrigatório'.

=========================================================================
3) EMPREGABILIDADE/SALARIO/EMPREGADORES -- NOVO, ESPECIFICO DO MSBA
=========================================================================
Fonte oficial: sc.edu/.../masters_business_analytics/career_impact.php
(pagina de resultados de carreira DEDICADA ao MSBA, nao generica):
  "85%" de colocacao profissional dentro de 6 meses da formatura (entre
  os que buscaram emprego ativamente)
  Salario inicial medio: "$83.000"
  Empregadores citados: NTT Data Services, Microsoft, EY, Harsco Rail,
  Stanley Black and Decker, TTI, Global Lending Services, G2O
  Technologies, Resurgent Capital Services, Epic, DXC Technology

=========================================================================
4) RANKINGS -- NAO ENCONTRADO RANKING ESPECIFICO DO MSBA
=========================================================================
'Ranking US News' encontrado e SO da escola em geral (Moore School,
MBA): #53 -- NAO usado nesta linha (seria contaminacao entre programas,
mesmo criterio de sempre nesta sessao). Nenhum ranking QS ou US News
ESPECIFICO de MSBA/Business Analytics foi encontrado para esta escola --
permanece "Verificar" para ambos.
Achado relevante, mas que nao e um "ranking" no formato do schema:
o departamento de Management Science ganhou o "2024 Global INFORMS UPS
Smith Prize" (melhor ensino em operations research/analytics/data
science, competindo com National University of Singapore e University
of Waterloo) -- adicionado como destaque curricular, ja que nao ha
coluna dedicada a premios.
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "U of South Carolina",
    "Estado": "SC",
    "Cidade": "Columbia",
}

UPDATES = {
    "GMAT/GRE Obrigatório": (
        "Obrigatório GMAT ou GRE (validade de até 5 anos); dispensável "
        "para candidatos com desempenho de graduação superior e/ou "
        "experiência profissional relevante. GRE esperado: mínimo 305 "
        "(fonte oficial FAQ: \"the MSBA program does not admit/deny "
        "solely based on GRE/GMAT scores\" -- avaliação é holística)"
    ),
    "Taxa Empregabilidade %": 85,  # oficial career_impact.php: colocação em 6 meses (job-seeking basis)
    "Salário Médio Pós USD": 83000,
    "Principais Empregadores": (
        "NTT Data Services, Microsoft, EY, Harsco Rail, Stanley Black "
        "and Decker, TTI, Global Lending Services, G2O Technologies, "
        "Resurgent Capital Services, Epic, DXC Technology"
    ),
    "Destaques Currículo": (
        "30 créditos, analytics; departamento de Management Science "
        "venceu o 2024 Global INFORMS UPS Smith Prize (melhor ensino em "
        "operations research/analytics/data science, disputado com "
        "National University of Singapore e University of Waterloo)"
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
        raise ValueError("Nenhuma linha encontrada para U of South Carolina (Columbia, SC). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (U of South Carolina, Columbia, SC)")

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
