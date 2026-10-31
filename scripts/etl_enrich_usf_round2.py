"""
Data enrichment (round 2): atualiza APENAS a linha da 'USF' (University of
South Florida, Tampa, FL) no msba_programs_database.xlsx.

ATENCAO CRITICA DE DESAMBIGUACAO (repetida do round 1): NAO confundir com
University of San Francisco (usfca.edu / myusf.usfca.edu). Durante esta
rodada, resultados de busca com "myUSF" -> "myusf.usfca.edu" (San
Francisco) apareceram misturados e foram DESCARTADOS integralmente (ex.:
paginas de "Health Insurance FAQ" e "Billing & Tuition" da usfca.edu).
Matching de escrita continua exigindo Nome da Universidade == 'USF' E
Estado == 'FL' E Cidade == 'Tampa'.

Foco exclusivo: MS in Artificial Intelligence (AI) in Business -- trilha
'Enterprise Integration' (presencial, STEM), sucessora do MS-BA/MS-AIB
documentada no round 1. Deep research em: custo de tuition p/
internacionais, bolsas/assistantships, GPA/GMAT/GRE/inglês, custo de vida.

TUITION: CONFIRMADO sem mudanca. $466,34/crédito residente ($13.990,20
total) e $1.030,20/crédito não-residente ($30.906,00 total), 30 créditos --
reconfirmado direto na pagina oficial do programa (fonte igual ao round 1,
ainda vigente para Fall 2026). Uma noticia da imprensa universitaria (USF
Oracle, "second out-of-state tuition hike") menciona uma taxa de
não-residente GERAL da universidade subindo p/ $537,01/crédito -- mas essa
e a taxa PADRAO de graduate (nao a taxa 'premium' especifica do Muma
College/MS-AIB, que ja e mais alta que isso: $1.030,20). Nao foi usada p/
sobrescrever o valor ja confirmado diretamente na pagina oficial do
programa (mais especifico e mais recente).

ACHADO CRITICO -- REQUISITOS DE ADMISSAO AINDA NAO PUBLICADOS PARA OS NOVOS
MAJORS: reportagem da imprensa (collegedunia.com, ago/2026) confirma que a
politica de GMAT/GRE para os 2 novos majors (Digital Transformation e
Enterprise Integration) "have not been released yet" e que a USF Muma
"is expected to publish firm application deadlines, GRE/GMAT policy and
international scholarship details... ahead of the Fall 2026 semester
start" -- checagem direta das paginas oficiais nesta rodada (application-
process.aspx, faq.aspx, e tentativas de admission-requirements.aspx que
retornaram 404) CONFIRMA que essas paginas ainda nao tem essa informacao
publicada, mesmo com o semestre ja tendo comecado. 'GMAT/GRE Obrigatório'
mantido com o texto do round 1 (politica do PREDECESSOR, carregada como
proxy), mas com ressalva explicita adicionada de que a politica dos novos
majors segue nao-anunciada. 'Deadline Aplicação' e 'Rounds Aplicação'
preenchidos com "admissão contínua/rolling" (confirmado p/ ambas as
paginas oficiais e pela reportagem).

CUSTO DE VIDA (fonte oficial nova):
https://tup-ofa.forest.usf.edu/finaid/coa/2526_on_campus -- Cost of
Attendance oficial da USF p/ GRADUATE, ano 2025-26 (nao achei uma versao
2026-27 publicada; usado como melhor fonte oficial disponivel), categoria
Non-Florida Resident, moradia NO CAMPUS (nao ha versao off-campus
disponibilizada nesta pagina), base de 20 creditos/ano (~9 meses,
Fall+Spring):
  - Housing: $9.224 / 9 = $1.024,89/mês -> 'Aluguel Privativo/mês'
  - Food: $5.600 / 9 = $622,22/mês -> 'Mercado/mês'
  - Transportation: $2.176 / 9 = $241,78/mês -> 'Transporte/mês'
  - Personal Allowances: $2.944 / 9 = $327,11/mês -> 'Outros Gastos/mês'
  - Books & Supplies: $612 (custo único, sem campo dedicado no schema)
'Custo Vida Total (programa)' PERMANECE "Verificar": a COA oficial cobre
so ~9 dos 12-16 meses de duracao do programa (nao ha breakdown do(s)
semestre(s) extra), entao somar daria um total sistematicamente
subestimado -- mesmo criterio conservador do round 1.

SEGURO SAUDE: NAO preenchido para o caso geral (nao-GA) -- as paginas
oficiais dedicadas (insurance-regulation.aspx, inscompliance.aspx) NAO
publicam o valor do premium em texto acessivel (redirecionam pro portal
AHP/UHC autenticado). Um valor de "$2.517/ano" apareceu em busca mas o
snippet misturava resultados com dominio usfca.edu (ver aviso de
desambiguacao acima) -- descartado por falta de confianca na fonte.
Mantido "Verificar".

BOLSAS E ASSISTANTSHIPS (novo, antes "Verificar"):
- https://www.usf.edu/admissions/international/admission-information/cost-of-attendance/assistantships.aspx
  Politica geral de GA da USF (aplica-se a todo grad, incl. Muma):
  competitivas, exige GPA geral e do programa >= 3.00, matricula em tempo
  integral (9 créditos outono/primavera, 6 verão), proficiência em inglês
  se aplicável. FTE < 0.25 NAO da direito a tuition waiver. Candidatura via
  portal "Work at USF" (RH) ou contato direto com o diretor do programa de
  pós-graduação.
- https://www.usf.edu/hr/benefits/insurance/graduate-assistants.aspx
  Seguro saúde SUBSIDIADO para GAs: responsabilidade do aluno de apenas
  $294 (Fall 2026) + $490 (Spring 2027) = $784/ano (vs. o valor cheio do
  SHIP, nao confirmado nesta rodada -- ver nota acima).
'Quão Realista (Intl)' preenchido como "Moderada" -- mesmo padrao
qualitativo encontrado em outras escolas do dataset (FSU, UC): GA existe e
pode isentar tuition total ou parcial, mas e explicitamente competitivo e
dependente de orcamento/vaga do departamento, sem numero de aprovacao
publicado.

CAMINHO TUITION RESIDENTE: preenchido como "Não" -- mesma lei estadual da
Flórida (Regulation 7.005) ja confirmada para a FSU nesta mesma sessao de
trabalho: alunos com visto F-1 nao sao elegiveis a tuition de residente
enquanto mantiverem esse status. Regra e estadual (nao especifica da USF),
entao reaproveitada com confianca.

ESTATISTICAS DE TURMA/EMPREGABILIDADE (Tamanho Turma, Ranking US News,
Ranking QS, Taxa Empregabilidade %, Salário Médio Pós USD, Principais
Empregadores, GPA Médio Admitidos, GMAT Médio, Taxa Aceitação %):
PERMANECEM "Verificar" -- nao apenas por falta de fonte, mas porque o
programa e uma REFORMULACAO QUE ACABOU DE LANCAR (Fall 2026, primeira
turma == a que esta entrando agora); estruturalmente nao pode existir
ainda taxa de empregabilidade/salario de egressos (nao ha egressos), nem
GPA/GMAT medio de admitidos publicado (perfil da 1a turma normalmente so e
divulgado depois de fechada a matricula, se for divulgado). Rankings de
programa tambem levam tempo para serem atribuidos a um programa novo.
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "USF",
    "Estado": "FL",
    "Cidade": "Tampa",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Aluguel Privativo/mês USD": 1024.89,  # 9224 / 9 (housing on-campus, COA oficial 2025-26, não-residente)
    "Mercado/mês USD": 622.22,  # 5600 / 9 (food)
    "Transporte/mês USD": 241.78,  # 2176 / 9 (transportation)
    "Outros Gastos/mês USD": 327.11,  # 2944 / 9 (personal allowances)
    "GMAT/GRE Obrigatório": (
        "Obrigatório, salvo dispensa (alumni da Muma College com GPA ≥ "
        "3.5, ou 3+ anos de experiência técnica relevante) -- política do "
        "programa PREDECESSOR (MS-AIB), usada como proxy. RESSALVA: a "
        "política de GMAT/GRE para os 2 novos majors (Digital "
        "Transformation e Enterprise Integration) ainda não foi anunciada "
        "publicamente (confirmado ago/2026, mesmo já com o semestre "
        "iniciado)"
    ),
    "Deadline Aplicação": (
        "Admissão contínua/rolling -- os novos majors (lançados Fall "
        "2026) não têm deadline fixo publicado; decisões seguem por "
        "ordem de chegada até fechar a capacidade da turma"
    ),
    "Rounds Aplicação": "Sem rounds -- admissão contínua/rolling",
    "Caminho Tuition Residente": (
        "Não -- alunos com visto F-1 não são elegíveis a tuition de "
        "residente na Flórida (FL Regulation 7.005, regra estadual, "
        "mesma aplicada a todas universidades públicas do estado)"
    ),
    "Bolsas Disponíveis": "Sim",
    "Tipos de Bolsa": (
        "Graduate Assistantships (GA) via portal 'Work at USF' (RH) ou "
        "contato direto com o diretor do programa; cobrem stipend + "
        "tuition waiver (parcial ou total conforme FTE; < 0.25 FTE não "
        "dá direito a waiver); seguro saúde subsidizado para GAs (~$294 "
        "Fall 2026 + $490 Spring 2027 de responsabilidade do aluno). "
        "Detalhes de bolsa internacional específica para os novos majors "
        "ainda não anunciados (mesma situação do GMAT/GRE)"
    ),
    "Processo TA/RA": (
        "Concorrência: GPA geral e do programa ≥ 3.00, matrícula em "
        "tempo integral (9 créditos outono/primavera, 6 no verão), "
        "proficiência em inglês se aplicável. Candidatura via portal "
        "'Work at USF' ou contato direto com o diretor do programa de "
        "pós-graduação; disponibilidade varia por financiamento do "
        "departamento -- caráter competitivo, não garantido"
    ),
    "Graduate Assistantships": (
        "Stipend + tuition waiver (parcial ou total conforme FTE); "
        "seguro saúde subsidizado (~$294 Fall + $490 Spring de "
        "responsabilidade do aluno). Competitivas, dependem de "
        "orçamento/vaga do departamento -- não garantidas com a admissão"
    ),
    "Quão Realista (Intl)": (
        "Moderada -- GA existe e pode isentar tuition total ou parcial, "
        "mas é explicitamente competitiva e depende de orçamento/vaga do "
        "departamento; sem taxa de aprovação publicada"
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
        raise ValueError("Nenhuma linha encontrada para USF (Tampa, FL). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (USF, Tampa, FL) -- NAO University of San Francisco")

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
