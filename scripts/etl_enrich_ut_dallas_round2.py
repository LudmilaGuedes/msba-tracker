# -*- coding: utf-8 -*-
"""
Data enrichment (round 2): atualiza APENAS a linha 'UT Dallas' (Naveen
Jindal School of Management, Richardson, TX) no msba_programs_database.xlsx.

Foco exclusivo: MS in Business Analytics and AI, trilha presencial/Flex
(autoritmada, full-time ou part-time, entrada Fall/Spring/Summer) -- a
mesma trilha ja documentada na rodada 1 (scripts/etl_enrich_ut_dallas.py).
NAO usados dados que pareciam ser das trilhas "Cohort" (36 creditos,
estrutura de tuition diferente) ou "Online Cohort", que sao ofertas
distintas do mesmo programa.

TUITION: RECONFIRMADA com fonte independente da trilha Flex especificamente
(nao mudou) -- $1.250/credito residente e $1.528/credito internacional,
30 creditos, totais $37.500 residente / $45.840 intl. Sem alteracao.

ATUALIZADO/DETALHADO com fonte oficial:
- Tipos de Bolsa: estava vago ("Dean's Excellence + competitive").
  jindal.utdallas.edu/student-resources/scholarships/graduate-scholarships/
  (oficial) detalha: Dean's Excellence Scholarship (exige envio OFICIAL de
  GMAT/GRE para concorrer, mesmo o teste sendo opcional para admissao
  regular), Dean's Impact Scholarship, Program Excellence Scholarship, e
  bolsas de doadores -- nenhuma com valor publico divulgado. Inscricao e
  via portal proprio, NAO automatica (contato:
  jsom-grad-scholarship@utdallas.edu).
- Quão Realista (Intl): mantido "Alta", mas detalhado o motivo real: existe
  um caminho FORMAL e conferido para tuition de residente via bolsa (ver
  abaixo) -- um beneficio financeiro concreto e incomum entre os programas
  ja pesquisados nesta base -- ainda que as bolsas de merito em si sejam
  competitivas, sem valor publico e exijam aplicacao separada.
- Caminho Tuition Residente: estava resumido ("Sim ($1k+ scholarship =
  in-state)"). CONFIRMADO como politica formal real (nao e boato) em
  finaid.utdallas.edu/aid-programs/exemptions-and-waivers/
  competitive-scholarships/ (oficial): bolsa competitiva institucional >=
  $1.000/ano academico (nao por semestre) da direito a pagar tuition de
  residente do Texas por ate 2 semestres consecutivos (Fall+Spring) ou 1
  semestre (Summer); exige manter GPA >= 3,000; NAO renova automaticamente;
  NAO e cumulativa entre bolsas; concedida a criterio exclusivo da
  universidade.
- GMAT/GRE Obrigatório: confirmado texto oficial ("not required... option
  to submit if desired") para admissao regular -- OPCIONAL, exceto para
  concorrer a Dean's Excellence Scholarship, que exige envio oficial.
  Codigos institucionais confirmados: GRE R6897, TOEFL 6897.
- Testes Inglês Aceitos: estava so a lista de testes sem notas ("TOEFL/
  IELTS/PTE/Duolingo"). CONFIRMADO com notas minimas oficiais
  (jindal.utdallas.edu/ms-business-analytics-faq/ e /admission-requirements/
  masters/): TOEFL 80 (iBT) / IELTS 6.5 / PTE 67 / Duolingo ACEITO mas nota
  minima nao encontrada em nenhuma fonte oficial.
- Documentos Necessários: detalhado -- confirmado que a exigencia de
  cartas de recomendacao e CONDICIONAL: 3 cartas e o padrao, mas reduz
  para 1 carta se o candidato tiver 2+ anos de experiencia profissional
  POS-graduacao (isso reconcilia numeros aparentemente conflitantes entre
  fontes distintas -- ambos "3" e "1" estao certos, dependendo do perfil).
  Acrescentados: personal statement (1 pagina), resume, transcripts.
- Deadline Aplicação: estava "Verificar". Preenchido com os deadlines
  regulares confirmados oficialmente (graduate-admissions.utdallas.edu/
  apply-to-ut-dallas/deadlines-and-fees/): Fall 1/mai, Spring 1/out,
  Summer 1/mar.
- Rounds Aplicação: estava "Verificar". A trilha Flex nao usa "rounds"
  competitivos -- e admissao por semestre com deadline unico por termo
  (ver Deadline Aplicação); preenchido para refletir essa estrutura.
- Principais Empregadores: estava generico ("Tech, Consulting"). Pagina
  oficial de career outcomes (infosystems.utdallas.edu/ms-business-
  analytics/career-outcomes) lista como logos de empregadores: Microsoft,
  Google, Amazon, Apple, Dell, JPMorgan Chase, Accenture, Deloitte --
  substituido pela lista nomeada oficial.
- Destaques Currículo: estava "11 tracks (Data Science, etc.)" -- essa
  contagem de 11 tracks NAO foi confirmada para a trilha Flex especifica
  (a pagina oficial da Flex nao lista concentracoes nomeadas; "tracks"
  nomeados como "Accounting Analytics"/"Data Science" parecem pertencer a
  trilha Cohort, uma oferta diferente -- ver nota da rodada 1 sobre nao
  misturar trilhas). Substituido pelo que a propria pagina oficial da Flex
  descreve: curriculo com tecnologias avancadas (Hadoop, Spark, R, Python,
  SAS), especializacao em IA (Predictive/Prescriptive Analytics, AI in
  Business, Large Language Models), opcao de internship.

NAO ALTERADO por falta de fonte oficial confiavel (mantido como estava,
incluindo "Verificar" onde ja era o caso):
- Taxa Empregabilidade % (83) e Salário Médio Pós USD (93.000): nenhum
  employment report oficial com numeros concretos foi encontrado para o
  MSBAi especificamente (a pagina oficial de career-outcomes nao expoe
  numeros no corpo da pagina, so PDFs antigos de 2016-2020 nao acessados).
  Uma fonte terceira nao-oficial citou uma faixa de $75.000-$85.000, mas
  por nao ser fonte primaria/oficial NAO foi usada para substituir os
  valores existentes (que tambem nao tem fonte registrada, mas nao ha base
  para dizer que estao errados).
- Fees Adicionais USD (300): fonte de um programa IRMAO (MS FTEC) sugere
  que isso corresponda a um "$300 non-refundable seat deposit" (taxa de
  reserva de vaga pos-admissao, nao recorrente), mas nao confirmado
  especificamente para a trilha BUAN Flex -- mantido sem alteracao.
- Tamanho Turma, GMAT Médio, Taxa Aceitação %, Ranking US News (especifico
  do programa presencial), Entrevista: nenhuma fonte oficial encontrada
  especifica do MSBAi Flex (numeros encontrados em buscas -- GMAT medio
  634, aceitacao 35-40%, turma 50 -- sao de fontes terceiras ou parecem
  ser do MBA, programa diferente; nao usados).
- TA/RA Disponível / Processo TA/RA / Graduate Assistantships: nenhuma
  fonte nova encontrada alem do que ja estava cadastrado.

NOTA ADICIONAL (fora do schema, sem coluna dedicada): a isencao da taxa de
aplicacao de $75 (fee waiver) e explicitamente indisponivel para um
candidato internacional externo -- so vale para alunos/alumni da propria
UT Dallas, cidadaos/residentes permanentes dos EUA em programa especifico,
ou militares/veteranos dos EUA. Fonte oficial:
graduate-admissions.utdallas.edu/apply-to-ut-dallas/deadlines-and-fees/.

Fontes principais (pesquisadas via WebSearch/WebFetch, set/2026):
- https://jindal.utdallas.edu/student-resources/scholarships/graduate-scholarships/
  (oficial -- tipos de bolsa, processo de aplicacao)
- https://finaid.utdallas.edu/aid-programs/exemptions-and-waivers/competitive-scholarships/
  (oficial -- politica formal de tuition de residente via bolsa >= $1.000)
- https://jindal.utdallas.edu/ms-business-analytics-faq/
  (oficial -- GMAT/GRE opcional, testes de ingles, cartas de recomendacao condicionais)
  (oficial -- requisitos de admissao)
- https://graduate-admissions.utdallas.edu/apply-to-ut-dallas/deadlines-and-fees/
  (oficial -- taxa de aplicacao $75, fee waiver so p/ alumni/militares, deadlines)
- https://infosystems.utdallas.edu/ms-business-analytics/career-outcomes
  (oficial -- logos de empregadores; sem numeros de salario/empregabilidade no corpo da pagina)
- https://infosystems.utdallas.edu/ms-business-analytics/
  (oficial -- confirma que e a pagina da trilha Flex; curriculo com Hadoop/
  Spark/R/Python/SAS, especializacao em IA, opcao de internship; NAO lista
  tracks/concentracoes nomeados)
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "UT Dallas",
    "Estado": "TX",
    "Cidade": "Richardson",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Tipos de Bolsa": (
        "Dean's Excellence Scholarship (exige envio OFICIAL de GMAT/GRE "
        "para concorrer, mesmo sendo opcional para admissão regular), "
        "Dean's Impact Scholarship, Program Excellence Scholarship, e "
        "bolsas de doadores -- nenhuma com valor público divulgado; "
        "inscrição via portal próprio, não automática "
        "(jsom-grad-scholarship@utdallas.edu)"
    ),
    "Quão Realista (Intl)": (
        "Alta -- existe um caminho FORMAL e confirmado para tuition de "
        "residente via bolsa competitiva ≥ $1.000/ano (ver 'Caminho "
        "Tuition Residente'), um benefício financeiro concreto e incomum "
        "entre os programas já pesquisados; ainda assim, as bolsas de "
        "mérito em si são competitivas, sem valor público divulgado e "
        "exigem aplicação separada (não são automáticas)"
    ),
    "Caminho Tuition Residente": (
        "Sim -- confirmado como política formal: bolsa competitiva "
        "institucional ≥ $1.000/ano acadêmico (não por semestre) dá "
        "direito a pagar tuition de residente do Texas por até 2 "
        "semestres consecutivos (Fall+Spring) ou 1 semestre (Summer); "
        "exige manter GPA ≥ 3,000; não renova automaticamente; não é "
        "cumulativa entre bolsas; concedida a critério exclusivo da "
        "universidade"
    ),
    "GMAT/GRE Obrigatório": (
        "Opcional para admissão regular (\"not required... option to "
        "submit if desired\"); OBRIGATÓRIO apenas para concorrer à Dean's "
        "Excellence Scholarship. Códigos institucionais: GRE R6897, "
        "TOEFL 6897"
    ),
    "Testes Inglês Aceitos": (
        "TOEFL mín. 80 (iBT) / IELTS mín. 6.5 / PTE mín. 67 / Duolingo "
        "aceito (nota mínima não encontrada em fonte oficial)"
    ),
    "Documentos Necessários": (
        "Personal statement (1 página), resume, transcripts, teste de "
        "proficiência em inglês (máx. 2 anos da candidatura, "
        "internacionais); 3 cartas de recomendação como padrão, "
        "reduzido para 1 carta se o candidato tiver 2+ anos de "
        "experiência profissional pós-graduação"
    ),
    "Deadline Aplicação": "Fall: 1/mai | Spring: 1/out | Summer: 1/mar (deadlines regulares)",
    "Rounds Aplicação": (
        "Não há rounds competitivos -- admissão por semestre (Fall/"
        "Spring/Summer) com um deadline regular por termo, não em rounds"
    ),
    "Principais Empregadores": "Microsoft, Google, Amazon, Apple, Dell, JPMorgan Chase, Accenture, Deloitte",
    "Destaques Currículo": (
        "Currículo com tecnologias avançadas (Hadoop, Spark, R, Python, "
        "SAS); especialização em IA com cursos em Predictive/Prescriptive "
        "Analytics, AI in Business e Large Language Models; opção de "
        "internship"
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
        raise ValueError("Nenhuma linha encontrada para UT Dallas (Richardson, TX). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (UT Dallas, Richardson, TX)")

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
