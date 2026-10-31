# -*- coding: utf-8 -*-
"""
Data enrichment: atualiza APENAS a linha da 'UT Austin' (McCombs School of
Business, Austin, TX) no msba_programs_database.xlsx.

Foco exclusivo: MS Business Analytics and AI (MSBAi), campus Austin,
presencial (on-campus), 10 meses, STEM-designated. Deep research (set/2026)
em tuition, bolsas/TA-RA, taxa de aplicação, requisitos completos de
admissão, empregabilidade/salário e duração/formato.

CONFIRMADO SEM ALTERACAO (auditoria bateu com a planilha, mantido como
estava): Nome Oficial do Programa "MS Business Analytics and AI (MSBAi)"
(nome oficial mudou de "MS Business Analytics (MSBA)" para "MS Business
Analytics and AI" em ago/2026 -- a planilha ja estava certa); Tuition Total
Intl USD$59.700 e Residente USD$55.600 (bate exatamente com a pagina oficial
de tuition-and-financial-aid, incluindo a taxa de pre-matricula de $2.000
ja embutida); Total Creditos 36; Duracao 10 meses; Formato Presencial;
Tamanho Turma 64 (confirmado na pagina oficial de class-stats para o campus
presencial, Class of 2027); GPA Medio Admitidos 3.55 e GMAT Medio 655
(ambos batem exatamente com a mesma pagina); Ranking QS 12 (QS Global
Business Masters Rankings 2026, categoria Business Analytics); Taxa
Aplicacao USD$90 (internacional; domestico e $65); GMAT/GRE Opcional com
recomendacao de >= 60o percentil; Testes de Ingles (TOEFL 100/IELTS 7,
Duolingo aceito sem minimo publicado); Deadlines dos 4 rounds (13/out/2026,
12/jan/2027 -- final para internacionais --, 30/mar/2027, 4/mai/2027).

CORRIGIDO/ATUALIZADO (fontes oficiais McCombs, ver lista abaixo):
- TA/RA Disponivel, Processo TA/RA, Graduate Assistantships: estavam
  "Verificar". A propria pagina oficial de tuition-and-financial-aid do
  MS Program (que cobre o MSBAi) afirma textualmente que "MS Program
  students are not typically eligible for university employment as
  teaching or research assistants" e que "outside employment is not
  generally recommended" dada a carga do curso -- ou seja, TA/RA
  essencialmente NAO existe como via de bolsa para este programa.
- Taxa Empregabilidade %: estava 91 (fonte nao identificada/desatualizada).
  O Employment Report OFICIAL da Class of 2025 (McCombs, publicado
  mar/2026) nao reporta uma taxa unica de "empregabilidade", mas reporta
  timing de ofertas: 81,4% dos formandos que buscavam emprego receberam a
  1a oferta ate 6 meses da formatura (36,3% ja na formatura). Substituido
  por 81,4 com nota explicita, por ser o dado oficial mais recente
  disponivel.
- Salario Medio Pos USD: estava 109.481 (fonte nao identificada). O
  Employment Report oficial da Class of 2025 reporta media geral de
  $104.798 (mediana $105.000, n=72 reportando). Para candidatos
  internacionais especificamente, a media e MAIOR: $117.734 (mediana
  $120.000, n=26) -- registrado aqui no comentario porque o schema so tem
  um campo de salario medio geral, sem quebra por nacionalidade.
- Principais Empregadores: atualizado com a lista oficial "Select
  Recruiting Employers 2025" (Capital One, Walmart, Deloitte, CGI, AMD,
  American Airlines, Wells Fargo, Discover, GEICO, Wayfair) e o
  detalhamento oficial por setor da Class of 2025 (Financial Services 26%,
  Technology 21%, Consulting 21%), substituindo a lista anterior (IBM,
  Dell, Toyota Connected; setores diferentes) que nao bateu com o
  relatorio oficial mais recente.
- Materias Pre-Requisito: mantido Calculo + Estatistica (ja confirmado) e
  ACRESCENTADO: experiencia/coursework obrigatorio em programacao (C,
  C++, Python ou Java) e Algebra Linear fortemente recomendada -- ambos
  confirmados na pagina oficial de "program readiness" mas ausentes na
  planilha.
- Destaques Curriculo: "Tech/Energy analytics" nao foi encontrado na
  pagina oficial atual; substituido pelas trilhas de eletivas
  efetivamente descritas hoje pelo McCombs (Financial Analytics OU Supply
  Chain and Marketing Analytics), mantendo a mencao a Capstone (nao
  recontrada nesta rodada, mas nao contradita).
- Entrevista: estava "Seletiva (nao universal)". A pagina oficial de
  application-process lista o video assessment via Kira Talent como
  "Step 3" do processo padrao (ate 30 min) -- ou seja, e uma etapa
  padrao para o funil de aplicacao, nao uma entrevista seletiva so para
  avancados. Corrigido para refletir isso; nao ha confirmacao de uma
  entrevista ao vivo adicional.
- Documentos Necessarios: detalhado com a contagem oficial de ensaios (2
  obrigatorios + 1 opcional, ate 500 palavras cada) e reafirmado o video
  assessment via Kira Talent (Step 3, ate 30 min).

NAO ALTERADO por falta de fonte confiavel (mantido "Verificar"): Custo/
Credito (Intl e Residente) -- o McCombs cobra em blocos por
semestre/verao, nao achei uma tarifa oficial por credito publicada, e
dividir o total por 36 creditos seria uma estimativa nao confirmada, o
que a convencao do projeto probe; Seguro Saude (programa); Ranking US
News (US News nao rankeia MS em Business Analytics separadamente, so MBA
geral); Taxa de Aceitacao % (nao divulgada publicamente).

NOTA ADICIONAL (fora do schema, sem coluna dedicada): a isencao de taxa de
aplicacao (fee waiver) da UT Austin Graduate School e explicitamente
INDISPONIVEL para candidatos internacionais -- so vale para cidadaos/
residentes permanentes dos EUA com necessidade financeira documentada.
Relevante para candidatos brasileiros mas sem campo no schema para
registrar isso; citado aqui e no resumo do terminal.

Fontes principais (pesquisadas manualmente/via WebSearch e WebFetch,
set/2026):
- https://www.mccombs.utexas.edu/graduate/specialized-masters/ms-business-analytics-and-artificial-intelligence/on-campus/admissions/tuition-financial-aid/
  (tuition total intl/residente, taxa de pre-matricula, bolsas de merito
  automaticas, TA/RA nao tipicamente disponivel para MS Program students)
- https://www.mccombs.utexas.edu/graduate/specialized-masters/ms-business-analytics-and-artificial-intelligence/on-campus/
  (nome oficial do programa, duracao 10 meses, 36 creditos, formato)
- https://www.mccombs.utexas.edu/graduate/specialized-masters/ms-business-analytics-and-artificial-intelligence/on-campus/admissions/
  (pre-requisitos: calculo, estatistica, programacao, algebra linear
  recomendada)
- https://www.mccombs.utexas.edu/graduate/specialized-masters/ms-business-analytics-and-artificial-intelligence/on-campus/admissions/application-process/
  (GMAT/GRE opcional, TOEFL/IELTS/DET, documentos, 4 rounds 2026-27,
  video assessment via Kira Talent como Step 3)
- https://www.mccombs.utexas.edu/graduate/specialized-masters/ms-business-analytics-and-artificial-intelligence/on-campus/admissions/faq/
  (TOEFL 100/IELTS 7 esperados; fee waiver processado pela UT Graduate
  Admissions)
- https://www.mccombs.utexas.edu/graduate/specialized-masters/ms-business-analytics/ms-business-analytics-on-campus/admissions/class-stats/
  (Tamanho Turma 64, GPA medio 3.55, GMAT medio 655, Class of 2027)
- https://www.mccombs.utexas.edu/wp-content/uploads/2026/03/25-MSBA-Employment-Report.pdf
  (relatorio oficial de emprego Class of 2025: salario medio $104.798,
  mediana $105.000; internacionais $117.734 media/$120.000 mediana;
  81,4% receberam 1a oferta ate 6 meses; empregadores e setores)
- https://www.topuniversities.com/business-masters-rankings/business-analytics
  (QS Global Business Masters Rankings 2026, Business Analytics: UT
  Austin McCombs #12)
- https://liberalarts.utexas.edu/government/graduate-program/prospective-students/admissions/application-fee-waiver-information.html
  e busca sobre "UT Austin graduate school application fee waiver
  eligibility" (fee waiver da UT Graduate Admissions indisponivel para
  candidatos internacionais)
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
    "Nome da Universidade": "UT Austin",
    "Estado": "TX",
    "Cidade": "Austin",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "TA/RA Disponível": (
        "Não -- MS Program students (inclui MSBAi) não são tipicamente "
        "elegíveis para emprego universitário como teaching/research "
        "assistant"
    ),
    "Processo TA/RA": (
        "Não há processo formal -- a própria página oficial de tuition/"
        "financial aid do MS Program afirma que outside employment não é "
        "geralmente recomendado dada a carga rigorosa do curso"
    ),
    "Graduate Assistantships": (
        "Não disponível para MS Program students como regra geral do "
        "McCombs (diferente de programas de PhD)"
    ),
    "Taxa Empregabilidade %": 81.4,  # Class of 2025: 81,4% receberam 1a oferta até 6 meses da formatura (relatório oficial); 36,3% já na formatura. Substitui o valor anterior (91), sem fonte identificada.
    "Salário Médio Pós USD": 104798,  # Class of 2025, relatório oficial: média geral $104.798 (mediana $105.000, n=72). Internacionais especificamente: média $117.734 / mediana $120.000 (n=26) -- sem campo próprio no schema para essa quebra.
    "Principais Empregadores": (
        "Capital One, Walmart, Deloitte, CGI, AMD, American Airlines, "
        "Wells Fargo, Discover, GEICO, Wayfair; setores (Class of 2025): "
        "Financial Services (26%), Technology (21%), Consulting (21%)"
    ),
    "Matérias Pré-Requisito": (
        "1 semestre de Cálculo + 1 semestre de Estatística (nível "
        "universitário); coursework ou experiência profissional em "
        "programação obrigatório (C, C++, Python ou Java); Álgebra "
        "Linear fortemente recomendada"
    ),
    "Destaques Currículo": (
        "Analytics aplicado à tomada de decisão empresarial; eletivas em "
        "Financial Analytics ou em Supply Chain and Marketing Analytics; "
        "Capstone"
    ),
    "Entrevista": (
        "Video assessment obrigatório via Kira Talent (até 30 min) -- "
        "Etapa 3 padrão do processo para todos os candidatos, não uma "
        "entrevista seletiva restrita a avançados; entrevista ao vivo "
        "adicional não confirmada"
    ),
    "Documentos Necessários": (
        "Transcripts oficiais, test scores (GMAT/GRE opcional; TOEFL/"
        "IELTS/Duolingo obrigatório para internacionais), 2 ensaios "
        "obrigatórios + 1 opcional (até 500 palavras cada), histórico "
        "profissional, 2 cartas de recomendação, taxa de aplicação, "
        "video assessment via Kira Talent (Etapa 3, até 30 min)"
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
        raise ValueError("Nenhuma linha encontrada para UT Austin (Austin, TX). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (UT Austin, Austin, TX)")

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
