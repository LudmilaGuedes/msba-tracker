# -*- coding: utf-8 -*-
"""
Data enrichment (round 3): auditoria + pesquisa aprofundada focada em UMA
unica linha -- West Texas A&M University (WTAMU), Canyon, TX -- MS in
Computer Information Systems and Business Analytics (MSCISBA), Paul and
Virginia Engler College of Business, STEM.

Pedido da usuaria: Data Enrichment de uma unica instituicao, cobrindo
tuition, bolsas, taxa de aplicacao, requisitos completos de aplicacao,
empregabilidade/salario/empregadores e duracao/formato. Substituir dados
incorretos, preencher lacunas, manter "Verificar" para o que nao for
encontrado.

Fontes oficiais (wtamu.edu, verificadas por WebFetch direto em set/2026):
- Pagina do programa: wtamu.edu/academics/college-business/academics/graduate-programs/programs/mscisba-program.html
- FAQ de admissao (especifica do MSCISBA): wtamu.edu/academics/college-business/academics/graduate-programs/frequently-asked-admissions-questions.html
- Requisitos/prazos gerais de pos-graduacao: wtamu.edu/academics/graduate-school/apply/graduate-admission-requirements-deadlines.html
- Teste de ingles (pos-graduacao): wtamu.edu/student-support/international-student-services/Applicants/international-undergraduate-test-scores.html
- Custos do programa: wtamu.edu/academics/college-business/academics/graduate-programs/program-costs.html
- Alunos internacionais/graduacao: wtamu.edu/student-support/international-student-services/Applicants/INSTRUCTIONS TO APPLY/international-graduate.html

ACHADOS:

1) FORMATO CORRIGIDO: o programa NAO e puramente presencial -- a pagina
   oficial descreve como "100% online ou hibrido". Porem, internacionais
   em visto F-1 so podem contar 1 materia online por semestre (regra
   federal de visto), entao na pratica cursam majoritariamente
   presencial -- por isso o registro anterior de "Presencial" nao estava
   errado para o publico-alvo, mas estava incompleto/impreciso sobre a
   natureza do programa. Corrigido para refletir a realidade hibrida +
   a restricao pratica para F-1.

2) DURACAO PREENCHIDA (estava "Verificar"): FAQ oficial do MSCISBA
   confirma "as little as 12 to 18 months" em tempo integral, com 33 a
   36 creditos dependendo de nivelamento em estatistica exigido.

3) GPA MINIMO PREENCHIDO (estava "Verificar"): pagina oficial de
   requisitos/prazos de pos-graduacao confirma, especificamente para o
   MSCISBA e todas as especializacoes de MBA: GPA cumulativo da
   graduacao >= 3.0.

4) GMAT/GRE REFINADO: a dispensa ja estava registrada como disponivel;
   agora com os criterios exatos confirmados (GPA >=3.0 na graduacao OU
   nos ultimos 60 creditos, mestrado/doutorado concluido, ou 10+ anos de
   experiencia profissional progressiva) e o piso para quem NAO se
   qualifica pra dispensa: GMAT >= 500 (GRE aceito, convertido para
   equivalente GMAT).

5) TESTES DE INGLES ATUALIZADOS: a pagina oficial de pos-graduacao lista
   os minimos atuais (confirmados em set/2026): TOEFL 79 iBT (escala
   antiga, antes de 21/jan/2026) ou 4.5 (nova escala) / IELTS 6.5 / PTE
   58. O registro anterior tinha "IELTS 6.0" -- corrigido para 6.5
   (minimo de POS-GRADUACAO confirmado na fonte oficial; 6.0 parece ter
   sido o minimo de GRADUACAO, que e diferente). Duolingo ACEITO,
   preenchido pela primeira vez: minimo 105 (nao estava no registro
   anterior).

6) DOCUMENTOS NECESSARIOS DETALHADOS (estava generico): FAQ oficial do
   MSCISBA lista especificamente: resume/CV, Statement of Goals (250-500
   palavras), 2 cartas de recomendacao (pelo menos 1 profissional),
   transcripts.

7) DEADLINE DE APLICACAO ENRIQUECIDO: alem do prazo de PRIORIDADE de
   bolsas (1/fev) ja registrado, adicionados os prazos GERAIS de
   admissao confirmados oficialmente: Fall 1/ago, Spring 1/dez, Summer
   1/mai.

8) ROUNDS APLICACAO PREENCHIDO (estava "Verificar"): confirmado que nao
   ha sistema de rounds numerados -- e admissao por prazo de semestre.

9) EXPERIENCIA PROFISSIONAL PREENCHIDA (estava "Verificar"): nao e
   requisito geral de admissao -- os "10+ anos de experiencia" citados
   na FAQ sao apenas UM dos criterios alternativos pra dispensa de GMAT,
   nao uma exigencia para ser admitido no programa.

NAO ALTERADO (ja confirmado correto ou nao encontrado apos pesquisa
extensa, sem inventar dado):
- Tuition (Custo/Credito Intl USD 980, Residente USD 490, Totais 32.340
  e 16.170) -- RE-CONFIRMADOS EXATAMENTE iguais na pagina oficial de
  custos do programa. Sem mudanca.
- Taxa de Aplicacao (USD 90) -- confirmado, e confirmado tambem que NAO
  ha isencao disponivel (pagina oficial de admissao internacional:
  "$90... no waivers available").
- Bolsas/Quao Realista/Caminho Tuition Residente -- ja estavam com
  pesquisa detalhada e correta de uma rodada anterior, nao contradito
  por nada encontrado agora.
- Taxa de Empregabilidade %, Salario Medio Pos USD, Principais
  Empregadores, Tamanho Turma, Taxa de Aceitacao %, GPA Medio Admitidos,
  GMAT Medio, Entrevista, Dificuldade Entrada, Processo TA/RA, Fees
  Adicionais: permanecem "Verificar". A unica informacao de
  emprego/salario encontrada nas paginas oficiais da WTAMU e agregada em
  nivel de "Paul and Virginia Engler College of Business" (majoritariamente
  dados de MBA, ex.: "MBA graduates average salary $56,000", "95% MBA
  employed within six months") -- NAO especifica o MSCISBA e, pelo
  precedente ja identificado nesta sessao (contaminacao MBA/MSBA na
  linha da University of Washington), esses numeros foram
  deliberadamente DESCARTADOS em vez de atribuidos ao programa errado.

Uso: python etl_enrich_wtamu_round3.py (executar a partir da raiz do
repositorio).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "West Texas A&M University (WTAMU)",
    "Estado": "TX",
    "Cidade": "Canyon",
}

UPDATES = {
    "Duração (meses)": "12–18 (tempo integral; 33 a 36 créditos conforme nivelamento em estatística exigido)",
    "Formato": (
        "Híbrido (online + presencial; internacionais em F-1 só podem "
        "contar 1 matéria online por semestre, cursando majoritariamente "
        "presencial na prática)"
    ),
    "GPA Mínimo": "3.0 (GPA cumulativo da graduação, confirmado oficialmente para o MSCISBA)",
    "GMAT/GRE Obrigatório": (
        "Dispensa disponível (GPA graduação ≥3.0, GPA ≥3.0 nos últimos "
        "60 créditos, mestrado/doutorado concluído, ou 10+ anos de "
        "experiência profissional progressiva); sem dispensa, exige "
        "GMAT ≥500 (GRE aceito, convertido para equivalente GMAT)"
    ),
    "Testes Inglês Aceitos": (
        "TOEFL iBT mín. 79 (antes de 21/jan/2026) ou 4.5 (nova escala) / "
        "IELTS mín. 6.5 / Duolingo mín. 105 / PTE mín. 58 (mínimos gerais "
        "de pós-graduação da WTAMU)"
    ),
    "Documentos Necessários": (
        "Resume/CV, Statement of Goals (250-500 palavras), 2 cartas de "
        "recomendação (pelo menos 1 profissional), transcripts, "
        "comprovação de proficiência em inglês (internacionais), taxa de "
        "aplicação ($90, sem isenção disponível)"
    ),
    "Deadline Aplicação": (
        "Admissão geral: Fall 1/ago; Spring 1/dez; Summer 1/mai. "
        "Prioridade de bolsas: 1/fev (para o próximo ano acadêmico)"
    ),
    "Rounds Aplicação": "Sem rounds numerados -- admissão dentro dos prazos por semestre (ver Deadline Aplicação)",
    "Experiência Profissional": (
        "Não obrigatória para admissão -- 10+ anos de experiência "
        "profissional progressiva é apenas um dos critérios alternativos "
        "para dispensa de GMAT, não um requisito geral de admissão"
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
        raise ValueError("Nenhuma linha encontrada para West Texas A&M University (Canyon, TX). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (West Texas A&M University, Canyon, TX)")

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

    print(f"\n{len(UPDATES)} campos atualizados na linha {row_idx} (West Texas A&M University):\n")
    for header in UPDATES:
        print(f"  - {header}:")
        print(f"      Antes:  {before[header]!r}")
        print(f"      Depois: {after[header]!r}")


if __name__ == "__main__":
    main()
