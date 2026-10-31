# -*- coding: utf-8 -*-
"""
Data enrichment (round 2): auditoria + pesquisa aprofundada focada em UMA
unica linha -- U of Rochester (Simon Business School), Rochester, NY --
Full-Time MS in Business Analytics, presencial, STEM.

Pedido da usuaria: Data Enrichment de uma unica instituicao, cobrindo
tuition, bolsas, taxa de aplicacao, requisitos completos de aplicacao,
empregabilidade/salario/empregadores e duracao/formato. Substituir dados
incorretos, preencher lacunas, manter "Verificar" para o que nao for
encontrado.

Fontes oficiais (simon.rochester.edu, verificadas por WebFetch/WebSearch
em set/2026):
- Tuition/financial aid: simon.rochester.edu/programs/full-time-ms/tuition-financial-aid
- Curriculo: simon.rochester.edu/programs/full-time-ms-in-business-analytics/curriculum
- Pagina do programa: simon.rochester.edu/programs/full-time-ms-in-business-analytics
- Admissoes: simon.rochester.edu/programs/full-time-ms/admissions
- FAQs de admissao: simon.rochester.edu/programs/full-time-ms/admissions/admissions-faqs
- Blog "Sneak Peek 2026 Full-Time Application Requirements": simon.rochester.edu/blog/admissions-blog/sneak-peek-2026-full-time-application-requirements
- Carreiras/empregadores: simon.rochester.edu/programs/full-time-ms/careers-outcomes e
  simon.rochester.edu/programs/full-time-ms-in-business-analytics/academics/career-paths

ACHADOS PRINCIPAIS:

1) TUITION: confirmado USD 68.000/ano para 2026-27 (ja estava correto em
   "Tuition Total (Intl)"). Escola PRIVADA -- sem distincao residente x
   nao-residente (mesma mensalidade pra todos). "Tuition Total
   (Residente)" estava "Verificar"; preenchido com o mesmo valor (68.000),
   seguindo a mesma convencao ja usada em todas as outras faculdades
   privadas desta base (ex.: MIT, Columbia, Georgetown -- residente =
   internacional quando a escola e privada).

2) TOTAL DE CREDITOS: pagina oficial do curriculo NAO divulga o programa
   em creditos, apenas em "13 cursos" (courses), com opcao de 2 semestres
   (sem estagio) ou 3 semestres (com estagio de verao). Como "curso" e
   "credito" nao sao a mesma unidade (nao da pra converter com
   confianca), "Total Creditos" e os dois campos "Custo/Credito"
   permanecem "Verificar" -- nao inventado. O detalhe de 13 cursos /
   estrutura por semestre foi registrado em "Destaques Curriculo" em vez
   disso.

3) *** ACHADO NOVO IMPORTANTE: O programa TAMBEM aceita entrada na
   PRIMAVERA (Spring), nao so Fall *** -- o campo "Inicio" estava
   registrado soh como "Fall", o que estava incompleto. A propria pagina
   do programa esta ativamente divulgando "Begin your degree in January
   2027!" com um processo de aplicacao proprio (2 rounds: 15/out/2026 e
   1/dez/2026), separado dos 4 rounds ja registrados para entrada em
   Fall. Corrigido "Inicio" para "Fall, Spring" e os campos de deadline
   para refletir ambos os ciclos.

4) MATERIAS PRE-REQUISITO preenchido (estava "Verificar"): confirmado
   oficialmente na FAQ de admissoes que o MS in Business Analytics NAO
   tem pre-requisitos formais e NAO exige experiencia previa em
   programacao.

5) TAXA DE APLICACAO: valor USD 90 confirmado correto (sem mudanca), mas
   NOVO detalhe encontrado: ha isencao de taxa disponivel para militares
   ativos/veteranos e "outros candidatos elegiveis" (criterio exato nao
   detalhado nas paginas encontradas) -- registrado em "Rounds
   Aplicacao", ja que nao ha campo proprio no schema pra isencao de taxa.

6) TESTES DE INGLES: confirmados EXATAMENTE iguais ao que ja estava
   registrado (TOEFL 100, IELTS 7.0/sem secao abaixo de 6.5, Duolingo
   125) -- sem mudanca.

7) EMPREGABILIDADE E SALARIO: confirmados EXATAMENTE iguais aos ja
   registrados (73% ofertas em 6 meses; USD 75.801 salario+bonus medio,
   relatorio oficial de abr/2026 sobre a turma de 2025) -- sem mudanca.

8) EMPREGADORES: lista oficial da pagina de "career paths" e bem mais
   completa que a ja registrada (que so tinha 4 nomes, todos confirmados
   como corretos, apenas incompletos). Expandida para incluir mais
   empregadores confirmados oficialmente.

9) GPA MINIMO: nenhuma pagina oficial divulga um GPA minimo formal para
   este programa (so agregadores terceiros com estimativas nao-oficiais,
   ex. "3.2-3.8" -- NAO usado por nao ser fonte primaria). Permanece
   "Verificar".

10) ENTREVISTA (de admissao, campo separado do processo de entrevista
    para GA/assistantship que ja estava preenchido): nao encontrada
    mencao oficial confirmando se ha ou nao entrevista de admissao.
    Permanece "Verificar".

NAO ALTERADO (fora do escopo pedido ou ja confirmado sem discrepancia):
Ranking US News, Ranking QS, Taxa de Aceitacao % (permanecem
"Verificar" -- nao encontrados com confianca); Bolsas
Disponiveis/Tipos de Bolsa/Quao Realista (Intl) -- ja confirmados
identicos na pagina oficial de tuition/financial aid (mérito automático,
>90% dos admitidos recebem, bolsa garantida de $15k para ex-alunos da
UofR, minimo $10k/ano para parceiros); Deadline/Rounds de Fall (datas
15/out, 5/jan, 1/mar, 1/mai ja estavam corretas).

Uso: python etl_enrich_rochester_round2.py (executar a partir da raiz do
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
    "Nome da Universidade": "U of Rochester",
    "Estado": "NY",
    "Cidade": "Rochester",
}

UPDATES = {
    "Início": "Fall, Spring",
    "Tuition Total (Residente) USD": 68000,
    "Destaques Currículo": (
        "Two tracks (internship/non-internship); 13 cursos -- track sem "
        "estágio = 2 semestres, track com estágio de verão = 3 semestres"
    ),
    "Matérias Pré-Requisito": (
        "Nenhum pré-requisito oficial -- não é exigida experiência prévia "
        "em programação (confirmado oficialmente para o MS in Business "
        "Analytics)"
    ),
    "Principais Empregadores": (
        "Bain & Company, Deloitte, Goldman Sachs, TikTok, Capital One, "
        "PwC, Morgan Stanley, Procter & Gamble, Alibaba Group, BlackRock"
    ),
    "Deadline Aplicação": (
        "Spring 2027: Round 1 15/out/2026 (decisão 15/dez); Round 2 "
        "1/dez/2026 (decisão em 2-3 semanas). Fall: Round 1 15/out "
        "(decisão 15/dez); Round 2 5/jan (decisão 15/mar); Round 3 1/mar "
        "(decisão 15/abr); Round 4 1/mai (decisão 1/jun)"
    ),
    "Rounds Aplicação": (
        "Spring 2027 (2 rounds) e Fall (4 rounds) -- ver Deadline "
        "Aplicação para as datas de cada termo. Taxa de aplicação ($90) "
        "tem isenção disponível para militares ativos/veteranos e outros "
        "candidatos elegíveis (critério exato não detalhado nas páginas "
        "oficiais encontradas -- confirmar diretamente)."
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
        raise ValueError("Nenhuma linha encontrada para U of Rochester (Rochester, NY). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (U of Rochester, Rochester, NY)")

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

    print(f"\n{len(UPDATES)} campos atualizados na linha {row_idx} (U of Rochester):\n")
    for header in UPDATES:
        print(f"  - {header}:")
        print(f"      Antes:  {before[header]!r}")
        print(f"      Depois: {after[header]!r}")


if __name__ == "__main__":
    main()
