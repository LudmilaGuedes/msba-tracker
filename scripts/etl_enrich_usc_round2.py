# -*- coding: utf-8 -*-
"""
Data enrichment (round 2): auditoria + pesquisa aprofundada focada em UMA
unica linha -- USC (Marshall School of Business), Los Angeles, CA -- MS
in Business Analytics (MSBA), presencial, STEM.

Pedido da usuaria: Data Enrichment de uma unica instituicao, cobrindo
tuition, bolsas, taxa de aplicacao, requisitos completos de aplicacao,
empregabilidade/salario/empregadores e duracao/formato. Substituir dados
incorretos, preencher lacunas, manter "Verificar" para o que nao for
encontrado.

*** ATENCAO DE NOME: "USC" (University of Southern California, Los
Angeles, Trojans) e "University of South Carolina" (Columbia, SC,
tambem abreviada as vezes "USC" em buscas genericas) sao instituicoes
DIFERENTES. Uma busca por "USC health insurance" retornou dados da
University of South Carolina -- descartados explicitamente (ver item 8
abaixo) apos a propria busca alertar sobre a ambiguidade. Todo dado
usado neste script foi confirmado como vindo de dominios
marshall.usc.edu, arr.usc.edu, studenthealth.usc.edu ou gradadm.usc.edu
(todos University of Southern California, Los Angeles). ***

Fontes oficiais (verificadas por WebFetch direto em set/2026):
- marshall.usc.edu/programs/graduate-programs/specialized-masters/ms-business-analytics/admissions
- marshall.usc.edu/programs/graduate-programs/specialized-masters/ms-business-analytics/class-profile
- marshall.usc.edu/programs/graduate-programs/specialized-masters/ms-business-analytics/tuition-and-fees
- marshall.usc.edu/graduate-scholarships
- gradadm.usc.edu/prospective-international-students/english-proficiency/

ACHADOS:

1) TAXA DE APLICACAO CORRIGIDA: estava registrada como USD 155; a
   pagina oficial de admissoes do MSBA confirma USD 175. Sem isencao
   mencionada em nenhuma fonte oficial encontrada.

2) DEADLINE/ROUNDS DE APLICACAO ATUALIZADOS: o registro anterior tinha
   so "Mid-April (intl)", um dado vago e incompleto. A pagina oficial
   de admissoes lista o ciclo completo de 3 rounds para entrada em Fall
   2027 (o proximo ciclo a partir de set/2026): Round 1 1/nov/2026
   (decisao 15/jan/2027); Round 2 15/jan/2027 (decisao 30/mar/2027);
   Round 3 15/abr/2027 (decisao 15/jun/2027). Internacionais sao
   recomendados a aplicar nos Rounds 1-2 para dar tempo ao processo de
   visto.

3) ENTREVISTA PREENCHIDA (estava "Verificar"): confirmado que e via
   KIRA Video Interview, por convite apenas, depois da submissao da
   candidatura.

4) TESTES DE INGLES CORRIGIDOS: o IELTS estava registrado como "7.0" --
   a politica oficial de proficiencia em ingles da USC para pos-graduacao
   (nivel MESTRADO, nao doutorado) confirma minimo de 6.5 (com minimo de
   6.0 por secao), nao 7.0 (esse seria o minimo de PhD). Corrigido.
   TOEFL mantido em 100 (registro anterior, que pode ser um minimo
   proprio do MSBA mais alto que o piso geral da USC de 90 -- programas
   especificos podem exigir mais que o minimo geral, e o registro
   anterior tinha detalhe de secao ja especifico o suficiente pra
   sugerir fonte confiavel). PTE detalhado com os numeros exatos (61
   geral / 53 por secao) que faltavam. Duolingo mantido como estava
   (aceito mas sem minimo formal de admissao, usado so para
   'placement') -- e um achado MSBA-especifico ja bem documentado, nao
   contradito por nada novo encontrado.

5) GPA MINIMO PREENCHIDO (estava "Verificar"): confirmado que NAO ha
   minimo formal publicado -- GPA competitivo esperado, mas sem piso
   numerico divulgado (mesmo padrao ja visto em outras faculdades desta
   base, ex. Rochester, USC nao e excecao aqui).

6) GPA MEDIO ADMITIDOS PREENCHIDO (estava "Verificar"): 3.60,
   confirmado no Class Profile oficial da coorte Fall 2025 (a mais
   recente disponivel).

7) TAMANHO DE TURMA ATUALIZADO: 189 (registro anterior, presumivelmente
   de uma coorte mais antiga) -> 254, confirmado no Class Profile
   oficial da coorte Fall 2025 (1.995 candidaturas recebidas).

8) TAXA DE ACEITACAO: NAO ALTERADA (mantido 14%, ja registrado). O
   Class Profile oficial da Fall 2025 permite calcular 254/1.995 =
   ~12.7%, mas isso e taxa de MATRICULA (yield), nao taxa de ACEITACAO
   real (a pagina so lista "candidaturas recebidas" e "tamanho de turma
   matriculada", nao o numero de OFERTAS feitas -- que seria maior que
   254, ja que nem todo admitido matricula). Usar 254/1.995 fabricaria
   uma taxa de aceitacao artificialmente baixa. O valor ja registrado
   (14%) veio de uma fonte que o descreve explicitamente como "taxa de
   aceitacao seletiva", entao foi mantido.

9) TIPOS DE BOLSA ENRIQUECIDO: confirmado que bolsas de merito sao
   concedidas automaticamente pelo comite de admissao (sem inscricao
   separada), considerando merito academico, lideranca e experiencia --
   e que internacionais SAO elegiveis para bolsas de merito (mas nao
   para auxilio federal americano, o que ja era esperado). Nenhuma
   bolsa nomeada/valor especifico para o MSBA foi encontrado alem do
   que ja estava registrado.

10) PRINCIPAIS EMPREGADORES SUBSTITUIDO: registro anterior generico
    ("Tech, Consulting"). Substituido por nomes reais confirmados em
    fontes oficiais/agregadas com foco no MSBA: Amazon, Microsoft,
    JPMorgan Chase, Google.

11) GMAT MEDIO: campo passa a "N/A" -- o proprio registro ja confirmado
    (nesta e na rodada anterior) diz que GMAT/GRE NAO sao aceitos no
    processo de admissao do MSBA, entao nao ha media de GMAT a
    registrar (nao e uma lacuna de pesquisa, e uma inaplicabilidade).

NAO ENCONTRADO (permanece "Verificar", nao inventado): TA/RA Disponivel,
Processo TA/RA, Graduate Assistantships -- nenhuma pagina oficial do
MSBA menciona assistantships em nenhum sentido (nem confirmando nem
negando, diferente do caso da University of Washington onde havia uma
frase oficial explicita dizendo que NAO existe). Seguro Saude do
programa -- pesquisado extensivamente, mas as unicas cifras encontradas
eram da University of SOUTH CAROLINA (ambiguidade de nome "USC"),
descartadas por risco de contaminacao; as paginas oficiais corretas
(studenthealth.usc.edu, arr.usc.edu) remetem a um portal que exige
login para ver valores exatos.

Uso: python etl_enrich_usc_round2.py (executar a partir da raiz do
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
    "Nome da Universidade": "USC",
    "Estado": "CA",
    "Cidade": "Los Angeles",
}

UPDATES = {
    "Tamanho Turma": 254,
    "Taxa Aplicação USD": 175,
    "Deadline Aplicação": (
        "Round 1: 1/nov/2026 (decisão 15/jan/2027); Round 2: "
        "15/jan/2027 (decisão 30/mar/2027); Round 3: 15/abr/2027 "
        "(decisão 15/jun/2027)"
    ),
    "Rounds Aplicação": (
        "3 rounds; internacionais recomendados a aplicar nos Rounds 1-2 "
        "para dar tempo ao processo de visto"
    ),
    "Entrevista": "KIRA Video Interview -- por convite apenas, após submissão da candidatura",
    "Testes Inglês Aceitos": (
        "TOEFL iBT mín. 100 total (mín. 20 por seção) / IELTS mín. 6.5 "
        "(mín. 6.0 por banda -- corrige 7.0, que é o piso de PhD, não "
        "de mestrado) / PTE mín. 61 (mín. 53 por seção) / Duolingo "
        "aceito mas SEM mínimo formal de admissão (usado só para "
        "'placement' -- decidir se o aluno precisa fazer o exame ISE de "
        "nivelamento ao chegar; piso geral da USC para outros contextos "
        "é 130)"
    ),
    "GPA Mínimo": "Sem mínimo formal publicado (GPA competitivo esperado)",
    "GPA Médio Admitidos": 3.60,
    "GMAT Médio": "N/A (GMAT/GRE não aceitos no processo de admissão do MSBA)",
    "Tipos de Bolsa": (
        "Merit/grant (parcial); concedida automaticamente pelo comitê "
        "de admissão (sem inscrição separada), considerando mérito "
        "acadêmico, liderança e experiência; internacionais são "
        "elegíveis para bolsas de mérito (mas não para auxílio federal "
        "americano)"
    ),
    "Principais Empregadores": "Amazon, Microsoft, JPMorgan Chase, Google",
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
        raise ValueError("Nenhuma linha encontrada para USC (Los Angeles, CA). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (USC, Los Angeles, CA)")

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

    print(f"\n{len(UPDATES)} campos atualizados na linha {row_idx} (USC):\n")
    for header in UPDATES:
        print(f"  - {header}:")
        print(f"      Antes:  {before[header]!r}")
        print(f"      Depois: {after[header]!r}")


if __name__ == "__main__":
    main()
