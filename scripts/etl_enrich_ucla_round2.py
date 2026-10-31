# -*- coding: utf-8 -*-
"""
Data enrichment (round 2): auditoria + pesquisa aprofundada focada em UMA
unica linha -- UCLA Anderson School of Management, Los Angeles, CA -- MS
in Business Analytics (MSBA), presencial, STEM.

Pedido da usuaria: Data Enrichment de uma unica instituicao, cobrindo
tuition, bolsas, taxa de aplicacao, requisitos completos de aplicacao,
empregabilidade/salario/empregadores e duracao/formato. Substituir dados
incorretos, preencher lacunas, manter "Verificar" para o que nao for
encontrado.

Fontes oficiais (anderson.ucla.edu, verificadas por WebFetch direto em
set/2026):
- Calendario e tarifas: .../admit-central/program-calendar-and-fees
- Academico/curriculo: .../academics
- Pre-requisitos: .../admissions/prerequisites
- FAQ de admissao: .../admissions/faq
- Outcomes/colocacao: .../career-impact/student-outcomes-placement
- Relatorio oficial de emprego do MSBA (PDF, Class of 2023):
  anderson.ucla.edu/sites/default/files/document/2025-04/2025MSBA%20Employment%20Report%20Updates.pdf

ACHADOS PRINCIPAIS:

1) TUITION ATUALIZADA (2026-27): pagina oficial de calendario/tarifas
   detalha por trimestre e confirma total de US$93.214,42, decomposto em
   US$81.600,00 de tarifa do programa (instrucional) + US$9.176,00 de
   UCSHIP (seguro saude) + US$2.438,42 de tarifas de campus. O valor
   anterior (US$75.445) estava desatualizado. Atualizado para
   US$81.600 (tuition "pura", separada do seguro e das tarifas de
   campus, que ja tem/ganham campos proprios no schema). Sem distincao
   residente/internacional -- programa self-support, todos pagam igual
   (ja confirmado antes, mantido).

2) TOTAL DE CREDITOS PREENCHIDO (estava "Verificar"): 48 unidades,
   confirmado na pagina oficial de academicos.

3) CUSTO/CREDITO CALCULADO (estava "Verificar" nos dois campos):
   US$81.600 / 48 = US$1.700,00/credito (igual para todos, self-support).

4) FEES ADICIONAIS PREENCHIDO (estava "Verificar"): US$2.438,42
   (tarifas de campus, separadas da tuition e do seguro-saude, conforme
   quebra oficial por trimestre).

5) SEGURO SAUDE ATUALIZADO: US$6.882 (valor antigo) -> US$9.176,00
   (UCSHIP 2025-26, valor mais recente confirmado oficialmente; a
   propria pagina avisa que pode mudar para 26-27, mas e o numero mais
   atual disponivel).

6) TAMANHO DE TURMA PREENCHIDO (estava "Verificar"): meta declarada
   oficialmente de 120 alunos para a Class of 2028 (FAQ oficial) --
   turma historica de referencia (Class of 2023, relatorio de emprego)
   tinha 93 alunos; registrado com nota sobre o crescimento.

7) *** CORRECAO IMPORTANTE: EXPERIENCIA PROFISSIONAL *** -- o registro
   anterior dizia "Minimo 2 anos de experiencia profissional full-time
   (exigido, nao apenas tipico)". O FAQ OFICIAL do programa contradiz
   isso taxativamente: "Q: Is work experience required before applying?
   A: Work experience is not required but it is recommended for
   admission." Corrigido para refletir que NAO e exigencia formal --
   e recomendada, com a media historica da turma (~2,1 anos, Class of
   2023) mantida como contexto.

8) GPA MINIMO PREENCHIDO (estava "Sem minimo formal publicado"): o FAQ
   oficial (fonte diferente da que gerou o registro anterior) cita
   explicitamente "Minimum GPA: 3.0+ on a 4.0 scale". Atualizado.

9) *** CORRECAO: DUOLINGO *** -- o registro anterior tinha "Duolingo
   min. 135". A pagina oficial de pre-requisitos/admissao NAO menciona
   Duolingo em nenhum momento -- apenas TOEFL e IELTS sao citados como
   testes aceitos, com detalhamento por secao do TOEFL. Uma busca
   adicional especifica sobre aceitacao de Duolingo no MSBA da UCLA
   Anderson nao encontrou nenhuma confirmacao oficial. Como a ausencia e
   consistente em 2 paginas oficiais diferentes (pre-requisitos E FAQ),
   o registro de Duolingo foi REMOVIDO por falta de confirmacao (nao e
   so uma lacuna -- e uma ausencia sistematica em fontes oficiais que
   deveriam mencionar se fosse aceito), com nota explicita da duvida.

10) MATERIAS PRE-REQUISITO PREENCHIDO (estava "Verificar"): algebra
    linear, calculo multivariavel, estatistica/probabilidade (nivel
    recomendado, nao pre-requisito formal obrigatorio); SQL, R, Python
    (Python e a linguagem predominante do curriculo).

11) PROCESSO TA/RA PREENCHIDO (estava "Verificar"): N/A -- o proprio
    campo "TA/RA Disponivel" ja confirmava "Nao" (programa self-support),
    entao nao ha processo a descrever.

12) PRINCIPAIS EMPREGADORES SUBSTITUIDO: o registro anterior era
    generico ("Tech, Consulting, Finance", nomes de categoria, nao de
    empresa). O relatorio OFICIAL de emprego do MSBA (Class of 2023,
    PDF direto do site, especifico do MSBA -- nao do MBA, evitando o
    erro de contaminacao ja identificado e corrigido nesta sessao para
    outras faculdades) lista dezenas de empregadores reais. Selecionada
    uma amostra representativa dos mais frequentes/reconheciveis.

13) TAXA DE EMPREGABILIDADE E SALARIO: RE-CONFIRMADOS, SEM MUDANCA -- a
    pagina oficial de outcomes (student-outcomes-placement) confirma
    EXATAMENTE os valores ja registrados: 98% empregados em ate 6 meses
    (Taxa Empregabilidade % = 98) e USD 110.332 de salario medio
    (Salario Medio Pos USD = 110332), ambos da Class of 2023 -- os
    numeros ja estavam certos.

14) GMAT/GRE: refinado com o limiar especifico da GMAT Focus Edition
    (665+) confirmado na pagina oficial de pre-requisitos, mantendo o
    restante do texto ja registrado (obrigatorio, sem dispensa geral).

15) DEADLINE APLICACAO: as datas ja registradas (Round 1: 9/jan, Round
    2: 6/mar, Round 3: 8/mai) foram RE-CONFIRMADAS como as datas oficiais
    da Class of 2027 (entrada Fall 2026) -- mas esse ciclo ja passou (a
    data de hoje e set/2026, o Round 3 desse ciclo foi em mai/2026). O
    ciclo seguinte (Class of 2028, Fall 2027) ainda nao tem datas
    publicadas oficialmente. Registrado com essa contextualizacao em vez
    de apresentar as datas antigas como se fossem as proximas.

NAO ENCONTRADO (permanece "Verificar", nao inventado): Taxa de
Aceitacao % -- o relatorio de emprego cita "Total Applicant Pool: 1314"
e "Size of Entering Class: 93" (Class of 2023), mas isso e proporcao de
MATRICULADOS/candidatos (afetada por yield), nao a taxa de ACEITACAO
real (que seria admitidos/candidatos) -- os dois numeros nao sao
equivalentes, entao nao foi calculado nem usado para nao fabricar uma
taxa de aceitacao a partir de um numero errado. Ranking US News e
Ranking QS ja estavam preenchidos e nao fazem parte do escopo desta
rodada.

Uso: python etl_enrich_ucla_round2.py (executar a partir da raiz do
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
    "Nome da Universidade": "UCLA",
    "Estado": "CA",
    "Cidade": "Los Angeles",
}

UPDATES = {
    "Total Créditos": 48,
    "Custo/Crédito (Intl) USD": 1700,
    "Custo/Crédito (Residente) USD": 1700,
    "Tuition Total (Intl) USD": 81600,
    "Tuition Total (Residente) USD": 81600,
    "Fees Adicionais USD": 2438.42,
    "Seguro Saúde (programa) USD": 9176,
    "Tamanho Turma": 120,
    "GPA Mínimo": "3.0+ (numa escala de 4.0) -- confirmado no FAQ oficial",
    "GMAT/GRE Obrigatório": (
        "Obrigatório (GMAT ou GRE); GMAT Focus Edition exige 665+; média "
        "da Classe de 2026 no percentil 91-93; perfil competitivo: GMAT "
        "710+ / GRE Quant 167+"
    ),
    "Experiência Profissional": (
        "Não obrigatória -- recomendada, mas não exigida (confirmado no "
        "FAQ oficial: 'work experience is not required but it is "
        "recommended'). Média histórica da turma: ~2,1 anos (Class of "
        "2023); registro anterior dizia erroneamente que era exigida"
    ),
    "Matérias Pré-Requisito": (
        "Álgebra linear, cálculo multivariável, estatística e "
        "probabilidade (nível recomendado, não pré-requisito formal "
        "obrigatório); SQL, R, Python (Python é a linguagem predominante "
        "do currículo)"
    ),
    "Testes Inglês Aceitos": (
        "TOEFL iBT mín. 100 (mínimos por seção: Writing 25, Speaking 24, "
        "Reading 21, Listening 17) / TOEFL PBT mín. 560 / IELTS mín. "
        "7.5 -- Duolingo NÃO mencionado em nenhuma página oficial de "
        "admissão (ausência em 2 fontes diferentes, provavelmente não "
        "aceito, diferente do registro anterior que citava 135). "
        "Dispensa para diploma de instituição verificada pelo WHED com "
        "inglês como único idioma de instrução (a partir de Fall 2025)"
    ),
    "Processo TA/RA": "N/A -- programa não oferece TA/RA (self-support), confirmado",
    "Principais Empregadores": (
        "Amazon, McKinsey & Company, Meta, Microsoft, JPMorgan Chase, "
        "Deloitte, Bain & Company, BCG X, TikTok, Kaiser Permanente, "
        "Tesla, Capital One (relatório oficial de emprego do MSBA, "
        "Class of 2023)"
    ),
    "Deadline Aplicação": (
        "Ciclo Class of 2027 (Fall 2026, já encerrado): Round 1 "
        "9/jan/2026; Round 2 6/mar/2026; Round 3 8/mai/2026. Ciclo "
        "seguinte (Class of 2028, Fall 2027) ainda não publicado "
        "oficialmente -- tende a seguir padrão semelhante (janeiro/"
        "março/maio)"
    ),
    "Destaques Currículo": (
        "Data analytics, ML, Capstone (equipes de ~5 alunos, projeto "
        "real patrocinado por empresa parceira); estágio de verão "
        "obrigatório (8-10 semanas) entre o 2º e o 3º trimestre; "
        "tamanho de turma em crescimento -- 93 alunos na Class of 2023, "
        "meta de 120 para a Class of 2028 (FAQ oficial)"
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
        raise ValueError("Nenhuma linha encontrada para UCLA (Los Angeles, CA). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (UCLA, Los Angeles, CA)")

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

    print(f"\n{len(UPDATES)} campos atualizados na linha {row_idx} (UCLA):\n")
    for header in UPDATES:
        print(f"  - {header}:")
        print(f"      Antes:  {before[header]!r}")
        print(f"      Depois: {after[header]!r}")


if __name__ == "__main__":
    main()
