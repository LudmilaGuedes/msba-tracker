# -*- coding: utf-8 -*-
"""
Data enrichment (round 2): ASU (W. P. Carey School of Business), Tempe, AZ
-- MS in Business Analytics, trilha Big Data (9 meses), presencial, STEM.

Fontes (verificadas por WebFetch/WebSearch em set/2026):
- wpcarey.asu.edu/masters-programs/business-analytics
- wpcarey.asu.edu/masters-programs/business-analytics/cost-financial-aid
- admission.asu.edu/international/graduate/english-proficiency
- admission.asu.edu/graduate/apply

ACHADOS:

1) TUITION ATUALIZADO: a pagina oficial mostra agora Big Data = US$41.000
   (residente) / US$67.250 (nao-residente) / US$71.000 (internacional);
   antes estava US$40.000 / US$67.500. Os valores sao "estimated tuition
   and fees" do programa inteiro, entao as fees ja estao incluidas.
   Custo/credito = total / 30: US$1.366,67 residente e US$2.366,67 intl.
   As outras trilhas (16 meses) custam US$59.000 / US$94.000 / US$98.000
   e nao sao o foco desta linha.

2) FEES ADICIONAIS: antes "Verificar", agora registrado como texto,
   "incluidas no total estimado", mais o deposito de US$500. Ja existe
   texto nesta coluna em outras linhas.

3) INICIO: agosto OU janeiro (antes so "August"). Prazos do Spring 2027
   adicionados.

4) FORMATO: tambem existe uma versao 100% online, sem trilhas. A linha
   continua sendo a presencial em Tempe.

5) INGLES PREENCHIDO com minimos oficiais da Graduate College: TOEFL iBT
   80 (antes de 21/jan/2026) ou 4 (nova escala), IELTS 6.5, PTE 60,
   Duolingo 105. Nao aceitos: TOEFL iBT Paper Edition e PTE Academic
   Online. A exclusao anterior de "TOEFL Home Edition antes de 21/jan/26"
   e "TOEFL Essentials" nao aparece mais na pagina oficial e foi
   substituida.

6) GPA MINIMO: "competitive applicants typically have" 3.0 nos ultimos 60
   creditos (Graduate College).

7) TAXA DE APLICACAO: a pagina do MSBA diz US$115 intl, mas a pagina geral
   da Graduate College diz US$120 para internacionais. O valor do programa
   (115) foi mantido e o conflito registrado em Rounds Aplicacao. Nenhuma
   isencao oficial para candidatos comuns; parcerias corporativas podem
   dispensar a taxa.

8) BOLSAS: as bolsas de merito sao "extremely competitive" e exigem uma
   aplicacao curta depois da admissao. As bolsas corporativas (US$2k-25k)
   sao para funcionarios de empresas parceiras. O GA e mencionado, sem
   detalhes. Realismo rebaixado de "Moderada" para "Baixa-Moderada".

9) EMPREGABILIDADE: 95% em 6 meses re-confirmado na pagina oficial. Os
   dados de terceiros (GPA 3.6, GMAT 690, US$131k, 5,3 anos de
   experiencia) misturam dados do MBA e nao foram usados. O salario
   existente (US$121.116) nao foi re-confirmado nem contradito por fonte
   oficial e foi mantido.

Continua "Verificar": Tamanho Turma, GPA/GMAT medios, Taxa Aceitacao,
Entrevista, Rankings, Processo TA/RA, custo de vida mensal, Seguro Saude.

Uso: python scripts/etl_enrich_asu_round2.py (a partir da raiz).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "ASU",
    "Estado": "AZ",
    "Cidade": "Tempe",
}

UPDATES = {
    "Formato": "Presencial (Tempe); existe também versão 100% online, sem trilhas",
    "Início": "August (Fall) ou January (Spring)",
    "Destaques Currículo": (
        "Trilha Big Data (9 meses, 30 créditos) ou trilhas opcionais de 3º semestre "
        "(16 meses): Cloud computing & tech consulting, Fintech, Marketing analytics, "
        "Supply chain analytics"
    ),
    "Custo/Crédito (Intl) USD": 2366.67,
    "Custo/Crédito (Residente) USD": 1366.67,
    "Tuition Total (Intl) USD": 71000,
    "Tuition Total (Residente) USD": 41000,
    "Fees Adicionais USD": (
        "Incluídas no total estimado ('estimated tuition and fees'); depósito de "
        "matrícula único de US$500, não reembolsável"
    ),
    "Tipos de Bolsa": (
        "Bolsas de mérito W. P. Carey em número limitado ('extremely competitive'; "
        "aplicação curta após a admissão); bolsas corporativas de US$2.000-25.000 "
        "(para funcionários de empresas parceiras); Graduate Assistantship mencionado "
        "sem detalhes"
    ),
    "Quão Realista (Intl)": (
        "Baixa-Moderada -- bolsas de mérito são poucas e muito competitivas; bolsas "
        "corporativas exigem vínculo com empresa parceira; GA sem detalhes publicados"
    ),
    "GPA Mínimo": "3.0 nos últimos 60 créditos ('competitive applicants typically have', Graduate College)",
    "Testes Inglês Aceitos": (
        "TOEFL iBT 80 (antes de 21/jan/2026) ou 4 (nova escala) / IELTS 6.5 / PTE 60 / "
        "Duolingo 105; NÃO aceita TOEFL iBT Paper Edition nem PTE Academic Online"
    ),
    "Deadline Aplicação": (
        "Fall 2027 (agosto): Round I 1/dez/2026; Round II intl prioridade 1/fev/2027; "
        "Round III intl final 1/abr/2027; Round IV final 1/jul/2027. Spring 2027 "
        "(janeiro): 1/jul, 1/set (intl prioridade), 1/nov (intl final), 1/dez/2026"
    ),
    "Rounds Aplicação": (
        "4 rounds por entrada (Fall e Spring) -- ver Deadline Aplicação. Taxa: US$115 "
        "intl na página do MSBA (a página geral da Graduate College cita US$120 -- "
        "conferir ao aplicar); sem isenção para candidatos comuns (parcerias "
        "corporativas podem dispensar)"
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
        raise ValueError("Nenhuma linha encontrada para ASU (Tempe, AZ). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (ASU, Tempe, AZ)")

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

    print(f"\n{len(UPDATES)} campos atualizados na linha {row_idx} (ASU):\n")
    for header in UPDATES:
        print(f"  - {header}:")
        print(f"      Antes:  {before[header]!r}")
        print(f"      Depois: {after[header]!r}")


if __name__ == "__main__":
    main()
