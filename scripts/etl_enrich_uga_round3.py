# -*- coding: utf-8 -*-
"""
Data enrichment (round 3): UGA (Terry College of Business), Athens, GA --
MS in Business Analytics ("Georgia Analytics"), presencial, STEM, 10 meses.

terry.uga.edu continua bloqueando fetch automatizado (HTTP 403), como nas
rodadas 1 e 2. Os dados da Terry vieram de snippets de busca que indexam
as paginas oficiais. As paginas da UGA fora da Terry foram lidas direto.

Fontes (set/2026):
- busfin.uga.edu/.../bursar-tuition-fees/2026-2027/ (fetch direto)
- grad.uga.edu/admissions/requirements/international-applications/
  english-language-proficiency-requirement/ (fetch direto)
- terry.uga.edu/msba/admissions/ e /msba/class-profile (via snippets)
- career.uga.edu/outcomes (sem dados do MSBA)

ACHADOS:

1) *** DUOLINGO NAO E ACEITO: "As of the Fall 2024 admissions cycle,
   Duolingo English Test scores are no longer accepted" (Graduate School).
   So TOEFL ou IELTS. Minimos: TOEFL 80 (>=20 em speaking e writing) na
   escala antiga, ou 4.5 (>=4 em speaking e writing) na nova escala a
   partir de 22/jan/2026; IELTS 6.5, com nenhuma banda abaixo de 6.0. ***

2) TUITION INTL CORRIGIDO: a tabela do bursar de 2026-27 tem tarifa propria
   para o MSBA. Com 12+ creditos o valor e fixo por semestre: US$18.528
   nao-residente (US$1.544/h abaixo de 12h) e US$8.940 residente (US$745/h).
   O programa tem 2 semestres (Fall + Spring):
     Intl:      2 x 18.528 = US$37.056 (antes 38.136 = 24 x 1.589, base desatualizada)
     Residente: 2 x 8.940  = US$17.880 (ja estava correto)
   Com o GA (ate ~90% da turma, ver rodada 2), a tuition cai para US$25
   por semestre.

3) PRAZOS ATUALIZADOS: o ciclo mais recente publicado (Fall 2026) tem 3
   rounds: 1/nov/2025, 15/fev/2026 (final para internacionais) e
   15/abr/2026 (final geral). O valor antigo "R1 Oct 10, R2 Jan 10, R3 Mar
   10" era de um ciclo anterior, e "4 rounds" estava errado. As datas do
   ciclo Fall 2027 ainda nao aparecem publicadas. Texto reescrito com o
   prazo mais cedo primeiro (regra do CLAUDE.md).

4) DOCUMENTOS: online application; GMAT, GRE ou Executive Assessment
   (dispensa disponivel); transcripts; essays; resume de 1 pagina; DUAS
   recomendacoes; TOEFL/IELTS.

5) PERFIL DA TURMA (Fall 2024): 121 alunos, GPA mediano 3.7, GRE mediano
   311, 17% internacionais. O GMAT medio nao e divulgado.

6) EXPERIENCIA: o comite "strongly prefers" pelo menos 1 estagio para
   recem-formados, ou experiencia full-time para quem se formou ha mais de
   1 ano.

Continua "Verificar": Salario Medio, Principais Empregadores, Taxa
Aceitacao, Entrevista, Rankings. O career.uga.edu nao traz dados do MSBA,
e os numeros achados (salario de US$120-126k, GMAT 641, 41% de aceitacao)
sao do MBA.

Uso: python scripts/etl_enrich_uga_round3.py (a partir da raiz).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "UGA",
    "Estado": "GA",
    "Cidade": "Athens",
}

UPDATES = {
    "Tamanho Turma": 121,
    "Destaques Currículo": (
        "SQL, R, Python; Capstone com empresas parceiras (Ivester Institute); "
        "pathway 'Double Dawg' (graduação + MSBA) para alunos da própria UGA"
    ),
    "Custo/Crédito (Intl) USD": 1544,
    "Tuition Total (Intl) USD": 37056,
    "GPA Médio Admitidos": 3.7,
    "GMAT/GRE Obrigatório": (
        "GMAT, GRE ou Executive Assessment; sem nota mínima definida; dispensa (waiver) "
        "disponível mediante solicitação"
    ),
    "GMAT Médio": "Não divulgado (só GRE: mediana 311, turma Fall 2024)",
    "Experiência Profissional": (
        "Fortemente preferida: pelo menos 1 estágio para quem vem direto da graduação, "
        "ou experiência full-time para quem se formou há mais de 1 ano"
    ),
    "Testes Inglês Aceitos": (
        "TOEFL iBT 80 (≥20 em speaking e writing; testes até 21/jan/2026) ou 4.5 na nova "
        "escala (≥4 em speaking e writing) / IELTS 6.5 (nenhuma banda abaixo de 6.0); "
        "Duolingo NÃO é aceito desde o ciclo Fall 2024"
    ),
    "Deadline Aplicação": (
        "Round 1: 1 nov 2025; Round 2: 15 fev 2026 (final para internacionais); Round 3: "
        "15 abr 2026 (final geral) -- ciclo Fall 2026; datas do Fall 2027 ainda não "
        "publicadas"
    ),
    "Rounds Aplicação": "3 rounds (ver Deadline Aplicação); internacionais só podem aplicar até o Round 2",
    "Documentos Necessários": (
        "Formulário online; GMAT, GRE ou Executive Assessment (dispensa possível); "
        "transcripts; essays; resume de 1 página; 2 cartas de recomendação; TOEFL/IELTS "
        "(não-nativos)"
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
        raise ValueError("Nenhuma linha encontrada para UGA (Athens, GA). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (UGA, Athens, GA)")

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

    print(f"\n{len(UPDATES)} campos atualizados na linha {row_idx} (UGA):\n")
    for header in UPDATES:
        print(f"  - {header}:")
        print(f"      Antes:  {before[header]!r}")
        print(f"      Depois: {after[header]!r}")


if __name__ == "__main__":
    main()
