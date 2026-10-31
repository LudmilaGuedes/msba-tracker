# -*- coding: utf-8 -*-
"""
Data enrichment (bolsas): Texas Tech (Rawls College of Business), Lubbock,
TX -- Master of Science in Data Science (MSDS). A linha foi adicionada
por etl_add_texas_tech_msds.py.

Pedido: pesquisar mais a fundo a possibilidade de bolsa.

Fontes (set/2026):
- depts.ttu.edu/gradschool/financial-support/graduateStudentUniversityFunding.php
  "Competitive scholarships/fellowships of at least $1,000 are expected to
  waive any out-of-state tuition charges for the academic award years";
  "TA/RA/GPTI appointments will waive out-of-state tuition charges for the
  terms that students are employed".
- Fact sheet do Texas Higher Education Coordinating Board (Competitive
  Scholarship Waiver, TEC 54.213): vale para qualquer "nonresident student",
  o que inclui internacionais, e para programas de graduacao ou
  pos-graduacao. A bolsa precisa ser de pelo menos US$1.000 individual (nao
  vale somar bolsas menores) e concedida por um comite da propria
  instituicao. O waiver cobre no maximo 12 meses por bolsa, e o total de
  waivers da instituicao fica limitado a 5% dos alunos matriculados.
- depts.ttu.edu/opmanual/OP62.40.php: o waiver para assistentes (TEC
  54.212) exige nomeacao de pelo menos meio periodo (0.5 FTE) como
  TA/RA/GPTI, cargo "related to their degree program" e matricula full-time.
  Isenta tambem a designated tuition, a differential tuition e varias fees.
- depts.ttu.edu/gradschool/financial-support/GeneralFellowships.php:
  fellowships da Graduate School de US$3.000 a US$5.000 por ano;
  auto-candidatura (o mesmo formulario vale para bolsas de faculdade e de
  departamento); nao exige admissao previa; prazo 2026-27 em 3/fev,
  resultado em 21/mar.
- depts.ttu.edu/rawlsbusiness/scholarships/: fellowships do Rawls via
  rawlsgrad@ttu.edu; Scholarship Universe; exige matricula full-time.

IMPACTO: com o waiver, o internacional paga a tarifa de residente de
business, ~US$768,67 por credito x 36 = ~US$27.672 em vez de ~US$42.432.
A economia de ~US$14.760 vem ALEM do valor da propria bolsa.

RESSALVA: o ano academico da TTU vai de 1/set a 31/ago. Um programa que
comeca no Summer atravessa dois anos academicos, entao pode ser preciso ter
uma bolsa qualificada em cada um. O waiver tambem vale por no maximo 12
meses por bolsa.

Uso: python scripts/etl_enrich_texas_tech_bolsas.py (a partir da raiz).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "Texas Tech",
    "Estado": "TX",
    "Cidade": "Lubbock",
}

UPDATES = {
    "Bolsas Disponíveis": "Sim",
    "Tipos de Bolsa": (
        "Fellowships da Graduate School (US$3-5k/ano; auto-candidatura, prazo 3/fev, "
        "não exige admissão prévia; o mesmo formulário considera bolsas do Rawls e do "
        "departamento); fellowships do Rawls (rawlsgrad@ttu.edu); GA/TA/RA de meio "
        "período. BÔNUS DO TEXAS: qualquer bolsa competitiva individual de ≥US$1.000 "
        "dá direito à tuition de residente (~US$27,7k em vez de ~US$42,4k)"
    ),
    "Quão Realista (Intl)": (
        "Moderada -- as fellowships aceitam auto-candidatura e basta UMA bolsa de "
        "≥US$1.000 para ganhar a tuition de residente (economia de ~US$14,8k); mas "
        "são competitivas, os waivers da universidade são limitados (máx. 5% dos "
        "alunos) e não há dado específico do MSDS"
    ),
    "TA/RA Disponível": (
        "Sim (institucional) -- depende de vaga; o cargo precisa ser relacionado ao "
        "curso, com pelo menos meio período (20h/semana), e o aluno precisa estar "
        "matriculado full-time"
    ),
    "Processo TA/RA": (
        "Buscar vaga no Rawls (rawlsgrad@ttu.edu) e nos departamentos. A nomeação de "
        "≥0.5 FTE (TA, RA ou GPTI) relacionada ao curso isenta a tuition de "
        "não-residente, a designated e a differential tuition e várias fees durante os "
        "semestres trabalhados"
    ),
    "Caminho Tuition Residente": (
        "Sim -- (1) bolsa competitiva individual de ≥US$1.000 da TTU (Texas Competitive "
        "Scholarship Waiver, TEC 54.213; vale para internacionais; máx. 12 meses por "
        "bolsa) ou (2) GA/TA/RA de ≥ meio período (TEC 54.212). Com isso, a tuition de "
        "internacional cai para a de residente (~US$27,7k)"
    ),
    "Graduate Assistantships": (
        "GA/TA/RA de meio período (~US$22.750/ano no PhD do Rawls; valor para "
        "mestrado não publicado) com waiver da tuition de não-residente e de várias "
        "fees; disponibilidade para o MSDS não confirmada -- perguntar ao programa"
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
        raise ValueError("Nenhuma linha encontrada para Texas Tech (Lubbock, TX). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Texas Tech, Lubbock, TX)")

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

    print(f"\n{len(UPDATES)} campos atualizados na linha {row_idx} (Texas Tech):\n")
    for header in UPDATES:
        print(f"  - {header}:")
        print(f"      Antes:  {before[header]!r}")
        print(f"      Depois: {after[header]!r}")


if __name__ == "__main__":
    main()
