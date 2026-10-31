# -*- coding: utf-8 -*-
"""
Cria a coluna "Cartas de Recomendação" (grupo Aplicação, logo depois de
"Documentos Necessários") e preenche a quantidade de cartas exigida por cada
um dos 139 programas. O schema passa de 54 para 55 colunas; no dashboard o
campo foi adicionado em msba-dashboard/src/types/columns.ts.

Formato do valor: começa sempre pela quantidade obrigatória (ex.: "2",
"0 (opcionais)", "1 (profissional)"), com uma nota curta quando há
detalhe relevante. "0" significa que nenhuma carta é exigida. Sem dado:
"Verificar".

FONTES:
- ~115 programas: quantidade que já estava em "Documentos Necessários",
  registrada no enrichment em lote de set/out-2026 (scripts/
  enrichment_batches/), feito em páginas oficiais.
- Pesquisados agora (20/out/2026), em páginas oficiais:
  Texas A&M (mays.tamu.edu MS Analytics: pede GPA, resume, essay,
  transcript e pré-requisitos, sem cartas); U of Iowa (tippie.uiowa.edu,
  checklist "Applying 101" sem cartas); LSU (online.lsu.edu: resume e
  transcripts; statement opcional; sem cartas); Illinois Tech (iit.edu:
  transcripts, CV, statement, inglês; sem cartas); Texas Tech MSDS
  (depts.ttu.edu/rawlsbusiness: sem cartas); Hofstra (hofstra.edu MS
  programs: transcripts, resume, GMAT/GRE condicional; sem cartas);
  St. Edward's (stedwards.edu deadlines-checklist: só Counseling e Ed.D.
  pedem cartas); Bryant (catalog.bryant.edu: 1 carta acadêmica ou
  profissional); Lehigh (business.lehigh.edu: "Two letters of
  recommendation"); Pitt Katz (business.pitt.edu, admissões dos MS:
  "Two letters of recommendation", ao menos 1 de professor preferível);
  Seton Hall (shu.edu: "Letter(s) of Recommendation", sem número).
- Mantidos como Verificar (sem fonte oficial acessível): VT Global BA
  (página com erro de certificado; a trilha MSBA-BA pede só contato de 3
  referências), UAH, SCSU, A&M-Texarkana (site cita cartas sem número),
  Clark (agregadores citam 2).
- Conflito registrado: U of Cincinnati (o lote dizia 1 carta de
  supervisor; a busca na página de admissão cita 2 recomendadores).

Uso: python scripts/etl_add_cartas_recomendacao.py (a partir da raiz).
"""

import sys
from copy import copy

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"
VERIFICAR = "Verificar"
NEW_HEADER = "Cartas de Recomendação"
AFTER_HEADER = "Documentos Necessários"

CARTAS = {
    ("MIT", "MA", "Cambridge"): "3",
    ("UCLA", "CA", "Los Angeles"): "2 (3ª opcional)",
    ("USC", "CA", "Los Angeles"): "0",
    ("UT Austin", "TX", "Austin"): "2 (pelo menos 1 acadêmica)",
    ("UC Davis", "CA", "Davis/San Francisco"): "2",
    ("UC San Diego", "CA", "La Jolla"): "2",
    ("U of Minnesota", "MN", "Minneapolis"): "2",
    ("Purdue", "IN", "West Lafayette"): "2",
    ("U of Rochester", "NY", "Rochester"): "0 (Simon não aceita cartas)",
    ("Notre Dame", "IN", "Notre Dame"): "1",
    ("UT Dallas", "TX", "Richardson"): "1 (mínimo)",
    ("Texas A&M", "TX", "Houston"): "0 (não listadas no site oficial)",
    ("U of Utah", "UT", "Salt Lake City"): "1",
    ("U of Cincinnati", "OH", "Cincinnati"): "1-2 (fontes conflitam: 1 de supervisor ou 2 recomendadores)",
    ("U of Florida", "FL", "Gainesville"): "2 ou mais (de supervisores)",
    ("Florida State", "FL", "Tallahassee"): "2 (3 preferível)",
    ("USF", "FL", "Tampa"): "2",
    ("UCF", "FL", "Orlando"): "3",
    ("Georgia Tech", "GA", "Atlanta"): "3",
    ("UGA", "GA", "Athens"): "1 (+1 opcional)",
    ("NC State", "NC", "Raleigh"): "2 (3ª opcional)",
    ("U of South Carolina", "SC", "Columbia"): "2",
    ("VT Global BA", "VA", "Blacksburg"): VERIFICAR,
    ("UT Knoxville", "TN", "Knoxville"): "2",
    ("U of Alabama", "AL", "Tuscaloosa"): "3",
    ("UAH", "AL", "Huntsville"): VERIFICAR,
    ("U of Louisville", "KY", "Louisville"): "2",
    ("LSU", "LA", "Baton Rouge"): "0 (não listadas no site oficial)",
    ("U of Arkansas", "AR", "Fayetteville"): "3",
    ("Columbia", "NY", "New York"): "3",
    ("Fordham", "NY", "New York"): "0 (1 opcional)",
    ("Baruch CUNY", "NY", "New York"): "2",
    ("U at Buffalo SUNY", "NY", "Buffalo"): "0",
    ("Stony Brook SUNY", "NY", "Long Island"): "3",
    ("Binghamton SUNY", "NY", "Binghamton"): "2",
    ("RIT", "NY", "Rochester"): "0 (opcionais)",
    ("Syracuse", "NY", "Syracuse"): "2",
    ("BU", "MA", "Boston"): "0",
    ("Northeastern", "MA", "Boston"): "2",
    ("UMass Amherst", "MA", "Amherst"): "2",
    ("UConn", "CT", "Hartford/Stamford"): "0",
    ("CMU", "PA", "Pittsburgh"): "1 (profissional)",
    ("Temple", "PA", "Philadelphia"): "1",
    ("Drexel", "PA", "Philadelphia"): "1",
    ("U of Maryland", "MD", "College Park"): "1",
    ("Johns Hopkins", "DC", "Washington"): "1 (profissional)",
    ("Georgetown", "DC", "Washington"): "1",
    ("GWU", "DC", "Washington"): "1",
    ("UIUC", "IL", "Champaign"): "2",
    ("Miami U Ohio", "OH", "Oxford"): "1",
    ("UW-Madison", "WI", "Madison"): "0 (até 2 opcionais)",
    ("U of Iowa", "IA", "Iowa City"): "0 (não listadas no site oficial)",
    ("Wichita State", "KS", "Wichita"): "0",
    ("Dakota State", "SD", "Madison"): "0",
    ("WashU", "MO", "St. Louis"): "0",
    ("UC Berkeley", "CA", "Berkeley"): "2",
    ("UC Irvine", "CA", "Irvine"): "2",
    ("ASU", "AZ", "Tempe"): "1 (só contato do recomendador, sem carta)",
    ("U of Arizona", "AZ", "Tucson"): "2-3 (formulário online curto)",
    ("CU Boulder", "CO", "Boulder"): "1 (pode enviar 2)",
    ("Cal Poly SLO", "CA", "San Luis Obispo"): "2",
    ("SF State", "CA", "San Francisco"): "1 (mínimo)",
    ("Cal Poly Pomona", "CA", "Pomona"): "2",
    ("Seattle U", "WA", "Seattle"): "0 (opcionais)",
    ("Baylor", "TX", "Waco"): "1",
    ("Oklahoma State", "OK", "Stillwater"): "3",
    ("U of Tulsa", "OK", "Tulsa"): "2-3",
    ("U of Central Oklahoma", "OK", "Edmond"): "0",
    ("SMU", "TX", "Dallas"): "0 (2 referências só com contato, sem carta)",
    ("TCU", "TX", "Fort Worth"): "0 (até 3 contatos de recomendadores; cartas não aceitas)",
    ("U of Michigan", "MI", "Ann Arbor"): "1 (formulário online)",
    ("Kent State", "OH", "Kent"): "2",
    ("UTSA", "TX", "San Antonio"): "0 (2 opcionais)",
    ("UT Arlington", "TX", "Arlington"): (
        "0 com admissão direta (GPA ≥ 3.0 nas últimas 60 horas); sem ela, exigidas "
        "(quantidade não informada)"
    ),
    ("U of Houston", "TX", "Houston"): "2",
    ("Texas State", "TX", "San Marcos"): "3",
    ("Illinois Tech", "IL", "Chicago"): "0 (não listadas no site oficial)",
    ("Florida Atlantic University (FAU)", "FL", "Boca Raton"): "0",
    ("University of North Texas (UNT)", "TX", "Denton"): "0 (3 referências opcionais, só contato)",
    ("Santa Clara University", "CA", "Santa Clara"): "2",
    ("Loyola Marymount University (LMU)", "CA", "Los Angeles"): "2",
    ("University of Illinois Chicago (UIC)", "IL", "Chicago"): "2",
    ("DePaul University", "IL", "Chicago"): "0 (opcionais)",
    ("Cornell University (SC Johnson)", "NY", "New York City (NYC)"): "2 (mín. 1 de supervisor)",
    ("Rensselaer Polytechnic Institute (RPI)", "NY", "Troy"): "2",
    ("Rutgers University–Camden", "NJ", "Camden"): "2",
    ("Rutgers Business School (Newark)", "NJ", "Newark"): "1",
    ("Seton Hall University", "NJ", "South Orange"): "1 ou mais (site diz 'letter(s)', sem número exato)",
    ("Southern Connecticut State University (SCSU)", "CT", "New Haven"): VERIFICAR,
    ("Bentley University", "MA", "Waltham"): "1",
    ("University of Pittsburgh (Katz)", "PA", "Pittsburgh"): "2 (pelo menos 1 de professor, preferível)",
    ("Wake Forest University", "NC", "Winston-Salem"): "0 (2 referências só com contato)",
    ("William & Mary (Raymond A. Mason School of Business)", "VA", "Williamsburg"): (
        "2 (via questionário; carta opcional)"
    ),
    ("George Mason University", "VA", "Fairfax"): "1 (profissional)",
    ("University of Washington", "WA", "Seattle"): "2",
    ("University of Denver", "CO", "Denver"): "0",
    ("University of Miami", "FL", "Coral Gables"): "2 (até 3)",
    ("Wayne State University", "MI", "Detroit"): "0 (1 suplementar, recomendada)",
    ("Oakland University", "MI", "Rochester (MI)"): "1 (só e-mail do recomendador)",
    ("American University", "DC", "Washington"): "1",
    ("University of North Florida (UNF)", "FL", "Jacksonville"): "0 com waiver; sem waiver, 2 profissionais",
    ("Florida Institute of Technology (Florida Tech)", "FL", "Melbourne"): "0",
    ("East Texas A&M University (ETAMU)", "TX", "Commerce"): "0",
    ("Texas A&M University-Texarkana (A&M-Texarkana)", "TX", "Texarkana"): VERIFICAR,
    ("University of Dallas", "TX", "Irving"): "0 (1 opcional)",
    ("West Texas A&M University (WTAMU)", "TX", "Canyon"): "2 (pelo menos 1 profissional)",
    ("Texas Tech", "TX", "Lubbock"): "0 (não listadas no site oficial)",
    ("Stevens Institute of Technology", "NJ", "Hoboken"): "2",
    ("Lehigh University", "PA", "Bethlehem"): "2",
    ("Babson College", "MA", "Wellesley"): "1",
    ("Brandeis University", "MA", "Waltham"): "1 (até 2)",
    ("Boston College", "MA", "Chestnut Hill"): "2",
    ("Worcester Polytechnic Institute (WPI)", "MA", "Worcester"): "3",
    ("Suffolk University", "MA", "Boston"): "1 (formulário online, não carta)",
    ("Clark University", "MA", "Worcester"): VERIFICAR,
    ("Hult International Business School", "MA", "Boston"): "1",
    ("Bryant University", "RI", "Smithfield"): "1",
    ("University of New Hampshire", "NH", "Durham"): "2",
    ("Hofstra University", "NY", "Hempstead"): "0 (não listadas no site oficial)",
    ("Fairfield University", "CT", "Fairfield"): "2 (formulário online)",
    ("Quinnipiac University", "CT", "Hamden"): "1 (profissional)",
    ("St. John's University", "NY", "Queens"): "1 (a página geral da Tobin cita 2)",
    ("Tulane University", "LA", "New Orleans"): "1",
    ("Emory University", "GA", "Atlanta"): "2",
    ("Duke University", "NC", "Durham"): "1 (no máximo 1)",
    ("Iowa State University", "IA", "Ames"): "2 (full-time; o site cita de 1 a 3)",
    ("UC Riverside", "CA", "Riverside"): "1",
    ("UNC Charlotte", "NC", "Charlotte"): "3",
    ("Chapman University", "CA", "Orange"): "1",
    ("Case Western Reserve University", "OH", "Cleveland"): "2",
    ("University of Dayton", "OH", "Dayton"): "0 (opcionais)",
    ("Loyola University Chicago", "IL", "Chicago"): "0",
    ("University of New Haven", "CT", "West Haven"): "2",
    ("UMass Lowell", "MA", "Lowell"): "2",
    ("Adelphi University", "NY", "Garden City"): "0 (1 opcional)",
    ("LIU Post (Long Island University)", "NY", "Brookville"): "2",
    ("Virginia Commonwealth University (VCU)", "VA", "Richmond"): "1",
    ("University of San Diego", "CA", "San Diego"): "Até 2 (não fica claro se são obrigatórias)",
    ("St. Edward's University", "TX", "Austin"): "0",
}


def main():
    rotate_backup()
    print("Backup rotacionado (política fixa) em: data/msba_programs_database_backup.xlsx")

    wb = openpyxl.load_workbook(DB_PATH)
    ws = wb[SHEET_DATA]
    headers = [cell.value for cell in ws[1]]
    if NEW_HEADER in headers:
        raise ValueError(f"A coluna '{NEW_HEADER}' já existe. Abortando.")

    src_idx = headers.index(AFTER_HEADER) + 1
    new_idx = src_idx + 1
    col = {h: i + 1 for i, h in enumerate(headers)}

    keys = []
    for r in range(2, ws.max_row + 1):
        keys.append((
            ws.cell(r, col["Nome da Universidade"]).value,
            ws.cell(r, col["Estado"]).value,
            ws.cell(r, col["Cidade"]).value,
        ))
    missing = [k for k in keys if k not in CARTAS]
    extra = [k for k in CARTAS if k not in keys]
    if missing or extra:
        raise ValueError(f"Linhas sem valor: {missing}; valores sem linha: {extra}")

    # Larguras das colunas à direita do ponto de inserção, para reaplicar depois do deslocamento.
    letter = openpyxl.utils.get_column_letter
    widths = {i: ws.column_dimensions[letter(i)].width for i in range(src_idx, ws.max_column + 1)}

    ws.insert_cols(new_idx)

    for i, w in widths.items():
        target = i if i == src_idx else i + 1
        ws.column_dimensions[letter(target)].width = w
    ws.column_dimensions[letter(new_idx)].width = 28

    verificar_fill = PatternFill(start_color="00FFF3CD", end_color="00FFF3CD", fill_type="solid")
    no_fill = PatternFill(fill_type=None)

    header_cell = ws.cell(1, new_idx, NEW_HEADER)
    ref_header = ws.cell(1, src_idx)
    header_cell.font = copy(ref_header.font)
    header_cell.fill = copy(ref_header.fill)
    header_cell.border = copy(ref_header.border)
    header_cell.alignment = copy(ref_header.alignment)

    counts = {}
    for offset, key in enumerate(keys):
        r = offset + 2
        value = CARTAS[key]
        ref = ws.cell(r, src_idx)
        cell = ws.cell(r, new_idx, value)
        cell.font = copy(ref.font)
        cell.border = copy(ref.border)
        cell.alignment = copy(ref.alignment)
        cell.fill = verificar_fill if value == VERIFICAR else no_fill
        bucket = value if value == VERIFICAR else value.split(" ")[0]
        counts[bucket] = counts.get(bucket, 0) + 1

    ws.auto_filter.ref = f"A1:{letter(ws.max_column)}{ws.max_row}"
    wb.save(DB_PATH)

    print(f"Coluna '{NEW_HEADER}' criada na posição {new_idx} (depois de '{AFTER_HEADER}'); "
          f"planilha agora tem {ws.max_column} colunas.")
    print(f"{len(keys)} programas preenchidos. Distribuição pela quantidade obrigatória:")
    for bucket, n in sorted(counts.items(), key=lambda kv: (-kv[1], kv[0])):
        print(f"  {bucket}: {n}")


if __name__ == "__main__":
    main()
