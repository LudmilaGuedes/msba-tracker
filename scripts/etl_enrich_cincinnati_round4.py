"""
Data enrichment (round 4): atualiza APENAS a linha da 'U of Cincinnati'
(Lindner College of Business, Cincinnati, OH) no msba_programs_database.xlsx.

Foco: detalhar como funcionam as Graduate Assistantships na UC (mecanica
geral de stipend + tuition scholarship), a partir de fonte oficial nova
indicada pela usuaria.

FONTES CONSULTADAS NESTA RODADA (ago/2026):

1. https://www.grad.uc.edu/fac-staff/handbook/financial-support.html
   Visao geral: UC oferece "merit-based graduate awards" na forma de
   Graduate Scholarships (com ou sem assistantship) + stipend awards
   (teaching/research assistantships), alem de bolsas externas e need-based
   aid. Nao lista processo competitivo especifico nem estatisticas.

2. https://www.grad.uc.edu/fac-staff/handbook/financial-support/assistantships.html
   Mecanica confirmada das Graduate Assistantships (regra geral da UC,
   aplicavel a qualquer assistantship, incluindo GIA/GAS do Lindner ja
   cadastrados):
     - Cobre stipend (pagamento por servico) + tuition scholarship
       parcial ou integral.
     - Stipend minimo FY2027 (nao-PhD, .5 FTE): $785,79 quinzenal.
     - Exige matricula em minimo 12 creditos de graduacao/semestre.
     - Assistantship deve ser compativel com a area academica do
       estudante; selecao feita pelo programa academico individual (nao
       ha processo centralizado nem estatisticas de competitividade
       publicadas).
     - Elegibilidade ao GSHI (Graduate Assistant/Fellow Health Insurance
       Award) exige ganhar >= $3.142,40/semestre.
   NAO menciona GIA/GAS pelo nome (esses sao especificos do Lindner,
   ja cadastrados a partir da fonte do enrichment anterior) -- esta
   pagina descreve a regra GERAL da universidade sob a qual GIA/GAS
   operam.

3. https://www.uc.edu/about/financial-aid/aid/scholarships/search.html
   Ferramenta de busca de bolsas (Scholarship Search / ScholarshipUniverse)
   e voltada para "incoming student UC scholarships" de forma geral, sem
   detalhar bolsas de mestrado/internacional especificamente -- sem novidade
   aproveitavel para os campos desta linha.

4. https://www.business.uc.edu/future-students/tuition-fees.html
   So linka de volta para "University scholarships" / "Scholarship search"
   gerais, sem informacao especifica ao MSBA -- sem novidade aproveitavel.

CONCLUSAO SOBRE A PERGUNTA DA USUARIA ("bolsas sao dificeis/possibilidade
real?"): nenhuma fonte oficial publica taxa de aprovacao ou nivel de
competitividade para assistantships/bolsas -- 'Quão Realista (Intl)' segue
como "Moderada" (ja cadastrado, sem novo dado que justifique mudanca).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "U of Cincinnati",
    "Estado": "OH",
    "Cidade": "Cincinnati",
}

UPDATES = {
    "Graduate Assistantships": (
        "Graduate Incentive Award (GIA) e Graduate Assistant Scholarship "
        "(GAS) -- bolsas financiadas pela universidade que cobrem parte "
        "ou toda a tuition e certas taxas. Seguem a regra geral de "
        "assistantship da UC: stipend minimo $785,79 quinzenal (nao-PhD, "
        ".5 FTE, FY2027) + tuition scholarship parcial ou integral; "
        "exige matricula em minimo 12 creditos/semestre; elegibilidade "
        "ao GSHI (seguro saude) se ganhar >= $3.142,40/semestre. Selecao "
        "feita pelo programa academico (Lindner), sem processo "
        "centralizado nem estatistica de competitividade publicada."
    ),
}


def main():
    rotate_backup()
    print(f"Backup rotacionado (política fixa) em: data/msba_programs_database_backup.xlsx")

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
        raise ValueError("Nenhuma linha encontrada para U of Cincinnati (Cincinnati, OH). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (U of Cincinnati, Cincinnati, OH)")

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

    print("\nCampos atualizados (Como Estava -> Como Ficou):")
    for header in UPDATES:
        print(f"  - {header}:")
        print(f"      Antes: {before[header]!r}")
        print(f"      Depois: {after[header]!r}")


if __name__ == "__main__":
    main()
