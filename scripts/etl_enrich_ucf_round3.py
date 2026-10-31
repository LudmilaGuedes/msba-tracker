"""
Data enrichment (round 3): atualiza APENAS a linha 'UCF' (Barry S. Miller
College of Business, Orlando, FL) no msba_programs_database.xlsx.

Foco: resposta a pergunta explicita da usuaria -- "essas bolsas sao
realistas [para internacional]?" -- que revelou uma lacuna critica no
round 2: o sistema A2O (citado no round 2 como canal de bolsas) e, na
pratica, quase todo FECHADO para nao-cidadaos americanos.

ACHADO CRITICO (corrige enfase do round 2):
https://www.ucf.edu/financial-aid/apply/international/ -- confirma
explicitamente: "The Office of Student Financial Assistance does not
award scholarships for non-U.S. citizens" (a excecao e um numero pequeno
e variavel de bolsas "periodically offered" a nao-cidadaos, listadas a
parte). Estudantes DACA/TPS/indocumentados tambem NAO sao elegiveis a
fundos de bolsa da UCF.
https://global.ucf.edu/international-students-scholars-and-professionals/scholarships/
-- confirma que ha APENAS UMA bolsa universitaria aberta especificamente a
internacionais (Student Government Fall International Scholarship); para
alem dela, o caminho realista de financiamento p/ grad internacional e
"assistantships, fellowships, and tuition support" via contato direto com
o diretor do programa e os emails gradfellowship@ucf.edu /
gradassistantship@ucf.edu -- NAO o sistema A2O.

Ou seja: o A2O (destacado no round 2 como um dos 3 canais de bolsa) NAO e,
na pratica, um canal realista para a maioria dos candidatos internacionais
ao MS-BA. Os canais que de fato se aplicam a internacionais sao: GTA
(Department of Economics, ja documentado) e Graduate Fellowships (College
of Graduate Studies, tambem ja documentado) -- nenhum dos dois tem
restricao de cidadania confirmada nas fontes consultadas.

'Quão Realista (Intl)' mantido como "Moderada" (nao ha dado novo que
justifique rebaixar para "Dificil"), mas com ressalva explicita sobre essa
restricao de cidadania no A2O, para nao induzir a usuaria a superestimar
o numero de canais de bolsa realmente acessiveis a ela.
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "UCF",
    "Estado": "FL",
    "Cidade": "Orlando",
}

UPDATES = {
    "Tipos de Bolsa": (
        "GTA automático (Department of Economics, vagas limitadas, "
        "competitivo -- candidatos ao MS-BA são automaticamente "
        "considerados, mas seleção não é garantida). Graduate "
        "Fellowships da College of Graduate Studies: > $2 milhões/ano "
        "por mérito acadêmico, sem obrigação de trabalho (contato: "
        "gradfellowship@ucf.edu). RESSALVA IMPORTANTE PARA "
        "INTERNACIONAIS: o sistema A2O (Access to Opportunities) e a "
        "maioria das bolsas do Office of Student Financial Assistance "
        "NÃO são abertas a não-cidadãos americanos -- só existe 1 bolsa "
        "universitária aberta especificamente a internacionais (Student "
        "Government Fall International Scholarship); GTA e Graduate "
        "Fellowships continuam sendo os canais realistas para "
        "internacional"
    ),
    "Quão Realista (Intl)": (
        "Moderada -- GTA (Dept. Economics) e Graduate Fellowships "
        "(College of Graduate Studies) são acessíveis a internacionais, "
        "mas ambos competitivos/vagas limitadas, sem taxa de aprovação "
        "publicada. O sistema A2O de bolsas do College of Business é, "
        "na prática, majoritariamente restrito a cidadãos americanos -- "
        "não conte com ele como canal principal"
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
        raise ValueError("Nenhuma linha encontrada para UCF (Orlando, FL). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (UCF, Orlando, FL) -- NAO 'UCF (Professional)'")

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
