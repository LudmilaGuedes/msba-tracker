"""
Data enrichment (RODADA 7): atualiza APENAS a linha 'Florida State'
(Herbert Wertheim College of Business, Tallahassee, FL) no
msba_programs_database.xlsx.

Rodada focada em bolsas/waivers para internacionais.

FONTES OFICIAIS (acessadas em 2026-10-19):
  - https://cge.fsu.edu/international-students/funding/oos-tuition-waiver
  - https://gradschool.fsu.edu/funding-awards/tuition-waivers-and-fees
  - https://wertheim.fsu.edu/msba
  - https://wertheim.fsu.edu/students/scholarships/graduate

CORREÇÕES:
- Caminho Tuition Residente: o waiver de out-of-state por mérito (GPA
  3.5/3.7) do Center for Global Engagement é SÓ para graduação ("the
  waivers are only available for undergraduate tuition. Graduate waivers
  are allocated by the academic departments"). Na pós, o waiver de
  out-of-state é reservado a graduate assistants e fellows full-time,
  enquanto houver vagas, e cobre só a tuition (fees continuam por conta
  do aluno).
- Graduate Assistantships / Quão Realista: stipend mínimo da Graduate
  School = US$ 24,45/hora; o waiver acaba se a assistantship terminar ou
  se a carga de créditos cair.
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

VERIFICAR = "Verificar"

MATCH = {
    "Nome da Universidade": "Florida State",
    "Estado": "FL",
    "Cidade": "Tallahassee",
}

UPDATES = {
    "Caminho Tuition Residente": (
        "Não -- alunos F-1 não são elegíveis a tuition de residente (FL "
        "Regulation 7.005). O waiver de out-of-state por mérito (GPA 3.5) "
        "da FSU é só para graduação; na pós, os waivers de out-of-state são "
        "distribuídos pelos departamentos e reservados a graduate "
        "assistants/fellows full-time, enquanto houver vagas. Ou seja, no "
        "MS-BA o único caminho para não pagar a tarifa de internacional é "
        "conseguir uma assistantship (prioridade para quem aplica até 15 jan)"
    ),
    "Graduate Assistantships": (
        "Competitivas, podem isentar até a tuition inteira (o waiver cobre só "
        "tuition, não fees, e acaba se a assistantship terminar ou a carga "
        "de créditos cair); stipend mínimo da Graduate School US$ 24,45/hora; "
        "carga horária e número de vagas do MS-BA não publicados"
    ),
    "Quão Realista (Intl)": (
        "Baixa a moderada -- a única forma de reduzir muito a tuition é uma "
        "assistantship (com waiver de out-of-state), sem número de vagas "
        "publicado; bolsas da Wertheim (incl. Danny Persaud para "
        "internacionais) sem valores divulgados. O "
        "waiver por GPA da FSU é só para graduação. Para concorrer: aplicar "
        "até 15 jan. Sem ajuda, a tuition (~US$ 42k) já é das mais baixas "
        "entre programas fortes"
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
        raise ValueError("Nenhuma linha encontrada para Florida State (Tallahassee, FL). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Florida State, Tallahassee, FL)")

    no_fill = PatternFill(fill_type=None)
    verificar_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    before = {}
    after = {}
    for header, new_value in UPDATES.items():
        col_idx = header_to_col[header]
        cell = ws.cell(row=row_idx, column=col_idx)
        before[header] = cell.value
        cell.value = new_value
        cell.fill = verificar_fill if new_value == VERIFICAR else no_fill
        after[header] = new_value

    wb.save(DB_PATH)

    print("\nCampos atualizados (Como Estava -> Como Ficou):")
    for header in UPDATES:
        print(f"  - {header}:")
        print(f"      Antes: {before[header]!r}")
        print(f"      Depois: {after[header]!r}")


if __name__ == "__main__":
    main()
