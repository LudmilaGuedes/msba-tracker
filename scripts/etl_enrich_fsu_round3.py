"""
Data enrichment (round 3): atualiza APENAS a linha da 'Florida State'
(Herbert Wertheim College of Business, Tallahassee, FL) no
msba_programs_database.xlsx.

Foco: responder pergunta explicita da usuaria -- "a Florida tem um caminho
igual ao do Texas, onde receber uma bolsa (ex: $2.000) da direito a pagar
tuition de residente?"

RESPOSTA: NAO. Confirmado em duas fontes oficiais da FSU:
- https://gradschool.fsu.edu/funding-awards/tuition-waivers-and-fees
  Waivers de out-of-state sao reservados para: funcionarios FSU/estado
  (ate 6 creditos), graduate assistants/fellows, e veteranos/militares.
  Nenhuma mencao a "receber bolsa de valor X libera tuition de residente".
- https://cge.fsu.edu/international-students/funding/oos-tuition-waiver
  Programa de waiver por MERITO ACADEMICO (GPA 3.5-3.7, nao valor de
  bolsa). CRITICO: alunos que ja receberam uma "out-of-state tuition
  reduction scholarship" da propria FSU admissions sao EXPLICITAMENTE
  EXCLUIDOS deste programa de waiver -- os dois beneficios nao se somam.

Ou seja, ao contrario do Texas (onde uma bolsa competitiva de ~$1.000
destrava a tarifa de residente por lei estadual), a Florida nao tem
mecanismo equivalente vinculado a valor de bolsa recebida. Atualiza
'Caminho Tuition Residente' (que ja estava "Não" desde o round2, pelo
motivo do status de visto F-1) para tambem deixar explicita essa
comparacao com o modelo Texas, ja que e um ponto de decisao relevante
para a usuaria ao comparar programas entre estados.
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "Florida State",
    "Estado": "FL",
    "Cidade": "Tallahassee",
}

UPDATES = {
    "Caminho Tuition Residente": (
        "Não -- alunos com visto F-1 não são elegíveis a tuition de "
        "residente (FL Regulation 7.005); só haveria via após mudança "
        "de status para residente permanente (I-485/green card), não "
        "prática durante um programa de 12 meses. DIFERENTE DO TEXAS: "
        "a Flórida não tem mecanismo de bolsa-valor-mínimo destravando "
        "tarifa de residente -- waivers de out-of-state na FSU são só "
        "por mérito acadêmico (GPA 3.5-3.7), assistantship, veterano ou "
        "funcionário; quem já tem bolsa de redução de out-of-state da "
        "própria FSU fica EXCLUÍDO do programa de waiver por mérito"
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
        raise ValueError("Nenhuma linha encontrada para Florida State (Tallahassee, FL). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Florida State, Tallahassee, FL)")

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
