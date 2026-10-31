"""
Remoção de linhas: remove 'Northwestern' (Evanston, IL) e 'Texas Tech'
(Lubbock, TX) do msba_programs_database.xlsx.

MOTIVO DA REMOÇÃO (confirmado com o usuário em set/2026, durante uma
pesquisa aprofundada sobre aceitação do Duolingo English Test):

Nenhuma das duas universidades oferece um programa efetivamente chamado
"MSBA" / "Master of Science in Business Analytics". A base as havia
registrado sob esse rótulo, mas o programa real encontrado em cada uma é:

  - Northwestern (McCormick): "MS in Machine Learning and Data Science"
    (MLDS) -- renomeado a partir do antigo "MS in Analytics" (MSiA) por
    volta de 2023. Não é um programa de business analytics hospedado
    numa business school.
  - Texas Tech (Rawls College of Business): não existe "MSBA" no Rawls;
    o mais próximo é "MS in Data Science" (ou "MS in Marketing Research
    and Analytics"), programas distintos do perfil buscado.

Como nenhum dos dois é de fato um MSBA, o usuário optou por remover as
duas linhas da base em vez de mantê-las rotuladas incorretamente.

Fonte: pesquisa via agentes (WebSearch/WebFetch) em set/2026 --
  Northwestern: https://www.mccormick.northwestern.edu/machine-learning-data-science/admissions/
  Texas Tech: https://www.depts.ttu.edu/gradschool/admissions/InternationalProspectiveStudents.php
"""

import sys

import openpyxl

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCHES = [
    {"Nome da Universidade": "Northwestern", "Estado": "IL", "Cidade": "Evanston"},
    {"Nome da Universidade": "Texas Tech", "Estado": "TX", "Cidade": "Lubbock"},
]


def main():
    rotate_backup()
    print("Backup rotacionado (política fixa) em: data/msba_programs_database_backup.xlsx")

    wb = openpyxl.load_workbook(DB_PATH)
    ws = wb[SHEET_DATA]

    headers = [cell.value for cell in ws[1]]
    header_to_col = {h: i + 1 for i, h in enumerate(headers)}

    col_uni = header_to_col["Nome da Universidade"]
    col_estado = header_to_col["Estado"]
    col_cidade = header_to_col["Cidade"]

    rows_to_delete = []

    for match in MATCHES:
        target_rows = []
        for row in ws.iter_rows(min_row=2):
            if (
                row[col_uni - 1].value == match["Nome da Universidade"]
                and row[col_estado - 1].value == match["Estado"]
                and row[col_cidade - 1].value == match["Cidade"]
            ):
                target_rows.append(row[0].row)

        if len(target_rows) == 0:
            raise ValueError(
                f"Nenhuma linha encontrada para {match['Nome da Universidade']} "
                f"({match['Cidade']}, {match['Estado']}). Abortando."
            )
        if len(target_rows) > 1:
            raise ValueError(
                f"Mais de uma linha corresponde ao filtro para {match['Nome da Universidade']} "
                f"(linhas {target_rows}). Abortando por segurança."
            )

        row_idx = target_rows[0]
        print(
            f"Linha alvo confirmada: {row_idx} "
            f"({match['Nome da Universidade']}, {match['Cidade']}, {match['Estado']})"
        )
        rows_to_delete.append(row_idx)

    # Deletar em ordem decrescente para não invalidar os índices das
    # linhas ainda não removidas.
    for row_idx in sorted(rows_to_delete, reverse=True):
        ws.delete_rows(row_idx, 1)
        print(f"Linha {row_idx} removida.")

    wb.save(DB_PATH)
    print("\nArquivo salvo. Northwestern e Texas Tech removidas do banco de dados.")


if __name__ == "__main__":
    main()
