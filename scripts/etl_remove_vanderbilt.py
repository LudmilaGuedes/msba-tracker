"""
Remoção de linha: remove a linha da 'Vanderbilt University (Owen +
Engineering)' (New York City) do msba_programs_database.xlsx.

MOTIVO DA REMOÇÃO (confirmado com o usuário em 26/ago/2026):

1. O programa mais próximo de um MSBA que a Vanderbilt oferece --
   M.S. in Business and Technology (MSBT), sediado no campus de Nova
   York, joint entre Owen Graduate School of Management + School of
   Engineering -- NÃO É chamado oficialmente de "Business Analytics"
   (a própria linha já trazia essa ressalva de uma pesquisa anterior:
   "equivalente mais próximo a MSBA; NÃO é MSBA puro").

2. ACHADO CRÍTICO E DECISIVO: a página oficial do programa
   (vanderbilt.edu/nyc) declara explicitamente que "students who
   require F-1 visa sponsorship are not eligible to enroll in this
   program at the New York City campus due to required federal
   approvals" -- ou seja, o programa EXCLUI ESTRUTURALMENTE, sem
   previsão de mudança, exatamente o perfil de aluno (visto F-1) para
   o qual esta planilha inteira foi construída. Diferente de outros
   casos já encontrados nesta planilha (ex.: LSU com admissões
   temporariamente fechadas), aqui não se trata de uma pausa
   temporária -- é uma barreira estrutural de elegibilidade.

Diante da combinação desses dois fatores (nome do programa não
corresponde ao alvo da pesquisa + exclusão ativa de alunos F-1), o
usuário optou por remover a linha da planilha em vez de mantê-la como
comparação inválida para o perfil de busca (estudante brasileira
precisando de visto F-1).

Fonte: https://www.vanderbilt.edu/nyc/academics-experiential-learning/
ms-business-and-tech/ (pesquisado ago/2026)
"""

import sys

import openpyxl

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "Vanderbilt University (Owen + Engineering)",
    "Cidade": "New York City",
}


def main():
    rotate_backup()
    print(f"Backup rotacionado (política fixa) em: data/msba_programs_database_backup.xlsx")

    wb = openpyxl.load_workbook(DB_PATH)
    ws = wb[SHEET_DATA]

    headers = [cell.value for cell in ws[1]]
    header_to_col = {h: i + 1 for i, h in enumerate(headers)}

    col_uni = header_to_col["Nome da Universidade"]
    col_cidade = header_to_col["Cidade"]

    target_rows = []
    for row in ws.iter_rows(min_row=2):
        if (
            row[col_uni - 1].value == MATCH["Nome da Universidade"]
            and row[col_cidade - 1].value == MATCH["Cidade"]
        ):
            target_rows.append(row[0].row)

    if len(target_rows) == 0:
        raise ValueError("Nenhuma linha encontrada para Vanderbilt University (New York City). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Vanderbilt University (Owen + Engineering), New York City)")

    ws.delete_rows(row_idx, 1)
    print(f"Linha {row_idx} removida.")

    wb.save(DB_PATH)
    print("\nArquivo salvo. Vanderbilt University (Owen + Engineering) removida do banco de dados.")


if __name__ == "__main__":
    main()
