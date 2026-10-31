"""
Data enrichment (round 3): atualiza a linha da 'University of North
Florida (UNF)' no msba_programs_database.xlsx, a pedido do usuário
(pergunta direta: "quais são os requisitos de aplicação?").

FONTE: página oficial de critérios de admissão do MSBA, acessada via
WebFetch em 2026-09-08:
https://www.unf.edu/coggin/graduate/msba/msba-admission-criteria.html

ACHADO NOVO: além das "2 cartas de recomendação profissionais" já
documentadas (exigidas apenas se o candidato não se qualificar para a
dispensa de teste padronizado E tiver pontuação insuficiente), a
mesma condição também exige uma "Letter of intent" (carta de
intenção, máximo 2 páginas) -- não estava documentada antes. Também
confirmado explicitamente: currículo profissional e transcripts
oficiais são exigidos de todos (não só de quem precisa da dispensa).

'Entrevista': a página oficial de critérios do MSBA NÃO menciona
entrevista em nenhum momento (diferente do GlobalMBA da mesma escola,
que tem entrevista seletiva confirmada) -- ATUALIZADO de 'Verificar'
para 'Não mencionada/não exigida', com a ressalva de que ausência de
menção não é uma confirmação 100% explícita de que nunca há
entrevista.
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
    "Nome da Universidade": "University of North Florida (UNF)",
    "Estado": "FL",
    "Cidade": "Jacksonville",
}

UPDATES = {
    "Documentos Necessários": (
        "Currículo profissional e transcripts oficiais (exigidos de "
        "todos); 2 cartas de recomendação profissionais + carta de "
        "intenção (máx. 2 páginas) -- exigidas apenas se Index Score "
        "ou teste padronizado ficarem abaixo do necessário; "
        "documentos de imigração para candidatos internacionais"
    ),
    "Entrevista": (
        "Não mencionada nas páginas oficiais de admissão do MSBA "
        "(diferente do GlobalMBA da mesma escola, que tem entrevista "
        "seletiva confirmada) -- provável que não seja exigida, mas "
        "ausência de menção não é confirmação 100% explícita"
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
        raise ValueError("Nenhuma linha encontrada para UNF. Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (UNF)")

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
