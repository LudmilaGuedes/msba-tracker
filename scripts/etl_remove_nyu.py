"""
Remoção de linha: remove a linha da 'New York University (NYU)' (Stern
School of Business, MS in Business Analytics and AI / MSBAi) do
msba_programs_database.xlsx.

MOTIVO DA REMOÇÃO (confirmado com o usuário em ago/2026, mesmo critério
já usado no caso Vanderbilt University nesta mesma planilha):

1. O programa é oficialmente PART-TIME e MODULAR (não full-time
   on-campus): estrutura de 6 módulos presenciais -- alguns em Nova
   York, outros em localidades internacionais rotativas -- num modelo
   descrito pela própria Stern como "executive-friendly". Isso diverge
   do formato "presencial" tradicional que esta planilha busca (a
   linha já trazia essa nota parcialmente, mas sem deixar claro o
   impacto no visto).

2. ACHADO CRÍTICO E DECISIVO: a página oficial de FAQ do programa
   declara explicitamente que, apesar de o programa ser STEM
   designated, "students are not eligible for OPT" devido à natureza
   part-time/modular -- ou seja, o campo 'STEM Designated' = 'Sim' já
   presente na planilha era tecnicamente verdadeiro mas MATERIALMENTE
   ENGANOSO para o propósito desta base de dados: o benefício prático
   de uma designação STEM (extensão de OPT pós-formatura) simplesmente
   NÃO se aplica aqui. Além disso, o visto F-1 só é válido para os
   períodos específicos de cada módulo (não o ano inteiro), exigindo
   I-20s separados por cluster de módulos -- uma estrutura de visto
   incomum e onerosa, incompatível com o perfil de busca (estudante
   brasileira em F-1 buscando um programa STEM full-time com OPT
   integral pós-formatura).

Diante da combinação desses dois fatores (formato part-time/modular,
divergente do "presencial" buscado + STEM sem benefício de OPT real),
o usuário optou por remover a linha da planilha em vez de mantê-la
como uma opção que pareceria, à primeira vista (via o campo 'STEM
Designated'), equivalente às demais linhas STEM da base, quando na
prática não oferece o mesmo benefício migratório.

Fonte: https://www.stern.nyu.edu/programs-admissions/ms-business-analytics/resources/frequently-asked-questions
(pesquisado ago/2026, via WebFetch — orçamento de WebSearch da sessão
esgotado)
"""

import sys

import openpyxl

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "New York University (NYU)",
    "Estado": "NY",
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
        raise ValueError("Nenhuma linha encontrada para New York University (NYU) (New York City, NY). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (New York University (NYU), New York City, NY)")

    ws.delete_rows(row_idx, 1)
    print(f"Linha {row_idx} removida.")

    wb.save(DB_PATH)
    print("\nArquivo salvo. New York University (NYU) removida do banco de dados.")


if __name__ == "__main__":
    main()
