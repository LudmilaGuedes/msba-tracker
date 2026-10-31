"""
Data enrichment: atualiza APENAS a linha 'Dakota State' (College of
Business & IS, Madison, SD) no msba_programs_database.xlsx.

Foco exclusivo: MS in Analytics and Applied Artificial Intelligence
(MSAA), trilha presencial (confirmado que o programa existe em formato
on-campus E online -- a planilha ja documentava a trilha correta,
inclusive com notas ja existentes tipo "GA (on-campus only)").

Tuition NAO foi alterada: a pagina oficial cita "$3.365/semestre" como
custo "estimado", sem distinguir residente/nao-residente, o que nao bate
com precisao com o valor por credito ja detalhado na planilha ($698/
credito, resultando em $20.943 para 30 creditos). Como a cifra nova e
vaga ("estimated cost") e a antiga e mais especifica e ja consistente
internamente, o valor antigo foi mantido.

Fontes principais (pesquisadas manualmente, fora deste script):
- https://dsu.edu/programs/msaa/ (formato dual on-campus/online confirmado,
  duracao flexivel de 2-5 anos, taxa de emprego/continuacao de estudos 92,9%)
- https://dsu.edu/programs/msaa/msaa-timelines-requirements.html (deadlines
  completos por semestre e status, documentos, pre-requisitos)
- busca por admissao (TOEFL iBT 76 / PBT 550 / IELTS 6.0 como referencia)
- busca por custo de vida em Madison, SD (aluguel medio ~$856/mes, 2026 --
  cidade bem acessivel)

Nao encontrado mesmo apos pesquisa profunda (mantido "Verificar"): salario
medio especifico do programa (so achei o salario medio ocupacional de
"computer systems analysts" no estado de SD, uma estatistica de mercado
generica que NAO foi usada por nao ser especifica dos formandos do MSAA),
tamanho de turma, ranking, GPA medio admitidos, GMAT medio, taxa de
aceitacao, entrevista.
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.stdout.reconfigure(encoding="utf-8")

from db_retention import rotate_backup  # noqa: E402

DB_PATH = "data/msba_programs_database.xlsx"
SHEET_DATA = "MSBA Programs"

VERIFICAR = "Verificar"

MATCH = {
    "Nome da Universidade": "Dakota State",
    "Estado": "SD",
    "Cidade": "Madison",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Deadline Aplicação": (
        "Fall: 15 ago (doméstico) / 15 abr (intl); Spring: 30 nov "
        "(doméstico) / 1 nov (intl); Summer: 30 abr (só doméstico)"
    ),
    "Duração (meses)": "24-60 (extremamente flexível; full-time ou part-time)",
    "Documentos Necessários": (
        "Application (taxa $40, já confirmada), Purpose Statement (200-300 "
        "palavras), transcripts oficiais; internacionais: transcripts "
        "traduzidos + comprovante de proficiência em inglês"
    ),
    "Matérias Pré-Requisito": (
        "Coursework prévio em ferramentas de dados (SAS, R, Python, Excel, "
        "SQL) e princípios estatísticos; cursos de nivelamento podem ser "
        "exigidos se não atendido"
    ),
    "Testes Inglês Aceitos": "TOEFL iBT 76 / TOEFL PBT 550 / IELTS 6.0 (PTE também aceito)",
    "Taxa Empregabilidade %": (
        "92,9% (emprego OU continuação de estudos combinados, não é taxa "
        "de emprego pura)"
    ),
    "Aluguel Privativo/mês USD": 856,  # aluguel médio em Madison, SD, 2026
}


def main():
    rotate_backup()
    print("Backup da versão anterior salvo em: data/msba_programs_database_backup.xlsx")

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
        raise ValueError("Nenhuma linha encontrada para Dakota State (Madison, SD). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Dakota State, Madison, SD)")

    no_fill = PatternFill(fill_type=None)

    before = {}
    after = {}
    for header, new_value in UPDATES.items():
        col_idx = header_to_col[header]
        cell = ws.cell(row=row_idx, column=col_idx)
        before[header] = cell.value
        cell.value = new_value
        cell.fill = no_fill  # nenhum novo valor aqui e "Verificar"
        after[header] = new_value

    wb.save(DB_PATH)

    print("\nCampos atualizados (Como Estava -> Como Ficou):")
    for header in UPDATES:
        print(f"  - {header}:")
        print(f"      Antes: {before[header]!r}")
        print(f"      Depois: {after[header]!r}")


if __name__ == "__main__":
    main()
