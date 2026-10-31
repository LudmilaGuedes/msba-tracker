"""
Data enrichment: atualiza APENAS a linha 'Wake Forest University' (School
of Business, Winston-Salem, NC) no msba_programs_database.xlsx.

Foco exclusivo: MS in Business Analytics (MSBA), presencial, STEM.

CORRECAO: a planilha tinha 'Deadline Aplicação': '1 nov 2025...' -- uma
data ja passada em relacao a data atual (ago/2026). Atualizado para o
ciclo de admissao 2026-2027 (5 rounds), com dados oficiais.

Fontes principais (pesquisadas manualmente, fora deste script):
- https://business.wfu.edu/masters-in-business-analytics/application-process/
  (5 rounds com datas exatas para o ciclo Summer 2027, entrevista por
  convite, deposito de matricula $1.500)
- https://business.wfu.edu/masters-in-business-analytics/class-profile/
  (turma de 104 alunos, GPA medio 3.5, GMAT medio 685, GRE medio 320 --
  media dos ultimos 3 anos; corrige o GMAT medio aproximado que ja estava
  na planilha)
- busca por admissao (GMAT/GRE obrigatorio salvo dispensa por GPA+cursos
  quantitativos, taxa de inscricao $100, TOEFL/IELTS)
- busca por outcomes (salario medio $85.806, principais empregadores)
- busca por custo de vida em Winston-Salem, NC (aluguel medio de moradia
  estudantil ~$1.124/mes, 2026)

Nao encontrado mesmo apos pesquisa profunda (mantido "Verificar"): taxa de
aceitacao, ranking, experiencia profissional media, GPA minimo formal,
matérias pre-requisito.
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
    "Nome da Universidade": "Wake Forest University",
    "Estado": "NC",
    "Cidade": "Winston-Salem",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Deadline Aplicação": (
        "Ciclo 2026-27 (5 rounds): R1 15 out/2026; R2 15 jan/2027 (final "
        "para F-1); R3 15 mar/2027; R4 15 mai/2027; R5 30 jun/2027"
    ),  # CORRECAO: valor antigo (1 nov 2025) ja estava no passado
    "Rounds Aplicação": "5 rounds",
    "Entrevista": (
        "Por convite apenas, após inscrição; virtual ou presencial (com "
        "aula e tour do campus)"
    ),
    "GMAT/GRE Obrigatório": (
        "Obrigatório, salvo dispensa (GPA ≥ 3.2 + 2 notas B+ em cursos "
        "quantitativos de instituição dos EUA)"
    ),
    "GMAT Médio": "685 (GMAT) / 320 (GRE) — média dos últimos 3 anos",  # atualização (planilha tinha "~660/~315")
    "GPA Médio Admitidos": 3.5,
    "Tamanho Turma": 104,  # média dos últimos 3 anos
    "Testes Inglês Aceitos": (
        "TOEFL iBT 80 (antes de 21/jan/2026) ou 4.5 (a partir de "
        "21/jan/2026, conforme fonte oficial); IELTS 6.5; dispensa possível "
        "para quem residiu/trabalhou nos EUA por período extenso"
    ),
    "Taxa Aplicação USD": 100,
    "Documentos Necessários": (
        "Application online, resume, transcript não oficial; depósito de "
        "matrícula de $1.500 (abatido na 1ª mensalidade) após admissão"
    ),
    "Salário Médio Pós USD": 85806,
    "Principais Empregadores": (
        "Wells Fargo, Bank of America, Dimensional Fund Advisors, "
        "Jefferies, JP Morgan, AlphaSights, UBS, Merrill Lynch, Barclays, "
        "Deloitte, Goldman Sachs"
    ),
    "Aluguel Privativo/mês USD": 1124,  # moradia estudantil em Winston-Salem, NC, 2026
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
        raise ValueError("Nenhuma linha encontrada para Wake Forest University (Winston-Salem, NC). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Wake Forest University, Winston-Salem, NC)")

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
