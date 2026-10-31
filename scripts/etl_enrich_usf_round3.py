"""
Data enrichment (RODADA 3): atualiza APENAS a linha 'USF' (University of
South Florida, Muma College of Business, Tampa, FL) no
msba_programs_database.xlsx. NÃO confundir com a University of San
Francisco (usfca.edu).

Programa: a USF não oferece mais MS in Business Analytics presencial. A
linhagem MS-BAIS -> MS in AI and Business Analytics virou o MS in
Artificial Intelligence in Business, com dois majors: Digital
Transformation (100% online) e Enterprise Integration (presencial,
full-time). Esta linha segue o Enterprise Integration.

FONTES OFICIAIS (acessadas em 2026-10-19):
  - https://www.usf.edu/business/graduate/ms-ai-business-enterprise-integration/index.aspx
  - https://www.usf.edu/business/graduate/ms-ai-business/index.aspx
  - https://www.usf.edu/business/graduate/masters.aspx
  - https://www.usf.edu/admissions/international/admission-information/graduate/academic-requirements.aspx

CONFIRMADO (sem alteração): 30 créditos (10 disciplinas de 3), 12-16
meses, presencial e full-time; US$ 466,34 / US$ 1.030,20 por crédito ->
US$ 13.990 / US$ 30.906; Duolingo 110, TOEFL 79/4.5, IELTS 6.5, PTE 53.

CORREÇÕES:
- STEM Designated: era "Sim (CIP 11.0501)", fora das opções do enum
  (Sim/Não) e fora do filtro do dashboard. A página do programa diz
  "STEM-designated"; o CIP não aparece mais.
- Início: a página atual diz "Fall or Spring" e não fala mais em
  aprovação pendente; o programa começou no Fall 2026 (masters.aspx).
- GPA Mínimo: 3.0 ("B" average em todo o trabalho tentado), regra da USF
  para pós-graduação.
- Testes Inglês: acrescentados Cambridge C1/C2, GRE/GMAT Verbal e a
  validade de 2 anos.
- Deadline: a página não publica prazo. Mantém 'Verificar', sem datas de
  outros programas da USF (o parser de rounds pegaria o mês errado).
- Rounds: "rolling" não tem fonte -> 'Verificar'.
- Documentos: cartas/statement vêm da página do programa predecessor;
  acrescentada tradução certificada e avaliação NACES do diploma.

NÃO ALTERADOS (continuam 'Verificar'): fees, seguro, empregabilidade,
salário, empregadores (programa novo, sem turma formada), rankings,
tamanho da turma, pré-requisitos, entrevista, prazos.
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
    "Nome da Universidade": "USF",
    "Estado": "FL",
    "Cidade": "Tampa",
}

UPDATES = {
    "STEM Designated": "Sim",
    "Início": (
        "Fall ou Spring (o MS in AI in Business começou no Fall 2026; a "
        "página atual da trilha presencial não fala mais em aprovação "
        "pendente)"
    ),
    "GPA Mínimo": 3,  # "B" average (3.00/4.00) em todo o trabalho tentado -- regra geral da pós da USF
    "Testes Inglês Aceitos": (
        "TOEFL iBT 79 (antes de 21/jan/2026) ou 4.5 (nova escala) / IELTS "
        "6.5 / Duolingo 110 / PTE-A 53 / TOEFL Essentials 8.5 / Cambridge "
        "B2 First 176 (subnotas 169) / C1 Advanced 180 / C2 Proficiency 200 "
        "/ GRE Verbal 153 ou GMAT Verbal 30. Notas de até 2 anos antes do "
        "início (mínimos da USF para pós; o programa não publica exigência "
        "própria)"
    ),
    "Deadline Aplicação": (
        "Verificar -- a página do programa não publica prazos (entrada no "
        "Fall ou no Spring); confirmar com muma-msaib@usf.edu"
    ),
    "Rounds Aplicação": VERIFICAR,
    "Documentos Necessários": (
        "Aplicação em usf.edu/application, taxa US$ 30; histórico de todas "
        "as instituições, com tradução certificada se não estiver em inglês "
        "e avaliação do diploma por membro NACES (internacionais); teste de "
        "inglês; 2 cartas de recomendação (acadêmicas ou profissionais), "
        "statement of purpose e experiência profissional relevante -- estes itens vêm da página do "
        "programa predecessor, a página nova não lista documentos"
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
        raise ValueError("Nenhuma linha encontrada para USF (Tampa, FL). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (USF, Tampa, FL)")

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
