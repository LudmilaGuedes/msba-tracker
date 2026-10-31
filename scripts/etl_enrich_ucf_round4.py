"""
Data enrichment (RODADA 4): atualiza APENAS a linha 'UCF' (Barry S. Miller
College of Business, Orlando, FL) no msba_programs_database.xlsx.

ATENÇÃO: 'UCF (Professional)' é OUTRO programa (MS in Management, BA
track) e NÃO é tocada -- o match exige o nome exato 'UCF'.

Programa: MS in Business Analytics, presencial, STEM.

FONTES OFICIAIS (acessadas em 2026-10-19):
  - https://www.ucf.edu/degree/business-analytics-ms/
  - https://search.cm.ucf.edu/api/v1/programs/1980/ (+ /deadlines/, /outcomes/)
  - https://business.ucf.edu/degree/ms-business-analytics/
  - https://studentaccounts.ucf.edu/tf-graduate/
  - https://graduate.ucf.edu/international-applicants/
  - https://graduate.ucf.edu/applying-to-ucf/

CONFIRMADO (sem alteração): 2026-27 US$ 369,65 / US$ 1.276,48 por crédito
(Student Accounts e Regulation 9.001) -> 30 créditos = US$ 11.090 /
US$ 38.294 (a página da Miller College mostra US$ 1.194,05, valor
desatualizado); 9 disciplinas, 1 ano / 3 termos full-time; prazos intl
15 jan / doméstico 1 jul, só Fall; 3 cartas, resume, goal statement
500-1.000 palavras; avaliação course-by-course só WES ou Josef Silny;
GPA 3.0; GRE/GMAT não exigido; GTA automático do Dept. of Economics.

CORREÇÕES:
- Taxa Aplicação: US$ 30 (era 31). Isenção só para categorias
  específicas (Fulbright internacional, McNair, GEM, funcionários UCF etc.).
- Testes Inglês: os mínimos TOEFL 91 / IELTS 7.0 / Duolingo 125 não
  aparecem em nenhuma página oficial. A College of Graduate Studies
  publica TOEFL 80 (ou 4 na nova escala) / IELTS 6.5 / Duolingo 120, e a
  página do programa não lista mínimo próprio. PTE não é aceito.
- Fees Adicionais: as fees por crédito já estão no valor por crédito;
  fora dele, International Student Fee US$ 50/termo x 3 + ID fee US$ 10
  = US$ 160.
- Formato: face-to-face no campus principal; existe opção part-time.
- Matérias Pré-Requisito: cursos recomendados (não obrigatórios) do
  catálogo.

NÃO ALTERADOS (continuam 'Verificar'): empregabilidade, salário e
empregadores (o feed oficial de outcomes do programa está vazio),
rankings, tamanho da turma, GPA/GMAT médios, taxa de aceitação,
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
    "Nome da Universidade": "UCF",
    "Estado": "FL",
    "Cidade": "Orlando",
}

UPDATES = {
    "Formato": (
        "Presencial (face-to-face no campus principal, Orlando); full-time "
        "em 1 ano / 3 termos ou part-time em 2 anos / 5 termos"
    ),
    "Fees Adicionais USD": 160,  # International Student Fee 50/termo x 3 + ID Service & Access 10/ano
    "Matérias Pré-Requisito": (
        "Recomendados, não obrigatórios: 1 matéria de matemática (ex.: "
        "cálculo I), 1 de estatística e economia (micro). Indicado para "
        "graduados em negócios, economia, computação, engenharia, "
        "matemática, estatística ou ciências naturais; outras áreas "
        "(inclusive ciências sociais) a critério do comitê, conforme o "
        "desempenho em matérias quantitativas"
    ),
    "Testes Inglês Aceitos": (
        "TOEFL iBT 80 (antes de 21/jan/2026) ou 4 (nova escala), código "
        "5233 / IELTS 6.5 / Duolingo 120 (mínimos da College of Graduate "
        "Studies; a página do programa não lista mínimo próprio). Só esses "
        "3 testes -- PTE não aceito; notas válidas por 2 anos e precisam "
        "chegar antes do prazo"
    ),
    "Taxa Aplicação USD": 30,
    "Documentos Necessários": (
        "Resume; goal statement (500-1.000 palavras); 3 cartas de "
        "recomendação (professores e/ou empregadores atuais/anteriores); "
        "avaliação course-by-course com GPA, só pela WES ou Josef Silny "
        "(formação fora dos EUA), + transcript oficial e certificado de "
        "conclusão; TOEFL/IELTS/Duolingo; GRE/GMAT não exigido. Taxa US$ 30 "
        "(isenção só para categorias específicas, como bolsistas Fulbright "
        "internacionais e McNair Scholars -- cgspayment@ucf.edu)"
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
        raise ValueError("Nenhuma linha encontrada para UCF (Orlando, FL). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (UCF, Orlando, FL)")

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
