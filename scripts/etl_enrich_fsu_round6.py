"""
Data enrichment (RODADA 6): atualiza APENAS a linha 'Florida State'
(Herbert Wertheim College of Business, Tallahassee, FL) no
msba_programs_database.xlsx.

Programa: MS in Business Analytics (MS-BA), presencial, full-time, STEM.

FONTES OFICIAIS (acessadas em 2026-10-19):
  - https://wertheim.fsu.edu/msba
  - https://wertheim.fsu.edu/graduate/program-costs
  - https://wertheim.fsu.edu/students/scholarships/graduate
  - https://wertheim.fsu.edu/program-rankings
  - https://gradschool.fsu.edu/admissions/international-admissions

CONFIRMADO (sem alteração): 2026-27, 33 créditos, US$ 479,32 / US$ 1.278,04
por crédito (fees incluídas) -> US$ 15.817,56 / US$ 42.175,32; lock-step
de 3 semestres, início só no Summer (fim de junho a início de maio);
prazos 15 jan (prioridade + assistantship) e 1 mar (final); taxa US$ 30
sem isenção; GMAT/GRE opcional; Duolingo 125; pré-requisitos cálculo,
estatística e programação; bolsas sem valores publicados.

CORREÇÕES:
- Testes Inglês: as notas do MS-BA são mais altas que o mínimo da
  Graduate School (TOEFL 4.0 / IELTS 6.5 / PTE 55); o Duolingo 125 é igual
  nos dois. A isenção vale só para diploma dos EUA ou de país de língua
  inglesa -- diploma brasileiro não isenta.
- Documentos: tradução certificada literal para documentos fora do
  inglês; comprovação financeira do 1º ano + fundos para o programa todo
  (enviada no pedido do I-20, depois da admissão).
- Dificuldade Entrada: a página do programa descreve a admissão como
  "highly competitive".

NÃO ALTERADOS (continuam 'Verificar'): Taxa Empregabilidade, Salário
Médio Pós, Principais Empregadores (a Wertheim não publica relatório de
colocação do MS-BA), Ranking US News/QS (a página de rankings só traz o
MS-MIS online, #18 US News 2026), Tamanho Turma, GPA/GMAT médios, Taxa
Aceitação, Entrevista.
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
    "Testes Inglês Aceitos": (
        "Mínimos do MS-BA (mais altos que os da Graduate School): TOEFL 600 "
        "(PBT, antes de 21/jan/2026) ou 5.0 (nova escala, a partir de "
        "21/jan/2026) / IELTS Academic 7.0 / PTE 66 / Duolingo 125 / "
        "Cambridge C1 Advanced 180 / Michigan 55 (testes feitos nos últimos "
        "2 anos). Isenção só para diploma dos EUA ou de país de língua "
        "inglesa -- diploma brasileiro não isenta"
    ),
    "Documentos Necessários": (
        "Applicant statement; resume/CV com datas e cargos (tempo integral "
        "ou parcial); 2 cartas de recomendação (3 preferível) de "
        "empregadores ou professores; 1 transcript oficial de cada "
        "instituição (internacionais: avaliação course-by-course NACES "
        "obrigatória antes da análise, e prova do diploma se o histórico não "
        "mostrar; tradução certificada literal se não estiver em inglês); "
        "GMAT/GRE opcional (códigos PN8K567 / 5219); teste de inglês; taxa "
        "US$ 30 (sem isenção na pós). Depois da admissão, para o I-20: "
        "comprovação financeira do 1º ano e de fundos para o programa todo"
    ),
    "Dificuldade Entrada": (
        "Moderada a alta (o programa descreve a admissão como \"highly "
        "competitive\"; GMAT/GRE opcional e sem experiência obrigatória)"
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
