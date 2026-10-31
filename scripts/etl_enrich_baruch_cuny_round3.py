"""
Data enrichment (round 3): atualiza APENAS a linha da 'Baruch CUNY'
(Zicklin School of Business, New York, NY) no msba_programs_database.xlsx.

Objetivo desta rodada: preencher empregabilidade, salário pós-programa
e principais empregadores, que estavam 'Verificar'.

FONTE PRINCIPAL (oficial, específica do programa MS Business Analytics,
não MBA nem média geral da escola): relatório oficial do Graduate
Career Management Center (GCMC) da Zicklin, turma 2023-2024:
  https://zicklin.baruch.cuny.edu/wp-content/uploads/sites/10/2025/03/Employment-Report-2024-Data-Business-Analytics.pdf
(acessado via download direto em 2026-09-08; o subdomínio
zicklin.baruch.cuny.edu segue com erro de certificado SSL em todas as
tentativas de acesso via ferramenta de fetch padrão -- mesma limitação
já documentada em etl_enrich_baruch_cuny.py -- contornado com download
direto do PDF).

DADOS DO RELATÓRIO (MS Business Analytics, turma 2023-2024):
- 149 formados, 124 responderam (83% resposta)
- Emprego na formatura: domésticos 84/94 empregados (89%); internacionais
  19/30 empregados (63%), 11/30 buscando emprego (37%)
  -> Total empregados: 84+19 = 103 de 124 respondentes = 83%
- Salários (autodeclarados, AMOSTRA MUITO PEQUENA: apenas 12 de 124 =
  10% responderam): low $60.000 / média $93.083 / mediana $83.500 /
  high $170.000
- Empregadores citados (títulos de vaga confirmados de formandos, e/ou
  empresas que postaram vagas para o programa no ano): Amazon, American
  Express, JPMorgan Chase & Co., Bank of America, UBS, Mastercard,
  LinkedIn, Fitch Ratings, JetBlue, Estée Lauder Companies, S&P Global,
  Northwell Health, Con Edison, National Grid, DTCC, Adobe Systems,
  Unilever, ZS, Bloomingdale's

CAMPOS ATUALIZADOS:
- 'Taxa Empregabilidade %': 83 (103/124 respondentes empregados na
  formatura; NÃO é taxa sobre os 149 formados totais, é sobre os que
  responderam a pesquisa -- única base disponível na fonte)
- 'Salário Médio Pós USD': 93083 (média/"mean" do relatório -- rotulado
  como 'Médio' no schema, que corresponde a média, não mediana; CAVEAT
  DE AMOSTRA documentado no comentário deste campo abaixo, pois é
  baseado em apenas 12 respondentes de 124, alta incerteza)
- 'Principais Empregadores': lista dos empregadores confirmados/que
  postaram vagas, conforme relatório oficial

CAMPOS **NÃO** ATUALIZADOS nesta rodada (permanecem 'Verificar'):
Ranking US News, Ranking QS, Taxa Aceitação %, GMAT Médio, GPA Médio
Admitidos, GPA Mínimo -- busca web não encontrou dados OFICIAIS
específicos do MS in Business Analytics (apenas números de MBA ou
médias gerais da Zicklin, que não devem ser atribuídos ao MSBA sem
confirmação -- risco de contaminar o dado com número de programa
diferente, mesmo erro que foi corrigido para tuition no round 1).
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
    "Nome da Universidade": "Baruch CUNY",
    "Estado": "NY",
    "Cidade": "New York",
}

UPDATES = {
    "Taxa Empregabilidade %": 83,  # 103 de 124 respondentes empregados na formatura (turma 2023-2024)
    "Salário Médio Pós USD": 93083,  # média self-reported; AMOSTRA PEQUENA (12 de 124 respondentes, 10%) -- baixa confiabilidade
    "Principais Empregadores": (
        "Amazon, American Express, JPMorgan Chase & Co., Bank of "
        "America, UBS, Mastercard, LinkedIn, Fitch Ratings, JetBlue, "
        "Estée Lauder Companies, S&P Global, Northwell Health, Con "
        "Edison, National Grid, DTCC, Adobe Systems, Unilever, ZS, "
        "Bloomingdale's (turma 2023-2024, relatório oficial GCMC)"
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
        raise ValueError("Nenhuma linha encontrada para Baruch CUNY (New York, NY). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Baruch CUNY, New York, NY)")

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
