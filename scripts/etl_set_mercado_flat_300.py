"""
Ajuste solicitado pela usuaria: definir 'Mercado/mês USD' = $300 em TODAS
as 111 linhas do msba_programs_database.xlsx, sem excecao.

IMPORTANTE -- ISTO E UMA SUPOSICAO PLANA (FLAT ASSUMPTION), NAO DADO
PESQUISADO: diferente de todo o resto deste dataset (onde cada valor
numerico tem uma fonte oficial rastreavel, documentada em um script
etl_enrich_*.py proprio), este valor de $300/mes para compras de mercado
foi escolhido pela usuaria como uma estimativa unica e uniforme para
TODAS as cidades, substituindo tanto celulas "Verificar" (72 linhas)
quanto valores JA PESQUISADOS anteriormente com fonte oficial (39
linhas -- ex.: UGA $350 de orcamento oficial da universidade, U of South
Carolina $350, Wake Forest, etc.). A decisao de usar um numero fixo em
vez de custo de vida especifico por cidade (que varia MUITO entre, por
exemplo, San Francisco/CA e Edmond/OK) foi explicita da usuaria -- este
script apenas executa essa decisao, nao reavalia se e realista para cada
cidade individualmente.

As 39 linhas com valor pesquisado anteriormente e que serao sobrescritas
permanecem documentadas nos seus respectivos scripts etl_enrich_*.py
originais (ex.: etl_enrich_uga.py, etl_enrich_u_of_south_carolina.py)
para quem quiser recuperar o dado pesquisado no futuro -- este script
NAO apaga esse historico, so sobrescreve o valor atual na planilha.
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"
CAMPO = "Mercado/mês USD"
VALOR_FIXO = 300


def main():
    rotate_backup()
    print("Backup rotacionado (política fixa) em: data/msba_programs_database_backup.xlsx")

    wb = openpyxl.load_workbook(DB_PATH)
    ws = wb[SHEET_DATA]

    headers = [cell.value for cell in ws[1]]
    header_to_col = {h: i + 1 for i, h in enumerate(headers)}
    if CAMPO not in header_to_col:
        raise ValueError(f"Coluna inexistente no schema: {CAMPO}")

    col_uni = header_to_col["Nome da Universidade"]
    col_campo = header_to_col[CAMPO]
    no_fill = PatternFill(fill_type=None)

    total = 0
    era_verificar = 0
    sobrescritas = []
    for row in ws.iter_rows(min_row=2):
        uni = row[col_uni - 1].value
        if not uni:
            continue
        cell = row[col_campo - 1]
        antes = cell.value
        if isinstance(antes, str) and antes.strip().lower() == "verificar":
            era_verificar += 1
        elif antes != VALOR_FIXO:
            sobrescritas.append((uni, antes))
        cell.value = VALOR_FIXO
        cell.fill = no_fill
        total += 1

    wb.save(DB_PATH)

    print(f"\n'{CAMPO}' definido como ${VALOR_FIXO} em {total} linhas.")
    print(f"  - Estavam 'Verificar': {era_verificar}")
    print(f"  - Tinham valor pesquisado (sobrescrito): {len(sobrescritas)}")
    if sobrescritas:
        print("\nLinhas com valor pesquisado que foram sobrescritas:")
        for uni, antes in sobrescritas:
            print(f"    {uni}: {antes!r} -> {VALOR_FIXO}")


if __name__ == "__main__":
    main()
