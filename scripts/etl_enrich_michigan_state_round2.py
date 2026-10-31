"""
Data enrichment (RODADA 2): atualiza APENAS a linha 'Michigan State'
(Broad College of Business, East Lansing, MI) no
msba_programs_database.xlsx.

Esta linha já passou por uma rodada de enrichment (script
etl_enrich_michigan_state.py), restando ~22/54 campos 'Verificar'.

LIMITAÇÃO SEVERA DE ACESSO (documentada para transparência): o domínio
broad.msu.edu (e o domínio raiz msu.edu) retornaram conteúdo VAZIO em
TODAS as tentativas via WebFetch nesta rodada -- inclusive na URL
específica do programa fornecida diretamente pelo usuário
(https://broad.msu.edu/masters/business-data-science-analytics/),
que também retornou vazio. O site parece ser renderizado inteiramente
via JavaScript client-side, sem conteúdo no HTML bruto que o WebFetch
recebe -- uma limitação técnica da ferramenta, não do link em si.
Contornado parcialmente navegando por subdomínios alternativos do MSU
que NÃO usam essa renderização (grad.msu.edu, oiss.isp.msu.edu), que
trouxeram apenas políticas GERAIS da universidade, não específicas do
programa MSBA.

'Processo TA/RA' preenchido com a política GERAL de assistantships da
Graduate School do MSU (não específica do Broad College, mas
aplicável por padrão salvo indicação contrária do departamento):
stipend pago quinzenalmente (faixa definida anualmente por nível de
cargo); pacote inclui tuition waiver + seguro saúde; exige GPA ≥3.0 e
bom standing acadêmico; quarter-time = ~10h/semana, half-time = ~20h/
semana; internacionais devem também se registrar no Office for
International Students and Scholars. RESSALVA: departamentos podem
ter padrões adicionais mais altos, e não há confirmação de que o
Broad College oferece essa modalidade especificamente para o MSBA
(já registrado na rodada 1 como "Sim (stipend + tuition waiver)",
não Verificar, não alterado).

'Aluguel Privativo/mês USD' preenchido: $750 -- estimativa de quarto
privativo em moradia compartilhada em East Lansing, MI (faixa $500-999
encontrada, mediana ~$750), moderada confiança (fonte de anúncios de
aluguel, não específica de estudantes).

CUIDADO APLICADO (dados NÃO encontrados mesmo após tentativas
extensivas de contornar o bloqueio do domínio principal, permanecem
'Verificar'): 'Fees Adicionais USD', 'Mercado/mês', 'Seguro Saúde
(programa) USD' (a página de seguro internacional do MSU confirma que
o SHIP é obrigatório, mas não divulga o valor do prêmio -- só um fee
de chegada antecipada de $8/dia, não uma mensalidade), 'Seguro
Residencial/mês', 'Utilities/mês', 'Celular/mês', 'WiFi/mês',
'Transporte/mês', 'Outros Gastos/mês', 'Custo Vida Total (programa)',
'Ranking US News' (a página de rankings do Broad College retornou
vazia, mesmo problema de renderização), 'Taxa Empregabilidade %',
'Salário Médio Pós USD', 'Principais Empregadores', 'GPA Mínimo',
'GPA Médio Admitidos', 'GMAT Médio', 'Taxa Aceitação %', 'Deadline
Aplicação', 'Rounds Aplicação', 'Entrevista'.

Fontes principais (pesquisadas via WebFetch, ago/2026):
- https://broad.msu.edu/masters/business-data-science-analytics/
  (URL fornecida pelo usuário -- INACESSÍVEL, conteúdo vazio em
  múltiplas tentativas)
- https://grad.msu.edu/funding/assistantships (política geral de GA
  do MSU, aplicável como base)
- https://oiss.isp.msu.edu/health-wellness-and-safety/us-health-care-insurance/
  (confirma SHIP obrigatório para internacionais, sem valor de
  prêmio)
- https://www.roomster.com/rooms-for-rent/east-lansing-mi-usa
  (faixa de aluguel de quarto privativo em East Lansing)
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
    "Nome da Universidade": "Michigan State",
    "Estado": "MI",
    "Cidade": "East Lansing",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Processo TA/RA": (
        "Política geral da Graduate School do MSU (não confirmada "
        "como específica do Broad/MSBA): stipend pago quinzenalmente; "
        "pacote inclui tuition waiver + seguro saúde; exige GPA ≥3.0; "
        "quarter-time ~10h/semana, half-time ~20h/semana; "
        "internacionais devem se registrar no Office for "
        "International Students and Scholars"
    ),
    "Aluguel Privativo/mês USD": 750,  # estimativa, quarto privativo em East Lansing, MI (faixa $500-999)
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
        raise ValueError("Nenhuma linha encontrada para Michigan State (East Lansing, MI). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Michigan State, East Lansing, MI)")

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
