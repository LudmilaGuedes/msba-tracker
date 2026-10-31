"""
Data enrichment (RODADA 2): atualiza APENAS a linha 'Dakota State'
(College of Business & IS, Madison, SD) no msba_programs_database.xlsx.

Esta linha já passou por uma rodada de enrichment (script
etl_enrich_dakota_state.py), restando ~23 campos 'Verificar'. Usuário
confirmou via AskUserQuestion (ago/2026) para prosseguir com uma
Rodada 2 focada nesses campos remanescentes.

'Custo/Crédito (Residente) USD' e 'Tuition Total (Residente) USD'
PREENCHIDOS: a página oficial de tuition da DSU (2026-27) confirma a
tarifa por crédito para masters ON-CAMPUS: $393,70/crédito (residente
SD) / $698,10/crédito (não-residente/internacional) -- essa última
bate EXATAMENTE com o valor já presente na planilha ($698), validando
cruzadamente que a rodada 1 já usou a fonte correta. 'Tuition Total
(Residente)' = 30 × $393,70 = $11.811,00 (também confere com o total
Intl já presente: 30 × $698,10 = $20.943, idêntico ao valor existente).

'GMAT Médio' preenchido com a única referência oficial disponível
(não é GMAT, é GRE, deixado claro no valor do campo): GRE NÃO é
exigido para admissão geral, mas é exigido especificamente para quem
solicita graduate assistantship (GA), com score preferencial de ~320
ou superior -- não é uma média de admitidos, é um patamar de
referência para candidatos a GA.

'Caminho Tuition Residente' preenchido com nota geral sobre residência
em SD (mesmo padrão de outros estados públicos já tratados nesta
sessão): estabelecer residência para fins de tuition normalmente exige
domicílio prévio, o que é inviável para uma aluna internacional em
F-1 -- tarifa não-residente/internacional é a relevante na prática.

CUIDADO APLICADO: mesmo após tentativas adicionais nesta rodada 2
(incluindo busca pela página de seguro saúde da DSU, que retornou erro
404), os seguintes campos permanecem 'Verificar' por falta de fonte
acessível: 'Tamanho Turma', 'Fees Adicionais USD' (a página de tuition
só menciona genericamente que "additional credit fees may apply" sem
valores), 'Seguro Saúde (programa) USD', todos os campos mensais de
custo de vida restantes (Mercado, Seguro Residencial, Utilities,
Celular, WiFi, Transporte, Outros Gastos, Custo Vida Total),
'Processo TA/RA', 'Ranking US News', 'Ranking QS', 'Salário Médio Pós
USD' (mesma conclusão da rodada 1: só existe estatística ocupacional
genérica de "computer systems analysts" em SD, não específica dos
formandos do MSAA, não usada), 'Principais Empregadores', 'GPA Médio
Admitidos', 'Taxa Aceitação %', 'Rounds Aplicação' (o programa usa
deadlines fixos por semestre/status, não rounds numerados), e
'Entrevista'.

Fontes principais (pesquisadas via WebFetch, ago/2026):
- https://dsu.edu/admissions/graduate/cost-aid/tuition-fees.html
  (tuition oficial 2026-27 por crédito, residente e não-residente,
  masters on-campus — validação cruzada com dados da rodada 1)
- https://dsu.edu/programs/msaa/msaa-timelines-requirements.html
  (confirmação de que GRE não é exigido, exceto para GA, com score
  preferencial ~320+)
- https://dsu.edu/admissions/international/index.html (sem dados de
  seguro saúde ou custo de vida disponíveis no conteúdo consultado)
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
    "Nome da Universidade": "Dakota State",
    "Estado": "SD",
    "Cidade": "Madison",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Residente) USD": 393.70,  # oficial 2026-27, masters on-campus
    "Tuition Total (Residente) USD": 11811.00,  # 30 x $393,70
    "GMAT Médio": "N/A (GRE, não GMAT): não exigido para admissão geral; exigido só para candidatos a GA, score preferencial ~320+",
    "Caminho Tuition Residente": "Não aplicável na prática para internacionais — tarifa não-residente/internacional é a relevante",
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
        raise ValueError("Nenhuma linha encontrada para Dakota State (Madison, SD). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Dakota State, Madison, SD)")

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
