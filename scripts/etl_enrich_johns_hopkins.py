"""
Data enrichment: atualiza APENAS a linha da 'Johns Hopkins' (Carey
Business School, Washington, DC) no msba_programs_database.xlsx.

Foco exclusivo: MS in Business Analytics and Artificial Intelligence
(MSBA&AI), formato full-time presencial no Hopkins Bloomberg Center
(555 Pennsylvania Ave NW, Washington DC — NÃO no campus principal de
Baltimore), STEM, 36 créditos, 12 meses. Deep research em: custo total
real de tuition para internacionais, bolsas/assistantships, requisitos
de GPA/GMAT/GRE/inglês, e custo de vida em Washington, DC.

NOTA IMPORTANTE: este é um programa MUITO NOVO — a ênfase em "Artificial
Intelligence" foi anunciada pela Carey Business School em set/2024, e o
programa full-time funciona no Hopkins Bloomberg Center, inaugurado em
2025. Isso explica por que não foi encontrada nenhuma fonte com
outcomes proprios do programa (taxa de emprego, ranking, lista de
empregadores especifica) — ele ainda não tem turmas formadas o
suficiente para gerar esses dados publicamente. Os poucos números
"gerais da Carey" encontrados (ex.: lista de empregadores de "alumni da
Johns Hopkins") NÃO foram usados aqui por não serem claramente
atribuíveis a este programa especifico — mesmo criterio de cautela ja
aplicado a outras universidades desta planilha quando uma fonte mistura
programas/turmas diferentes.

TUITION (já estava correto, apenas verificado e complementado): $91.500
para o formato full-time em DC (confirmado por múltiplas fontes,
inclusive página oficial). Como é universidade PRIVADA, 'Custo/Crédito'
e 'Tuition Total (Residente)' (estavam 'Verificar') PREENCHIDOS com os
mesmos valores do Intl: $91.500 / 36 créditos = $2.541,67/crédito.

Custo de vida: preenchidos os campos com fontes razoavelmente
convergentes para a região do campus (Penn Quarter/Chinatown, DC):
'Aluguel Privativo/mês' ($1.150, quarto privativo mobiliado perto do
campus/metro), 'Mercado/mês' ($320) e 'Transporte/mês' ($136, passe de
Metro) vieram de estimativas de orçamento estudantil para DC;
'Utilities/mês' ($175) é o ponto médio da faixa encontrada ($150-200).
'Seguro Saúde (programa)', 'Seguro Residencial/mês', 'Celular/mês',
'WiFi/mês' e 'Outros Gastos/mês' permanecem 'Verificar' — sem fonte
especifica encontrada — e por isso 'Custo Vida Total (programa) USD'
também permanece 'Verificar'.

Fontes principais (pesquisadas via busca web, ago/2026):
- Busca web cruzada (carey.jhu.edu, Poets&Quants, yocket): programa
  full-time em Washington DC, 36 créditos, test-optional, GPA médio
  ~3.53 e GMAT/GRE médio ~674 (perfil típico de candidatos aceitos),
  TOEFL 100/IELTS 7.0/PTE 70/Duolingo 130
- Busca web: taxa de aplicação $100, deadlines Fall 2026 (Round 1: 8
  out 2025, Round 2: 14 jan 2026, Round 3: 11 mar 2026), bolsas
  merit-based automáticas (incl. Dean's Scholarship) para candidatos
  full-time
- https://carey.jhu.edu/locations/Washington-dc (localização exata do
  campus: Hopkins Bloomberg Center, Penn Quarter/Chinatown)
- Busca web: estimativas de custo de vida em Washington DC (aluguel de
  quarto privativo perto do campus, mercado, utilities, transporte)

Páginas oficiais carey.jhu.edu retornaram HTTP 403 (bloqueio de bot) em
fetch direto — dados complementados via busca web e fontes terceiras
(Poets&Quants, yocket, etc.), sem contradição entre si.
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
    "Nome da Universidade": "Johns Hopkins",
    "Estado": "DC",
    "Cidade": "Washington",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Intl) USD": 2541.67,  # $91.500 / 36 créditos
    "Custo/Crédito (Residente) USD": 2541.67,  # privada, mesma taxa do Intl
    "Tuition Total (Residente) USD": 91500,  # privada, mesma taxa do Intl
    "Aluguel Privativo/mês USD": 1150,  # quarto privativo mobiliado, Penn Quarter/Chinatown (perto do campus)
    "Mercado/mês USD": 320,
    "Utilities/mês USD": 175,
    "Transporte/mês USD": 136,  # passe de Metro
    "Bolsas Disponíveis": "Sim",
    "Tipos de Bolsa": (
        "Consideração automática para bolsas merit-based (incl. Dean's "
        "Scholarship) para todos os candidatos full-time; bolsas "
        "endowed adicionais também disponíveis"
    ),
    "GPA Médio Admitidos": 3.53,  # perfil típico de candidatos aceitos (fonte: Poets&Quants)
    "GMAT/GRE Obrigatório": "Test-optional (GMAT/GRE não obrigatório para nenhum candidato)",
    "GMAT Médio": 674,  # perfil típico de candidatos aceitos, GMAT/GRE combinado (fonte: Poets&Quants)
    "Testes Inglês Aceitos": (
        "TOEFL iBT mín. 100 / IELTS mín. 7.0 / PTE mín. 70 / Duolingo "
        "mín. 130 — exigido para não-cidadãos dos EUA cuja graduação "
        "não foi em inglês"
    ),
    "Deadline Aplicação": "Fall 2026: Round 1: 8 out 2025; Round 2: 14 jan 2026; Round 3: 11 mar 2026",
    "Taxa Aplicação USD": 100,
    "Rounds Aplicação": "Round 1, Round 2, Round 3 (ver Deadline Aplicação)",
    "Documentos Necessários": "Histórico acadêmico (transcripts), resume, essays, cartas de recomendação",
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
        raise ValueError("Nenhuma linha encontrada para Johns Hopkins (Washington, DC). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Johns Hopkins, Washington, DC)")

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
