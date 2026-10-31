"""
Data enrichment: atualiza APENAS a linha da 'Fordham' (Gabelli School of
Business, New York, NY) no msba_programs_database.xlsx.

Foco exclusivo: MS in Business Analytics, presencial, STEM, 30 créditos
(10 disciplinas: 5 obrigatórias + 5 eletivas), 12 meses tempo integral
(3 semestres). Deep research em: custo total real de tuition para
internacionais, bolsas/assistantships, requisitos de GPA/GMAT/GRE/
inglês, e custo de vida em Nova York.

TUITION (já estava correto, apenas complementado): $66.030/ano ÷ 30
créditos = $2.201/crédito, já presente em 'Custo/Crédito (Intl)' e
'Tuition Total (Intl)'. Como é universidade PRIVADA sem distinção de
residência, PREENCHIDOS os campos '(Residente)' equivalentes com os
MESMOS valores (mesmo critério já usado para outras privadas desta
planilha, ex.: University of Dallas, Seattle U).

'Seguro Saúde (programa) USD' permanece 'Verificar': o único valor
encontrado ($2.567/ano) é explicitamente rotulado como do ciclo
2017-2018 pela própria fonte -- quase 8 anos desatualizado, num mercado
(NYC) com inflação de custos de saúde muito acima da média nacional.
Diferente de outros casos deste projeto onde um dado antigo mas plausível
foi usado com ressalva, aqui o gap temporal e a volatilidade do mercado
tornam o número pouco confiável o suficiente para preferir 'Verificar' a
um valor provavelmente muito abaixo do real atual.

'Ranking US News' e 'Ranking QS' permanecem 'Verificar': os únicos
números encontrados (US News #68, QS #62) vieram de uma fonte terceira
que os apresenta ao lado da citação explícita "QS Top MBA ranking" --
ou seja, parecem ser rankings do programa de MBA da Gabelli, não do
MSBA, e foram descartados para não atribuir um dado do programa errado
a esta linha (mesmo critério de cautela já aplicado a outras
universidades nesta planilha quando uma fonte mistura programas).

Fontes principais (pesquisadas via busca web, ago/2026):
- https://bulletin.fordham.edu/gabelli-graduate/ms/business-analytics/
  (estrutura curricular: 5 obrigatórias + 5 eletivas, 4 trilhas de
  especialização — ML/AI, Advanced Analytics, Analytics Consulting, IT
  Ethics & Cybersecurity)
- Busca web cruzada (múltiplas fontes incl. agregadores): GMAT/GRE
  opcional, sem GPA mínimo formal, 2 cartas de recomendação, statement
  of purpose 500-750 palavras, taxa de aplicação $130 (isenta no Round
  1), deadlines Round 2 16 jan / Round 3 20 mar (Fall), Summer intake 6
  mar; bolsas merit $5k-tuition integral (69% dos alunos recebem, já
  presente na planilha)
- Busca web: taxa de emprego 80% em 6 meses (já presente na planilha,
  confirmada), salário médio $90.000, principais empregadores (Google,
  Amazon, JPMorgan Chase, Deloitte, EY, PwC, IBM, Bloomberg)
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
    "Nome da Universidade": "Fordham",
    "Estado": "NY",
    "Cidade": "New York",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Residente) USD": 2201,  # privada, mesma taxa do Intl (sem distinção de residência)
    "Tuition Total (Residente) USD": 66030,  # privada, mesma taxa do Intl
    "GPA Mínimo": "Nenhum mínimo formal (revisão holística, caso a caso)",
    "Matérias Pré-Requisito": "Nenhum pré-requisito formal específico; aptidão quantitativa/analítica forte é esperada",
    "Deadline Aplicação": (
        "Fall Round 1 (taxa isenta, data rolling); Round 2: 16 jan; "
        "Round 3: 20 mar; Summer intake: 6 mar"
    ),
    "Taxa Aplicação USD": 130,  # isenta para candidaturas no Round 1
    "Rounds Aplicação": "Round 1 (taxa isenta), Round 2 (16 jan), Round 3 (20 mar) para Fall; intake de Summer com deadline 6 mar",
    "Documentos Necessários": (
        "Histórico acadêmico oficial de todas as instituições, 2 "
        "cartas de recomendação (acadêmica/profissional), statement of "
        "purpose (500–750 palavras), GMAT/GRE (opcional)"
    ),
    "Salário Médio Pós USD": 90000,
    "Principais Empregadores": "Google, Amazon, JPMorgan Chase, Deloitte, EY, PwC, IBM, Bloomberg",
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
        raise ValueError("Nenhuma linha encontrada para Fordham (New York, NY). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Fordham, New York, NY)")

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
