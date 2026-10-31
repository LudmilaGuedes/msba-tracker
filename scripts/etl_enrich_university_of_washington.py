"""
Data enrichment: atualiza APENAS a linha da 'University of Washington'
(Michael G. Foster School of Business, Seattle, WA) no
msba_programs_database.xlsx.

Foco exclusivo: MS in Business Analytics (MSBA), presencial (evening &
weekend, carga full-time), STEM, 46 créditos, 12 meses. Deep research em:
custo total real de tuition para internacionais, bolsas/assistantships,
requisitos de GPA/GMAT/GRE/inglês, e custo de vida em Seattle, WA.

TUITION: a página oficial de "Program Fees & Financing" (foster.uw.edu,
ciclo 2026-27) fornece o billing trimestral completo: Residente $12.043
tuition + $403 fees/trimestre (verão $382); Não-residente $15.341
tuition + mesmas fees. Isso permite DECOMPOR o total anunciado
($49.763 residente / $62.955 não-residente) em tuition pura vs. fees
obrigatórias: tuition pura = $48.172 (residente) / $61.364 (não-
residente); fees obrigatórias = $1.591/ano (iguais para todos).
PREENCHIDO 'Custo/Crédito' dividindo a tuition pura por 46 créditos
($1.047 residente / $1.334 não-residente/intl) e 'Fees Adicionais' com
os $1.591 de taxas trimestrais somadas.

Custo de vida: 'Seguro Saúde (programa)' preenchido com $1.965 (ISHIP --
International Student Insurance Health Plan --, estimativa oficial anual
do próprio Foster). 'Transporte/mês' ($132) e 'Outros Gastos/mês' ($209)
vieram da estimativa oficial de custo de vida da mesma página (Local
transportation $1.581/ano e Personal expenses $2.508/ano, divididos por
12). 'Aluguel Privativo/mês' preenchido com $900 (quarto privativo
mobiliado, University District, perto do campus -- múltiplos anúncios
convergentes). 'Mercado/mês' e 'Utilities/mês' permanecem 'Verificar':
a página oficial só fornece um valor combinado de "Room and board"
($23.199/ano) sem separar aluguel de comida -- como o aluguel já foi
obtido de fonte de mercado independente, subtrair não seria confiável
(o "board" oficial provavelmente reflete um orçamento mais generoso que
o de um aluno freelance cozinhando em casa). Por faltarem esses
componentes, 'Custo Vida Total (programa) USD' também permanece
'Verificar'.

'TA/RA Disponível' preenchido com nota de que o MSBA é descrito pela
própria UW como um "fee-based graduate program" (financiado
majoritariamente por mensalidade dos alunos) -- nenhuma menção a
assistantships nas fontes oficiais, ao contrário de programas de
pesquisa tradicionais.

Fontes principais (pesquisadas via busca web, ago/2026):
- https://foster.uw.edu/academics/degree-programs/master-science-
  business-analytics/program-fees-financing-msba/ (tuition oficial 2026-
  27, billing trimestral, ISHIP, estimativas de custo de vida)
- https://foster.washington.edu/academics/degree-programs/master-
  science-business-analytics/admissions-msba (deadlines Round 1 20 out/
  Round 2 5 jan, taxa de aplicação $90, documentos, TOEFL/IELTS/
  Duolingo, GPA mínimo, pré-requisitos, entrevista em vídeo)
- Busca web: taxa de emprego 91% em 3 meses e salário médio $140.354
  para a turma mais recente do MSBA (distinto dos números do MBA,
  não usados aqui)
- Busca web: aluguel de quarto privativo mobiliado no University
  District, Seattle WA
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
    "Nome da Universidade": "University of Washington",
    "Estado": "WA",
    "Cidade": "Seattle",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Intl) USD": 1334,  # tuition pura (sem fees) / 46 créditos: $61.364 / 46
    "Custo/Crédito (Residente) USD": 1047,  # tuition pura (sem fees) / 46 créditos: $48.172 / 46
    "Tuition Total (Intl) USD": 61364,  # oficial 2026-27 ($62.955 total - $1.591 de fees obrigatórias)
    "Tuition Total (Residente) USD": 48172,  # oficial 2026-27 ($49.763 total - $1.591 de fees obrigatórias)
    "Fees Adicionais USD": 1591,  # taxas trimestrais obrigatórias somadas (iguais para residente e não-residente)
    "Aluguel Privativo/mês USD": 900,  # quarto privativo mobiliado, University District (perto do campus)
    "Seguro Saúde (programa) USD": 1965,  # ISHIP, estimativa oficial anual (Foster/UW)
    "Transporte/mês USD": 132,  # estimativa oficial de custo de vida: $1.581/ano
    "Outros Gastos/mês USD": 209,  # estimativa oficial de custo de vida, "Personal expenses": $2.508/ano
    "TA/RA Disponível": (
        "Improvável — o MSBA é descrito pela própria UW como um "
        "'fee-based graduate program' (financiado majoritariamente por "
        "mensalidade dos alunos); nenhuma menção a assistantships nas "
        "fontes oficiais"
    ),
    "Bolsas Disponíveis": "Sim (pool limitado)",
    "Tipos de Bolsa": (
        "Pool limitado de bolsas do próprio programa, consideradas "
        "automaticamente na revisão da candidatura; maioria dos alunos "
        "financia via fundos próprios, apoio do empregador ou "
        "empréstimos; bolsas privadas externas (ex.: FastWeb, "
        "Scholarships.com) recomendadas como complemento"
    ),
    "Quão Realista (Intl)": "Baixa — pool de bolsas do programa é limitado; a maioria dos alunos financia via fundos próprios/empréstimos",
    "Testes Inglês Aceitos": (
        "TOEFL iBT mín. 92 (antes de 21/jan/2026) ou 4.5 (nova escala) "
        "/ IELTS mín. 7.0 (7.5 recomendado para internacionais) / "
        "Duolingo mín. 120–125"
    ),
    "Taxa Aplicação USD": 90,
    "Rounds Aplicação": "Round 1: 20 out; Round 2: 5 jan; início único em junho",
    "Deadline Aplicação": "Round 1: 20 out; Round 2: 5 jan",
    "Matérias Pré-Requisito": (
        "Nenhum pré-requisito formal específico; status 'test-optional' "
        "(dispensa de GRE/GMAT) exige 2+ anos de experiência "
        "profissional OU graduação (undergrad/grad) em área STEM com "
        "GPA ≥ 3.3"
    ),
    "Experiência Profissional": "Não obrigatória para admissão, mas necessária (2+ anos) para dispensa de GRE/GMAT, salvo graduação STEM com GPA ≥ 3.3",
    "Taxa Empregabilidade %": 91,  # dentro de 3 meses pós-formatura, turma mais recente do MSBA
    "Salário Médio Pós USD": 140354,  # turma mais recente do MSBA (distinto do MBA)
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
        raise ValueError("Nenhuma linha encontrada para University of Washington (Seattle, WA). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (University of Washington, Seattle, WA)")

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
