"""
Data enrichment: atualiza APENAS a linha da 'U of Michigan' (Ross
School of Business, Ann Arbor, MI) no msba_programs_database.xlsx.

Foco exclusivo: Master of Business Analytics (MBAn), presencial, STEM,
~36 créditos (35,75 exatos, conforme catálogo), 10 meses (3 termos
consecutivos: Summer, Fall, Winter). Deep research em: custo total real
de tuition para internacionais, bolsas/assistantships, requisitos de
GPA/GMAT/GRE/inglês, e custo de vida em Ann Arbor, MI.

NOTA IMPORTANTE: as páginas oficiais michiganross.umich.edu e
rossweb.bus.umich.edu retornaram HTTP 403 (bloqueio de bot) em todas as
tentativas de fetch direto. Os dados usados aqui vêm de busca web e
fontes terceiras. CUIDADO ESPECIAL foi tomado para NÃO usar dados do
programa de MBA full-time da Ross (um programa MUITO maior e mais caro,
com outcomes de emprego/salário bem diferentes) que aparecem
repetidamente misturados nos resultados de busca sobre "Michigan Ross"
-- nenhum dado de emprego, salário ou empregador do MBA foi usado nesta
linha do MBAn.

TUITION: 'Tuition Total (Intl)' já presente ($66.625) foi MANTIDO sem
alteração -- uma fonte terceira (tabela comparativa de tuition,
atualizada set/2025) cita $69.290 para não-residentes, um valor
diferente mas próximo (~4%); sem confirmação oficial direta (bloqueio
403) para decidir qual e mais atual, o valor ja presente na planilha
foi preservado por cautela. 'Tuition Total (Residente)' (estava
'Verificar') PREENCHIDO com $64.290 (mesma fonte comparativa). 'Custo/
Crédito' calculado dividindo por 36 créditos (já presente em 'Total
Créditos'). 'Fees Adicionais' preenchido com $1.853 = taxas
obrigatórias por termo ($120,60 Summer + $241,19 x 2 Fall/Winter =
$602,98) + International Student Services Fee ($500 x 2 Fall/Winter +
$250 Summer = $1.250, específica de vistos F/J).

Custo de vida: 'Transporte/mês' preenchido com $0 -- alunos da
University of Michigan viajam gratuitamente na AAATA (The Ride,
sistema de ônibus de Ann Arbor/Ypsilanti) usando o M-Card. Demais
campos mensais (Aluguel, Mercado, Utilities, Celular, WiFi) preenchidos
com estimativas de orçamento estudantil para Ann Arbor, MI. 'Seguro
Saúde (programa)', 'Seguro Residencial/mês' e 'Outros Gastos/mês'
permanecem 'Verificar' -- sem fonte especifica -- e por isso 'Custo
Vida Total (programa) USD' também permanece 'Verificar'.

'Taxa Empregabilidade %', 'Salário Médio Pós USD', 'Principais
Empregadores' e 'Ranking US News/QS' permanecem 'Verificar': existe um
relatório oficial de emprego especifico do MBAn (michiganross.umich.edu/
graduate/master-of-business-analytics/careers/employment-data), mas o
fetch direto foi bloqueado (403) e a busca web só retornou dados
PARCIAIS (ex.: bônus médio de assinatura $13.000, 23% dos candidatos
com bônus) sem uma taxa de emprego ou salário médio completo e
claramente atribuível ao MBAn -- preferiu-se não preencher com dados
incompletos ou arriscar misturar com o MBA.

Fontes principais (pesquisadas via busca web, ago/2026):
- Busca web: GRE/GMAT dispensado automaticamente com GPA de graduação
  >= 3.30; sem nota mínima fixa se enviado; TOEFL recomendado 100+/
  IELTS 7.0+/DET 130+ (sem mínimo obrigatório, dispensado com
  graduação integralmente em inglês)
- Busca web: deadlines Round 1 (12 set 2025), Round 2 (24 out 2025),
  Final internacional (5 jan 2026), Round 6/final doméstico (~15 mai
  2026, "One-Year Master's Programs")
- Busca web: tabela comparativa de tuition (set/2025) -- $64.290
  residente / $69.290 não-residente; fees obrigatórias por termo
  ($120,60 Summer + $241,19 Fall/Winter) e International Student
  Services Fee ($500/termo Fall-Winter, $250 Summer)
- Busca web: AAATA (The Ride) gratuito para alunos da U-M com M-Card;
  faixas de custo de vida em Ann Arbor, MI (aluguel, mercado,
  utilities, celular/wifi)
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
    "Nome da Universidade": "U of Michigan",
    "Estado": "MI",
    "Cidade": "Ann Arbor",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Intl) USD": 1850.69,  # $66.625 (já presente, mantido) / 36 créditos
    "Custo/Crédito (Residente) USD": 1786.67,  # $64.290 / 36 créditos
    "Tuition Total (Residente) USD": 64290,  # fonte comparativa terceira, set/2025
    "Fees Adicionais USD": 1853,  # taxas obrigatórias por termo ($603) + International Student Services Fee ($1.250)
    "Aluguel Privativo/mês USD": 900,  # quarto privativo, Ann Arbor MI (faixa de mercado encontrada)
    "Mercado/mês USD": 350,
    "Utilities/mês USD": 75,
    "Celular/mês USD": 50,
    "WiFi/mês USD": 60,
    "Transporte/mês USD": 0,  # AAATA (The Ride) gratuito para alunos U-M com M-Card
    "GMAT/GRE Obrigatório": (
        "GMAT ou GRE aceito, sem nota mínima fixa; dispensado "
        "automaticamente com GPA de graduação ≥ 3.30"
    ),
    "Testes Inglês Aceitos": (
        "TOEFL recomendado 100+ / IELTS recomendado 7.0+ / DET "
        "recomendado 130+ — sem nota mínima obrigatória; dispensado "
        "com graduação integralmente em instituição de língua inglesa"
    ),
    "Deadline Aplicação": (
        "Round 1: 12 set 2025; Round 2: 24 out 2025; Final "
        "internacional: 5 jan 2026; Final doméstico (Round 6): ~15 mai "
        "2026"
    ),
    "Rounds Aplicação": "Múltiplos rounds (Round 1 a Round 6); internacionais têm deadline final mais cedo (5 jan) que domésticos",
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
        raise ValueError("Nenhuma linha encontrada para U of Michigan (Ann Arbor, MI). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (U of Michigan, Ann Arbor, MI)")

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
