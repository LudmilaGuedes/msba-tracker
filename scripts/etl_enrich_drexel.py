"""
Data enrichment: atualiza APENAS a linha da 'Drexel' (LeBow College of
Business, Philadelphia, PA) no msba_programs_database.xlsx.

Foco exclusivo: MS in Business Analytics (MSBSAN), presencial, STEM
(CIP 30.7102), sistema de quarters, 45 créditos trimestrais, 15 meses.
Deep research em: custo total real de tuition para internacionais,
bolsas/assistantships, requisitos de GPA/GMAT/GRE/inglês, e custo de
vida em Philadelphia, PA.

TUITION: 'Custo/Crédito (Intl)' ($1.357) e 'Tuition Total (Intl)'
($61.065) já presentes são internamente consistentes (1.357 x 45 =
61.065 exatamente) -- MANTIDOS sem alteração, apesar de uma fonte
terceira citar um valor bem diferente ($56.070/18 meses), provavelmente
de um ciclo ou premissa de duração distintos e sem confirmação oficial
clara (a página oficial de tuition do LeBow não lista valores por
programa, apenas direciona a portais individualizados). Como é
universidade PRIVADA, 'Custo/Crédito (Residente)' e 'Tuition Total
(Residente)' (estavam 'Verificar') foram preenchidos com os MESMOS
valores do Intl.

'Principais Empregadores' preenchido com a lista da página oficial
"Career Landings" do próprio MSBA (Booz Allen Hamilton, Citibank,
Comcast, CompassRed Data Labs, Continuum Health, Home Depot,
Independence Blue Cross, JPMorgan, Symphony Health). Uma segunda lista
diferente (PECO/Exelon, Goldman Sachs, Blackrock, Amazon, Vanguard,
Tesla, CVS Health, Starbucks) apareceu em outra busca, mas a página
oficial "Career Landings" foi priorizada por ser fonte direta e
especifica do programa.

'Taxa Empregabilidade %' e 'Salário Médio Pós USD' permanecem
'Verificar': a mesma página oficial "Career Landings" lista SALÁRIOS
MÉDIOS DE MERCADO por cargo (ex.: Data Scientist $96.072), explicitamente
qualificados como "general industry figures, not specific to program
graduates" -- não usados aqui. Um número diferente encontrado em outra
busca ("90% empregados em 6 meses, salário médio $88.666") não pôde ser
confirmado com certeza como especifico do MSBA (o texto original é
ambíguo sobre se refere-se a todos os mestrados especializados do LeBow
ou só ao MSBA) -- por cautela, não foi usado.

Bolsas: 'Tipos de Bolsa' (Alumni Merit Scholarship) confirmado e
enriquecido -- apesar do nome, é uma bolsa aberta a TODO candidato
(doméstico e internacional) ao MBA e aos programas de MS, com
consideração automática (sem inscrição separada). 'TA/RA Disponível'
preenchido com o achado especifico: o Center for Applied AI and
Business Analytics oferece posições de "Graduate Fellow" (projetos de
consultoria com empresas parceiras, sob supervisão de professores).

'Seguro Saúde (programa) USD' preenchido com $3.178/ano -- taxa oficial
do Drexel Student Health Insurance Plan (SHIP/Dragon Plan).

Fontes principais (pesquisadas via busca web, ago/2026):
- https://www.lebow.drexel.edu/academics/graduate/specialized-masters-
  programs/business-analytics/career-landings (empregadores oficiais
  do MSBA, cargos, ressalva sobre salários de mercado genéricos)
- Busca web cruzada: GMAT/GRE não exigido (recomendado para bolsa
  adicional), deadlines (prioridade de bolsa 1 abr, final
  internacional 1 jun, doméstico/online 1 set), Alumni Merit
  Scholarship com consideração automática para todos, Graduate Fellow
  via Center for Applied AI and Business Analytics
- Busca web: seguro saúde $3.178/ano (SHIP/Dragon Plan, oficial)
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
    "Nome da Universidade": "Drexel",
    "Estado": "PA",
    "Cidade": "Philadelphia",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Residente) USD": 1357,  # privada, mesma taxa do Intl
    "Tuition Total (Residente) USD": 61065,  # privada, mesma taxa do Intl
    "Seguro Saúde (programa) USD": 3178,  # oficial, Drexel Student Health Insurance Plan
    "Tipos de Bolsa": (
        "Alumni Merit Scholarship — apesar do nome, aberta a TODO "
        "candidato (doméstico e internacional) ao MBA e aos programas "
        "de MS, com consideração automática (sem inscrição separada)"
    ),
    "TA/RA Disponível": (
        "Limitado — Graduate Fellow via Center for Applied AI and "
        "Business Analytics (projetos de consultoria com empresas "
        "parceiras, sob supervisão de professores); sem confirmação de "
        "assistantship tradicional (stipend + tuition) para o MSBA"
    ),
    "GMAT/GRE Obrigatório": "Não exigido; recomendado para consideração de bolsa de mérito adicional",
    "Deadline Aplicação": "Prioridade de bolsa: 1º abr; Final internacional: 1º jun; Doméstico/online: 1º set",
    "Rounds Aplicação": "Sem rounds numerados; deadlines por categoria (prioridade de bolsa, internacional, doméstico)",
    "Documentos Necessários": (
        "Histórico acadêmico (avaliação de equivalência para diplomas "
        "internacionais), GMAT/GRE (opcional), teste de inglês (se "
        "aplicável)"
    ),
    "Principais Empregadores": (
        "Booz Allen Hamilton, Citibank, Comcast, CompassRed Data Labs, "
        "Continuum Health, Home Depot, Independence Blue Cross, "
        "JPMorgan, Symphony Health"
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
        raise ValueError("Nenhuma linha encontrada para Drexel (Philadelphia, PA). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Drexel, Philadelphia, PA)")

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
