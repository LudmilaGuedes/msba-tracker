"""
Data enrichment: atualiza APENAS a linha da 'Miami U Ohio' (Farmer
School of Business, Oxford, OH) no msba_programs_database.xlsx.

ATENCAO: "Miami U Ohio" (publica, Oxford, OH) e uma instituicao
DIFERENTE de "University of Miami" (privada, Coral Gables, FL, ja
cadastrada em outra linha desta planilha) -- varias buscas trazem
resultados misturados dos dois nomes; todos os dados usados aqui foram
verificados como vindos especificamente de paginas oficiais miamioh.edu
ou de buscas que citam "Miami University" em Oxford/Ohio explicitamente.

Foco exclusivo: MS in Business Analytics (MSBA), presencial, STEM/OPT,
30 creditos, 9-10 meses (tempo integral, um ano academico). Deep
research em: custo total real de tuition para internacionais,
bolsas/assistantships, requisitos de GPA/GMAT/GRE/ingles, e custo de
vida em Oxford, OH.

TUITION: 'Custo/Credito (Intl)' ($995) e 'Tuition Total (Intl)'
($29.850) ja presentes foram CONFIRMADOS por duas buscas independentes.
Achado importante: a tarifa geral de pos-graduacao de Miami University
(publica) DIFERE por residencia ($598/credito residente vs $1.475/
credito nao-residente) -- mas o MSBA e um dos "programas profissionais"
que usa uma tarifa PROPRIA e FLAT (igual para todos, independente de
residencia), confirmado por fonte que descreve esse grupo de programas
explicitamente. Por isso, PREENCHIDOS 'Custo/Credito (Residente)' e
'Tuition Total (Residente)' (estavam 'Verificar') com os MESMOS valores
do Intl, e 'Caminho Tuition Residente' documentando esse regime flat
especifico do MSBA (distinto da tarifa geral da universidade).

ACHADO CRITICO: a pagina oficial do programa confirma "Scholarships and
assistantships are not available for this program" -- SEM bolsas nem
assistantships para o MSBA (mesmo padrao ja encontrado em Binghamton
SUNY e UMass Amherst nesta planilha). CORRIGIDO/PREENCHIDO 'Bolsas
Disponiveis', 'Tipos de Bolsa', 'Quao Realista (Intl)', 'TA/RA
Disponivel', 'Processo TA/RA' e 'Graduate Assistantships' refletindo
essa ausencia confirmada oficialmente.

'Principais Empregadores' preenchido com a lista ESPECIFICA de
empregadores de egressos recentes do MSBA (nao generica da
universidade): Bain & Company, Grange Insurance, ServiceNow, EY, P&G,
Polen Capital, Empire Marketing Strategies, GW, Warner Bros. Discovery,
Terillium, A&F, MarketVision Research, Baker Tilly, Progressive
Insurance.

Custo de vida: 'Custo Vida Total (programa) USD' preenchido com
$21.730 -- estimativa PRORATEADA a partir do "Estimated Living
Expenses" oficial anual (calendario cheio, 12 meses) de $26.076 para
alunos internacionais de pos-graduacao em Oxford ($26.076 x 10/12).
Essa e uma estimativa GERAL da universidade (nao decomposta em
aluguel/mercado/etc. pela fonte, nem especifica do MSBA), usada aqui
apenas para o total agregado -- os campos MENSAIS individuais
permanecem 'Verificar' porque a fonte nao os decompoe com confianca
suficiente. 'Aluguel Privativo/mes' preenchido separadamente com $600
(quarto privativo, faixa "economica" de mercado em Oxford, OH).

Fontes principais (pesquisadas via busca web, ago/2026):
- Busca web cruzada (yocket, ameerkhatri.com): tuition $995/credito
  flat confirmado 2x; GPA sem minimo formal declarado (so para
  pre-requisitos: C minimo + GPA>=3.0 nesses cursos); GMAT/GRE
  opcional; deadline de prioridade 1 dez, revisao continua apos;
  "Scholarships and assistantships are not available for this
  program"; empregadores especificos do MSBA
- https://miamioh.edu/global-initiatives/international-student-
  scholar-services/resources/money-matters/tutition-fees/graduate.html
  (COA oficial internacional, confirma tarifa geral $598/$1.475 por
  credito -- DIFERENTE da tarifa flat do MSBA -- e estimativa de
  gastos de vida anual $26.076)
- Busca web: aluguel de quarto privativo em Oxford, OH
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
    "Nome da Universidade": "Miami U Ohio",
    "Estado": "OH",
    "Cidade": "Oxford",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Residente) USD": 995,  # confirmado: MSBA usa tarifa flat própria, igual à do Intl
    "Tuition Total (Residente) USD": 29850,  # flat, mesma taxa do Intl
    "Caminho Tuition Residente": (
        "Flat — o MSBA usa uma tarifa própria por crédito ($995), "
        "igual para residentes e não-residentes, DIFERENTE da tarifa "
        "geral de pós-graduação da universidade ($598 residente/"
        "$1.475 não-residente, aplicável a outros programas)"
    ),
    "Bolsas Disponíveis": "Não — confirmado oficialmente que bolsas e assistantships não estão disponíveis para este programa",
    "Tipos de Bolsa": "N/A",
    "Quão Realista (Intl)": "Baixa — bolsas e assistantships não disponíveis para este programa (confirmado oficialmente)",
    "TA/RA Disponível": "Não — confirmado oficialmente que assistantships não estão disponíveis para este programa",
    "Processo TA/RA": "N/A — sem assistantships neste programa",
    "Graduate Assistantships": "Não disponíveis (confirmado oficialmente)",
    "Aluguel Privativo/mês USD": 600,  # quarto privativo, faixa econômica de mercado em Oxford, OH
    "Custo Vida Total (programa) USD": 21730,  # estimativa oficial de gastos de vida anual ($26.076) prorateada para 10 meses
    "GMAT/GRE Obrigatório": "Opcional (não exigido)",
    "Matérias Pré-Requisito": "Pré-requisitos específicos (se aplicável) exigem nota mínima C em cada curso e GPA ≥ 3.0 nesses cursos",
    "Rounds Aplicação": "Prioridade 1º dez; revisão contínua (rolling) após essa data; decisão em 2-4 semanas",
    "Documentos Necessários": (
        "Histórico acadêmico (transcript), resume/CV, personal "
        "statement, 1 carta de recomendação (dispensada para "
        "graduados do Miami PRIME ou do Farmer School of Business)"
    ),
    "Principais Empregadores": (
        "Bain & Company, Grange Insurance, ServiceNow, EY, P&G, Polen "
        "Capital, Empire Marketing Strategies, GW, Warner Bros. "
        "Discovery, Terillium, A&F, MarketVision Research, Baker "
        "Tilly, Progressive Insurance"
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
        raise ValueError("Nenhuma linha encontrada para Miami U Ohio (Oxford, OH). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Miami U Ohio, Oxford, OH)")

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
