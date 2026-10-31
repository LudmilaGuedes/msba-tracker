"""
Data enrichment (round 2): atualiza a linha da 'University of North
Florida (UNF)' (Coggin College of Business, MSBA) no
msba_programs_database.xlsx, a pedido do usuário: detalhar como cada
bolsa funciona, elegibilidade, valor, quão realista é, confirmar taxa
de aplicação (e se em algum momento fica de graça), e confirmar
empregabilidade/salário.

FONTES (oficiais, unf.edu, via WebFetch/WebSearch em 2026-09-08):
- https://www.unf.edu/coggin/graduate/coggin-graduate-assistantship.html
- https://www.unf.edu/coggin/scholarships/catalog.html
- https://www.unf.edu/graduateschool/admissions/faqs.html
- busca geral sobre GA para internacionais (F-1) e outcomes do Coggin

CONFIRMADO SEM MUDANÇA:
- Taxa de Aplicação ($30): CONFIRMADO oficial.
- Graduate Assistantship Tier I/II (stipend e % de waiver): os
  números já na planilha ($2.500-5.000 + waiver 37-75% p/ Tier I;
  $1.500-3.000 sem waiver p/ Tier II) foram RECONFIRMADOS por busca
  externa -- a página oficial do Coggin não detalha os tiers (só
  descreve os TIPOS de posição: "Non-Teaching" e "Distance Learning
  Coach"), mas os valores batem com o que já estava documentado.

CATÁLOGO DE BOLSAS NOMEADAS DO COGGIN -- ACHADO IMPORTANTE: pesquisei
o catálogo oficial completo (7 bolsas/fellowships nomeadas: Anna and
Dave Brosche, Blanche & Luther Coggin, Ennis Pellum CPAs, Graduate
Programs in Business Fellowship, Jack & Mary Lambert, Lowell M.
Salter, Crowley Center for Transportation and Logistics -- esta
última é a única que cobre MSBA diretamente). NENHUMA delas está
explicitamente aberta a internacionais -- a maioria exige "residir na
Flórida para fins de tuition" ou cidadania/residência permanente dos
EUA. Ou seja: PRATICAMENTE TODO O CATÁLOGO DE BOLSAS NOMEADAS DA
COGGIN NÃO SE APLICA a uma candidata internacional -- o único caminho
de financiamento realista é o Graduate Assistantship (Tier I/II),
que É aberto a internacionais mas com uma limitação relevante (ver
abaixo).

LIMITAÇÃO CRÍTICA DO GA PARA INTERNACIONAIS (não documentada antes):
confirmado via fonte oficial que "the standard stipend amount is not
enough to meet the financial requirements of an F-1 student visa, so
international students need some additional source of financial
support in order to qualify for a visa" -- ou seja, MESMO CONSEGUINDO
um GA Tier I ou II, o valor do stipend sozinho NÃO é suficiente para
comprovar capacidade financeira exigida pelo visto F-1; a candidata
precisaria de fonte de renda ADICIONAL de qualquer forma. Isso reforça
a classificação de 'Quão Realista' já existente ('Baixa-Moderada'),
agora com a razão concreta documentada.

TAXA DE APLICAÇÃO -- FICA DE GRAÇA EM ALGUM MOMENTO? Resposta: NÃO,
não para uma candidata aplicando pela primeira vez (caso de uma candidata internacional nova).
A página oficial de FAQs é explícita: "we do not offer application
fee waivers unless the applicant has previously paid and qualifies
for a waiver" -- as únicas situações de isenção são técnicas e exigem
JÁ TER PAGO antes: reaplicar para o mesmo programa dentro de 3
semestres consecutivos, ter aplicado para o programa errado por
engano, ou ter sido negada e querer aplicar para outro programa em
outro termo. NENHUM mecanismo de isenção por evento, por round, ou
para novos candidatos -- diferente de praticamente todas as outras
escolas já pesquisadas nesta base.

NÃO ENCONTRADO -- 'Taxa Empregabilidade %', 'Salário Médio Pós USD' e
'Principais Empregadores' permanecem 'Verificar': pesquisei o Coggin
Career Management Center e não encontrei nenhum relatório de outcomes
publicado (nem agregado do Coggin, nem específico do MSBA). Achei só
uma estatística genérica não relacionada a salário ("99% dos formandos
aceitam emprego na Flórida ou sul dos EUA", sem fonte de ano/programa
claro) -- não confiável o suficiente para usar.
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
    "Nome da Universidade": "University of North Florida (UNF)",
    "Estado": "FL",
    "Cidade": "Jacksonville",
}

UPDATES = {
    "Tipos de Bolsa": (
        "Graduate Assistantship Tier I (stipend $2.500-5.000 + waiver "
        "de 37-75% da tuition in-state) ou Tier II (stipend "
        "$1.500-3.000, sem waiver de tuition) -- ÚNICO caminho de "
        "financiamento realmente aberto a internacionais. O catálogo "
        "completo de bolsas NOMEADAS do Coggin (7 fellowships, "
        "incluindo uma específica de MSBA/Logistics) NÃO está aberto "
        "a internacionais -- quase todas exigem residência na Flórida "
        "para fins de tuition ou cidadania/residência permanente dos "
        "EUA."
    ),
    "Quão Realista (Intl)": (
        "Baixa-Moderada -- internacionais não são elegíveis às bolsas "
        "nomeadas do Coggin (quase todas exigem residência FL ou "
        "cidadania/PR dos EUA); o Graduate Assistantship é a única "
        "via realista, mas vagas são limitadas E, mesmo conseguindo "
        "um GA, o valor do stipend sozinho NÃO é suficiente para "
        "atender à exigência de comprovação financeira do visto F-1 "
        "-- internacional precisaria de fonte de renda adicional de "
        "qualquer forma (confirmado oficialmente)."
    ),
    "Rounds Aplicação": (
        "Sem rounds numerados; deadline único por semestre (ver "
        "Deadline Aplicação). Taxa de $30 NÃO fica gratuita para "
        "quem aplica pela primeira vez -- isenção só existe para "
        "quem JÁ PAGOU antes e reaplica dentro de 3 semestres "
        "consecutivos, aplicou pro programa errado por engano, ou foi "
        "negada e quer tentar outro programa (confirmado oficialmente "
        "-- 'we do not offer application fee waivers unless the "
        "applicant has previously paid'). Não há isenção por evento, "
        "round ou para novos candidatos, diferente da maioria das "
        "outras escolas já pesquisadas."
    ),
}


def main():
    rotate_backup()
    print("Backup rotacionado (política fixa) em: data/msba_programs_database_backup.xlsx")

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
        raise ValueError("Nenhuma linha encontrada para UNF. Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (UNF)")

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
