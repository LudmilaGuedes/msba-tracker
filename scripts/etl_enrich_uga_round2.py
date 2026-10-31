"""
Data enrichment (round 2): atualiza APENAS a linha da 'UGA' (Terry College
of Business, Athens, GA) no msba_programs_database.xlsx.

Foco exclusivo: MS in Business Analytics (MSBA, marca "Georgia Analytics"),
presencial, STEM, 10 meses (já confirmado no round 1). Deep research em:
custo de tuition p/ internacionais (fees adicionais, novo), bolsas/
assistantships (mecanismo detalhado), GPA/GMAT/GRE/inglês (sem novidade),
custo de vida (sem novidade) -- a pedido da usuaria, mesmo processo das
rodadas anteriores desta sessão.

BLOQUEIO PERSISTENTE: terry.uga.edu continua bloqueando fetch automatizado
(HTTP 403) em TODAS as URLs tentadas nesta rodada (class-profile,
employment-stats, MSBA guidebook PDF, brochure PDF) -- mesmo problema já
relatado no round 1. Toda a pesquisa desta rodada veio de snippets de
busca (Google) que indexam essas páginas, não de fetch direto.

FEES ADICIONAIS (novo, antes "Verificar"):
https://busfin.uga.edu/.../bursar-mandatory-fees-26-27/ -- taxas
obrigatórias 2026-27 p/ campus Athens, categoria "Athens - MD" (todos os
graduate/degree-seeking): Activity $116 + Athletic $73 + ConnectUGA $40 +
Green $4 + Health $290 + Recreation $24 + Student Center $100 +
Technology $152 + Transportation $173 = $972/semestre. Programa dura 2
semestres (Fall+Spring) -> $972 x 2 = $1.944.

TAXA DE APLICACAO (novo, antes "Verificar"):
https://grad.uga.edu/admissions/application-fee -- taxa padrão da UGA
Graduate School (MSBA se aplica através dela): $75 primeira aplicação /
$25 cada aplicação adicional. Usado $75 (primeira aplicação, caso comum).

TAXA DE EMPREGABILIDADE (novo, antes "Verificar"):
Confirmado que "Georgia Analytics" É a marca do MSBA da Terry (a página do
LinkedIn "linkedin.com/school/georgiaanalytics/" se identifica como
"M.S. Business Analytics - University of Georgia Terry..."). Múltiplos
snippets de busca citam: "100% of Georgia Analytics graduates received an
offer within 1 month post graduation (as of the past two years)".
Preenchido 100 (campo é tipo "percent" no schema -- valor numérico puro,
sem texto). IMPORTANTE PARA LEITURA FUTURA: a definição da fonte é
"recebeu oferta em até 1 mês pós-formatura", não necessariamente
"empregado em X meses" (metodologia mais comum em outras linhas do
dataset) -- essa diferença de metodologia fica registrada aqui no
script, já que a célula em si não comporta texto explicativo.

ASSISTANTSHIPS -- RESOLUCAO PARCIAL DA AMBIGUIDADE DO ROUND 1:
o round 1 tinha descartado o dado "até 90% da turma de MSBA de DOIS ANOS
recebe GA" por contradizer a duração de 10 meses já confirmada. Nesta
rodada, a MESMA frase ("up to 90 percent of the incoming two-year class...
tuition reduced to $25 per semester... 50% reduction for the mandatory
student health insurance premium") apareceu de forma consistente e
repetida em MÚLTIPLAS buscas diretamente associada a paginas
especificamente do MSBA (terry.uga.edu/msba/admissions/,
msba/admissions/class-profile/, MSBA guidebook) -- inclusive uma sintese
de busca concluiu explicitamente que a frase "relates to... particularly
the MBA and MSBA programs". Interpretacao mais provavel: e um paragrafo
de financial aid PADRONIZADO/reaproveitado entre os programas de mestrado
da Terry (MBA sim e 2 anos), copiado tambem para a pagina do MSBA sem
ajustar "two-year" -- mas os NUMEROS (90%, $25/semestre, desconto de 50%
no seguro) parecem se aplicar tambem ao MSBA dado onde a frase foi
encontrada. Usado com RESSALVA EXPLICITA na celula (nao removida a
ambiguidade, apenas documentada).

NAO ENCONTRADO (mantido "Verificar", mesmo apos buscas adicionais):
Tamanho Turma, Ranking US News, Ranking QS, Salário Médio Pós USD,
Principais Empregadores, GPA Médio Admitidos, GMAT Médio, Taxa Aceitação
%, Entrevista. Rankings/GMAT/GPA/taxa de aceitação/salário QUE FORAM
encontrados nas buscas sao explicitamente do programa de MBA da Terry
(GMAT medio 641, GPA 3.55, aceitacao 41%, salario $120-126k) -- NAO
usados nesta linha, mesmo criterio ja aplicado no round 1 (nao contaminar
MSBA com dados de outro programa).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "UGA",
    "Estado": "GA",
    "Cidade": "Athens",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Fees Adicionais USD": 1944,  # taxas obrigatórias UGA 2026-27: $972/semestre x 2 semestres
    "Taxa Aplicação USD": 75,  # taxa padrão UGA Graduate School (primeira aplicação)
    "Taxa Empregabilidade %": 100,  # "Georgia Analytics" (MSBA): "100% received an offer within 1 month post graduation" -- definição da fonte difere de "empregado em X meses" usado em outras linhas; ver docstring
    "Processo TA/RA": (
        "Todos os admitidos são automaticamente considerados (sem "
        "candidatura separada); até ~90% da turma recebe Graduate "
        "Assistantship, que reduz a tuition a $25/semestre + stipend + "
        "desconto de 50% no seguro saúde obrigatório. RESSALVA: a fonte "
        "usa a expressão 'incoming two-year class', possivelmente "
        "boilerplate reaproveitado da página do MBA (que dura 2 anos) -- "
        "o MSBA dura 10 meses (já confirmado) -- mas a mesma cifra "
        "aparece de forma consistente em múltiplas páginas oficiais "
        "específicas do MSBA"
    ),
    "Graduate Assistantships": (
        "GA reduz tuition a $25/semestre (efetivamente full-tuition) + "
        "stipend + desconto de 50% no seguro saúde obrigatório; até ~90% "
        "da turma recebe (ver ressalva sobre 'two-year' em 'Processo "
        "TA/RA')"
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
        raise ValueError("Nenhuma linha encontrada para UGA (Athens, GA). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (UGA, Athens, GA)")

    no_fill = PatternFill(fill_type=None)

    before = {}
    after = {}
    for header, new_value in UPDATES.items():
        col_idx = header_to_col[header]
        cell = ws.cell(row=row_idx, column=col_idx)
        before[header] = cell.value
        cell.value = new_value
        cell.fill = no_fill
        after[header] = new_value

    wb.save(DB_PATH)

    print("\nCampos atualizados (Como Estava -> Como Ficou):")
    for header in UPDATES:
        print(f"  - {header}:")
        print(f"      Antes: {before[header]!r}")
        print(f"      Depois: {after[header]!r}")


if __name__ == "__main__":
    main()
