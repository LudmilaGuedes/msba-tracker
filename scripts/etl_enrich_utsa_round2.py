"""
Data enrichment (RODADA 2): atualiza APENAS a linha 'UTSA' (Carlos
Alvarez College of Business, San Antonio, TX) no
msba_programs_database.xlsx.

Esta linha já passou por uma rodada de enrichment (script
etl_enrich_utsa.py), que deixou a tuition 'Verificar' por cautela
(a única tarifa encontrada na rodada 1 era claramente do formato
100% online). Usuário confirmou via AskUserQuestion (ago/2026) para
prosseguir com uma Rodada 2, priorizando tuition.

TUITION: PERMANECE 'Verificar' mesmo após pesquisa adicional --
múltiplas tentativas de localizar a tabela de tuition por crédito
específica da Alvarez College of Business (via future.utsa.edu,
onestop.utsa.edu, e páginas gerais de tuition) retornaram 404 ou
conteúdo sem os números. O catálogo confirma que existe uma
"Differential Tuition" aplicada a cursos de Management Science (MS),
sugerindo uma tarifa diferenciada para o MSBA, mas sem o valor exato
divulgado nas páginas acessíveis nesta rodada.

VALIDAÇÃO CRUZADA: a página oficial correta do programa (localizada
nesta rodada em future.utsa.edu, distinta e mais completa que a fonte
usada na rodada 1) CONFIRMA exatamente os valores já presentes: GPA
mínimo 3.0, GMAT/GRE não exigido, TOEFL 79/IELTS 6.5/Duolingo 100 --
todos batem com o que já estava na planilha, sem necessidade de
correção.

'Deadline Aplicação' CORRIGIDO: o valor anterior ("Fall: 1 mai;
Spring: 1 nov") divergia significativamente da página oficial
específica do programa, que lista para o ciclo 2027-28: prioridade
internacional 1 jun; doméstico 1 ago. Como a nova fonte é
explicitamente a página de admissões do próprio programa (mais
específica e atual que a fonte da rodada 1, cuja origem exata não
estava documentada), o valor foi CORRIGIDO.

'Documentos Necessários' preenchido: transcripts oficiais de todas as
instituições; purpose statement; resume; avaliação de credencial (para
educação internacional); cartas de recomendação opcionais (2
recomendadas).

'Matérias Pré-Requisito' preenchido: nenhum pré-requisito formal além
do bacharelado; preparo quantitativo adequado esperado; candidatos sem
background em business são explicitamente bem-vindos ("non-business
backgrounds encouraged").

CUIDADO APLICADO (dados NÃO encontrados mesmo após tentativas
adicionais, permanecem 'Verificar'): 'Custo/Crédito' (ambos), 'Tuition
Total' (ambos), 'Fees Adicionais', 'Tamanho Turma', 'GPA Médio
Admitidos', 'GMAT Médio' (não aplicável -- teste não exigido), 'Taxa
Aceitação %', 'Ranking US News', 'Ranking QS', 'Taxa Empregabilidade
%', 'Salário Médio Pós USD', 'Principais Empregadores' (mesma cautela
da rodada 1: os únicos números de emprego/salário encontrados eram de
um relatório rotulado como específico de MBA, não usados), 'Processo
TA/RA' (além do que já estava preenchido em 'Graduate Assistantships',
não Verificar, não tocado), 'Caminho Tuition Residente', 'Quão
Realista (Intl)', 'Entrevista', e a maior parte dos campos mensais de
custo de vida em San Antonio, TX.

Fontes principais (pesquisadas via WebFetch, ago/2026):
- https://future.utsa.edu/programs/master/business-analytics/
  (página oficial correta e completa do programa -- admissão,
  deadlines corrigidos, documentos, validação cruzada de GPA/GMAT/
  inglês, bolsas/GTA/GRA já consistentes com rodada 1)
- https://catalog.utsa.edu/graduate/business/operationsanalytics/
  (30 créditos confirmados, currículo completo, menção a "Differential
  Tuition" sem valor exato, programa agora sob o Department of
  Operations and Analytics)
- Múltiplas tentativas de páginas de tuition (future.utsa.edu/tuition,
  onestop.utsa.edu) sem sucesso em obter valores por crédito
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
    "Nome da Universidade": "UTSA",
    "Estado": "TX",
    "Cidade": "San Antonio",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Deadline Aplicação": "Ciclo 2027-28: prioridade internacional 1 jun; doméstico 1 ago",
    "Documentos Necessários": (
        "Transcripts oficiais de todas as instituições; purpose "
        "statement; resume; avaliação de credencial (educação "
        "internacional); cartas de recomendação opcionais (2 "
        "recomendadas)"
    ),
    "Matérias Pré-Requisito": (
        "Nenhum pré-requisito formal além do bacharelado; preparo "
        "quantitativo adequado esperado; candidatos sem background "
        "em business explicitamente bem-vindos"
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
        raise ValueError("Nenhuma linha encontrada para UTSA (San Antonio, TX). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (UTSA, San Antonio, TX)")

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
