"""
Data enrichment: atualiza APENAS a linha da 'U of Tulsa' (Collins
College of Business, Tulsa, OK) no msba_programs_database.xlsx.

LIMITAÇÃO SEVERA DE ACESSO (documentada para transparência): dezenas
de tentativas de localizar a página oficial ESPECÍFICA do programa
"MS in Business Analytics" (via navegação a partir da home, do college
de business, do diretório de programas, do bulletin acadêmico, e via
múltiplos padrões de URL plausíveis) falharam -- todas retornaram 404
ou conteúdo vazio via WebFetch. O orçamento de WebSearch da sessão
também está esgotado. Como resultado, este enrichment ficou limitado
a páginas GERAIS da universidade (tuition/fees, assistantships) que
não são especificas do programa, mas que se aplicam a ele por padrão
(nenhuma exceção de tarifa para "Business Analytics" foi encontrada na
lista de programas com tarifas diferenciadas).

TUITION: 'Custo/Crédito (Intl) USD' ATUALIZADO de $1.548 para $1.610 --
a página oficial de tuition 2026-27 (geral da universidade) lista
"$1.610/semester hour" como a tarifa PADRÃO para a maioria dos
programas de mestrado, com uma lista explícita de EXCEÇÕES (MBA
online, MS Nursing online, DNP, Athletic Training, etc.) que NÃO inclui
Business Analytics -- forte indício de que o MSBA segue a tarifa
padrão. CONFIANÇA MODERADA: não foi possível confirmar isso numa
página específica do programa (inacessível nesta rodada), por isso
documentado como inferência a partir da tarifa geral + ausência na
lista de exceções, não como confirmação direta program-specific.
'Custo/Crédito (Residente) USD' preenchido com o MESMO valor -- a
Universidade de Tulsa é privada com tuition flat (já documentado:
'Caminho Tuition Residente' = 'N/A (privada)').

'Total Créditos' e 'Tuition Total' (ambos os campos de residência)
PERMANECEM 'Verificar' -- sem a página específica do programa, não foi
possível confirmar o número de créditos exigidos, e multiplicar um
custo/crédito por um número de créditos não confirmado inventaria um
total. 'Fees Adicionais USD' também permanece 'Verificar' pelo mesmo
motivo (as fees oficiais encontradas são por SEMESTRE -- University
Fee $658, International Student Services Fee $206 -- mas sem saber o
número de semestres do programa, não há como totalizar com confiança).

'Taxa Aplicação USD' preenchido: $60 (oficial, não-reembolsável).

'Processo TA/RA' preenchido com detalhe oficial e ESPECÍFICO
(university-wide, mas presumivelmente aplicável ao MSBA, já que a
própria planilha já registrava "full tuition + stipend" para este
programa antes desta rodada): stipend de pelo menos $14.349 para
contrato de 9 meses (outono+primavera) + remissão de 100% da tuition
por semestre para mestrado + seguro saúde estudantil incluído;
candidatura via seleção de "Yes" na pergunta de assistantship do
formulário de admissão; exige manutenção de GPA ≥3.0 para reter o
benefício.

CUIDADO APLICADO (dados NÃO encontrados, permanecem 'Verificar'):
'Duração (meses)', 'Tamanho Turma', 'Seguro Saúde (programa) USD'
(o pacote de GA inclui seguro saúde, mas isso não é o mesmo que o
prêmio-base para quem NÃO tem GA), 'GPA Mínimo', 'GPA Médio
Admitidos', 'GMAT Médio', 'Taxa Aceitação %', 'Matérias
Pré-Requisito', 'Documentos Necessários', 'Rounds Aplicação',
'Entrevista', 'Ranking US News', 'Ranking QS', 'Salário Médio Pós
USD', e todos os campos mensais de custo de vida específicos (a única
cifra de custo de vida encontrada foi um valor agregado de "room and
board" no campus, $8.653/semestre com plano de refeições em
acomodação dupla -- não corresponde ao que o campo 'Aluguel Privativo/
mês' pede, que é aluguel privativo fora do campus, então não foi
usado para evitar misturar duas situações de moradia diferentes).

Fontes principais (pesquisadas via WebFetch, ago/2026):
- https://utulsa.edu/tuition-aid/tuition-costs/cost-graduate/ (tuition
  geral 2026-27, lista de exceções por programa, fees obrigatórias,
  taxa de aplicação, orçamento agregado de room & board)
- https://utulsa.edu/tuition-aid/graduate-financial-assistance/assistantships-fellowships/
  (detalhamento oficial de stipend, remissão de tuition, seguro
  saúde incluído, processo de candidatura, requisito de GPA)
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
    "Nome da Universidade": "U of Tulsa",
    "Estado": "OK",
    "Cidade": "Tulsa",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Intl) USD": 1610,  # ATUALIZADO: tarifa padrão oficial 2026-27 (confiança moderada, inferida — não confirmada em página específica do programa)
    "Custo/Crédito (Residente) USD": 1610,  # privada, tarifa flat (igual ao Intl)
    "Taxa Aplicação USD": 60,  # oficial, não-reembolsável
    "Processo TA/RA": (
        "Stipend de pelo menos $14.349 (contrato de 9 meses, outono+"
        "primavera) + remissão de 100% da tuition por semestre "
        "(mestrado) + seguro saúde estudantil incluído; candidatura "
        "via seleção de \"Yes\" na pergunta de assistantship do "
        "formulário de admissão (com opção de anexar CV); exige "
        "manutenção de GPA ≥3.0 para reter o benefício"
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
        raise ValueError("Nenhuma linha encontrada para U of Tulsa (Tulsa, OK). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (U of Tulsa, Tulsa, OK)")

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
