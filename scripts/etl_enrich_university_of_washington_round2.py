"""
Data enrichment (RODADA 2): atualiza APENAS a linha da 'University of
Washington' (Michael G. Foster School of Business, Seattle, WA) no
msba_programs_database.xlsx.

Esta linha já foi enriquecida em rodada anterior desta mesma sessão
(script etl_enrich_university_of_washington.py), restando 14/54 campos
'Verificar'. Usuário confirmou via AskUserQuestion (ago/2026) para
prosseguir com uma Rodada 2 focada APENAS nesses 14 campos
remanescentes, sem re-verificar o que já foi preenchido.

NOTA DE METODOLOGIA: orçamento de busca web da sessão esgotado (200/200
WebSearch); toda a pesquisa desta rodada foi feita via WebFetch,
navegando diretamente pelas páginas oficiais do programa MSBA em
foster.uw.edu.

'GPA Médio Admitidos' PREENCHIDO: 3.5 -- GPA médio oficial da Class of
2026 (página de Class Profile do MSBA).

'Ranking QS' PREENCHIDO: '#7 em outcomes de alumni; #10 em
custo-benefício (value for money)' -- QS World University Rankings
2025, dados apresentados diretamente na página oficial do programa
MSBA (não de uma página genérica do Foster/MBA, diferente de casos
anteriores nesta sessão onde rankings encontrados eram de outro
programa). 'Ranking US News' permanece 'Verificar' -- não encontrado.

'Caminho Tuition Residente' PREENCHIDO com nota de inaplicabilidade
prática: Washington exige domicílio prévio de 1 ano para status de
residente estadual, o que é inviável para uma aluna internacional em
visto F-1 num programa de 12 meses.

'Graduate Assistantships' e 'Processo TA/RA' PREENCHIDOS com 'Não
disponível' -- a página oficial de tuition/financing do MSBA não
menciona nenhuma modalidade de assistantship, consistente com o já
registrado em rodada anterior de que o MSBA é descrito pela própria UW
como um "fee-based graduate program" (financiado majoritariamente por
mensalidade).

'Dificuldade Entrada' PREENCHIDO com avaliação qualitativa baseada nos
dados quantitativos disponíveis (GPA médio 3.5, GRE médio ~317 da
Class of 2026, sem GMAT/GRE mínimo oficial) -- 'Taxa Aceitação %'
permanece 'Verificar' pois não foi encontrada em nenhuma fonte oficial
consultada, então a avaliação de dificuldade é necessariamente
qualitativa/hedged.

'Custo Vida Total (programa) USD' PREENCHIDO com $30.078 -- total
oficial de custos estimados anuais (Room & board $23.199 + Books &
supplies $825 + Local transportation $1.581 + Personal expenses $2.508
+ Health insurance $1.965), da página oficial de Program Fees &
Financing, baseada no orçamento OSFA (Office of Student Financial Aid)
da UW. VALIDAÇÃO CRUZADA: os valores mensais já preenchidos em rodada
anterior batem quase exatamente com este mesmo orçamento oficial
('Transporte/mês' $132 × 12 = $1.584 ≈ $1.581 'Local transportation';
'Outros Gastos/mês' $209 × 12 = $2.508 = exatamente 'Personal
expenses'; 'Seguro Saúde (programa)' $1.965 = exatamente 'Health
insurance') -- confirma que a rodada anterior já usou esta mesma fonte
oficial. 'Tuition Total' também bate: Intl $61.364 + Fees $1.591 =
$62.955 = 'total program cost non-residents' oficial; Residente
$48.172 + $1.591 = $49.763 = 'total program cost WA residents' oficial.
Nenhuma correção necessária nesses campos (fora do escopo desta
rodada de qualquer forma).

CUIDADO: 'Mercado/mês', 'Utilities/mês', 'Celular/mês', 'WiFi/mês' e
'Seguro Residencial/mês' PERMANECEM 'Verificar' -- a página oficial de
custo de vida só fornece "Room and board" como valor COMBINADO
($23.199/ano), sem separar aluguel de alimentação/utilities. Como
'Aluguel Privativo/mês' ($900) já foi preenchido em rodada anterior
com uma fonte DIFERENTE (pesquisa geral de aluguel em Seattle, não
este orçamento oficial combinado), calcular "Mercado" como a diferença
entre o Room & Board oficial e o Aluguel já preenchido misturaria duas
fontes distintas numa proporção inventada -- violaria o princípio já
estabelecido nesta sessão de nunca inventar uma decomposição. Por isso
esses campos ficam 'Verificar' em vez de estimados por subtração.

DISCREPÂNCIA REGISTRADA (NÃO CORRIGIDA -- fora do escopo aprovado desta
rodada, que toca apenas os 14 campos que já estavam 'Verificar'): a
própria página oficial do MSBA lista "Average Starting Salary:
$106.000", que diverge do valor já preenchido em 'Salário Médio Pós
USD' ($140.354, de rodada anterior). Ambos parecem ser de fontes
oficiais do programa, mas de anos/relatórios possivelmente diferentes
-- registrado aqui para uma eventual Rodada 3 revisar com mais
contexto, não alterado nesta rodada por estar fora do escopo aprovado
pelo usuário.

Fontes principais (pesquisadas via WebFetch, ago/2026):
- https://foster.uw.edu/academics/degree-programs/master-science-business-analytics/
  (rankings QS 2025, salário médio $106k conflitante, deadlines)
- https://foster.uw.edu/academics/degree-programs/master-science-business-analytics/student-experience-msba/class-profile/
  (GPA médio 3.5, GRE médio 317, tamanho de turma 99, 67% internacional
  -- Class of 2026)
- https://foster.uw.edu/academics/degree-programs/master-science-business-analytics/program-fees-financing-msba/
  (orçamento OSFA oficial completo, custo total do programa por
  residência, ISIP/ISHIP para internacionais, ausência de menção a
  assistantships)
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
    "GPA Médio Admitidos": 3.5,  # oficial, Class of 2026
    "Ranking QS": "#7 em outcomes de alumni; #10 em custo-benefício (value for money) — QS World University Rankings 2025 (página oficial do MSBA)",
    "Caminho Tuition Residente": (
        "Não aplicável na prática para internacionais em programa de "
        "12 meses — WA exige domicílio prévio de 1 ano para status "
        "de residente estadual, inviável para alunos em F-1"
    ),
    "Graduate Assistantships": "Não disponível — sem menção a assistantships nas fontes oficiais (MSBA é \"fee-based graduate program\")",
    "Processo TA/RA": "N/A — não disponível (ver 'TA/RA Disponível'); nenhuma menção a processo de TA/RA nas fontes oficiais",
    "Dificuldade Entrada": "Moderada — GPA médio 3.5, GRE médio ~317 (Class of 2026); GMAT/GRE sem mínimo oficial; taxa de aceitação não divulgada publicamente",
    "Custo Vida Total (programa) USD": 30078,  # oficial, orçamento anual OSFA (Room&board+Books+Transporte+Pessoal+Seguro Saúde)
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
