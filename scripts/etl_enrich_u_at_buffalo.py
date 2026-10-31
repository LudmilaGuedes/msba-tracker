"""
Data enrichment: atualiza APENAS a linha da 'U at Buffalo SUNY'
(School of Management, Buffalo, NY) no msba_programs_database.xlsx.

NOTA DE METODOLOGIA: orçamento de busca web da sessão esgotado
(200/200 WebSearch); pesquisa feita via WebFetch, navegando
diretamente pelo site oficial (management.buffalo.edu,
buffalo.edu/studentaccounts). A tabela de tuition NÃO-residente/
internacional está atrás de uma aba interativa que não foi capturada
pelo WebFetch (apenas a aba "Resident" veio no conteúdo extraído) --
por isso a tuition internacional usa o total JÁ PRESENTE na planilha
(não alterado), com o custo por crédito derivado dele.

CORREÇÃO DE NOME DE PROGRAMA: 'Nome Oficial do Programa' atualizado de
'MS in Business Analytics' para 'Master of Science in Business
Analytics and Applied AI' -- nome oficial atual confirmado no site da
School of Management (mesmo padrão de atualização de branding já visto
em outras escolas nesta sessão, ex.: Pitt). STEM confirmado, com nota
oficial explícita de 24 meses de OPT adicional para portadores de F-1
(já registrado na planilha em 'Destaques Currículo': 'STEM + 24mo
OPT').

'Duração (meses)' CORRIGIDO de 12 para 10 -- a página oficial do
programa declara explicitamente "10 months" para o formato full-time
on-campus, não 12.

TUITION: 'Custo/Crédito (Residente)' PREENCHIDO com $471 -- tarifa
oficial de tuition (Fall 2026, SEM incluir fees, que são reportadas
separadamente na mesma fonte). 'Tuition Total (Residente)' = 30 × $471
= $14.130. 'Custo/Crédito (Intl)' PREENCHIDO POR DERIVAÇÃO do total já
presente e não alterado na planilha ($36.178 ÷ 30 créditos =
$1.205,93) -- a tabela de tuition não-residente do site oficial da UB
está atrás de uma aba JavaScript que não foi capturada pelo WebFetch
em nenhuma tentativa, então o total pré-existente foi mantido como
está (não havia indicação de que fosse um valor de baixa confiança) e
apenas o custo por crédito foi derivado dele.

'Taxa Aplicação USD' preenchido com $100 (confirmado oficialmente,
dispensada para militares/veteranos dos EUA e cônjuges). 'Documentos
Necessários' preenchido com a lista oficial (aplicação online, resume,
transcripts não-oficiais em PDF de todas as instituições; a página
menciona requisitos adicionais para internacionais sem detalhá-los
nesta página específica). 'Deadline Aplicação' preenchido com o
deadline de PRIORIDADE para bolsas/assistantships (15 jan) -- a
página de admissões geral não lista um deadline final claro além
disso, sugerindo admissão contínua (rolling) fora do prazo de
prioridade.

'Processo TA/RA' preenchido, confirmando e detalhando a nota já
existente em 'TA/RA Disponível' ("competitivo, 20hrs"): seleção
competitiva, exige matrícula full-time, ~20h/semana de trabalho
administrativo/ensino/pesquisa, remuneração = stipend + bolsa de
tuition (valores exatos não divulgados publicamente); mesmo deadline
de prioridade de 15 jan.

'Principais Empregadores' preenchido com uma amostra da lista oficial
de +50 empregadores de egressos (BNY Mellon, Chubb, Comcast, Centene
Corp., BlueCross BlueShield, entre outros, em setores financeiro, de
saúde e tecnologia).

CUIDADO APLICADO: a página do programa cita "#1 ranked business school
in SUNY" e "top 10% of AACSB-accredited business schools" -- essas NÃO
são métricas no formato esperado pelos campos 'Ranking US News' ou
'Ranking QS' desta planilha (não são rankings numerados dessas
publicações específicas), então NÃO foram usadas para preencher esses
campos, que permanecem 'Verificar'. 'Fees Adicionais USD' também
permanece 'Verificar': a fonte oficial dá um valor "por crédito"
implícito ($184,86) derivado de um único exemplo de carga horária
(12 créditos/semestre), mas não confirma se essa é uma tarifa
verdadeiramente por-crédito ou uma taxa fixa por semestre para
regime full-time -- calcular o total exigiria assumir a estrutura
exata de semestres do programa, o que não foi confirmado; documentado
aqui em vez de estimado, para não empilhar suposição sobre suposição
(mesmo princípio já estabelecido nesta sessão). NOTA PARCIAL PARA
RODADA FUTURA: a "International Student Fee" oficial é de $100/
semestre, mencionada explicitamente na fonte, mas não incorporada a
nenhum campo estruturado por não haver confirmação da fee
comprehensive completa para complementá-la.

Fontes principais (pesquisadas via WebFetch, ago/2026):
- https://management.buffalo.edu/degree-programs/master-of-science-ms/business-analytics-applied-ai.html
  (nome oficial atual, formato, duração, STEM+OPT, rankings
  qualitativos, links de subpáginas)
- .../admissions.html (GMAT/GRE test-optional, taxa de aplicação $100,
  documentos exigidos)
- .../program.html (30 créditos confirmados oficialmente)
- .../funding-scholarships-fellowships.html (GA/TA/RA: processo,
  20h/semana, stipend+tuition scholarship, deadline de prioridade 15
  jan; Graduate Opportunity Program para EOP/HEOP/SEEK)
- .../career-preparation/employment.html (lista de +50 empregadores de
  egressos; sem estatísticas de salário/taxa de emprego)
- https://www.buffalo.edu/studentaccounts/tuition-and-fees/fall.html
  (tuition residente oficial Fall 2026: $471/crédito + fees ~$184,86/
  crédito; International Student Fee $100/semestre; aba de tuition
  não-residente inacessível via WebFetch)
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
    "Nome da Universidade": "U at Buffalo SUNY",
    "Estado": "NY",
    "Cidade": "Buffalo",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Nome Oficial do Programa": "Master of Science in Business Analytics and Applied AI",
    "Duração (meses)": 10,  # CORREÇÃO: oficial full-time, não 12
    "Custo/Crédito (Residente) USD": 471,  # oficial Fall 2026 (tuition, sem fees)
    "Tuition Total (Residente) USD": 14130,  # 30 x $471
    "Custo/Crédito (Intl) USD": 1205.93,  # derivado do total já presente ($36.178 / 30, não alterado)
    "Taxa Aplicação USD": 100,  # oficial, dispensada para militares/veteranos dos EUA e cônjuges
    "Documentos Necessários": (
        "Aplicação online via Application Status Portal; resume; "
        "transcripts não-oficiais (PDF) de todas as instituições "
        "cursadas; internacionais têm requisitos adicionais "
        "referenciados na página de admissões (não detalhados nesta "
        "fonte)"
    ),
    "Deadline Aplicação": "Prioridade para bolsas/assistantships: 15 jan; admissão aparentemente contínua (rolling) fora desse prazo",
    "Processo TA/RA": (
        "Seleção competitiva; exige matrícula full-time; ~20h/semana "
        "de trabalho administrativo, ensino ou pesquisa; remuneração "
        "= stipend + bolsa de tuition (valores não divulgados "
        "publicamente); mesmo deadline de prioridade de 15 jan"
    ),
    "Principais Empregadores": "BNY Mellon, Chubb, Comcast, Centene Corp., BlueCross BlueShield (entre +50 empregadores listados oficialmente)",
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
        raise ValueError("Nenhuma linha encontrada para U at Buffalo SUNY (Buffalo, NY). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (U at Buffalo SUNY, Buffalo, NY)")

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
