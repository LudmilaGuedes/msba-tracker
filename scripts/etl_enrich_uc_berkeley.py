"""
Data enrichment + DEDUPLICAÇÃO: consolida as DUAS linhas de UC Berkeley
no msba_programs_database.xlsx em uma só.

ACHADO DA AUDITORIA: a planilha tinha DUAS linhas para o mesmo programa
-- 'UC Berkeley' (linha original) e 'UC Berkeley (Haas/IEOR)' -- ambas
descrevendo o Master of Analytics (MAnalytics), hospedado no
departamento de IEOR (College of Engineering), com os MESMOS deadlines
(Priority 15 dez, Final 7 jan), mesma duração (12 meses) e mesmo
formato. A linha "(Haas/IEOR)" tinha inclusive um erro estrutural: o
campo 'Business School' continha o NOME DO PROGRAMA repetido em vez do
nome da escola. Confirmado com o usuário (26/ago/2026): consolidar em
uma única linha, mantendo a linha original ('UC Berkeley', nome mais
limpo e já usado em outras linhas desta planilha) e removendo a
duplicata. Dados que só existiam na linha duplicada (ex.: nomes
específicos de bolsas, detalhes de currículo, internship de verão)
foram MESCLADOS na linha mantida antes da remoção; nenhuma informação
foi perdida.

Foco exclusivo: Master of Analytics, presencial, STEM, no IEOR (College
of Engineering), 12 meses. Deep research em: custo total real de
tuition para internacionais, bolsas/assistantships, requisitos de GPA/
GMAT/GRE/inglês, e custo de vida em Berkeley, CA.

TUITION: CORRIGIDO para o ciclo 2026-2027 (mais atual que o valor de
2025-26 já presente, $37.250, que era só a "Program Fee" de Fall+Spring
daquele ciclo anterior). A página oficial de custos
(analytics.berkeley.edu/costs-and-financing) confirma para 2026-27:
Program Fee $38.000/semestre (Fall e Spring) + $3.600 (Summer, termo
reduzido de capstone/internship) = $79.600 de tuition pura no ano. A
mesma página afirma explicitamente que NÃO HÁ diferenciação de tuition
entre residentes da Califórnia e não-residentes/internacionais.
CORRIGIDO/PREENCHIDO 'Tuition Total' (Intl e Residente) com $79.600.
'Custo/Crédito' permanece 'Verificar' -- o programa cobra uma "Program
Fee" fixa por semestre, não uma tarifa por crédito/unidade, e o total
de créditos do programa não foi encontrado com confiança suficiente
para calcular um valor "por crédito" que a própria universidade não
divulga.

'Fees Adicionais USD' preenchido com $4.793 -- soma de TODAS as taxas
obrigatórias exceto o seguro saúde (que tem campo próprio no schema):
Berkeley Campus Fee ($904,75 x 2 semestres + $466 verão), Document
Management Fee ($119, só Fall), IREF ($141 x 2), Transit Fee ($236 x
2), Access to Student Services Fee ($672 x 2), International Student
Service Fee ($100 x 3 termos, específica de vistos F/J). 'Seguro Saúde
(programa) USD' preenchido separadamente com $4.103 (SHIP, oficial,
cobrado automaticamente salvo dispensa).

Bolsas: consolidadas e detalhadas com nomes oficiais -- Faculty's
Distinguished Scholar Award (cobertura integral, EXCLUSIVA para
residentes da Califórnia com diploma UC/CSU -- não disponível a
internacionais), Master of Analytics Excellence Scholarship, Master of
Analytics Opportunity Grant (necessidade financeira/diversidade), KONE
Fellowship for Future Analytics Leaders (foco em mobilidade urbana/
sustentabilidade). Internacionais não são elegíveis a empréstimos
federais dos EUA, mas podem acessar crédito privado.

Fontes principais (pesquisadas via busca web, ago/2026):
- https://analytics.berkeley.edu/costs-and-financing/ (tuition oficial
  2026-27, breakdown completo de fees, bolsas nomeadas)
- https://analytics.berkeley.edu/faq/ (pré-requisitos: Álgebra Linear,
  linguagem de programação, Probabilidade/Estatística -- equivalentes
  MATH 54/STAT 134/STAT 135; sem entrevista planejada, mas possível via
  Zoom; documentos exigidos)
- Busca web: GPA mínimo 3.0, GRE exigido salvo diploma UC com GPA>=3.5,
  GMAT não mencionado/não exigido, TOEFL 90/IELTS 7.0, taxa de
  aplicação $155 (internacional) / $135 (doméstico)
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

VERIFICAR = "Verificar"

MATCH_KEEP = {
    "Nome da Universidade": "UC Berkeley",
    "Estado": "CA",
    "Cidade": "Berkeley",
}

MATCH_DUPLICATE = {
    "Nome da Universidade": "UC Berkeley (Haas/IEOR)",
    "Estado": "CA",
    "Cidade": "Berkeley",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos na linha mantida.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Destaques Currículo": (
        "IEOR-housed, quant-heavy, capstone; machine learning, risk "
        "modeling, simulation, Python; internship de verão (200h)"
    ),
    "Tuition Total (Intl) USD": 79600,  # CORREÇÃO: ciclo 2026-27 oficial ($38.000 x 2 semestres + $3.600 verão)
    "Tuition Total (Residente) USD": 79600,  # tuition flat — sem diferenciação de residência (oficial)
    "Fees Adicionais USD": 4793,  # Campus Fee + Doc Mgmt + IREF + Transit + Access to Student Services + Intl Service Fee
    "Seguro Saúde (programa) USD": 4103,  # SHIP oficial 2026-27, auto-cobrado salvo dispensa
    "Tipos de Bolsa": (
        "Faculty's Distinguished Scholar Award (cobertura integral, "
        "EXCLUSIVA para residentes da Califórnia com diploma UC/CSU — "
        "não disponível a internacionais); Master of Analytics "
        "Excellence Scholarship; Master of Analytics Opportunity Grant "
        "(necessidade financeira/diversidade); KONE Fellowship for "
        "Future Analytics Leaders (mobilidade urbana/sustentabilidade)"
    ),
    "Quão Realista (Intl)": (
        "Moderada — elegíveis às bolsas de mérito/necessidade "
        "(Excellence, Opportunity Grant, KONE), mas não à Distinguished "
        "Scholar Award (só residentes da Califórnia) nem a empréstimos "
        "federais dos EUA; crédito privado é uma alternativa"
    ),
    "Caminho Tuition Residente": (
        "N/A — tuition já é flat (mesma taxa para residentes e não-"
        "residentes); a Faculty's Distinguished Scholar Award cobre "
        "integralmente para residentes da Califórnia com diploma UC/"
        "CSU, mas é uma bolsa específica, não uma via de tuition"
    ),
    "GPA Mínimo": 3.0,
    "GMAT/GRE Obrigatório": (
        "GRE exigido, salvo para quem tem diploma UC com GPA "
        "cumulativo/avançado ≥ 3.5 (dispensa automática); GMAT não é "
        "mencionado como opção aceita"
    ),
    "Testes Inglês Aceitos": "TOEFL mín. 90 (iBT) / IELTS mín. 7.0 — exigido para não-cidadãos/não-residentes permanentes dos EUA",
    "Matérias Pré-Requisito": (
        "Álgebra Linear, linguagem de programação, e Probabilidade ou "
        "Estatística (equivalentes UC Berkeley: MATH 54, STAT 134/135); "
        "fortemente recomendado completar até a submissão da "
        "candidatura"
    ),
    "Taxa Aplicação USD": 155,  # internacional ($135 doméstico)
    "Documentos Necessários": (
        "Application online, GPA (cumulativo e avançado), GRE, TOEFL/"
        "IELTS (internacionais), 2 cartas de recomendação, resume/CV, "
        "statement of purpose, submissão em áudio/vídeo (1-2 min "
        "respondendo pergunta sobre preparo de carreira)"
    ),
    "Entrevista": "Não planejada por padrão, mas pode ocorrer via Zoom a critério do comitê de admissão",
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

    def find_rows(match):
        rows = []
        for row in ws.iter_rows(min_row=2):
            if (
                row[col_uni - 1].value == match["Nome da Universidade"]
                and row[col_estado - 1].value == match["Estado"]
                and row[col_cidade - 1].value == match["Cidade"]
            ):
                rows.append(row[0].row)
        return rows

    keep_rows = find_rows(MATCH_KEEP)
    dup_rows = find_rows(MATCH_DUPLICATE)

    if len(keep_rows) != 1:
        raise ValueError(f"Esperava exatamente 1 linha para 'UC Berkeley', encontrou {keep_rows}. Abortando.")
    if len(dup_rows) != 1:
        raise ValueError(f"Esperava exatamente 1 linha duplicada 'UC Berkeley (Haas/IEOR)', encontrou {dup_rows}. Abortando.")

    keep_idx = keep_rows[0]
    dup_idx = dup_rows[0]
    print(f"Linha mantida: {keep_idx} (UC Berkeley) | Linha duplicada a remover: {dup_idx} (UC Berkeley (Haas/IEOR))")

    no_fill = PatternFill(fill_type=None)
    verificar_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    before = {}
    after = {}
    for header, new_value in UPDATES.items():
        col_idx = header_to_col[header]
        cell = ws.cell(row=keep_idx, column=col_idx)
        before[header] = cell.value
        cell.value = new_value
        cell.fill = verificar_fill if new_value == VERIFICAR else no_fill
        after[header] = new_value

    # Remove a linha duplicada. Como dup_idx > keep_idx neste caso, a remoção
    # não afeta o índice de keep_idx (linhas abaixo dela são deslocadas).
    if dup_idx < keep_idx:
        raise ValueError("Linha duplicada está acima da linha mantida; lógica de remoção precisa ser revisada antes de prosseguir.")
    ws.delete_rows(dup_idx, 1)
    print(f"Linha {dup_idx} (UC Berkeley (Haas/IEOR)) removida.")

    wb.save(DB_PATH)

    print("\nCampos atualizados na linha mantida (Como Estava -> Como Ficou):")
    for header in UPDATES:
        print(f"  - {header}:")
        print(f"      Antes: {before[header]!r}")
        print(f"      Depois: {after[header]!r}")


if __name__ == "__main__":
    main()
