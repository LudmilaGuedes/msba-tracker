"""
Data enrichment: atualiza APENAS a linha 'UC San Diego' (Rady School of
Management, La Jolla, CA) no msba_programs_database.xlsx.

Foco exclusivo: MS in Business Analytics (MSBA), presencial, STEM.

CORRECAO IMPORTANTE: a planilha tinha 'Ranking QS': 4. Pesquisa confirma
(duas fontes independentes) que isso e uma leitura errada de "#4 na
California" (QS Business Analytics Rankings 2026, ranking REGIONAL). O
ranking GLOBAL do programa no mesmo QS Business Analytics Rankings 2026 e
#40 -- valor corrigido para manter consistencia com o resto da planilha
(que usa sempre o ranking global, ex.: USC = 14).

CORRECAO: tuition estava em $1.295/credito ($64.750 total), valor
desatualizado. Fonte oficial (TritonLink, coorte Fall 2025) mostra
$1.497/unidade x 50 unidades = $74.850. Programa e "self-supporting":
mesma taxa para residentes, nao-residentes e internacionais (nao ha NR
Supplemental Tuition), entao Intl e Residente recebem o mesmo valor.

Fontes principais (pesquisadas manualmente, fora deste script):
- https://students.ucsd.edu/finances/fees/registration/2025-26/rady.html
  (tuition $1.497/unidade; Student Services Fee $430/trimestre; UC San
  Diego Fees $2.064,48/trimestre; total de taxas obrigatorias anuais
  (exclui tuition e seguro saude) = $8.483,44)
- https://rady.ucsd.edu/programs/ms-business-analytics/tuition-aid.html
  (bolsa merito automatica para todos os admitidos, doméstico e intl, sem
  aplicacao separada; UC Scholarship de $10.000 para egressos de qualquer
  campus UC)
- http://rady.ucsd.edu/programs/ms-business-analytics/admissions/
  (GMAT/GRE opcional -- confirma valor ja existente na planilha; TOEFL
  minimo 85 [competitivo 100+], IELTS 7.0, PTE 65, Duolingo 120; taxa de
  inscricao $150 -- corrige valor antigo de $100; entrevista obrigatoria
  so por convite, plataforma assincrona; pre-requisitos de calculo e
  probabilidade/estatistica, programacao recomendada)
- busca por QS Business Analytics Rankings 2026 (topuniversities.com,
  confirmado por 2 buscas independentes: rank global #40)
- busca por employment report 2022 (88% empregabilidade, $100.094 salario
  medio -- valores ja identicos aos que estavam na planilha, confirmados,
  sem alteracao)

Nao encontrado mesmo apos pesquisa profunda (mantido "Verificar"): GPA
minimo formal, ranking US News especifico do MSBA, deadlines/rounds do
ciclo 2026-27 ("coming soon" no site oficial), TA/RA e Graduate
Assistantships (nao confirmados para este programa self-supporting), e
todos os componentes de custo de vida mensal (aluguel, mercado, seguro
residencial, utilities, celular, wifi, transporte, outros) -- as fontes
disponiveis davam apenas faixas amplas e inconsistentes (ex.: aluguel
variando de $1.050 a $2.914/mes conforme o tipo de moradia), sem numero
oficial e especifico o suficiente para preencher com confianca; por isso
'Custo Vida Total (programa) USD' tambem permanece "Verificar".
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.stdout.reconfigure(encoding="utf-8")

from db_retention import rotate_backup  # noqa: E402

DB_PATH = "data/msba_programs_database.xlsx"
SHEET_DATA = "MSBA Programs"

VERIFICAR = "Verificar"

MATCH = {
    "Nome da Universidade": "UC San Diego",
    "Estado": "CA",
    "Cidade": "La Jolla",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Intl) USD": 1497,  # correção (planilha tinha 1295)
    "Custo/Crédito (Residente) USD": 1497,  # correção (planilha tinha 1295)
    "Tuition Total (Intl) USD": 74850,  # correção: 1497 x 50 unidades (planilha tinha 64750)
    "Tuition Total (Residente) USD": 74850,  # self-supporting: mesma taxa para todos
    "Fees Adicionais USD": 8483,  # taxas obrigatórias anuais (Student Services + UCSD campus fees), exclui seguro saúde
    "Tipos de Bolsa": (
        "Merit (automático, sem aplicação separada); UC Scholarship de "
        "$10.000 para egressos de qualquer campus UC"
    ),
    "Ranking QS": 40,  # CORREÇÃO: planilha tinha 4, que era o rank regional "#4 na Califórnia", não o global
    "Testes Inglês Aceitos": (
        "TOEFL 85 mín. (competitivo 100+) / IELTS 7.0 / PTE 65 / "
        "Duolingo 120 (todos válidos por 2 anos)"
    ),
    "Taxa Aplicação USD": 150,  # correção (planilha tinha 100)
    "Documentos Necessários": (
        "Resume/CV, transcritos (com tradução certificada se necessário), "
        "2 cartas de recomendação, Statement of Purpose (máx. 500 "
        "palavras), avaliação curso-a-curso para diploma estrangeiro "
        "(exigida após admissão condicional)"
    ),
    "Entrevista": (
        "Obrigatória, apenas por convite; plataforma assíncrona online "
        "(casos selecionados podem ter entrevista ao vivo)"
    ),
    "Matérias Pré-Requisito": (
        "Cálculo e probabilidade/estatística em nível universitário "
        "(preferencialmente notas A/B); programação em Python, R, SQL "
        "e/ou MATLAB recomendada"
    ),
}


def main():
    rotate_backup()
    print("Backup da versão anterior salvo em: data/msba_programs_database_backup.xlsx")

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
        raise ValueError("Nenhuma linha encontrada para UC San Diego (La Jolla, CA). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (UC San Diego, La Jolla, CA)")

    no_fill = PatternFill(fill_type=None)

    before = {}
    after = {}
    for header, new_value in UPDATES.items():
        col_idx = header_to_col[header]
        cell = ws.cell(row=row_idx, column=col_idx)
        before[header] = cell.value
        cell.value = new_value
        cell.fill = no_fill  # nenhum novo valor aqui e "Verificar"
        after[header] = new_value

    wb.save(DB_PATH)

    print("\nCampos atualizados (Como Estava -> Como Ficou):")
    for header in UPDATES:
        print(f"  - {header}:")
        print(f"      Antes: {before[header]!r}")
        print(f"      Depois: {after[header]!r}")


if __name__ == "__main__":
    main()
