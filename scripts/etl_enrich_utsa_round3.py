# -*- coding: utf-8 -*-
"""
Data enrichment (RODADA 3): atualiza APENAS a linha 'UTSA' (Carlos Alvarez
College of Business, San Antonio, TX) no msba_programs_database.xlsx.

CONTEXTO: 2 rodadas anteriores (etl_enrich_utsa.py, _round2.py) deixaram a
tuition "Verificar" por nao conseguirem localizar a tabela de tarifas
oficial (paginas retornavam 404 ou sem numeros). Esta rodada usou uma
abordagem diferente (tabela oficial em PDF do Financial Affairs/Bursar) e
conseguiu o valor exato.

RESOLVIDO -- TUITION (fonte: PDF oficial FY26-27,
utsa.edu/financialaffairs/services/fiscal-services/tuition-fees/docs/
fy26-27-graduate.pdf, tabela de tarifas por Semester Credit Hour para o
programa PRESENCIAL do Alvarez College of Business -- NAO confundir com o
programa 100% ONLINE, que tem tarifa unica de $775/SCH sem diferenciacao
por residencia, ja descartado nas rodadas anteriores):
  - Tuition base (Statutory + Incremental + Board Designated):
      Residente: $370,82/SCH | Nao-residente/Intl: $1.429,47/SCH
  - + Differential Tuition do Alvarez College of Business (graduate):
      $140,45/SCH (adicional, mesma taxa para os dois grupos)
  - TOTAL EFETIVO por credito:
      Residente: $511,27/SCH | Intl: $1.569,92/SCH
  - Total para 30 creditos (so tuition, SEM mandatory fees semestrais):
      Residente: ~$15.338 | Intl: ~$47.098

NAO incluido no total acima: "mandatory fees" (Athletics, Library, Student
Services etc.), cobrados a parte, por semestre matriculado (nao por
credito) -- ex.: ~$1.545/semestre a 12 SCH/semestre para residente. Como o
programa tem duracao variavel (11 ou 16 meses, ~3-4 semestres) e a carga de
creditos por semestre varia por aluno, nao foi possivel calcular um total
agregado confiavel -- por isso 'Fees Adicionais USD' permanece "Verificar"
em vez de uma estimativa (a faixa aproximada seria $4.000-$6.000 no total
do programa, mas isso NAO e um numero confirmado o suficiente para
preencher o campo).

RESOLVIDO -- Caminho Tuition Residente: estava "Verificar". CONFIRMADO
oficialmente (graduateschool.utsa.edu/tuition-funding/assistantships/,
secao "In-State Tuition Eligibility Requirements"): "Graduate students
awarded a competitive scholarship of at least $1,000 will receive in-state
tuition for the academic year" -- mesma logica ja confirmada em UT Dallas e
Texas A&M (nao e boato, e politica formal do UT System).

RESOLVIDO -- Processo TA/RA: estava "Verificar". CONFIRMADO oficialmente
(business.utsa.edu/students/financial-aid/assistantships.html): candidatura
via formulario online competitivo, departamentos contratam diretamente
quem selecionam; compensacao para mestrado cobre TUITION + $1.000/mes (ate
9 meses) + ate 19h/semana de trabalho; elegibilidade exige bom desempenho
academico, GPA minimo 3.0, minimo 6 creditos/semestre longo. A pagina nao
exclui candidatos internacionais nem confirma explicitamente elegibilidade
-- registrado com essa ressalva.

ATUALIZADO -- Quão Realista (Intl): estava "Verificar". Com dois caminhos
de financiamento agora confirmados oficialmente (bolsa competitiva >=$1.000
=> tuition residente; e GTA/GRA cobrindo tuition + $1.000/mes), classificado
como "Moderada" -- caminhos reais existem e sao documentados oficialmente,
mas nenhuma taxa de sucesso/seletividade especifica foi divulgada, e a
elegibilidade de internacionais para o GTA/GRA nao e explicitamente
confirmada.

NAO ALTERADO por falta de fonte oficial confiavel (mantido "Verificar"):
- Taxa Empregabilidade % e Salário Médio Pós USD: um valor ($79.000
  mediana, 80% empregados em 3 meses) foi encontrado, mas em
  mbaguidance.com -- um agregador especifico de programas de MBA, nao
  confirmado como sendo do MSBA. Risco de mistura de programas, nao usado.
- Principais Empregadores, Ranking US News, Ranking QS, GPA Médio
  Admitidos, Taxa Aceitação %, Tamanho Turma, Rounds Aplicação, Entrevista:
  nenhuma fonte oficial encontrada nesta rodada.
- GMAT Médio: nao aplicavel (teste nao exigido, ja documentado).

Fontes principais (pesquisadas via WebSearch/WebFetch, set/2026):
- https://www.utsa.edu/financialaffairs/services/fiscal-services/tuition-fees/docs/fy26-27-graduate.pdf
  (oficial -- tabela de tarifas por SCH, residente/nao-residente,
  Differential Tuition do Alvarez College of Business)
- https://graduateschool.utsa.edu/tuition-funding/assistantships/
  (oficial -- politica de in-state tuition via bolsa competitiva >=$1.000)
- https://business.utsa.edu/students/financial-aid/assistantships.html
  (oficial -- processo e cobertura de GTA/GRA)
- https://online.utsa.edu/programs/master/business-analytics/ (programa
  ONLINE, tarifa $775/SCH sem diferenciacao -- confirmado que NAO se aplica
  a esta linha, ja descartado corretamente nas rodadas anteriores)
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "UTSA",
    "Estado": "TX",
    "Cidade": "San Antonio",
}

CUSTO_CREDITO_RESIDENTE = 511.27
CUSTO_CREDITO_INTL = 1569.92
TOTAL_CREDITOS = 30

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Residente) USD": CUSTO_CREDITO_RESIDENTE,
    "Custo/Crédito (Intl) USD": CUSTO_CREDITO_INTL,
    "Tuition Total (Residente) USD": round(CUSTO_CREDITO_RESIDENTE * TOTAL_CREDITOS),  # 15338
    "Tuition Total (Intl) USD": round(CUSTO_CREDITO_INTL * TOTAL_CREDITOS),  # 47098
    "Caminho Tuition Residente": (
        "Sim -- confirmado oficialmente: bolsa competitiva de pelo menos "
        "$1.000 dá direito a tuition de residente pelo ano acadêmico "
        "(\"Graduate students awarded a competitive scholarship of at "
        "least $1,000 will receive in-state tuition for the academic "
        "year\")"
    ),
    "Processo TA/RA": (
        "Candidatura via formulário online competitivo; departamentos "
        "contratam diretamente quem selecionam. Compensação (mestrado): "
        "cobre tuition + $1.000/mês (até 9 meses) + até 19h/semana. "
        "Elegibilidade: bom desempenho acadêmico, GPA mínimo 3.0, mínimo "
        "6 créditos/semestre longo. Elegibilidade de internacionais não "
        "confirmada explicitamente na fonte"
    ),
    "Quão Realista (Intl)": (
        "Moderada -- dois caminhos de financiamento confirmados "
        "oficialmente (bolsa competitiva ≥$1.000 → tuition de residente; "
        "GTA/GRA cobrindo tuition + $1.000/mês), mas sem taxa de "
        "sucesso/seletividade divulgada, e elegibilidade de "
        "internacionais para GTA/GRA não confirmada explicitamente"
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
        raise ValueError("Nenhuma linha encontrada para UTSA (San Antonio, TX). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (UTSA, San Antonio, TX)")

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
