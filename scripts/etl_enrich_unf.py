"""
Data enrichment: atualiza APENAS a linha da 'University of North Florida (UNF)'
(Jacksonville, FL) no msba_programs_database.xlsx, com base em pesquisa web
focada no programa 'Master of Science in Business Analytics (MSBA)' (presencial,
STEM, Coggin College of Business).

Esta linha ja havia passado por rodada(s) anteriores de pesquisa bastante
profunda (varios campos ja tinham confirmacoes oficiais e ressalvas detalhadas
sobre bolsas/GA para internacionais). Esta rodada foi uma auditoria de
verificacao + tentativa de preencher lacunas remanescentes.

Resultado da auditoria: a maioria dos dados existentes foi CONFIRMADA como
correta e atual para 2026-2027 (tuition por credito, tuition total, testes de
ingles, deadlines, taxa de aplicacao e sua politica de isencao, formula de
dispensa de GMAT/GRE) -- nenhuma mudanca necessaria nesses campos. Duas
correcoes/enriquecimentos concretos foram encontrados (ver UPDATES). Os campos
que continuavam "Verificar" (Taxa Empregabilidade %, Salario Medio Pos USD,
Principais Empregadores, Tamanho Turma, GPA Medio Admitidos, GMAT Medio, Taxa
Aceitacao %, Ranking US News, Ranking QS, Dificuldade Entrada) permanecem
"Verificar": nenhuma fonte oficial ou confiavel com esses numeros especificos
para o MSBA (distintos do MBA/da universidade em geral) foi encontrada nesta
rodada -- estimativas fracas de agregadores de baixa confianca (ex.: uma alegacao
solta de "salario inicial de $65.000" sem fonte oficial) foram descartadas por
nao atenderem ao padrao de confirmacao oficial usado no restante da planilha.

Fontes principais (pesquisadas manualmente, fora deste script):
- https://www.unf.edu/coggin/graduate/msba/index.html (pagina oficial do programa;
  confirma formato "In-Person and Online", duracao "12-20 Months", parceiros de
  projetos aplicados: FIS, Regency Centers, TEA, Crowley)
- https://www.unf.edu/coggin/graduate/msba/msba-admission-criteria.html
- https://www.unf.edu/graduateschool/academics/programs/business-analytics.html
- https://www.unf.edu/controller/student-financial-services/graduate-tuition.html
  (tarifas oficiais 2026-2027: $493.53/credito residente, $1,126.88/credito
  nao-residente -- confirma exatamente os valores ja registrados na planilha)
- https://www.unf.edu/regulations-policies/11-tuitionfees/11-0010R.html (mesma
  confirmacao, com detalhamento da composicao da tarifa)
- https://www.unf.edu/graduateschool/tuition/exclusive-rate.html (Exclusive UNF
  Graduate Tuition Rate, para ex-alunos de graduacao da UNF: $353.78/credito
  residente, $987.13/credito nao-residente -- os valores de $408/$275 que
  estavam registrados na planilha estavam desatualizados/incorretos: $408.10 e
  na verdade apenas o componente "tuition" da tarifa padrao, nao a tarifa
  exclusiva completa)
"""

import sys

sys.stdout.reconfigure(encoding="utf-8")

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "University of North Florida (UNF)",
    "Estado": "FL",
    "Cidade": "Jacksonville",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Caminho Tuition Residente": (
        "A única tarifa reduzida encontrada (\"Exclusive UNF Graduate Tuition "
        "Rate\") é exclusiva para ex-alunos de GRADUAÇÃO da própria UNF "
        "(matriculados em até 5 anos após o bacharelado, com matrícula "
        "contínua): $353,78/crédito residente ou $987,13/crédito não-residente "
        "(CORRIGIDO: valor anterior de \"$408→$275/crédito\" estava desatualizado "
        "-- $408,10 é apenas o componente \"tuition\" dentro da tarifa padrão de "
        "$493,53/crédito, não o valor da tarifa exclusiva). Não se aplica a "
        "candidata internacional sem bacharelado na UNF. Caminho de residência "
        "estadual da FL propriamente dito não encontrado."
    ),
    "Destaques Currículo": (
        "Concentrações em Sports Analytics, FinTech, Marketing Analytics, "
        "Econometrics e Logistics/Supply Chain Management; projetos aplicados "
        "com empresas parceiras (FIS, Regency Centers, TEA, Crowley) -- "
        "CONFIRMADO/ATUALIZADO: TEA adicionada à lista oficial de parceiros de "
        "projetos aplicados (pagina oficial do MSBA); AACSB accredited"
    ),
}


def main():
    rotate_backup()
    print(f"Backup rotacionado em: {DB_PATH.parent / 'msba_programs_database_backup.xlsx'}")

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
        raise ValueError("Nenhuma linha encontrada para University of North Florida (UNF). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (University of North Florida (UNF), Jacksonville, FL)")

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
