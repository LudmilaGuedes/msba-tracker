# -*- coding: utf-8 -*-
"""
Data enrichment: atualiza APENAS a linha 'Texas A&M University-Texarkana
(A&M-Texarkana)' (College of Business, Engineering and Technology,
Texarkana, TX) no msba_programs_database.xlsx.

CONTEXTO: programa MUITO NOVO -- lançado no Fall 2026 (já documentado na
planilha). A linha nunca tinha passado por uma rodada de enrichment
dedicada, só a adição inicial mínima (etl_add_texas_programs.py), com
quase tudo "Verificar". Muitos campos permanecem "Verificar" nesta rodada
não por falha de busca, mas porque o programa é recém-lançado e
genuinamente NÃO TEM histórico ainda (sem formandos, sem dados de
emprego/salário reais, sem GPA médio de admitidos consolidado, sem
página de admissões dedicada ainda indexada no catálogo).

CORRIGIDO -- Formato: estava "Presencial" sem ressalva. A página oficial
do programa (tamut.edu/academic-programs/business-and-data-analytics/)
confirma que o MSBDA é oferecido em 3 modalidades: "100% Online, Hybrid,
Face-to-Face". Corrigido para deixar claro que esta linha se refere
especificamente à opção presencial (a relevante para visto F-1), não que
o programa só exista nesse formato.

RECONFIRMADO sem alteração: Duração 15 meses ("as little as 15 months to
complete" -- fonte oficial, bate com o já cadastrado); 30 créditos; STEM
Designated (com nota adicional: OPT estendido de até 36 meses confirmado
oficialmente -- acrescentado a Destaques Currículo); Taxa Aplicação USD
$50 (confirmado, sem isenção encontrada para internacionais).

PREENCHIDO com fonte oficial, mas com RESSALVA DE NÍVEL (dados
institucionais gerais, não confirmados como específicos do MSBDA -- o
catálogo ainda não lista o programa na seção de admissões de Business,
reforçando que é recém-lançado):
- Testes Inglês Aceitos: TOEFL 71 (iBT)/550 (PBT)/213 (CBT), IELTS 6.0,
  PTE 50, Duolingo 105 -- nível institucional geral (catalog.tamut.edu),
  não confirmado se há exigência mais alta específica do MSBDA.
- GPA Mínimo: 2,5 (overall ou últimos 60 créditos) -- mínimo
  institucional geral para admissão à universidade; catálogo não
  especifica corte próprio do MSBDA.
- Documentos Necessários: resume/CV, carta de intenção (2-4 páginas),
  cartas de recomendação, transcript oficial de todas as instituições --
  documentos padrão de mestrado a nível institucional; não confirmado se
  são exatamente os mesmos exigidos especificamente pelo MSBDA.

NÃO ALTERADO -- Tuition (Custo/Crédito e Tuition Total, ambos
residente/intl): PERMANECE "Verificar" DELIBERADAMENTE. Encontradas
apenas peças soltas e não confiáveis: uma "Foreign Student Tuition Rate"
de $412,00/SCH está rotulada no catálogo como sendo do ANO LETIVO
2014-2015 (claramente desatualizada); uma "Differential Designated
Tuition" de $46,43/SCH para MBA/MS Accounting é "constante através de
Fall 2025", sem confirmação de que o MSBDA (lançado em Fall 2026) usa a
mesma taxa. A tabela real e atual está numa calculadora/PDF do Business
Office que não pôde ser extraída. Somar essas peças produziria uma
estimativa não confiável -- mantido "Verificar" pela mesma disciplina já
aplicada a outras linhas desta sessão (ex.: UT Arlington).

NÃO ALTERADO por ser programa novo sem histórico (mantido "Verificar",
resposta esperada e correta, não falha de busca): Tamanho Turma, Fees
Adicionais, Bolsas Disponíveis/Tipos de Bolsa/Quão Realista/TA-RA/
Caminho Tuition Residente/Graduate Assistantships, Ranking US News/QS,
Taxa Empregabilidade %, Salário Médio Pós, Principais Empregadores, GPA
Médio Admitidos, GMAT/GRE Obrigatório (não confirmado se exigido
especificamente pelo MSBDA), GMAT Médio, Taxa Aceitação %, Experiência
Profissional, Matérias Pré-Requisito específicas, Deadline Aplicação,
Rounds Aplicação, Entrevista, Dificuldade Entrada.

Fontes principais (pesquisadas via WebSearch/WebFetch, set/2026):
- https://www.tamut.edu/academic-programs/business-and-data-analytics/index.html
  (oficial -- nome do programa, 3 modalidades, 30 créditos/15 meses, STEM + OPT 36 meses)
- https://catalog.tamut.edu/admissions/international-admissions/
  (oficial -- taxa de aplicação $50 sem isenção p/ intl, testes de inglês institucionais)
- https://www.tamut.edu/admissions/graduate/index.html
  (oficial -- GPA mínimo institucional 2,5, documentos gerais de mestrado)
- https://catalog.tamut.edu/tuition-fees/ e /differential-designated-tuition/
  (oficial, mas com dado desatualizado de 2014-2015 -- não usado)
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "Texas A&M University-Texarkana (A&M-Texarkana)",
    "Estado": "TX",
    "Cidade": "Texarkana",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Formato": "Presencial (o programa também é oferecido em 100% Online e Híbrido -- esta linha refere-se especificamente à opção presencial)",
    "Destaques Currículo": (
        "Três pilares de analytics (descritivo, preditivo, prescritivo); "
        "Python, SQL, Tableau; programa lançado no Fall 2026; STEM "
        "Designated com elegibilidade a OPT estendido de até 36 meses "
        "(confirmado oficialmente)"
    ),
    "Testes Inglês Aceitos": (
        "TOEFL 71 (iBT) / 550 (PBT) / 213 (CBT); IELTS 6.0; PTE 50; "
        "Duolingo 105 -- nível institucional geral; não confirmado se há "
        "exigência específica mais alta para o MSBDA"
    ),
    "GPA Mínimo": (
        "2,5 (overall ou últimos 60 créditos) -- mínimo institucional "
        "geral para admissão à universidade; catálogo não especifica "
        "corte próprio do MSBDA"
    ),
    "Documentos Necessários": (
        "Resume/CV, carta de intenção (2-4 páginas), cartas de "
        "recomendação, transcript oficial de todas as instituições -- "
        "documentos padrão de mestrado a nível institucional; não "
        "confirmado se são exatamente os exigidos especificamente pelo "
        "MSBDA"
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
        raise ValueError("Nenhuma linha encontrada para Texas A&M University-Texarkana. Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Texas A&M University-Texarkana, Texarkana, TX)")

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
