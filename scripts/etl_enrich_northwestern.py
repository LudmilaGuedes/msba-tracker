"""
Data enrichment: atualiza APENAS a linha da 'Northwestern' (McCormick
School of Engineering, Evanston, IL) no msba_programs_database.xlsx.

RESSALVA IMPORTANTE SOBRE O PROGRAMA (decisão documentada, sem
escalar ao usuário -- mesmo critério já usado nos casos Rutgers-Newark
e Pitt Katz nesta sessão): confirmado via pesquisa que a Kellogg
School of Management da Northwestern NÃO oferece nenhum "Master in
Business Analytics" ou programa equivalente autônomo -- seus únicos
mestrados são Full-Time MBA, Evening & Weekend MBA, Executive MBA e
Master in Management, nenhum deles focado em analytics. O programa
mais próximo de um MSBA em toda a Northwestern é o "MS in Machine
Learning and Data Science" (MSDS/MLDS), sediado no McCormick School of
ENGINEERING (não na escola de negócios) -- já era essa a linha
presente na planilha antes desta rodada. Sem exclusão de F-1 ou
ausência de OPT identificada (STEM confirmado, com benefício de OPT
aplicável normalmente) -- prosseguido como "equivalente mais próximo"
sem necessidade de remoção ou escalonamento.

Foco: MS in Machine Learning and Data Science, full-time, presencial,
15 meses / 5 trimestres (já confirmado), turma de 55-60 alunos, início
em setembro.

TUITION: CORREÇÃO CRÍTICA. O valor anterior de 'Tuition Total (Intl)'
($45.946) estava MUITO abaixo do valor real -- a página oficial de
tuition 2026-27 confirma $23.662 POR TRIMESTRE, mais $287 de health
services fee e $125 de student activity fee por trimestre (também por
trimestre). Para os 5 trimestres do programa: tuition pura ≈ 5 ×
$23.662 = $118.310 (ESTIMATIVA: a página menciona que o último
trimestre de outono tem um aumento de ~5%, o que elevaria o total para
~$119.493 -- não há um total único oficial declarado, este é um
cálculo aproximado documentado como tal). CORRIGIDO 'Tuition Total
(Intl)' para $118.310 (base, sem o possível aumento do último
trimestre). Como a Northwestern é privada com tuition flat (já
documentado: 'Caminho Tuition Residente' = 'N/A (privada)'), 'Tuition
Total (Residente)' PREENCHIDO com o MESMO valor.

'Fees Adicionais USD' preenchido com $2.060 -- soma de ($287 + $125) =
$412/trimestre × 5 trimestres, oficial. 'Custo/Crédito' e 'Total
Créditos' PERMANECEM 'Verificar' -- a estrutura de tuition é por
TRIMESTRE (não por crédito), e o número exato de créditos por
trimestre não foi confirmado nas fontes consultadas.

'Seguro Saúde (programa) USD' ATUALIZADO de $5.919 para $5.996 -- a
página oficial confirma o prêmio anual mais atual e específico
(1/set/2026 a 31/ago/2027) do NU-SHIP, obrigatório salvo comprovação
de cobertura alternativa comparável.

'GPA Mínimo' preenchido: 3.0 cumulativo (graduação ou pós-graduação,
escala 4.0 ou equivalente internacional). 'Matérias Pré-Requisito'
preenchido: SEM pré-requisitos formais de disciplinas -- o programa
declara explicitamente isso -- mas recomenda forte base em matemática
e programação; candidatos idealmente vêm de engenharia, ciência,
estatística ou business.

'Documentos Necessários' preenchido: 3 cartas de recomendação;
personal statement; transcripts não-oficiais de todas as instituições;
resume/CV; portfólio ou links de GitHub demonstrando projetos de
código; vídeo de candidato (opcional); TOEFL/IELTS (internacionais sem
diploma em instituição de língua inglesa).

'Deadline Aplicação' refinado com mais contexto: aplicações abrem 1
set; prioridade para bolsas até 1 dez; deadline final 15 jan (já
presente); decisões de admissão saem em meados de março. 'Rounds
Aplicação' preenchido: sem rounds numerados, apenas os 2 marcos acima
(prioridade de bolsa + final).

'Tamanho Turma' preenchido: 55-60 alunos (oficial). 'Taxa
Empregabilidade %' preenchido: 100 -- "100 percent of graduates have
received a job offer within three months of graduation" (dado oficial
e específico do programa).

CUIDADO APLICADO (dados NÃO encontrados ou NÃO alterados por já
estarem preenchidos e não contraditos): 'Testes Inglês Aceitos' (já
tinha valores específicos "IELTS 7.5/TOEFL 95" de rodada anterior; a
nova fonte só confirma que o teste é exigido sem repetir os números --
mantido sem alteração por não haver contradição), 'GMAT/GRE
Obrigatório' (já 'Aceito', não Verificar, sem menção nova que
contradiga), 'Total Créditos', 'Custo/Crédito' (ver nota acima),
'GPA Médio Admitidos', 'Taxa Aceitação %', 'Ranking US News', 'Ranking
QS', 'Salário Médio Pós USD', 'Principais Empregadores', 'Entrevista'
(não mencionado), 'Taxa Aplicação USD' (já preenchido, não tocado),
'Processo TA/RA', 'TA/RA Disponível', 'Bolsas Disponíveis'/'Tipos de
Bolsa' (a fonte só confirma "um pequeno número de bolsas por ano",
sem valores -- não preenchido para não sugerir uma quantidade
específica não confirmada), e todos os campos mensais de custo de
vida em Evanston, IL.

Fontes principais (pesquisadas via WebFetch, ago/2026):
- https://www.kellogg.northwestern.edu/ (confirma ausência de MSBA/
  MBAn na Kellogg School of Management)
- https://www.mccormick.northwestern.edu/machine-learning-data-science/
  (formato, duração, tamanho de turma, currículo, deadlines,
  100% empregabilidade)
- .../admissions/index.html (GPA, documentos, ausência de
  pré-requisitos formais, teste de inglês exigido sem mínimos
  numéricos nesta fonte)
- .../admissions/tuition-financial-aid.html (tuition oficial 2026-27
  por trimestre, fees por trimestre, seguro saúde NU-SHIP atualizado,
  bolsas sem valores)
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
    "Nome da Universidade": "Northwestern",
    "Estado": "IL",
    "Cidade": "Evanston",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Tuition Total (Intl) USD": 118310,  # CORREÇÃO: 5 x $23.662/trimestre, oficial (era 45.946, muito abaixo do real)
    "Tuition Total (Residente) USD": 118310,  # privada, tarifa flat (igual ao Intl)
    "Fees Adicionais USD": 2060,  # ($287 + $125) x 5 trimestres, oficial
    "Seguro Saúde (programa) USD": 5996,  # ATUALIZADO: NU-SHIP oficial 2026-27 (era 5.919)
    "GPA Mínimo": "3.0 cumulativo (graduação ou pós-graduação, escala 4.0 ou equivalente internacional)",
    "Matérias Pré-Requisito": (
        "Sem pré-requisitos formais; recomendada forte base em "
        "matemática e programação; candidatos idealmente de "
        "engenharia, ciência, estatística ou business"
    ),
    "Documentos Necessários": (
        "3 cartas de recomendação; personal statement; transcripts "
        "não-oficiais de todas as instituições; resume/CV; "
        "portfólio ou GitHub com projetos de código; vídeo de "
        "candidato (opcional); TOEFL/IELTS (internacionais sem "
        "diploma em instituição de língua inglesa)"
    ),
    "Deadline Aplicação": "Abre 1 set; prioridade de bolsa até 1 dez; final 15 jan; decisões em meados de março",
    "Rounds Aplicação": "Sem rounds numerados; marcos: prioridade de bolsa (1 dez) e deadline final (15 jan)",
    "Tamanho Turma": "55-60",  # oficial
    "Taxa Empregabilidade %": 100,  # oficial, específico do programa (oferta em até 3 meses pós-formatura)
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
        raise ValueError("Nenhuma linha encontrada para Northwestern (Evanston, IL). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Northwestern, Evanston, IL)")

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
