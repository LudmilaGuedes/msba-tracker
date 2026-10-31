"""
Data enrichment: atualiza APENAS a linha da 'UC Irvine' (Paul Merage
School of Business, Irvine, CA) no msba_programs_database.xlsx.

Foco: Master of Science in Business Analytics, formato full-time
on-campus, 4 trimestres (Summer→Spring, ~11 meses, já confirmado),
52 unidades (já confirmado), STEM. Programa "self-support" (autofinan-
ciado) -- CONFIRMADO PELO PDF OFICIAL da tuition: a mesma tarifa se
aplica a residentes da Califórnia e não-residentes/internacionais, sem
premium de não-residência, exatamente como já anotado em 'Caminho
Tuition Residente' = 'Self-support (sem premium)'.

TUITION: CORRIGIDO com base no PDF oficial de tuition 2026-27 da
University Registrar da UCI (coorte de entrada Verão 2026, trilha
padrão de 4 trimestres): 'Program Fee' (tuition pura) = $19.075 por
trimestre × 4 = $76.300 TOTAL -- mesma tarifa para residentes e
não-residentes. O valor anterior de 'Tuition Total (Intl)' ($68.920)
não batia com esse total oficial nem com o total incluindo fees
($87.336,58) -- substituído por $76.300 em AMBOS os campos (Intl e
Residente), já que não há diferenciação de tarifa.

O PDF também discrimina as fees obrigatórias separadamente do seguro
saúde, permitindo preencher os dois campos do schema com precisão sem
misturar as duas coisas: 'Seguro Saúde (programa) USD' = $8.825,28
(GSHIP oficial, já incluindo o premium prorata de verão de $1.053,42
específico para quem entra no trimestre de Verão). 'Fees Adicionais
USD' = $2.211,30 (soma de Student Services Access Fee $1.096 + Assoc.
Grad Students Fee $56 + Student Center Fee $615,30 + Bren Events
Center Fee $92 + Recreation Fee $352 -- todas oficiais, mandatory,
SEM incluir o seguro saúde já contabilizado separadamente). 'Custo/
Crédito' derivado dividindo a tuition pura pelos 52 créditos já
confirmados: $1.467,31 (mesmo valor para Intl e Residente).

'GMAT/GRE Obrigatório' preenchido: NÃO exigido, mas recomendado para
candidatos com GPA de graduação abaixo de 3.3. 'GPA Mínimo' preenchido
com a mesma lógica: sem mínimo formal declarado.

'Testes Inglês Aceitos' refinado: TOEFL/IELTS exigido apenas para
candidatos internacionais (mínimos específicos não divulgados
publicamente na página consultada). 'Taxa Aplicação USD' preenchido:
$150 (dispensada para militares/veteranos em serviço ativo dos EUA).
'Documentos Necessários' preenchido com a lista oficial: transcripts
não-oficiais (com tradução para inglês se necessário), 2 cartas de
recomendação, resume, 2 essays (1 em vídeo, 1 escrito), TOEFL/IELTS
oficial (se aplicável). 'Entrevista' preenchido: pode ser solicitada
por convite, não é obrigatória por padrão.

'Deadline Aplicação' e 'Rounds Aplicação' preenchidos com os 5 rounds
oficiais + deadline final, para o ciclo de Verão 2026: Round 1 (1
nov); Round 2 (1 dez); Round 3 (1 fev, prioridade para moradia,
recomendado para internacionais); Round 4 (1 abr, deadline final
recomendado para internacionais); Round 5 (15 mai, deadline estendido
para internacionais); deadline final absoluto (5 jul).

'Tipos de Bolsa' REFINADO com achado IMPORTANTE de elegibilidade: das
3 bolsas de mérito oficiais, apenas a "Merage Scholarship" (até
$10.000, sem necessidade de aplicação, revisão holística) está
amplamente disponível, INCLUINDO para internacionais. As outras duas
NÃO se aplicam ao perfil desta busca: "California Residency
Scholarship" ($10.000) é EXCLUSIVA para residentes domésticos da
Califórnia; "Anteaker Pathway Scholarship" ($2.000 garantido) é
EXCLUSIVA para ex-alunos de GRADUAÇÃO da própria UC Irvine -- o valor
"$2k guaranteed" que já estava na planilha se refere a essa última
bolsa, que NÃO é garantida para uma candidata externa/internacional
sem bacharelado na UCI. Essa ressalva foi adicionada para não sugerir
elegibilidade indevida.

CUIDADO APLICADO (dados NÃO encontrados, permanecem 'Verificar'):
'Tamanho Turma', 'GPA Médio Admitidos', 'GMAT Médio', 'Taxa
Aceitação %', 'Ranking US News', 'Ranking QS', 'Salário Médio Pós
USD' (a página de career impact só disponibiliza um "Employment
Report" para download, não exibido em texto), 'Matérias
Pré-Requisito', e todos os campos mensais de custo de vida específicos
de aluguel/mercado/utilities/etc. (a página de financial aid só
forneceu um orçamento agregado de 9 meses com categorias amplas --
Books/supplies $2.009, Room & board $12.831-29.479, Personal $3.489,
Transportation $3.663 -- sem granularidade suficiente para preencher
os campos mensais individuais sem inventar uma decomposição).
'Taxa Empregabilidade %' e 'Principais Empregadores' já estavam
preenchidos (90%; Experian/Disney/Amazon) e não foram tocados por não
serem 'Verificar' nem estarem incorretos -- uma lista mais longa de
empregadores foi encontrada (Amazon, Edwards Lifesciences, Experian,
EY, Facebook, Google, Huawei, Hulu, Ingram Micro, Northrop Grumman),
mas não substituiu a já existente por não ser uma correção, apenas uma
expansão.

Fontes principais (pesquisadas via WebFetch, ago/2026):
- https://merage.uci.edu/programs/masters/master-science-business-analytics/index.html
  (formato, STEM, currículo, empregadores de destino, links)
- .../admissions.html (GMAT/GRE recomendado condicional, taxa de
  aplicação $150, documentos, entrevista por convite, 5 rounds +
  deadline final)
- .../financial-aid.html (3 bolsas de mérito com elegibilidades
  distintas, orçamento agregado de custo de vida)
- .../career-impact.html (lista expandida de empregadores; sem
  estatísticas de emprego/salário em texto)
- https://www.reg.uci.edu/fees/2026-2027/businessanalytics.pdf (PDF
  oficial da University Registrar: tuition e fees discriminados por
  trimestre e por trilha de duração, 2026-27, "self-support" sem
  diferenciação de residência)
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
    "Nome da Universidade": "UC Irvine",
    "Estado": "CA",
    "Cidade": "Irvine",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Intl) USD": 1467.31,  # 76.300 / 52, oficial 2026-27, self-support (mesmo valor p/ residente)
    "Custo/Crédito (Residente) USD": 1467.31,  # self-support, mesmo valor
    "Tuition Total (Intl) USD": 76300,  # CORREÇÃO: oficial (Program Fee puro, 4 trimestres), era 68.920
    "Tuition Total (Residente) USD": 76300,  # self-support, mesmo valor
    "Fees Adicionais USD": 2211.30,  # oficial: Student Services + Assoc Grad + Student Center + Bren Events + Recreation (sem seguro saúde)
    "Seguro Saúde (programa) USD": 8825.28,  # GSHIP oficial 2026-27, incl. premium prorata de verão
    "GMAT/GRE Obrigatório": "Não exigido; recomendado para candidatos com GPA de graduação abaixo de 3.3",
    "GPA Mínimo": "Sem mínimo formal declarado",
    "Testes Inglês Aceitos": "TOEFL/IELTS exigido apenas para candidatos internacionais (mínimos específicos não divulgados publicamente)",
    "Taxa Aplicação USD": 150,  # dispensada para militares/veteranos em serviço ativo dos EUA
    "Documentos Necessários": (
        "Transcripts não-oficiais (com tradução para inglês se "
        "necessário); 2 cartas de recomendação; resume; 2 essays (1 "
        "em vídeo, 1 escrito); TOEFL/IELTS oficial (se aplicável)"
    ),
    "Entrevista": "Pode ser solicitada por convite; não é obrigatória por padrão",
    "Deadline Aplicação": "Ciclo Verão 2026: Round 1 (1 nov); Round 2 (1 dez); Round 3 (1 fev); Round 4 (1 abr, recomendado final p/ internacionais); Round 5 (15 mai, estendido p/ internacionais); final absoluto (5 jul)",
    "Rounds Aplicação": "5 rounds (1 nov / 1 dez / 1 fev / 1 abr / 15 mai) + deadline final absoluto (5 jul)",
    "Tipos de Bolsa": (
        "Merage Scholarship (até $10.000, sem aplicação, revisão "
        "holística) — disponível para internacionais. California "
        "Residency Scholarship ($10.000) — EXCLUSIVA para residentes "
        "domésticos da CA, não aplicável. \"$2k guaranteed\" "
        "(Anteaker Pathway Scholarship) — EXCLUSIVA para ex-alunos de "
        "GRADUAÇÃO da própria UC Irvine, não garantida para candidata "
        "externa/internacional"
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
        raise ValueError("Nenhuma linha encontrada para UC Irvine (Irvine, CA). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (UC Irvine, Irvine, CA)")

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
