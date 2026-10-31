"""
Data enrichment: atualiza APENAS a linha da 'U of Maryland' (Robert H.
Smith School of Business, College Park, MD) no msba_programs_database.xlsx.

Foco: Master of Science in Business Analytics & AI, full-time
presencial, 3 semestres (~12 meses, estimativa documentada -- a fonte
oficial confirma "3 semesters" mas não a duração exata em meses;
assumido 12 meses como estimativa razoável para um programa Fall+
Spring+Fall ou Fall+Spring+Summer, sem confirmação oficial adicional),
30 créditos (já confirmado).

'GPA Médio Admitidos' preenchido: 3.3 (oficial, média de admitidos).
'GPA Mínimo' preenchido: 3.0 recomendado; ensaio opcional exigido se
abaixo de 3.0.

'GMAT/GRE Obrigatório' preenchido com achado IMPORTANTE para o perfil
de busca: o teste É exigido para o programa de Business Analytics
(diferente de MS Accounting/Management/Supply Chain, que não exigem),
mas dispensável com GPA ≥3.2 em graduação quantitativa OU 24+ meses de
experiência profissional relevante pós-graduação (revisão automática,
decisão em até 10 dias úteis). CUIDADO: existe uma dispensa adicional
chamada "Maryland Advantage" que TAMBÉM reduz a tuition, mas ela é
EXPLICITAMENTE EXCLUSIVA para "applicants with an undergraduate and/or
master's degree in the U.S." -- ou seja, NÃO se aplica a uma candidata
com diploma obtido no Brasil. Documentado para não sugerir
elegibilidade indevida a essa dispensa mais vantajosa.

'Testes Inglês Aceitos' refinado com mínimos oficiais: TOEFL iBT 96,
IELTS Academic 7.0, PTE Academic 68; dispensado para diploma de
instituição de língua inglesa qualificada.

'Matérias Pré-Requisito' preenchido: estatística + data analytics ou
cálculo, com nota B ou superior (requisito específico da trilha
Business Analytics, distinto de outras trilhas do mesmo conjunto de
masters do Smith).

'Taxa Aplicação USD' preenchido: $90 (dispensável para militares/
veteranos com comprovação). 'Documentos Necessários' preenchido com a
lista oficial: taxa de aplicação, resume, 1 carta de recomendação,
transcripts oficiais, essay sobre qualificações/interesse no programa,
GMAT/GRE (salvo dispensa), TOEFL/IELTS (diploma não-inglês). 'Deadline
Aplicação' e 'Rounds Aplicação' preenchidos com a estrutura oficial de
6 prazos para Fall 2027: Round 1 (15 out 2026); Round 2 (15 nov 2026);
Prioridade internacional (15 jan 2027); Prioridade doméstica (15 fev
2027); Final internacional (12 mar 2027, com nota de que
internacionais devem submeter até 15 mar); Final doméstico (1 abr
2027, revisão contínua depois).

'Taxa Empregabilidade %' e 'Salário Médio Pós USD' preenchidos com
dados do Employment Report OFICIAL e ESPECÍFICO do programa (Class of
2025, n=107): 81,3% aceitaram ofertas de emprego em até 6 meses;
salário médio $86.922 / mediana $85.000 -- RESSALVA IMPORTANTE
registrada no valor do campo e aqui: apenas 26,6% dos formandos
reportaram dados de salário, abaixo do patamar de 70% que a própria
Smith School considera necessário para uma estatística confiável (o
programa marca esse dado como de confiabilidade reduzida). Preferido
sobre uma cifra mais antiga e igualmente de baixa amostra ($77.000,
2024, 31% reporting) encontrada na página geral do programa, por ser
mais recente. 'Tamanho Turma' preenchido: 115 (Class of 2025, mesmo
relatório, experiência profissional média de 1,58 anos). 'Principais
Empregadores' preenchido com uma amostra da lista oficial: Bank of
America, Capital One, Deloitte LLP, Goldman Sachs & Company, Morgan
Stanley, Freddie Mac.

CUIDADO APLICADO (dados NÃO encontrados, permanecem 'Verificar'):
'Custo/Crédito (Residente)', 'Tuition Total (Residente)', 'Fees
Adicionais USD', 'Seguro Saúde (programa) USD' -- múltiplas tentativas
de páginas de tuition da UMD (tanto do Smith School quanto da Graduate
School geral) retornaram erro 404; o valor Intl já presente na
planilha ($2.073/crédito, $62.175 total) foi mantido sem alteração por
não ter sido contradito por nenhuma fonte nova. 'Processo TA/RA',
'Ranking US News', 'Ranking QS', 'Taxa Aceitação %', 'Entrevista' (não
mencionado como exigência formal em nenhuma fonte oficial consultada),
e todos os campos mensais de custo de vida em College Park, MD.

Fontes principais (pesquisadas via WebFetch, ago/2026):
- https://www.rhsmith.umd.edu/node/39241 (nome oficial, formato,
  duração, GPA médio, currículo, "Maryland Advantage" mencionado pela
  primeira vez, salário 2024 de baixa amostra)
- https://www.rhsmith.umd.edu/programs/business-masters/admissions
  (GPA mínimo, critério completo de dispensa de GMAT/GRE, "Maryland
  Advantage" explicitamente doméstico, mínimos de inglês, 6 deadlines,
  taxa de aplicação, documentos, pré-requisitos por trilha)
- https://www.rhsmith.umd.edu/offices/career-services/employment-reports/business-analytics-report
  (Class of 2025: 81,3% empregabilidade, salário médio/mediana com
  ressalva de baixa taxa de resposta, tamanho de turma 115,
  empregadores)
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
    "Nome da Universidade": "U of Maryland",
    "Estado": "MD",
    "Cidade": "College Park",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Duração (meses)": 12,  # estimativa: 3 semestres a partir de Fall, não confirmado com precisão oficial
    "GPA Mínimo": "3.0 recomendado; ensaio opcional exigido se abaixo de 3.0",
    "GPA Médio Admitidos": 3.3,  # oficial
    "GMAT/GRE Obrigatório": (
        "Exigido, mas dispensável com GPA ≥3.2 em graduação "
        "quantitativa OU 24+ meses de experiência profissional "
        "relevante (revisão automática, decisão em até 10 dias "
        "úteis). ATENÇÃO: a dispensa \"Maryland Advantage\" (que "
        "também reduz tuition) é EXCLUSIVA para diploma de "
        "graduação/mestrado obtido nos EUA — não se aplica a diploma "
        "brasileiro"
    ),
    "Testes Inglês Aceitos": "TOEFL iBT mín. 96; IELTS Academic mín. 7.0; PTE Academic mín. 68; dispensado para diploma de instituição de língua inglesa qualificada",
    "Matérias Pré-Requisito": "Estatística + data analytics ou cálculo, nota B ou superior",
    "Taxa Aplicação USD": 90,  # dispensável para militares/veteranos com comprovação
    "Documentos Necessários": (
        "Taxa de aplicação $90; resume; 1 carta de recomendação; "
        "transcripts oficiais de todas as instituições; essay sobre "
        "qualificações/interesse no programa; GMAT/GRE (salvo "
        "dispensa); TOEFL/IELTS (diploma não obtido em inglês)"
    ),
    "Deadline Aplicação": (
        "Fall 2027: Round 1 (15 out 2026); Round 2 (15 nov 2026); "
        "Prioridade internacional (15 jan 2027); Prioridade doméstica "
        "(15 fev 2027); Final internacional (12 mar 2027, submeter "
        "até 15 mar); Final doméstico (1 abr 2027, revisão contínua "
        "depois)"
    ),
    "Rounds Aplicação": "6 prazos: 15 out, 15 nov, 15 jan (prior. intl), 15 fev (prior. doméstico), 12 mar (final intl), 1 abr (final doméstico)",
    "Taxa Empregabilidade %": 81.3,  # Class of 2025 (n=107), oficial
    "Salário Médio Pós USD": 85000,  # mediana oficial Class of 2025 (média: $86.922); apenas 26,6% reportaram — abaixo do patamar de 70% considerado confiável pela própria escola
    "Tamanho Turma": 115,  # Class of 2025, oficial
    "Principais Empregadores": "Bank of America, Capital One, Deloitte LLP, Goldman Sachs & Company, Morgan Stanley, Freddie Mac",
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
        raise ValueError("Nenhuma linha encontrada para U of Maryland (College Park, MD). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (U of Maryland, College Park, MD)")

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
