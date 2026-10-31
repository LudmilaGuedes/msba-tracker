"""
Data enrichment: atualiza APENAS a linha da 'Kent State' (College of
Business Administration, Kent, OH) no msba_programs_database.xlsx.

Foco: MS in Business Analytics, trilha IN-PERSON no campus de Kent
(distinta da versão 100% online, também oferecida -- não tocada
aqui), 30 créditos mínimos (já confirmado), STEM ("3-foci STEM program
with OPT for international students", confirmado oficialmente).

TUITION: PREENCHIDO com CONFIANÇA MODERADA. Múltiplas páginas oficiais
específicas de tuition/custos retornaram 404 ou conteúdo sem os
valores. A única fonte que trouxe números foi a página geral de
tuition da Kent State, que ela mesma rotula como "historical rate" (não
confirmado como vigente para 2026-27): Ohio Resident ≈ $631,00/crédito;
Non-Ohio Resident/Internacional ≈ $1.176,05/crédito -- "Graduate
Student rates are based on Kent campus Graduate Rate" (aplicável de
forma geral, sem tabela específica do College of Business). PREENCHIDO
'Custo/Crédito' com esses valores e 'Tuition Total' (30 créditos):
Residente ≈ $18.930; Intl ≈ $35.281,50. RESSALVA EXPLÍCITA: são
valores históricos/aproximados, não uma tarifa 2026-27 confirmada
como vigente -- recomendada confirmação direta com o Financial,
Billing and Enrollment Center antes de decisões financeiras.

'Testes Inglês Aceitos' JÁ ESTAVA CORRETO -- a página oficial do
programa confirma EXATAMENTE os mesmos valores já presentes na
planilha (TOEFL 79/IELTS 6.5/PTE 58/Duolingo 110), validando
cruzadamente sem necessidade de alteração.

'Matérias Pré-Requisito' preenchido: conhecimento fundamental em
estatística, matemática, sistemas de informação de negócios e álgebra;
deficiências podem ser corrigidas com cursos específicos de
nivelamento (BA 24056, CIS 24053, MATH 11012, MGMT 24163).

'Deadline Aplicação' REFINADO com deadlines DOMÉSTICOS adicionais que
não estavam na planilha (que só tinha os internacionais): Fall --
internacional 15 mar / doméstico 15 jun; Spring -- internacional 1 out
/ doméstico 1 dez. 'Rounds Aplicação' preenchido: sem rounds numerados;
aplicações após o deadline são avaliadas conforme disponibilidade de
vagas.

'Documentos Necessários' preenchido com a lista oficial: transcripts
oficiais; resume; goal statement; 2 cartas de recomendação.

CUIDADO APLICADO: a página oficial do programa lista estatísticas de
"Career Outcomes" (ex.: "Data Scientists: 33,5% crescimento, $112.590
salário mediano"), mas essas são projeções OCUPACIONAIS GERAIS do
mercado de trabalho dos EUA (provavelmente do BLS), NÃO estatísticas
específicas dos formandos do MSBA de Kent State -- por isso NÃO foram
usadas para preencher 'Taxa Empregabilidade %' ou 'Salário Médio Pós
USD', que permanecem 'Verificar' (mesmo critério de cautela já
aplicado a outras universidades desta sessão, ex.: Dakota State).

CUIDADO APLICADO (dados NÃO encontrados, permanecem 'Verificar'):
'Fees Adicionais USD', 'Seguro Saúde (programa) USD', 'Tamanho
Turma', 'GPA Médio Admitidos', 'GMAT Médio' (não aplicável -- GRE/
GMAT não são exigidos), 'Taxa Aceitação %', 'Ranking US News',
'Ranking QS', 'Principais Empregadores' (nomes específicos não
encontrados, apenas categorias ocupacionais genéricas), 'Taxa
Aplicação USD', 'Entrevista', 'Bolsas Disponíveis'/'Tipos de Bolsa'
(a página só menciona genericamente "financial aid opportunities",
sem detalhamento), 'Processo TA/RA', 'TA/RA Disponível', 'Caminho
Tuition Residente', e todos os campos mensais de custo de vida em
Kent, OH.

Fontes principais (pesquisadas via WebFetch, ago/2026):
- https://www.kent.edu/business/msba (formato, STEM, currículo, GPA,
  GMAT/GRE não exigido, mínimos de inglês validados, pré-requisitos
  com cursos de nivelamento específicos, deadlines domésticos e
  internacionais, estatísticas ocupacionais genéricas não usadas)
- https://www.kent.edu/business/isba/graduate-programs (sem dados
  adicionais de taxa de aplicação, tamanho de turma, ou bolsas)
- https://www.kent.edu/tuition (única fonte com números de tuition,
  rotulados pela própria página como "historical rate")
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
    "Nome da Universidade": "Kent State",
    "Estado": "OH",
    "Cidade": "Kent",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Residente) USD": 631.00,  # confiança moderada, rotulado "historical rate" pela fonte
    "Custo/Crédito (Intl) USD": 1176.05,  # confiança moderada, rotulado "historical rate" pela fonte
    "Tuition Total (Residente) USD": 18930.00,  # 30 x $631,00 (estimativa, ver ressalva)
    "Tuition Total (Intl) USD": 35281.50,  # 30 x $1.176,05 (estimativa, ver ressalva)
    "Matérias Pré-Requisito": (
        "Conhecimento fundamental em estatística, matemática, "
        "sistemas de informação de negócios e álgebra; deficiências "
        "corrigíveis via cursos de nivelamento (BA 24056, CIS 24053, "
        "MATH 11012, MGMT 24163)"
    ),
    "Deadline Aplicação": "Fall: internacional 15 mar / doméstico 15 jun; Spring: internacional 1 out / doméstico 1 dez",
    "Rounds Aplicação": "Sem rounds numerados; após deadline, avaliado conforme disponibilidade de vagas",
    "Documentos Necessários": "Transcripts oficiais; resume; goal statement; 2 cartas de recomendação",
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
        raise ValueError("Nenhuma linha encontrada para Kent State (Kent, OH). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Kent State, Kent, OH)")

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
