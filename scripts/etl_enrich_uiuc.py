"""
Data enrichment: atualiza APENAS a linha da 'UIUC' (Gies College of
Business, Champaign, IL) no msba_programs_database.xlsx.

Foco: MS in Business Analytics (MSBA), on-campus, trilha "Traditional
Track" (2 semestres, ~9 meses -- consistente com 'Duração (meses)' = 9
já presente na planilha; distinta da "Foundational Track", 3
semestres/~12-15 meses, para quem precisa de nivelamento adicional, e
da versão 100% online MSBAi, nenhuma das duas tocada aqui), 36
créditos (já confirmado), STEM (até 36 meses de OPT).

TUITION: CORRIGIDO. O valor anterior de 'Tuition Total (Intl)'
($66.000) não batia com nenhuma das duas trilhas oficiais e parecia
uma estimativa aproximada sem fonte clara. A página oficial de tuition
2026-27 traz um breakdown LIMPO por trilha e por residência, separando
tuition pura de fees: Traditional Track, 2 semestres --
Internacional: tuition $26.106/semestre + fees $2.847/semestre;
Residente: tuition $19.035/semestre + fees $2.769/semestre. CORRIGIDO
'Tuition Total (Intl)' para $52.212 (2 × $26.106, tuition pura) e
PREENCHIDO 'Tuition Total (Residente)' com $38.070 (2 × $19.035).
'Custo/Crédito' derivado dividindo por 36: $1.450,33 (Intl) / $1.057,50
(Residente). 'Fees Adicionais USD' preenchido com $5.694 (2 ×
$2.847, fees oficiais da trilha internacional -- a trilha residente
tem fees ligeiramente menores, $5.538, mas o valor internacional foi
usado por ser o foco desta pesquisa). NOTA: há também um depósito de
matrícula de $1.000 (não-reembolsável, creditado à tuition), não
incluído em nenhum campo estruturado por ser um valor único e não uma
mensalidade/fee recorrente.

'GMAT/GRE Obrigatório' refinado com o critério oficial completo de
dispensa: alumni ou aluno atual da UIUC; 2+ anos de experiência
full-time em business/analytics; OU GPA de graduação ≥3.3 em áreas
quantitativas (business, ciência da computação, engenharia,
matemática) de instituição credenciada. O comitê pode ainda solicitar
scores mesmo com dispensa concedida, e a ausência de scores "pode
impactar negativamente" a admissibilidade.

'GPA Mínimo' preenchido: nenhum mínimo formal para admissão, mas GPA
≥3.3 é o limiar de dispensa de GMAT/GRE; ensaio opcional fortemente
recomendado para candidatos com GPA ≤3.0. 'Testes Inglês Aceitos'
refinado com dois patamares oficiais: Full-Status Admission (TOEFL
103/IELTS 7.5/Duolingo 135) vs. Limited-Status Admission (TOEFL 79/
IELTS 6.5/Duolingo 115), scores válidos apenas após 1/ago/2024.

'Taxa Aplicação USD' preenchido: $90. 'Documentos Necessários'
preenchido com a lista oficial completa (transcripts, diploma,
"coursework form" de habilidades quantitativas -- explicitamente NÃO
um formulário de pré-requisito, apenas informativo --, resume 2
páginas, statement of purpose 500 palavras, 2 personal statements
curtos, video essay de 3 perguntas/3 min, 2 cartas de recomendação,
GMAT/GRE se não dispensado, teste de inglês para internacionais).
'Entrevista' preenchido: comitê pode conduzir um pequeno número de
entrevistas por convite, se precisar de mais informação; decisão pode
ser tomada com ou sem entrevista.

'Taxa Empregabilidade %' preenchido com 89,8 -- dado OFICIAL e
ESPECÍFICO de graduados INTERNACIONAIS (2025 Illini Success Report),
mais relevante para o perfil desta busca do que a taxa geral de
domésticos (77,8% no mesmo relatório 2025; uma versão anterior, o
relatório de 2024, citava "100% para domésticos" numa página diferente
do site -- não usada aqui por ser menos atual e por não cobrir
internacionais). 'Principais Empregadores' preenchido: Amazon,
Caterpillar, PwC, TikTok, State Farm.

CUIDADO APLICADO (dados NÃO encontrados, permanecem 'Verificar'):
'Salário Médio Pós USD' (a página direciona a um PDF de "Career
Outcomes Report" não acessado), 'Tamanho Turma', 'GPA Médio
Admitidos', 'GMAT Médio', 'Taxa Aceitação %', 'Ranking US News',
'Ranking QS' (nenhuma menção encontrada nas páginas do programa),
'Matérias Pré-Requisito' (o "coursework form" oficial é explicitamente
informativo, não um requisito eliminatório -- não preenchido para não
sugerir uma barreira que não existe), 'Seguro Saúde (programa) USD' e
todos os campos mensais de custo de vida.

Fontes principais (pesquisadas via WebFetch, ago/2026):
- https://giesbusiness.illinois.edu/academics/graduate/business-analytics
  (formato, duração, STEM/OPT, currículo, links de subpáginas)
- .../admissions (GPA, dispensa de GMAT/GRE, mínimos de inglês em 2
  patamares, deadlines, taxa de aplicação $90, documentos, entrevista)
- .../tuition-fees (tuition oficial 2026-27, separado por trilha
  Traditional/Foundational e por residência, com tuition e fees
  discriminados separadamente)
- .../msba-career-outcomes (taxas de emprego 2025 separadas por
  domésticos/internacionais, empregadores)
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
    "Nome da Universidade": "UIUC",
    "Estado": "IL",
    "Cidade": "Champaign",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Intl) USD": 1450.33,  # 52.212 / 36, oficial 2026-27 (tuition pura, Traditional Track)
    "Custo/Crédito (Residente) USD": 1057.50,  # 38.070 / 36
    "Tuition Total (Intl) USD": 52212,  # CORREÇÃO: 2 x $26.106 (tuition pura), era $66.000 sem fonte clara
    "Tuition Total (Residente) USD": 38070,  # 2 x $19.035, oficial
    "Fees Adicionais USD": 5694,  # 2 x $2.847 (fees oficiais, trilha internacional)
    "GMAT/GRE Obrigatório": (
        "Dispensável com: alumni/aluno atual UIUC; 2+ anos de "
        "experiência full-time em business/analytics; OU GPA de "
        "graduação ≥3.3 em área quantitativa (business, CS, "
        "engenharia, matemática); comitê pode ainda solicitar scores "
        "mesmo com dispensa"
    ),
    "GPA Mínimo": "Nenhum mínimo formal; GPA ≥3.3 é o limiar de dispensa de GMAT/GRE; ensaio opcional recomendado para GPA ≤3.0",
    "Testes Inglês Aceitos": (
        "Full-Status Admission: TOEFL 103/IELTS 7.5/Duolingo 135; "
        "Limited-Status Admission: TOEFL 79/IELTS 6.5/Duolingo 115 "
        "(scores válidos apenas após 1/ago/2024)"
    ),
    "Taxa Aplicação USD": 90,
    "Documentos Necessários": (
        "Transcripts não-oficiais com legenda de GPA; diploma; "
        "coursework form de habilidades quantitativas (informativo, "
        "não eliminatório); resume (máx. 2 páginas); academic "
        "statement of purpose (máx. 500 palavras); 2 personal "
        "statements curtos (máx. 250 palavras cada); video essay (3 "
        "perguntas, 3 min); 2 cartas de recomendação; GMAT/GRE (salvo "
        "dispensa); teste de inglês (internacionais)"
    ),
    "Entrevista": "Comitê pode conduzir um pequeno número de entrevistas por convite, se necessário; decisão pode ser tomada com ou sem entrevista",
    "Taxa Empregabilidade %": 89.8,  # oficial, específico de graduados INTERNACIONAIS (2025 Illini Success Report)
    "Principais Empregadores": "Amazon, Caterpillar, PwC, TikTok, State Farm",
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
        raise ValueError("Nenhuma linha encontrada para UIUC (Champaign, IL). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (UIUC, Champaign, IL)")

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
