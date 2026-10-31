"""
Data enrichment (RODADA 4): atualiza APENAS a linha 'U of Florida'
(Warrington College of Business, Gainesville, FL) no
msba_programs_database.xlsx.

Programa: MS in Business Analytics, presencial (on-campus), STEM.

FONTES OFICIAIS (acessadas em 2026-10-19):
  - https://warrington.ufl.edu/graduate/masters-business-analytics/
  - https://warrington.ufl.edu/admissions-aid/cost-aid/graduate/
  - https://admissions.ufl.edu/apply/graduate/international
  - https://warrington.ufl.edu/recruiters-partners/career-outcomes/
  - https://grad.ufl.edu/gss/funding/mcnair/

CONFIRMADO (sem alteração): tuition US$ 19.171 / US$ 52.174 (36 créditos
-> US$ 532,53 / US$ 1.449,28 por crédito); professional development fee
US$ 650 (intl); 10 meses, módulos de 7 semanas, só Fall, on-campus;
prazos intl 1 jan (prioridade) / 15 fev (final), doméstico 1 mai; taxa
US$ 30; GPA 3.0; GMAT/GRE opcional; TOEFL 4.5 / IELTS 6.5 / UF ELI;
Duolingo NÃO aceito; entrevista para quem estudou fora dos EUA; salário
médio US$ 82.166 + bônus US$ 5.600; US News #30.

CORREÇÕES:
- Principais Empregadores: a página do programa agora lista empregadores
  dos formados (Accenture, Bank of America, Deloitte, EY, ISM, KPMG,
  NextEra Energy, Protiviti, PwC, RSM), o que confirma a lista antiga.
- Graduate Assistantships / Tipos de Bolsa: a página de custos diz que
  assistantships da Graduate School costumam ser reservadas a PhD; no
  MSBA, as vagas de research assistant vêm do próprio programa, após o 1º
  semestre.
- Taxa de aplicação: sem isenção para internacionais (a única isenção
  da Graduate School é para McNair Scholars, programa só dos EUA).
- Inglês: a isenção vale só para quem estudou em país de língua inglesa
  ou em instituição com ensino só em inglês -- diploma brasileiro não
  isenta.

NÃO ALTERADOS: Taxa Empregabilidade (a Warrington publica taxa só de MAcc,
MSRE, MSF e MBA; o MS-ISOM tem salário de US$ 81.453, mas é outro
programa); Ranking QS; Tamanho Turma; GPA/GMAT médios; Taxa Aceitação.
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
    "Nome da Universidade": "U of Florida",
    "Estado": "FL",
    "Cidade": "Gainesville",
}

UPDATES = {
    "Tipos de Bolsa": (
        "Bolsas para alguns admitidos, avisadas junto com a oferta de "
        "admissão (sem inscrição separada, valores não divulgados); depois "
        "do 1º semestre, bolsas e vagas de research assistant oferecidas "
        "pelo próprio programa a alunos com bom desempenho. Assistantships "
        "da Graduate School costumam ser reservadas a PhD. Internacionais "
        "não têm acesso a empréstimos federais/estaduais"
    ),
    "Graduate Assistantships": (
        "Possível após o 1º semestre (research assistant do próprio MSBA, "
        "baseado em desempenho); as da Graduate School costumam ser só "
        "para PhD"
    ),
    "Principais Empregadores": (
        "Accenture, Bank of America, Deloitte, Ernst & Young, ISM, KPMG, "
        "NextEra Energy, Protiviti, PwC, RSM (página oficial do programa). "
        "Sponsors de projetos: Accenture, AT&T, Bank of America, CITY HOME, "
        "Disney, PwC, Tampa Electric"
    ),
    "Testes Inglês Aceitos": (
        "TOEFL iBT (inclusive Home Edition) 4.5 na nova escala, com mínimo de "
        "4.0 em reading, listening, speaking e writing, enviado "
        "eletronicamente pela ETS (código 5812) / IELTS Academic 6.5, só via "
        "IELTS E-Delivery / UF ELI Exit Certificate. Duolingo NÃO é aceito "
        "(PTE também não); não aceita combinar notas de testes diferentes. "
        "Isenção só para quem estudou em país de língua inglesa ou em "
        "instituição com ensino só em inglês -- diploma brasileiro não isenta"
    ),
    "Documentos Necessários": (
        "Statement of purpose (400-600 palavras); resume/CV (educação "
        "pós-ensino médio, empregos, afiliações internacionais, "
        "financiamento de pesquisa); 2+ cartas de recomendação de supervisor "
        "(não aceita de colegas/líderes estudantis); transcripts oficiais ou "
        "não oficiais de todas as instituições (cópias escaneadas com "
        "assinaturas e selos), com tradução certificada linha a linha se não "
        "estiverem em inglês; avaliação NACES opcional; teste de inglês; "
        "GMAT/GRE opcional; taxa US$ 30, sem isenção para internacionais "
        "(só McNair Scholars, programa dos EUA); entrevista para quem "
        "estudou fora dos EUA"
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
        raise ValueError("Nenhuma linha encontrada para U of Florida (Gainesville, FL). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (U of Florida, Gainesville, FL)")

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
