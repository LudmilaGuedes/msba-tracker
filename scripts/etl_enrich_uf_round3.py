"""
Data enrichment (RODADA 3): atualiza APENAS a linha 'U of Florida'
(Warrington College of Business, Gainesville, FL) no
msba_programs_database.xlsx.

Programa: MS in Business Analytics, presencial (on-campus), STEM.

FONTES OFICIAIS (acessadas em 2026-10-17):
  - https://warrington.ufl.edu/graduate/masters-business-analytics/
  - https://warrington.ufl.edu/admissions-aid/cost-aid/graduate/
  - https://admissions.ufl.edu/apply/graduate/international
  - https://warrington.ufl.edu/recruiters-partners/career-outcomes/
  - https://internationalcenter.ufl.edu/iss/obtaining-f-1-status/start-up-cost/

CONFIRMADO (sem alteração): tuition US$ 19.171 (residente FL) / US$ 52.174
(out-of-state/internacional); professional development fee US$ 650
(internacional); 36 créditos, 10 meses, módulos de 7 semanas, início só
no Fall, on-campus; GMAT/GRE não exigido; taxa US$ 30; GPA 3.0
(upper-division, só para diplomas dos EUA); entrevista para quem estudou
fora dos EUA; salário médio US$ 82.166 + bônus de US$ 5.600; Duolingo
NÃO aceito (UF aceita só TOEFL, IELTS e UF ELI); US News #30.

CORREÇÕES:
- Seguro Saúde: obrigatório; prêmio médio anual 2026-27 = US$ 3.185
  (era US$ 3.150).
- Caminho Tuition Residente: o texto "Near-resident pricing" não tem
  fonte oficial. Internacional F-1 paga a tarifa out-of-state.
- GPA Médio (3.59), GMAT Médio (691) e Taxa de Aceitação (17%): vieram da
  coleta inicial sem fonte registrada; a Warrington não publica perfil da
  turma do MSBA, e um GMAT médio de 691 não combina com um programa que
  não exige GMAT (valor provavelmente de quando o teste era obrigatório).
  Voltam para 'Verificar'.
- Bolsas: a página de custos diz que o MSBA oferece "select scholarship
  and research assistant opportunities to students who have attended
  the program for at least a semester"; a página do programa diz que
  alguns admitidos recebem bolsa junto com a oferta de admissão. As duas
  informações foram combinadas.
- Inglês: subnotas mínimas do TOEFL (4.0 em reading, listening, speaking
  e writing), envio eletrônico obrigatório (código 5812), IELTS só via
  E-Delivery; não aceita somar notas de testes diferentes.
- Documentos: tradução certificada linha a linha para diplomas fora do
  inglês; avaliação NACES é opcional na UF.
- Deadline no formato "Round N:" (prioridade internacional = Round 1).

NÃO ALTERADOS: Taxa Empregabilidade (a Warrington só publica taxa de
colocação de MAcc, MSRE, MSF e MBA); Ranking QS; Tamanho Turma.
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
    "Formato": (
        "Presencial (on-campus, full-time, módulos de 7 semanas); existe "
        "também um Online MS in Business Analytics separado -- não é o foco "
        "desta linha"
    ),
    "Seguro Saúde (programa) USD": 3185,  # prêmio médio anual 2026-27, obrigatório
    "Tipos de Bolsa": (
        "Bolsas para alguns admitidos, avisadas junto com a oferta de "
        "admissão (sem inscrição separada, valores não divulgados); depois "
        "do 1º semestre, bolsas e vagas de research assistant/assistantship "
        "para alunos com bom desempenho, oferecidas pelo departamento. "
        "Internacionais não têm acesso a empréstimos federais/estaduais"
    ),
    "Caminho Tuition Residente": (
        "Não aplicável para F-1 -- internacional paga a tarifa out-of-state "
        "(US$ 52.174)"
    ),
    "GPA Médio Admitidos": VERIFICAR,
    "GMAT Médio": VERIFICAR,
    "Taxa Aceitação %": VERIFICAR,
    "GMAT/GRE Obrigatório": (
        "Não exigido (opcional; notas podem reforçar a candidatura na "
        "análise holística; código UF 5812)"
    ),
    "Testes Inglês Aceitos": (
        "TOEFL iBT (inclusive Home Edition) 4.5 na nova escala, com mínimo de "
        "4.0 em reading, listening, speaking e writing, enviado "
        "eletronicamente pela ETS (código 5812) / IELTS Academic 6.5, só via "
        "IELTS E-Delivery / UF ELI Exit Certificate. Duolingo NÃO é aceito; "
        "não aceita combinar notas de testes diferentes"
    ),
    "Deadline Aplicação": (
        "Round 1: 1 jan 2027 (prioridade, internacional); Round 2: 15 fev "
        "2027 (prazo final, internacional); doméstico: 1 mai 2027 (Fall). "
        "Para ingresso (Fall 2028): ~1 jan 2028 (prioridade)"
    ),
    "Rounds Aplicação": (
        "Sem rounds numerados formais: prioridade internacional 1 jan, prazo "
        "final internacional 15 fev, doméstico 1 mai; só entrada no Fall"
    ),
    "Documentos Necessários": (
        "Statement of purpose (400-600 palavras); resume/CV; 2+ cartas de "
        "recomendação de supervisor (não aceita de colegas/líderes "
        "estudantis); transcripts oficiais ou não oficiais de todas as "
        "instituições (cópias escaneadas com assinaturas e selos), com "
        "tradução certificada linha a linha se não estiverem em inglês; "
        "avaliação NACES opcional; teste de inglês; GMAT/GRE opcional; taxa "
        "US$ 30; entrevista para quem estudou fora dos EUA"
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
