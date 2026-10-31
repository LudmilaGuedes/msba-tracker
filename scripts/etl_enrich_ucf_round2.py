"""
Data enrichment (round 2): atualiza APENAS a linha 'UCF' (Barry S. Miller
College of Business, Orlando, FL) no msba_programs_database.xlsx.

ATENCAO (repetida do round 1): existe uma segunda linha, 'UCF
(Professional)', que e um programa DIFERENTE e NAO deve ser tocada. O
matching abaixo exige Nome da Universidade == 'UCF' (exato) E Estado ==
'FL' E Cidade == 'Orlando'.

Foco exclusivo: Master of Science in Business Analytics, presencial, STEM.
Deep research em: custo de tuition p/ internacionais (reconfirmado, sem
mudanca), bolsas/assistantships, GPA/GMAT/GRE/inglês (sem mudanca), custo
de vida local -- a pedido da usuaria, seguindo o mesmo processo aplicado a
FSU e USF nesta mesma sessao de trabalho.

TUITION: reconfirmado sem mudanca direto na pagina oficial
(business.ucf.edu/degree/ms-business-analytics/): $369,65/crédito
residente, $1.194,05/crédito não-residente, 30 créditos -- os totais já
cadastrados ($11.090 / $35.822) continuam corretos.

SEGURO SAUDE (novo, antes "Verificar"):
https://studenthealth.ucf.edu/payments/intl/ -- UCF International Student
Health Insurance Plan (SHIP) 2026-27, via UnitedHealthcare/Academic
HealthPlans (AHP), efetivo a partir de 15/ago/2026:
  Fall $958 + Spring $876 + Summer $683 = Annual $2.517 (bate exato)
Usado o valor ANUAL ($2.517), já que a duração do programa é 12 meses
(cobre Fall+Spring+Summer integralmente).

BOLSAS (enriquecido -- antes só "GTA automático", agora com mecanismo e
fonte explícitos):
- https://business.ucf.edu/scholarships/: sistema "A2O: Access to
  Opportunities" (ucf.academicworks.com) -- candidatura única por ano,
  faz match automático com bolsas do College of Business + bolsas
  universitárias abertas a todos os cursos + bolsas externas. Confirmado
  que graduate students são elegíveis, mas SEM valores específicos
  divulgados publicamente nem critério dedicado ao MS-BA.
- https://graduate.ucf.edu/fellowships/: College of Graduate Studies
  distribui MAIS DE $2 MILHÕES/ano em Graduate Fellowships por mérito
  acadêmico, pagas via Office of Student Financial Assistance, SEM
  obrigação de trabalho (diferente de GTA) -- mecanismo adicional, não
  documentado antes nesta linha.
- Reconfirmado e refinado o mecanismo de GTA já cadastrado: as vagas são
  ofertadas especificamente pelo Department of Economics (não é um pool
  geral do College of Business), em número LIMITADO e avaliadas de forma
  COMPETITIVA -- candidatos ao MS-BA são automaticamente considerados,
  mas "qualified candidates will be contacted if selected" (ou seja,
  seleção != certeza).

CAMINHO TUITION RESIDENTE (novo, antes "Verificar"): preenchido como "Não"
-- mesma lei estadual da Flórida (Regulation 7.005) já confirmada nesta
sessão de trabalho para FSU e USF: alunos com visto F-1 não são elegíveis
a tuition de residente. Regra é estadual, não específica da UCF, então
reaproveitada com confiança.

DESTAQUES CURRICULO (enriquecido -- antes só "Analytics interdisciplinar",
frase genérica demais):
Fonte oficial (business.ucf.edu/degree/ms-business-analytics/): formato
full-time (1 ano/3 termos) ou part-time (2 anos/5 termos), 9
disciplinas/30 créditos; estatísticas de mercado citadas pela própria
página -- "11% projected growth of Management Analyst jobs (2021-2031)"
e "86% of corporate recruiters plan to hire MS BA grads".

NAO ENCONTRADO mesmo apos pesquisa adicional (mantido "Verificar"):
Tamanho Turma, Ranking US News, Ranking QS, Taxa Empregabilidade %,
Salário Médio Pós USD, Principais Empregadores, GPA Médio Admitidos,
GMAT Médio, Taxa Aceitação %, Rounds Aplicação, Entrevista, Fees
Adicionais, Seguro Residencial/mês, Utilities/mês, Celular/mês, WiFi/mês.
Um ranking "#112 (US News, MBA)" apareceu em busca, mas é do programa de
MBA da UCF, NÃO do MS-BA -- descartado para não contaminar esta linha com
dado de outro curso (mesmo critério já usado no round 1 para não usar
dados do MS em Data Analytics). O handbook oficial do programa
(Business-Analytics-MS-Handbook-2023-2024.pdf) foi localizado mas o PDF
não pôde ser lido de forma confiável pela ferramenta de fetch.
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "UCF",
    "Estado": "FL",
    "Cidade": "Orlando",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Seguro Saúde (programa) USD": 2517,  # UCF SHIP (AHP/UnitedHealthcare) 2026-27, anual: 958+876+683
    "Tipos de Bolsa": (
        "GTA automático (Department of Economics, vagas limitadas, "
        "competitivo -- candidatos ao MS-BA são automaticamente "
        "considerados, mas seleção não é garantida). Sistema A2O: "
        "Access to Opportunities (ucf.academicworks.com) -- candidatura "
        "única/ano, faz match com bolsas do College of Business + bolsas "
        "universitárias + externas. Graduate Fellowships da College of "
        "Graduate Studies: > $2 milhões/ano por mérito acadêmico, sem "
        "obrigação de trabalho (diferente de GTA)"
    ),
    "Processo TA/RA": (
        "GTA de 0,5 FTE (20h/semana), até 2 disciplinas; inclui stipend + "
        "waiver de tuition (9 créditos) + seguro saúde quando ≥ 0,5 FTE. "
        "Vagas ofertadas pelo Department of Economics (não um pool geral "
        "do College of Business), em número limitado; candidatos ao "
        "MS-BA são automaticamente considerados, contatados se "
        "selecionados -- competitivo, não garantido"
    ),
    "Caminho Tuition Residente": (
        "Não -- alunos com visto F-1 não são elegíveis a tuition de "
        "residente na Flórida (FL Regulation 7.005, regra estadual, "
        "mesma aplicada a todas universidades públicas do estado)"
    ),
    "Destaques Currículo": (
        "Formato full-time (1 ano / 3 termos) ou part-time (2 anos / 5 "
        "termos), 9 disciplinas / 30 créditos; analytics interdisciplinar; "
        "página oficial cita crescimento projetado de 11% em vagas de "
        "Management Analyst (2021-2031) e que 86% dos recrutadores "
        "corporativos pretendem contratar formados do MS-BA"
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
        raise ValueError("Nenhuma linha encontrada para UCF (Orlando, FL). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (UCF, Orlando, FL) -- NAO 'UCF (Professional)'")

    no_fill = PatternFill(fill_type=None)

    before = {}
    after = {}
    for header, new_value in UPDATES.items():
        col_idx = header_to_col[header]
        cell = ws.cell(row=row_idx, column=col_idx)
        before[header] = cell.value
        cell.value = new_value
        cell.fill = no_fill
        after[header] = new_value

    wb.save(DB_PATH)

    print("\nCampos atualizados (Como Estava -> Como Ficou):")
    for header in UPDATES:
        print(f"  - {header}:")
        print(f"      Antes: {before[header]!r}")
        print(f"      Depois: {after[header]!r}")


if __name__ == "__main__":
    main()
