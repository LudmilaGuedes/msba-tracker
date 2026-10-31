# -*- coding: utf-8 -*-
"""
Data enrichment (round 6): atualiza APENAS a linha da 'U of Cincinnati'
(Lindner College of Business, Cincinnati, OH) no msba_programs_database.xlsx.

Esta e a 6a rodada de enrichment desta linha (ver etl_enrich_cincinnati.py
ate _round5.py). As rodadas anteriores ja cobriram com solidez: tuition
(fonte: email direto da admissao da UC, set/2026), bolsas/TA-RA, requisitos
de admissao (GPA minimo, GMAT/GRE opcional, testes de ingles com Duolingo),
documentos e deadlines. Esta rodada foca EXCLUSIVAMENTE nos campos que
ainda estavam "Verificar" ou sem fonte identificada, via Deep Research
(WebSearch/WebFetch) em fontes oficiais (business.uc.edu, grad.uc.edu).

ATUALIZADO com fonte oficial:
- Ranking US News: estava "Verificar". Confirmado em
  business.uc.edu/about/rankings.html: U.S. News & World Report (2024)
  rankeia a especialidade "Business Analytics" da Lindner em #30 (ranking
  por especialidade, nao o ranking geral do MBA).
- Salario Medio Pos USD: estava 92.498, sem fonte registrada nesta base.
  business.uc.edu/career-ready/career-outcomes.html (relatorio oficial de
  outcomes por programa, ano letivo 2024-25) tem uma linha propria para
  "MS-Business Analytics": mediana e media = $107.500 (min $53.000, max
  $157.000; ~80% dos formandos responderam). Substituido pelo valor oficial
  mais recente e especifico ao programa.
- Destaques Curriculo: estava generico ("STEM, Data Analytics").
  business.uc.edu/programs-degrees/graduate/specialized-masters/
  business-analytics.html (pagina oficial do programa) descreve a
  estrutura real: 40 creditos (ate 10 de business foundations, 17 de core
  analytics, 1 de capstone, 12 de eletivas), possibilidade de cursar
  simultaneamente 1 dos 20+ graduate certificates da Lindner para dupla
  credencial, e projetos hands-on com parceiros corporativos (Cincinnati
  sedia sedes de 8 empresas Fortune 500).

NAO ALTERADO por falta de fonte oficial (mantido como estava, incluindo
"Verificar" onde ja era o caso) -- dados de terceiros nao-oficiais
(mim-essay.com, campusreel.org) foram encontrados para varios destes
campos, mas NAO foram usados por nao serem fonte primaria/oficial e por
poderem estar desatualizados ou incorretos:
- Tamanho Turma: nao encontrado em fonte oficial. Mantido "Verificar".
- Fees Adicionais USD: paginas oficiais de tuition/admissoes nao mencionam
  taxa de programa unica alem da tuition por credito. Mantido "Verificar".
- GPA Medio Admitidos: so achado em fonte terceira (mim-essay.com: "3.3+"),
  nao ha class profile oficial publicada para o MSBA especificamente.
  Mantido "Verificar".
- GMAT Medio: idem (terceiro cita "310 GRE / 650 GMAT", nao oficial).
  Mantido "Verificar".
- Taxa Aceitacao %: so achada em fonte terceira (campusreel.org: "~21,7%").
  Mantido "Verificar".
- Entrevista: nenhuma pagina oficial de admissoes/checklist menciona
  entrevista (nem confirma, nem descarta). Mantido "Verificar".
- Taxa Empregabilidade %: valor atual (92) ja estava na planilha sem fonte
  registrada. Uma fonte terceira (mim-essay.com) cita "96% placement" --
  direcionalmente proximo, mas NAO oficial. A pagina oficial de outcomes
  (career-outcomes.html) publica salario por programa mas NAO uma taxa de
  "placement rate" isolada por programa. Como nao ha fonte oficial para
  confirmar OU contradizer o valor existente, ele foi mantido sem alteracao
  (nao ha base para dizer que esta errado).
- Principais Empregadores: mantida a lista existente (Amazon, Deloitte,
  P&G, Kroger, Google, Disney). Fonte terceira nao-oficial (mim-essay.com)
  cita uma lista mais ampla (adicionando Apple, U.S. Bank, IBM, Macy's,
  Morgan Stanley), mas por nao ser oficial nao foi incorporada.
- Taxa Aplicacao USD ($70): confirmado como correto em fonte oficial
  (grad.uc.edu/admissions/criteria/fees.html) -- sem alteracao necessaria.

NOTA ADICIONAL (fora do schema, sem coluna dedicada): existe um programa
oficial chamado "Lindner Launch" (business.uc.edu/future-students/
admissions-requirements/graduate/lindner-launch.html) que isenta a taxa de
aplicacao e dispensa ensaio/cartas de recomendacao -- mas e EXCLUSIVO para
ex-alunos de GRADUACAO da propria Lindner (GPA UC >= 3.2, formado ha ate 3
anos ou formando no proximo ano letivo). Nao se aplica a uma candidata
internacional externa. A unica isencao geral de taxa de aplicacao da UC
(grad.uc.edu) e para militares/veteranos dos EUA -- internacionais nao sao
elegiveis por nenhuma via encontrada. Registrado aqui e no resumo do
terminal por ser relevante a candidata, mas sem campo no schema.

Fontes principais (pesquisadas via WebSearch/WebFetch, set/2026):
- https://www.business.uc.edu/career-ready/career-outcomes.html
  (oficial -- salario por programa, MS-Business Analytics 2024-25)
- https://www.business.uc.edu/programs-degrees/graduate/specialized-masters/business-analytics.html
  (oficial -- estrutura curricular: creditos, capstone, certificates, parcerias)
- https://www.business.uc.edu/future-students/admissions-requirements/graduate/lindner-launch.html
  (oficial -- fee waiver "Lindner Launch", exclusivo p/ ex-alunos Lindner)
- https://www.business.uc.edu/future-students/admissions-requirements/graduate.html
  (oficial -- checklist de admissao, sem mencao a entrevista)
- https://www.grad.uc.edu/admissions/criteria/fees.html
  (oficial -- taxa de aplicacao $70 intl confirmada; fee waiver so p/ veteranos)
- https://www.business.uc.edu/about/rankings.html
  (oficial -- US News 2024, especialidade Business Analytics: #30)
- https://www.mim-essay.com/cincinnati-msba
  (TERCEIRO, nao oficial -- GPA 3.3+, GMAT 650/GRE 310, 96% placement,
  lista estendida de empregadores -- consultado mas NAO usado nos campos
  por falta de confirmacao oficial)
- https://www.campusreel.org/mba-programs/guides/university-of-cincinnati-lindner-mba-profile-acceptance-rate-gmat-scores
  (TERCEIRO, nao oficial -- acceptance rate ~21,7% -- consultado mas NAO
  usado por falta de confirmacao oficial)
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "U of Cincinnati",
    "Estado": "OH",
    "Cidade": "Cincinnati",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Ranking US News": "#30 (US News & World Report 2024, especialidade Business Analytics)",
    "Salário Médio Pós USD": 107500,  # career-outcomes.html oficial, MS-Business Analytics 2024-25: mediana=média $107.500 (min $53.000, max $157.000, ~80% resposta). Substitui 92.498 (sem fonte registrada).
    "Destaques Currículo": (
        "40 créditos (até 10 de business foundations, 17 de core "
        "analytics, 1 de capstone, 12 de eletivas); possibilidade de "
        "cursar simultaneamente 1 dos 20+ graduate certificates da "
        "Lindner para dupla credencial; projetos hands-on com parceiros "
        "corporativos (Cincinnati sedia sedes de 8 empresas Fortune 500)"
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
        raise ValueError("Nenhuma linha encontrada para U of Cincinnati (Cincinnati, OH). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (U of Cincinnati, Cincinnati, OH)")

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
