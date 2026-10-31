"""
Data enrichment (round 2): atualiza a linha da 'Illinois Tech' (Stuart
School of Business) no msba_programs_database.xlsx, a pedido do
usuário: confirmar tuition, bolsas (e quão realista é pra
internacional), empregabilidade/salário, taxa de aplicação e
requisitos.

FONTES (fordham.edu-style: iit.edu, oficiais, acessadas via
WebFetch/WebSearch em 2026-09-08):
- https://www.iit.edu/academics/programs/business-analytics-ms
  (página do programa -- requisitos de admissão)
- https://www.iit.edu/student-accounting/tuition-and-fees/current-tuition-and-fees/current-stuart-school-business
  (tuition e fees oficiais)
- https://www.iit.edu/stuart/admission-aid/scholarships-and-aid
  (bolsas)
- https://elevate.iit.edu/graduates-outcomes/ e
  https://careeroutcomes.iit.edu/graduates/ (outcomes -- ver nota
  abaixo)

CONFIRMADO SEM MUDANÇA (já estava correto na planilha):
- Custo/Crédito: $1.925/crédito -- CONFIRMADO oficial, 2026-27.
- Tuition Total: 30 créditos x $1.925 = $57.750 -- bate exato.
- Seguro Saúde: $2.510/ano -- CONFIRMADO oficial (billed no Fall,
  com opção de waiver e pricing parcial).
- Fees Adicionais ($2.775): CONFIRMADO ser a soma de Activity Fee
  ($125/semestre, full-time) + Student Service Fee ($800/semestre,
  coorte "entrando a partir de Summer 2023", full-time) = $925/
  semestre x 3 semestres (programa de 14 meses = Fall+Spring+Summer
  ou equivalente) = $2.775. Bate exatamente com o valor já na
  planilha -- fórmula documentada aqui pela primeira vez.
- GPA Mínimo (3.0): CONFIRMADO oficial.
- GMAT/GRE (não exigido): CONFIRMADO oficial ("No GMAT or GRE scores
  required").
- Taxa de Aplicação ($0): CONFIRMADO oficial -- Illinois Tech não
  cobra taxa de aplicação para nenhum programa de mestrado do Stuart,
  incluindo o MSBA (política permanente da escola, não vinculada a
  round -- já documentado no lote anterior).
- Documentos Necessários (transcript oficial, CV, professional
  statement, prova de inglês se aplicável): CONFIRMADO, sem
  documentos adicionais encontrados além dos já listados.

ENRIQUECIDO (detalhe novo sobre bolsas, antes só "Institutional
scholarships"):
- Bolsas: TODOS os alunos de pós-graduação do Stuart são
  AUTOMATICAMENTE avaliados para bolsa de mérito no momento da
  admissão (sem aplicação separada) -- valores são "tiered" (em
  faixas) conforme o histórico acadêmico, mas NENHUM valor ou faixa
  em dólares é publicado oficialmente (diferente da Fordham, que
  publica $5k-$30k). Internacionais também podem ser considerados
  para ajuda baseada em necessidade financeira durante o processo de
  admissão (não é comum -- a maioria das escolas só oferece
  need-based aid via FAFSA, que exclui internacionais). Uma
  estatística geral da pós-graduação do Stuart (não específica do
  MSBA) indica que cerca de 30-40% dos alunos de pós-graduação
  recebem bolsa ou assistantship por ano, tipicamente cobrindo
  tuition de 1-2 semestres -- CITADO com essa ressalva de escopo (não
  é garantido ser a mesma taxa para o MSBA especificamente).

NÃO CONFIRMADO (fontes não divulgam dado específico do MSBA --
'Taxa Empregabilidade %' e 'Salário Médio Pós USD' permanecem
'Verificar'): a página de outcomes do Illinois Tech
(careeroutcomes.iit.edu / elevate.iit.edu) é um DASHBOARD
interativo com filtro por escola/programa, sem dados estáticos
extraíveis via fetch simples (a URL de outcomes redireciona para uma
plataforma de terceiros que exige JS/login). Uma pesquisa geral
encontrou apenas estatísticas AGREGADAS do Stuart School inteiro (ex:
"95% grau de emprego de pós-graduação", salário médio ~$109.596 para
uma turma de referência) -- mas essas cifras não são atribuíveis
especificamente ao MSBA (podem incluir MBA e outros mestrados de
maior remuneração, como Finance), então, seguindo o mesmo princípio
já aplicado no Baruch/Fordham (não misturar dado de programa
diferente), NÃO foram escritas nesses dois campos. Também notei que a
própria página do Illinois Tech avisa: "Stuart School of Business
Data for 2025 will be available near the end of April" -- ou seja,
mesmo os dados agregados mais recentes ainda não estavam publicados
no momento desta pesquisa.
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
    "Nome da Universidade": "Illinois Tech",
    "Estado": "IL",
    "Cidade": "Chicago",
}

UPDATES = {
    "Tipos de Bolsa": (
        "Bolsa de mérito com avaliação automática de TODOS os "
        "candidatos no momento da admissão (sem aplicação separada); "
        "valores em faixas ('tiered') conforme histórico acadêmico, "
        "SEM valor ou faixa em dólares publicado oficialmente. "
        "Internacionais também podem ser considerados p/ ajuda "
        "baseada em necessidade financeira durante a admissão (raro "
        "entre as escolas do banco). Estatística geral (não "
        "específica do MSBA) do Stuart School: ~30-40% dos alunos de "
        "pós-graduação recebem bolsa/assistantship por ano, "
        "tipicamente cobrindo 1-2 semestres de tuition."
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
        raise ValueError("Nenhuma linha encontrada para Illinois Tech. Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Illinois Tech)")

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
