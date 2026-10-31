"""
Data enrichment (round 2): atualiza a linha da 'Florida Atlantic
University (FAU)' (College of Business, MSBA Traditional/presencial)
no msba_programs_database.xlsx, a pedido do usuário: confirmar
tuition, bolsas, taxa de aplicação, requisitos e
empregabilidade/salário.

IMPORTANTE: FAU tem DOIS programas de Business Analytics -- o
"Traditional" (presencial, é o que está na planilha) e o "Professional
Master's" (online). São tuitions e páginas DIFERENTES. Toda a pesquisa
abaixo foi filtrada para o Traditional; descartei explicitamente
números do programa online (ex: "$19.456/ano" e "~$1.000/crédito" que
apareceram numa busca inicial eram do programa ONLINE, não do
presencial).

FONTES (oficiais, fau.edu/business.fau.edu, via WebFetch/WebSearch em
2026-09-08):
- https://www.fau.edu/controllers-office/student-services/tuition-breakdown/
- https://business.fau.edu/masters-phd/business-analytics/traditional/admissions/
- https://business.fau.edu/masters-phd/business-analytics/traditional/career-outlook
- https://business.fau.edu/departments/information-technology-operations-management/our-graduates/where-are-they-now/

CONFIRMADO SEM MUDANÇA (bateu exato com a fonte oficial):
- Custo/Crédito Residente ($371,82) e Não-residente/Intl ($1.125,06):
  CONFIRMADO oficial para o ano letivo 2026-2027 (tabela geral de
  pós-graduação da FAU, que se aplica ao MSBA -- não há tarifa
  diferenciada por programa dentro da mesma categoria de mestrado).
- Tuition Total: 30 x $1.125,06 = $33.752 (Intl) e 30 x $371,82 =
  $11.155 (Residente) -- batem exatos.
- Taxa de Aplicação ($30): CONFIRMADO oficial.
- GRE/GMAT: CONFIRMADO -- combinado (V+Q) mín. 295 no GRE OU mín. 495
  no GMAT, válido até 5 anos; dispensa se diploma de pós-graduação
  prévio OU GPA de admissão >= 3.3. Já estava certo.
- GPA Mínimo (3.0) e prova de inglês (TOEFL 79 iBT / IELTS 6.0):
  CONFIRMADO oficial, sem mudança.
- Documentos: página oficial NÃO menciona carta de recomendação,
  ensaio ou resume como exigidos -- consistente com o que já estava
  na planilha (nenhum documento faltando identificado).

ENRIQUECIDO (bolsas):
- Achei um mecanismo específico não documentado antes: o
  departamento ITOM (que administra o MSBA) às vezes concede uma
  "Provost Fellowship" UMA VEZ POR ANO para o MELHOR aplicante do
  MSBA daquele ciclo -- extremamente competitiva (é uma bolsa, não
  um processo aberto a todos). Adicionado a 'Tipos de Bolsa'.
- Confirmado: internacionais só são elegíveis a bolsas de MÉRITO
  (sem exigência de "necessidade financeira", que é exclusiva para
  cidadãos/residentes via FAFSA) -- consistente com o padrão de GTA
  já documentado.

NÃO ENCONTRADO -- 'Taxa Empregabilidade %', 'Salário Médio Pós USD' e
'Principais Empregadores' permanecem 'Verificar': pesquisei a fundo,
incluindo a página oficial "Career Outlook" do PRÓPRIO programa
Traditional MSBA -- ela só tem estatísticas NACIONAIS genéricas (ex:
"$72.560 salário médio nacional para mestrado em Business Analytics",
citando Payscale -- isso NÃO é um dado da FAU, é uma referência de
mercado que a página cita, mesmo erro que teria sido cometido se eu
tivesse usado esse número como se fosse da FAU). A página "Where are
they now" do departamento ITOM tem uma lista de empregadores reais,
mas é EXCLUSIVA de graduados de MIS (graduação) e MSITM (MS em IT
Management, programa diferente) -- nenhum registro de graduados do
MSBA especificamente. Não escrevi nenhum desses números/nomes na
linha do MSBA para não repetir o erro de atribuir dado de programa
errado (mesmo cuidado já aplicado em Baylor/CU Boulder/Baruch/
Fordham).
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
    "Nome da Universidade": "Florida Atlantic University (FAU)",
    "Estado": "FL",
    "Cidade": "Boca Raton",
}

UPDATES = {
    "Testes Inglês Aceitos": (
        "TOEFL 79 (iBT) / 548 (Paper-Based) -- IELTS 6.0 -- Duolingo "
        "110 (CORRIGIDO: página do programa não mencionava, mas a "
        "página oficial de admissões do College of Business confirma "
        "que é aceito) -- alternativa: Level VI (Academic Track) do "
        "Intensive English Institute da própria FAU. PTE NÃO é aceito."
    ),
    "Tipos de Bolsa": (
        "Fellowships do Graduate College (anuais); bolsas via Office "
        "of Student Financial Aid; Graduate Teaching Assistantships "
        "(GTA) no departamento ITOM; 'Provost Fellowship' concedida "
        "PELO ITOM uma vez por ano ao MELHOR candidato do MSBA daquele "
        "ciclo (extremamente competitiva, não é processo aberto a "
        "todos). Internacionais só são elegíveis a bolsas de MÉRITO "
        "(sem exigência de necessidade financeira, exclusiva de "
        "cidadãos/residentes via FAFSA)."
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
        raise ValueError("Nenhuma linha encontrada para FAU. Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (FAU)")

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
