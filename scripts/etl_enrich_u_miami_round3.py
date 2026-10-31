"""
Data enrichment (RODADA 3): atualiza APENAS a linha 'University of Miami'
(University of Miami Business School, Coral Gables, FL) no
msba_programs_database.xlsx. NÃO confundir com 'Miami U Ohio' (Oxford, OH)
-- o match exige nome + estado + cidade.

Programa: MS in Business Analytics, presencial, STEM.

FONTES OFICIAIS (acessadas em 2026-10-19):
  - https://www.mbs.miami.edu/graduate/find-and-compare-programs/business-analytics/requirements.html
  - https://www.mbs.miami.edu/graduate/find-and-compare-programs/business-analytics/cost-of-attendance.html
  - https://www.mbs.miami.edu/graduate/graduate-how-to-apply.html
  - https://bulletin.miami.edu/graduate-academic-programs/business/specialized-master-degrees/business-analytics-ms/
    (Academic Bulletin 2026-2027)

CONFIRMADO (sem alteração): US$ 2.561/crédito x 33 = US$ 84.523 + fees
US$ 1.438 (a página diz "Academic Year 2027-2028" e chama os valores de
aproximações); GMAT/GRE opcional; TOEFL 94 (só iBT) / IELTS 7.0 /
Duolingo 125; taxa US$ 100; 2 essays; WES; entrevista por convite;
bolsas de mérito automáticas (74% do Fall 2025); 94% de empregabilidade.

CORREÇÕES:
- Documentos: o Bulletin 2026-27 diz "At least two letters of
  recommendation are required. Up to three may be submitted" (antes:
  "número não especificado").
- Deadline: o Round 1 do ciclo 2026-27 (1 out 2026) já passou; para a
  usuária (Fall 2028) o Round 1 deve cair por volta de 1 out 2027.
- Caminho Tuition Residente: universidade privada, mesma tuition para
  todos -- não existe tarifa de residente.
- Testes Inglês: admissão condicional possível com nota mínima + o
  Graduate Business English Certificate Course.
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
    "Nome da Universidade": "University of Miami",
    "Estado": "FL",
    "Cidade": "Coral Gables",
}

UPDATES = {
    "Caminho Tuition Residente": (
        "Não se aplica -- universidade privada, mesma tuition para residentes "
        "e internacionais (US$ 2.561/crédito)"
    ),
    "Testes Inglês Aceitos": (
        "TOEFL 94 (só iBT) / IELTS 7.0 / Duolingo 125. Dispensa se tiver "
        "diploma de instituição dos EUA ou de país de língua inglesa. "
        "Admissão condicional possível com nota mínima + Graduate Business "
        "English Certificate Course"
    ),
    "Deadline Aplicação": (
        "Round 1: 1 out 2026 (prioridade de bolsa); Round 2: 1 dez 2026; "
        "Round 3: 1 fev 2027; Round 4: 2 abr 2027 (prazo para "
        "internacionais); Final: 1 jun 2027. Para ingresso (Fall 2028): "
        "~1 out 2027 (Round 1, prioridade de bolsa)"
    ),
    "Documentos Necessários": (
        "Aplicação online + taxa US$ 100; statement of purpose / 2 essays "
        "(1 comum + 1 específico do programa); transcripts de todas as "
        "instituições pelo portal (avaliação WES + tradução certificada para "
        "diplomas estrangeiros); resume; 2 cartas de recomendação "
        "obrigatórias (até 3), só pelo formulário da aplicação, com e-mail "
        "profissional do recomendador; TOEFL/IELTS/DET; GMAT/GRE opcional "
        "(ajuda na bolsa); entrevista por convite"
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
        raise ValueError("Nenhuma linha encontrada para University of Miami (Coral Gables, FL). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (University of Miami, Coral Gables, FL)")

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
