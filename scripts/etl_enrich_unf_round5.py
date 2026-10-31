"""
Data enrichment (RODADA 5): atualiza APENAS a linha 'University of North
Florida (UNF)' (Coggin College of Business, Jacksonville, FL) no
msba_programs_database.xlsx.

Programa: Master of Science in Business Analytics (MSBA), presencial,
STEM.

FONTES OFICIAIS (acessadas em 2026-10-19):
  - https://www.unf.edu/coggin/graduate/msba/index.html
  - https://www.unf.edu/coggin/graduate/msba/msba-admission-criteria.html
  - https://www.unf.edu/controller/student-financial-services/graduate-tuition.html
  - https://www.unf.edu/graduateschool/academics/program-contacts-deadlines.html
  - https://www.unf.edu/catalog/admissions/graduate/International-Students.html
  - https://www.unf.edu/intlctr/international-students/forms.html

CONFIRMADO (sem alteração): 2026-27 US$ 493,53 / US$ 1.126,88 por crédito
(fees por crédito incluídas) -> 30 créditos = US$ 14.806 / US$ 33.806;
STEM; início Fall/Spring/Summer; GPA 3.0; waiver de GMAT/GRE e Index
Score; taxa US$ 30 sem isenção para novos candidatos; Duolingo 110.

CORREÇÕES:
- Formato: a página do programa diz que as aulas começam às 18h ou mais
  tarde (noturno) e que a duração é de 12 a 20 meses.
- Fees Adicionais: US$ 3.778 não tem origem rastreável e não bate com a
  tabela 2026-27. Fora do valor por crédito, as fees fixas são Wall
  Street Journal Fee US$ 10/semestre x 3 + Osprey Card US$ 10/ano = US$ 40.
- Seguro Saúde: estimativa oficial do International Center para
  Fall + Spring 2026-27 = US$ 2.781 (era US$ 2.100). Não inclui o Summer.
- Deadline: os prazos internacionais (1 jun Fall / 1 out Spring) estão
  no catálogo da pós; texto reescrito começando pelo prazo mais cedo do
  Fall (regra do projeto).
- Documentos: internacionais enviam Affidavit of Financial Support +
  Bank Certification; avaliação course-by-course por agência aprovada.
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
    "Nome da Universidade": "University of North Florida (UNF)",
    "Estado": "FL",
    "Cidade": "Jacksonville",
}

UPDATES = {
    "Formato": (
        "Presencial, aulas noturnas (a partir das 18h), em Jacksonville; "
        "também existe versão online separada, sem concentrações"
    ),
    "Fees Adicionais USD": 40,  # WSJ Fee 10/semestre x 3 + Osprey Card 10/ano; demais fees já no valor por crédito
    "Seguro Saúde (programa) USD": 2781,  # estimativa oficial Fall + Spring 2026-27 (UNF/Aetna); não inclui Summer
    "Deadline Aplicação": (
        "1 jun (Fall, internacional); Fall doméstico: 1 ago; Spring: 1 out "
        "(internacional) / 1 dez (doméstico); Summer: 15 abr. Para a "
        "usuária (Fall 2028): ~1 jun 2028"
    ),
    "Documentos Necessários": (
        "Aplicação online (abre 1 ano antes) + taxa US$ 30; transcripts "
        "enviados direto pela instituição, com avaliação course-by-course "
        "por agência aprovada (internacionais); resume/CV; GMAT/GRE ou "
        "waiver automático; 2 cartas de recomendação profissionais + letter "
        "of intent (máx. 2 páginas) SÓ se não tiver waiver e o Index Score "
        "ficar abaixo do exigido; internacionais: Affidavit of Financial "
        "Support + Bank Certification"
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
        raise ValueError("Nenhuma linha encontrada para UNF (Jacksonville, FL). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (UNF, Jacksonville, FL)")

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
