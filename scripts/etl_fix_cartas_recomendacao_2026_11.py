"""
Correção de 'Cartas de Recomendação' após a verificação de 2026-11-01 dos
programas marcados como "0 cartas" (conferência contra as páginas oficiais).

Match por Nome + Estado + Cidade. Rodar da raiz do repositório:
    python scripts/etl_fix_cartas_recomendacao_2026_11.py [--dry-run]
"""

import sys

import openpyxl

from db_retention import DB_PATH, rotate_backup

SHEET = "MSBA Programs"
FIELD = "Cartas de Recomendação"

FIXES = [
    # iit.edu/admissions-aid/graduate-admission/.../admission-requirements-and-checklist
    (("Illinois Tech", "IL", "Chicago"),
     "1 (regra geral da Illinois Tech: 'in most cases one letter of recommendation is "
     "required for master's degree programs'; a página do MSBA não especifica)"),
    # business.wisc.edu/graduate/masters/business-analytics/admissions
    (("UW-Madison", "WI", "Madison"),
     "0 (1 opcional: 'You may submit a letter of recommendation... it is not required')"),
    # cob.unt.edu/masters/admission
    (("University of North Texas (UNT)", "TX", "Denton"),
     "0 cartas (referências profissionais/acadêmicas só com contato, via GradCAS; "
     "número não informado)"),
    # hofstra.edu/academics/colleges/zarb/domestic-what-i-need-to-apply.html
    (("Hofstra University", "NY", "Hempstead"),
     "0 (2 opcionais, enviadas pelo recomendador no sistema online)"),
    # engineering.wayne.edu/data-science-program/admissions/
    (("Wayne State University", "MI", "Detroit"),
     "0 (1 suplementar opcional; recomendada se GPA < 3.0)"),
]


def main() -> None:
    dry = "--dry-run" in sys.argv
    if not dry:
        rotate_backup()
    wb = openpyxl.load_workbook(DB_PATH)
    ws = wb[SHEET]
    header = [c.value for c in ws[1]]
    col = {name: i + 1 for i, name in enumerate(header)}

    for (nome, estado, cidade), value in FIXES:
        rows = [
            r for r in range(2, ws.max_row + 1)
            if ws.cell(r, col["Nome da Universidade"]).value == nome
            and ws.cell(r, col["Estado"]).value == estado
            and ws.cell(r, col["Cidade"]).value == cidade
        ]
        if len(rows) != 1:
            raise SystemExit(f"{nome}: esperava 1 linha, achei {len(rows)}")
        cell = ws.cell(rows[0], col[FIELD])
        print(f"[{nome}]\n  ANTES: {cell.value}\n  DEPOIS: {value}\n")
        cell.value = value

    if dry:
        print("dry-run: nada gravado")
        return
    wb.save(DB_PATH)
    print(f"gravado em {DB_PATH}")


if __name__ == "__main__":
    main()
