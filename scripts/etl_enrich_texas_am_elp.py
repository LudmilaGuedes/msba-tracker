# -*- coding: utf-8 -*-
"""
Acrescenta ao `Testes Inglês Aceitos` do Texas A&M (MS in Analytics, Houston, TX)
a via de dispensa por GRE/GMAT da Graduate & Professional School: quem tira
GRE Verbal >= 146 ou GMAT Verbal >= 22 (GMAT Focus Verbal >= 77) fica isento
de TOEFL/IELTS (catalog.tamu.edu/graduate/admission/ e
grad.tamu.edu/knowledge-center/general-faqs/elp-faq). O Duolingo continua
nao aceito (so valeu no Spring 2022).

Uso: python scripts/etl_enrich_texas_am_elp.py (a partir da raiz).
"""

import sys

import openpyxl

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

TARGET = ("Texas A&M", "TX", "Houston")
COL = "Testes Inglês Aceitos"
NEW = (
    "TOEFL iBT mín. 80 (antes de 21 jan 2026) ou 4.5 na nova escala com mín. 4.0 por seção / "
    "IELTS Academic mín. 6.0 (mínimos da Graduate School da Texas A&M; programas podem exigir mais); "
    "isenção de TOEFL/IELTS com GRE Verbal mín. 146 ou GMAT Verbal mín. 22 (GMAT Focus Verbal mín. 77) "
    "— regra da Graduate School, confirmar com o programa; Duolingo não aceito atualmente "
    "(só foi aceito temporariamente no Spring 2022)"
)


def main():
    rotate_backup()
    wb = openpyxl.load_workbook(DB_PATH)
    ws = wb["MSBA Programs"]
    col = {c.value: i + 1 for i, c in enumerate(ws[1])}
    rows = [
        r[0].row for r in ws.iter_rows(min_row=2)
        if (r[col["Nome da Universidade"] - 1].value, r[col["Estado"] - 1].value, r[col["Cidade"] - 1].value) == TARGET
    ]
    if len(rows) != 1:
        raise ValueError(f"Esperava 1 linha para {TARGET}, encontrou {len(rows)}. Abortando.")
    cell = ws.cell(row=rows[0], column=col[COL])
    print(f"{COL}: {cell.value!r}\n  -> {NEW!r}")
    cell.value = NEW
    wb.save(DB_PATH)


if __name__ == "__main__":
    main()
