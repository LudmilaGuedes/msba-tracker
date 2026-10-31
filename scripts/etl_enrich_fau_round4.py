"""
Data enrichment (RODADA 4): atualiza APENAS a linha 'Florida Atlantic
University (FAU)' (College of Business, Boca Raton, FL) no
msba_programs_database.xlsx.

Programa: Master of Science in Business Analytics (Traditional),
presencial, STEM.

FONTES OFICIAIS (acessadas em 2026-10-19):
  - https://business.fau.edu/masters-phd/business-analytics/traditional/admissions/
  - https://business.fau.edu/masters-phd/business-analytics/traditional/career-outlook
  - https://www.fau.edu/shs/insurance/international
  - https://www.fau.edu/graduate/resources/graduate-assistantship-information/faq-students/index.php
  - https://www.fau.edu/graduate/documents/gpc/10-29-2024/gc-ay-2025-2026-gta-gra-offer-letter.pdf
    (modelo oficial de carta de oferta de GTA/GRA da Graduate College, 2025-26)

CONFIRMADO (sem alteração): GPA ~3.0 nos últimos 60 créditos; GMAT 495 /
GRE 295 (dispensa com mestrado ou pedido com GPA >= 3.3); TOEFL 79 /
IELTS 6.0; prazos intl 15 fev / 15 jul / 15 jan; taxa US$ 30; nenhuma
carta de recomendação; admissão condicional; salário de US$ 72.560 da
página é média do Payscale (não usado); seguro obrigatório, preço não
publicado.

CORREÇÕES:
- Assistantships: a planilha dizia "waiver total é raro, geralmente
  parcial". O modelo de carta de oferta da Graduate College diz que
  GTA/GRA recebe waiver de tuition de até 27 créditos por ano, e que
  "International and out-of-state students who work at least 0.25 FTE
  (10 hours per week) will also have the non-resident fee waived". Ou
  seja: quem consegue a vaga (>= 10 h/semana) não paga tuition nem a
  non-resident fee -- só as fees obrigatórias. Com 0.5 FTE (20 h), a FAU
  paga 85% do seguro saúde. O difícil é conseguir a vaga, não o tamanho
  do waiver. Stipend definido pelo departamento (sem valor publicado);
  internacionais trabalham no máximo 20 h/semana.
- Caminho Tuition Residente: acrescentado esse caminho via assistantship.
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
    "Quão Realista (Intl)": (
        "Baixa a moderada -- GTA/GRA é competitiva e as vagas são decididas "
        "pelo departamento (ITOM), mas quem consegue (>= 10 h/semana) tem "
        "tuition e non-resident fee isentas; bolsas de mérito sem valores "
        "publicados; Provost Fellowship só para o melhor candidato do ciclo"
    ),
    "Processo TA/RA": (
        "Sem aplicação central: contatar o programa/departamento (ITOM) e o "
        "Graduate Advising do College of Business (cobgradadvising@fau.edu); "
        "algumas vagas também aparecem no Handshake e no site de RH da FAU. "
        "Renovação depende de desempenho e de verba"
    ),
    "Caminho Tuition Residente": (
        "F-1 não pode ser residente na Flórida, mas GTA/GRA de pelo menos "
        "0.25 FTE (10 h/semana) isenta a non-resident fee (US$ 717,37/crédito) "
        "e a tuition (até 27 créditos/ano); sem assistantship, internacional "
        "paga a tarifa de não-residente"
    ),
    "Graduate Assistantships": (
        "GTA/GRA: stipend definido pelo departamento (sem valor publicado); "
        "waiver de tuition de até 27 créditos/ano + non-resident fee isenta "
        "para internacionais com >= 0.25 FTE (10 h/semana); fees "
        "obrigatórias não cobertas; com 0.5 FTE (20 h/semana, o máximo para "
        "internacionais) a FAU paga 85% do seguro saúde"
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
        raise ValueError("Nenhuma linha encontrada para FAU (Boca Raton, FL). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (FAU, Boca Raton, FL)")

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
