"""
Data enrichment: atualiza APENAS a linha da 'Illinois Tech' (Stuart
School of Business, Chicago, IL) no msba_programs_database.xlsx.

Foco exclusivo: MS in Business Analytics, presencial, STEM, 30 créditos
(10 disciplinas de 3 créditos cada). Deep research em: custo total real
de tuition para internacionais, bolsas/assistantships, requisitos de
GPA/GMAT/GRE/inglês, e custo de vida em Chicago, IL.

ACHADO DA AUDITORIA: 'Duração (meses)' estava 24 -- a página oficial do
programa confirma que essa é a duração do formato PART-TIME (2 anos/6
semestres); o formato FULL-TIME (presencial, o foco deste enrichment) é
de 12-16 meses. CORRIGIDO para 14 (ponto médio da faixa oficial de
tempo integral).

TUITION: 'Total Créditos' preenchido com 30 (10 disciplinas x 3
créditos, confirmado oficialmente). 'Custo/Crédito' preenchido com
$1.925 -- tarifa oficial 2026-27 do Stuart School (mais atual que o
valor de $1.851 de uma busca referente ao ciclo 2025-26). Como é
universidade PRIVADA, sem distinção de residência mencionada em nenhuma
fonte, 'Custo/Crédito (Residente)' e 'Tuition Total (Residente)' foram
preenchidos com os MESMOS valores do Intl. 'Tuition Total' = 30 x
$1.925 = $57.750. 'Fees Adicionais' preenchido com $2.775 = (Activity
Fee $125 + Student Service Fee $800) x 3 semestres (estimativa de
tempo integral) -- a UPass Fee ($155/semestre, transporte) foi excluída
por ser opcional (opt-in), e tratada separadamente em 'Transporte/mês'.
'Seguro Saúde (programa)' preenchido com $2.510/ano (oficial, cobrado
no semestre de Fall).

'Ranking US News' e 'Ranking QS' permanecem 'Verificar': o único
ranking encontrado (#24 nos EUA, TFE Times 2025) é do programa **M.S.
in Marketing Analytics** -- um programa IRMÃO mas DIFERENTE do MS in
Business Analytics (nosso alvo) na mesma escola. Não usado aqui para
evitar atribuir o ranking de um programa diferente a esta linha.

Custo de vida: 'Aluguel Privativo/mês' preenchido com $750 -- faixa
específica de moradia perto do campus da IIT ($550-975/mês),
consideravelmente mais barata que a média geral de Chicago.
'Transporte/mês' preenchido com $31 -- Ventra U-Pass (opt-in, oferece
CTA ilimitado), $155/semestre ÷ 5 meses.

Fontes principais (pesquisadas via busca web, ago/2026):
- https://www.iit.edu/academics/programs/business-analytics-ms
  (duração 12-16 meses full-time / 2 anos part-time, 30 créditos,
  formato, currículo, GPA mínimo 3.0, GMAT/GRE não exigido, documentos)
- https://www.iit.edu/student-accounting/tuition-and-fees/
  current-tuition-and-fees/current-stuart-school-business (tuition
  oficial 2026-27, breakdown completo de fees por semestre, seguro
  saúde, UPass)
- Busca web: aluguel perto do campus da IIT, Ventra U-Pass para alunos
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

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Duração (meses)": 14,  # CORREÇÃO: planilha tinha 24 (duração do formato part-time); full-time oficial é 12-16 meses
    "Total Créditos": 30,  # 10 disciplinas x 3 créditos (oficial)
    "Custo/Crédito (Intl) USD": 1925,  # oficial 2026-27
    "Custo/Crédito (Residente) USD": 1925,  # privada, mesma taxa do Intl
    "Tuition Total (Intl) USD": 57750,  # 30 créditos x $1.925
    "Tuition Total (Residente) USD": 57750,  # privada, mesma taxa do Intl
    "Fees Adicionais USD": 2775,  # (Activity Fee $125 + Student Service Fee $800) x 3 semestres (tempo integral)
    "Seguro Saúde (programa) USD": 2510,  # oficial, cobrado no semestre de Fall
    "Aluguel Privativo/mês USD": 750,  # faixa de moradia perto do campus da IIT ($550-975/mês)
    "Transporte/mês USD": 31,  # Ventra U-Pass (opt-in), $155/semestre ÷ 5 meses
    "GPA Mínimo": 3.0,
    "GMAT/GRE Obrigatório": "Não exigido",
    "Matérias Pré-Requisito": "Nenhum pré-requisito formal listado",
    "Documentos Necessários": (
        "Histórico acadêmico oficial, CV/resume, professional "
        "statement (statement of purpose), teste de inglês "
        "(TOEFL/IELTS, se aplicável)"
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
        raise ValueError("Nenhuma linha encontrada para Illinois Tech (Chicago, IL). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Illinois Tech, Chicago, IL)")

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
