"""
Data enrichment: atualiza APENAS a linha da 'WashU' (Olin Business
School, St. Louis, MO) no msba_programs_database.xlsx.

Foco exclusivo: MS in Business Analytics, formato ON-CAMPUS (presencial,
full-time), STEM, 39 créditos, 3 semestres (18 meses). Deep research em:
custo total real de tuition para internacionais, bolsas/assistantships,
requisitos de GPA/GMAT/GRE/inglês, e custo de vida em St. Louis, MO.

ATENCAO: a Olin Business School tem TAMBEM uma versao ONLINE do MS in
Business Analytics (33 creditos, formato diferente, tuition por
credito) -- multiplas fontes de busca misturam os dois programas.
Todos os dados de tuition usados aqui vieram exclusivamente da pagina
oficial "on-campus" (nao a "online"), confirmada explicitamente como
"three-semester program" full-time presencial em St. Louis.

TUITION: CORRIGIDO para o ciclo 2026-27 oficial. A pagina oficial
on-campus confirma tuition total FLAT de $101.799 ($33.933/semestre x
3), maior que o valor de $95.970 ja presente na planilha (provavelmente
de um ciclo anterior). CORRIGIDO 'Tuition Total' (Intl e Residente,
universidade PRIVADA sem distincao de residencia) para $101.799.
'Custo/Credito' (estava 'Verificar' em ambos) PREENCHIDO dividindo por
39 creditos = $2.610/credito. 'Fees Adicionais' preenchido com $4.050 =
(Program Fee $900 + Student Government Fee $100 + Health and Wellness
Fee $350) x 3 semestres -- a Health and Wellness Fee e uma taxa de
ACESSO a servicos de saude no campus, DIFERENTE do premio de seguro
saude propriamente dito (que tem campo proprio no schema e foi
pesquisado separadamente).

'Seguro Saude (programa) USD' preenchido com $2.923 -- premio oficial
2026-27 do WashU Student Health Insurance Plan (SHIP/UHC), obrigatorio
para todos os alunos internacionais (sem opcao de dispensa, ao
contrario de alunos domesticos).

'Transporte/mes' preenchido com $0 -- WashU oferece U-Pass GRATUITO a
todos os alunos, dando acesso ilimitado ao MetroBus/MetroLink
(economia estimada de ~$936/ano por aluno).

Fontes principais (pesquisadas via busca web, ago/2026):
- https://olin.washu.edu/programs/specialized-masters/ms-in-business-
  analytics/cost-aid-scholarships.php (tuition oficial on-campus
  2026-27, breakdown de fees, bolsas merit automaticas, Yellow Ribbon,
  Olin Veterans Association Fellowships)
- https://olin.washu.edu/programs/specialized-masters/ms-in-business-
  analytics/index.php (formato, 5 concentracoes -- Accounting,
  Customer, Fintech, Healthcare, Supply Chain Analytics -- curso de
  experiential learning obrigatorio)
- Busca web cruzada: GPA medio 3.4 (media geral dos specialized
  master's da Olin), GMAT/GRE recomendado mas nao exigido, rounds de
  aplicacao (R3 5 jan, R4 1 mar, R5 15 abr)
- Busca web: seguro saude $2.923/ano (SHIP oficial 2026-27); U-Pass
  gratuito para alunos WashU; aluguel de quarto privativo em St. Louis
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
    "Nome da Universidade": "WashU",
    "Estado": "MO",
    "Cidade": "St. Louis",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Intl) USD": 2610,  # $101.799 (oficial 2026-27) / 39 créditos
    "Custo/Crédito (Residente) USD": 2610,  # privada, mesma taxa do Intl
    "Tuition Total (Intl) USD": 101799,  # CORREÇÃO: planilha tinha $95.970 (ciclo anterior); oficial 2026-27
    "Tuition Total (Residente) USD": 101799,  # privada, mesma taxa do Intl
    "Fees Adicionais USD": 4050,  # (Program Fee $900 + Student Gov Fee $100 + Health/Wellness Fee $350) x 3 semestres
    "Seguro Saúde (programa) USD": 2923,  # SHIP oficial 2026-27, obrigatório para internacionais
    "Aluguel Privativo/mês USD": 650,  # quarto privativo, St. Louis MO (faixa de mercado encontrada)
    "Transporte/mês USD": 0,  # U-Pass gratuito para alunos WashU (MetroBus/MetroLink ilimitado)
    "GPA Médio Admitidos": 3.4,  # média geral dos specialized master's da Olin
    "GMAT/GRE Obrigatório": "Recomendado, mas não exigido (opcional)",
    "Tipos de Bolsa": (
        "Merit (automático, consideração para todo admitido, sem "
        "inscrição separada); bolsa para ex-alunos da WashU admitidos "
        "à Olin; Women's Leadership Scholarship; Olin Veterans "
        "Association Fellowships; participação no Yellow Ribbon "
        "Program (veteranos)"
    ),
    "Rounds Aplicação": "Round 3: 5 jan; Round 4: 1º mar; Round 5: 15 abr (intake de Fall)",
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
        raise ValueError("Nenhuma linha encontrada para WashU (St. Louis, MO). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (WashU, St. Louis, MO)")

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
