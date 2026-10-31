"""
Data enrichment: atualiza APENAS a linha 'UCF (Professional)' (Professional
MS in Management / Business Analytics track, Orlando, FL) no
msba_programs_database.xlsx.

ATENCAO: distinta da linha 'UCF' (MS in Business Analytics tradicional, ja
enriquecida anteriormente). O matching exige Nome da Universidade ==
'UCF (Professional)' E Estado == 'FL' E Cidade == 'Orlando'.

ACHADO CRITICO DA AUDITORIA: a planilha tinha 'STEM Designated': 'Sim', mas
apos pesquisa exaustiva e repetida (varias buscas + fetch direto da pagina
oficial do programa), NENHUMA fonte confirma STEM/CIP code especificamente
para esta trilha ('Professional MSM / Business Analytics'). A pagina geral
do MSM nem menciona STEM. Dado o peso critico deste campo para elegibilidade
de OPT em F-1, o valor foi revertido para 'Verificar' em vez de manter uma
afirmacao que nao consegui reconfirmar -- ISSO PRECISA DE CONFIRMACAO MANUAL
com a UCF Global (escritorio de alunos internacionais) antes de qualquer
decisao baseada nesse dado.

Fontes principais (pesquisadas manualmente, fora deste script):
- https://studentaccounts.ucf.edu/tf-profmsmgmt-ba/ (flat fee $30.000/ano, inclui
  creditos, material, laptop, orientacao, estacionamento, alumni -- confirma valor
  ja existente na planilha)
- Busca por "PMSM/Business Analytics" (formato cohort presencial, 24-32 alunos,
  aulas seg/qua noite no UCF Executive Development Center)
- Busca por admissao Business Analytics (GPA 3.0, GMAT/GRE nao obrigatorio)
- Reaproveitado do enriquecimento anterior da linha 'UCF': orcamento oficial de
  custo de vida da UCF 2026-27 (mesma universidade, mesmo campus), prorateado
  para os 20 meses estimados deste programa (vs. 12 meses do MSBA tradicional)
"""

import shutil
import sys
from datetime import datetime

import openpyxl
from openpyxl.styles import PatternFill

sys.stdout.reconfigure(encoding="utf-8")

DB_PATH = "data/msba_programs_database.xlsx"
SHEET_DATA = "MSBA Programs"

VERIFICAR = "Verificar"

MATCH = {
    "Nome da Universidade": "UCF (Professional)",
    "Estado": "FL",
    "Cidade": "Orlando",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "STEM Designated": VERIFICAR,  # CORRECAO/reversao: "Sim" nao pode ser reconfirmado, ver nota acima
    "Duração (meses)": (
        "~20 (estimado por analogia com a trilha Professional MSM/HR do "
        "mesmo formato; não confirmado oficialmente para a trilha "
        "Business Analytics)"
    ),
    "Total Créditos": 30,
    "Destaques Currículo": (
        "Formato cohort presencial (24-32 alunos), aulas seg/qua 18h-21h50 "
        "no UCF Executive Development Center (Downtown Orlando); flat fee "
        "inclui todos os créditos, apostilas, laptop, orientação, "
        "estacionamento e associação vitalícia ao UCF Alumni Association"
    ),
    "Fees Adicionais USD": (
        "Incluído na tuition (flat fee cobre créditos, material, laptop, "
        "estacionamento e orientação)"
    ),
    "Aluguel Privativo/mês USD": 1507,  # mesmo orçamento oficial UCF usado na linha "UCF" (mesma universidade/campus)
    "Mercado/mês USD": 686,
    "Transporte/mês USD": 122,
    "Outros Gastos/mês USD": 478,
    "Custo Vida Total (programa) USD": 55860,  # (1507+686+122+478) x ~20 meses deste programa
    "GPA Mínimo": 3.0,
    "GMAT/GRE Obrigatório": "Não obrigatório",
    "Matérias Pré-Requisito": (
        "Recomendado, não obrigatório: matemática, estatística, economia, "
        "ou cursos de business analytics como econometria"
    ),
}


def main():
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    backup_path = f"data/msba_programs_database.backup-{timestamp}.xlsx"
    shutil.copy2(DB_PATH, backup_path)
    print(f"Backup criado em: {backup_path}")

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
        raise ValueError("Nenhuma linha encontrada para UCF (Professional) (Orlando, FL). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (UCF (Professional), Orlando, FL)")

    verificar_fill = PatternFill(start_color="00FFF3CD", end_color="00FFF3CD", fill_type="solid")
    no_fill = PatternFill(fill_type=None)

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
