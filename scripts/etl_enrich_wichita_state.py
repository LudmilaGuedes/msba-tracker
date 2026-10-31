"""
Data enrichment: atualiza APENAS a linha 'Wichita State' (W. Frank Barton
School of Business, Wichita, KS) no msba_programs_database.xlsx.

Foco exclusivo: MS in Business Analytics, presencial/híbrido, STEM.

CORRECAO PRINCIPAL: a planilha tinha 'Tuition Total (Intl)': $13.732 -- um
valor muito proximo da tarifa de RESIDENTE, nao da internacional/nao-
residente. A tarifa oficial 2026-27 confirma $903,59/credito (nao-residente)
+ $22,33/credito (taxas obrigatorias) + $76/credito (taxa do Barton School)
= ~$1.001,92/credito, resultando em ~$30.058 para 30 creditos -- bem mais
alto que o valor antigo. E possivel que o valor antigo refletisse o efeito
da Global Select Scholarship (que reduz a tarifa para quem tem GPA >= 3.3),
mas por convencao deste banco de dados o campo de tuition documenta o valor
de tabela (sticker price), com bolsas registradas separadamente em 'Tipos de
Bolsa' -- por isso o valor foi corrigido para o de tabela.

Fontes principais (pesquisadas manualmente, fora deste script):
- https://www.wichita.edu/services/tuitionfees/index.php (tarifa oficial
  2026-27 por credito, residente e nao-residente)
- busca por taxa do Barton School por credito ($76) e taxas obrigatorias
  ($22,33), somadas a tarifa base
- https://www.wichita.edu/academics/business/msba/tuition-scholarships-aid.php
  (Global Select Scholarship, deadlines de bolsa do Barton e fellowships)
- https://www.wichita.edu/academics/business/msba/FAQ_MSBA.php (duracao,
  pre-requisito de estatistica, documentos)
- busca por assistantships da WSU (GTA 20h/semana, $2.900-3.956/semestre +
  ate 100% de waiver de tuition, nao combinavel com Global Select)
- busca por custo de vida em Wichita, KS (aluguel medio de 1-quarto ~$725/
  mes, 2026)

Nao encontrado mesmo apos pesquisa profunda (mantido "Verificar"): taxa de
empregabilidade, salario medio e ranking especificos do MSBA (so encontrei
dados gerais da WSU e do mercado de trabalho de Wichita, que NAO foram
usados por nao serem especificos deste programa).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.stdout.reconfigure(encoding="utf-8")

from db_retention import rotate_backup  # noqa: E402

DB_PATH = "data/msba_programs_database.xlsx"
SHEET_DATA = "MSBA Programs"

VERIFICAR = "Verificar"

MATCH = {
    "Nome da Universidade": "Wichita State",
    "Estado": "KS",
    "Cidade": "Wichita",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Intl) USD": 1001.92,  # 903,59 (tuition) + 22,33 (taxas) + 76 (Barton)
    "Custo/Crédito (Residente) USD": 466.25,  # 367,92 + 22,33 + 76
    "Tuition Total (Intl) USD": 30058,  # CORRECAO: 30 x 1.001,92 (valor antigo de $13.732 parecia usar tarifa de residente)
    "Tuition Total (Residente) USD": 13988,  # 30 x 466,25
    "Duração (meses)": "18-24",
    "Matérias Pré-Requisito": "Curso de fundamentos de estatística (prévio ou cursado no 1º semestre, se não atendido)",
    "Documentos Necessários": (
        "Resume, transcripts, GRE/GMAT (opcional), comprovante de "
        "proficiência em inglês, certificação de suporte financeiro "
        "(candidatos internacionais)"
    ),
    "Testes Inglês Aceitos": "TOEFL iBT 79 / TOEFL PBT 550 / IELTS 6.5 / PTE 58 / Duolingo 115",
    "Deadline Aplicação": (
        "Bolsas do Barton School: 1 fev; Fellowships: 1 jul; aplicações "
        "tardias têm taxa de inscrição maior (deadline geral de admissão "
        "não encontrado)"
    ),
    "Processo TA/RA": "Vagas divulgadas via Handshake; GTA de 20h/semana",
    "Graduate Assistantships": (
        "GTA/GRA de 20h/semana: $2.900-3.956/semestre + até 100% de waiver "
        "de tuition; não pode ser combinado com a Global Select Scholarship"
    ),
    "Aluguel Privativo/mês USD": 725,  # mercado de aluguel de Wichita, KS, 2026
}


def main():
    rotate_backup()
    print("Backup da versão anterior salvo em: data/msba_programs_database_backup.xlsx")

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
        raise ValueError("Nenhuma linha encontrada para Wichita State (Wichita, KS). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Wichita State, Wichita, KS)")

    no_fill = PatternFill(fill_type=None)

    before = {}
    after = {}
    for header, new_value in UPDATES.items():
        col_idx = header_to_col[header]
        cell = ws.cell(row=row_idx, column=col_idx)
        before[header] = cell.value
        cell.value = new_value
        cell.fill = no_fill  # nenhum novo valor aqui e "Verificar"
        after[header] = new_value

    wb.save(DB_PATH)

    print("\nCampos atualizados (Como Estava -> Como Ficou):")
    for header in UPDATES:
        print(f"  - {header}:")
        print(f"      Antes: {before[header]!r}")
        print(f"      Depois: {after[header]!r}")


if __name__ == "__main__":
    main()
