"""
Data enrichment: atualiza APENAS a linha 'Texas Tech' (Rawls College of
Business, Lubbock, TX) no msba_programs_database.xlsx.

Foco exclusivo: MS in Data Science, presencial, STEM. Confirmado por
pesquisa que o Rawls College NAO tem um programa com o titulo exato
"MS in Business Analytics" -- o MS in Data Science (36 creditos, STEM,
12 meses) e a melhor e mais proxima correspondencia, o que valida a
escolha ja feita nesta linha.

ACHADO DA AUDITORIA: a planilha afirmava categoricamente 'Bolsas
Disponiveis': 'Nao', 'TA/RA Disponivel': 'Nao', 'Graduate Assistantships':
'Nao oferecido'. A pesquisa encontrou evidencia de que o Rawls College
OFERECE assistantships institucionalmente (~$22.750/ano, meio periodo,
confere status in-state e waiver de tuition out-of-state), o que
contradiz a certeza absoluta desses campos. Como nao foi possivel
confirmar se o MSDS especificamente participa desse programa, os campos
foram revertidos para "Verificar" com uma nota explicativa, em vez de
manter uma negativa categorica que uma fonte independente contradiz.

Fontes principais (pesquisadas manualmente, fora deste script):
- https://www.depts.ttu.edu/rawlsbusiness/graduate/ms/datascience/ (confirma
  que nao ha MS in Business Analytics separado; MSDS e a opcao correta)
- busca por admissao (TOEFL 79 iBT / IELTS 6.5 / PTE 60, GMAT/GRE opcional,
  taxa de inscricao $75, documentos)
- busca por assistantships do Rawls College (institucional, nao confirmado
  para o MSDS especificamente)
- busca por custo de vida em Lubbock, TX (aluguel medio de apartamento
  estudantil ~$720/mes, frequentemente com utilities inclusas, 2026)
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
    "Nome da Universidade": "Texas Tech",
    "Estado": "TX",
    "Cidade": "Lubbock",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Bolsas Disponíveis": VERIFICAR,  # CORRECAO/reversao: "Não" contradito por evidência institucional do Rawls College
    "Tipos de Bolsa": (
        "Rawls College oferece GA/RA institucionalmente (~$22.750/ano, "
        "meio período); não confirmado se o MSDS especificamente participa"
    ),
    "Quão Realista (Intl)": VERIFICAR,
    "TA/RA Disponível": VERIFICAR,
    "Processo TA/RA": VERIFICAR,
    "Graduate Assistantships": (
        "Rawls College oferece GA/RA institucionalmente (~$22.750/ano, meio "
        "período, waiver de tuition out-of-state e status in-state); não "
        "confirmado se disponível especificamente para o MSDS"
    ),
    "Testes Inglês Aceitos": "TOEFL iBT 79 / TOEFL Essentials 8.5 / IELTS 6.5 / PTE 60",
    "GMAT/GRE Obrigatório": "Opcional (não obrigatório, mas pode ser enviado)",
    "Taxa Aplicação USD": 75,
    "Documentos Necessários": "Application online, transcripts não oficiais, resume/CV",
    "Deadline Aplicação": "Summer: 1 mai (confirmar deadlines de Fall/Spring diretamente com o programa)",
    "Aluguel Privativo/mês USD": 720,  # mercado de aluguel estudantil de Lubbock, TX, 2026
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
        raise ValueError("Nenhuma linha encontrada para Texas Tech (Lubbock, TX). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Texas Tech, Lubbock, TX)")

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
