"""
Data enrichment: atualiza APENAS a linha do 'MIT' (Sloan School of
Management, Cambridge, MA) no msba_programs_database.xlsx.

Foco exclusivo: Master of Business Analytics (MBAn), presencial, STEM.
ATENCAO: distinto do MBA geral de Sloan (programa maior, com estatisticas de
admissao proprias) -- varios resultados de busca misturavam dados do MBA
com o MBAn; apenas dados confirmados como especificos do MBAn foram usados.

ACHADO CRITICO DA AUDITORIA: a planilha tinha 'TA/RA Disponivel': 'Nao
(programa 1 ano)' e 'Graduate Assistantships': 'Nao'. A pagina oficial de
tuition/financial aid do MBAn contradiz isso explicitamente: RA/TA existem,
limitados a 10h/semana. Corrigido.

Fontes principais (pesquisadas manualmente, fora deste script):
- https://mitsloan.mit.edu/master-of-business-analytics/admissions/tuition-and-financial-aid
  (tuition, RA/TA, fellowships -- confirma e corrige dados da planilha)
- Application requirements do MBAn (resume, 3 cartas, ensaio 500 palavras, taxa $150,
  GMAT/GRE opcional, sem exigencia de TOEFL/IELTS para quem estudou em ingles)
- https://sfs.mit.edu/graduate-students/cost-of-attendance/ (orcamento oficial de
  custo de vida 2026-2027, opcao de 12 meses, que bate exatamente com a duracao do MBAn)
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
    "Nome da Universidade": "MIT",
    "Estado": "MA",
    "Cidade": "Cambridge",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Tuition Total (Residente) USD": 96884,  # privada, mesma tarifa fixa para todos os admitidos
    "Aluguel Privativo/mês USD": 1900,  # orcamento oficial MIT 2026-27, opcao 12 meses (22800/12)
    "Mercado/mês USD": 870,  # 10440/12
    "Seguro Saúde (programa) USD": 5148,  # premio anual oficial 2026-27
    "Transporte/mês USD": 316,  # 3792/12
    "Outros Gastos/mês USD": 995,  # livros (1548/12=129) + pessoal (10392/12=866)
    "Custo Vida Total (programa) USD": 54120,  # soma do orcamento oficial (12 meses, exclui tuition)
    "TA/RA Disponível": (
        "Sim (RA/TA disponíveis, limitado a 10h/semana; combina salário + "
        "auxílio de tuition)"
    ),  # CORRECAO: planilha tinha "Não (programa 1 ano)", contradito pela fonte oficial
    "Graduate Assistantships": (
        "Sim (RA/TA limitados a 10h/semana; número de vagas restrito)"
    ),  # CORRECAO: planilha tinha "Não"
    "Processo TA/RA": "Sem processo formal público; contato via Sloan Student Funding Office",
    "Testes Inglês Aceitos": (
        "Não exigido para quem estudou em instituição de ensino em inglês; "
        "para os demais, proficiência avaliada na entrevista (sem TOEFL/"
        "IELTS/Duolingo obrigatório)"
    ),  # CORRECAO: planilha tinha "TOEFL/IELTS" (implicando exigencia de teste)
    "Documentos Necessários": (
        "Resume (1 página), 3 cartas de recomendação (2 acadêmicas + 1 "
        "profissional), histórico acadêmico, GMAT/GRE (opcional), ensaio de "
        "500 palavras ('por que MBAn no MIT Sloan')"
    ),
    "Taxa Aplicação USD": 150,
    "Deadline Aplicação": (
        "Round 1: ~out/2026; Round 2: ~jan/2027; Round 3: ~fev/2027 "
        "(ciclo de ingresso 2027)"
    ),
    "Rounds Aplicação": "3 rounds",
    "Entrevista": (
        "Obrigatória para todos os candidatos convidados; também usada para "
        "avaliar proficiência em inglês de não-nativos"
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
        raise ValueError("Nenhuma linha encontrada para MIT (Cambridge, MA). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (MIT, Cambridge, MA)")

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
