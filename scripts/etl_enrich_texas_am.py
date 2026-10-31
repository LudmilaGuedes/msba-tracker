"""
Data enrichment: atualiza APENAS a linha 'Texas A&M' (Mays Business School,
College Station, TX) no msba_programs_database.xlsx.

ACHADO CRITICO DA AUDITORIA: o programa standalone 'MS in Analytics' (o que
ja constava nesta linha, distinto do dual-degree 'MBA + MS Analytics' que
esse sim e full-time e 100% em College Station) e na verdade:
  - Ministrado no CityCentre HOUSTON (nao no campus principal de College
    Station), com opcao de assistir ao vivo remotamente.
  - Formato HIBRIDO (presencial em Houston OU livestream), noturno
    (terca/quinta), voltado a profissionais e recem-formados.
  - 5 semestres em formato lock-step (nao pode ser acelerado nem estendido).
Isso e uma diferenca relevante do perfil "presencial full-time" que o
schema deste projeto pressupoe para estudantes F-1. A cidade foi corrigida
de 'College Station' para 'Houston'. O campo 'Formato' foi mantido como
'Presencial' (ha opcao presencial real em Houston), mas o usuario deve
avaliar se o formato hibrido/noturno atende ao perfil desejado.

Tambem corrigido: 'Ranking QS': 7 -- nenhuma fonte confirma um ranking QS
para este programa; o unico ranking encontrado e do US News, especifico
para a categoria de programas ONLINE de negocios (nao QS). O valor antigo
foi revertido para "Verificar" por falta de confirmacao.

Fontes principais (pesquisadas manualmente, fora deste script):
- https://mays.tamu.edu/graduate/masters/ms-programs/analytics/ (formato, local,
  publico-alvo, estrutura de 5 semestres, custo total $70.000 -- ja incluindo
  taxas, software, livros e orientacao)
- Busca por rankings (US News: top-10/top-16 nacional entre programas ONLINE de
  business; nao ha ranking QS confirmado)
- Busca por admissao (GMAT/GRE opcional com dispensa, deadlines por round)
- Busca por custo de vida em Houston (aluguel medio de 1-quarto ~$1.199/mes, 2026)
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
    "Nome da Universidade": "Texas A&M",
    "Estado": "TX",
    "Cidade": "College Station",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Cidade": "Houston",  # CORRECAO CRITICA: programa e ministrado no CityCentre Houston, nao em College Station
    "Ranking QS": VERIFICAR,  # CORRECAO: valor antigo (7) sem fonte confirmavel; nao existe ranking QS publicado
    "Destaques Currículo": (
        "ATENÇÃO: formato híbrido -- aulas presenciais no CityCentre "
        "Houston às terças/quintas à noite, com opção de assistir ao vivo "
        "remotamente; NÃO é ministrado no campus de College Station apesar "
        "do nome da universidade. 5 semestres em formato lock-step (não "
        "pode ser acelerado nem estendido). Voltado a profissionais e "
        "recém-formados. Ranqueado entre os top 10-16 nacionais pelo US "
        "News na categoria de programas online de negócios (não há ranking "
        "QS confirmado para este programa)"
    ),
    "Duração (meses)": "~20 (5 semestres em formato lock-step; não pode ser acelerado nem estendido)",
    "Tuition Total (Intl) USD": 70000,  # atualização: valor mais recente, já all-inclusive
    "Tuition Total (Residente) USD": 70000,
    "Fees Adicionais USD": (
        "Incluído no total (tuition, taxas, licenças de software, livros e "
        "orientação já estão no valor de $70.000)"
    ),
    "GMAT/GRE Obrigatório": "Opcional (dispensa disponível; se enviado, validade de 5 anos)",
    "Deadline Aplicação": (
        "Abertura: 1 set; Round 1: 3 dez; Round 2 (consideração para "
        "bolsa): 1 abr; Final: 25 jun"
    ),
    "Rounds Aplicação": "Múltiplos rounds: R1 3 dez, R2 (bolsa) 1 abr, Final 25 jun",
    "Documentos Necessários": "Resume, histórico acadêmico (não oficial na inscrição), entrevista avaliativa",
    "Entrevista": "Sim (entrevista avaliativa faz parte do processo seletivo)",
    "Aluguel Privativo/mês USD": 1199,  # mercado de aluguel de Houston (nao College Station), 2026
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
        raise ValueError("Nenhuma linha encontrada para Texas A&M (College Station, TX). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Texas A&M, College Station -> Houston, TX)")

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
