"""
Data enrichment: atualiza APENAS a linha 'UCLA' (Anderson School of
Management, Los Angeles, CA) no msba_programs_database.xlsx.

Foco exclusivo: Master of Science in Business Analytics, presencial, STEM.

Confirmado (nao alterado): 'Ranking QS': 1 ja estava correto -- UCLA divide
o 1o lugar mundial com MIT Sloan no QS 2026 Business Analytics Masters
Ranking (100/100 pontos ambos), confirmado por multiplas fontes
independentes. Diferente do caso anterior de Texas A&M, aqui o numero
antigo se confirmou verdadeiro.

Fontes principais (pesquisadas manualmente, fora deste script):
- Multiplas fontes confirmam tuition de $75.445 para a Classe de 2027 (um
  unico resultado isolado citava $93.214,42, que parece ser custo total de
  atendimento incluindo vida, nao apenas tuition -- nao foi usado por falta
  de corroboracao)
- Programa e "self-support" (ja documentado em 'Caminho Tuition Residente'),
  entao a tuition residente e igual a internacional
- busca por admissao (GMAT/GRE obrigatorio, percentil 91-93 medio, 2+ anos
  de experiencia EXIGIDOS -- nao apenas tipico como a planilha registrava --,
  3 rounds com datas, taxa $200 ja confirmada, entrevista por convite)
- Custo Vida Total calculado diretamente dos campos mensais que ja estavam
  preenchidos na planilha (aluguel+mercado+seguro residencial+utilities+
  celular+wifi+transporte+outros) x 15 meses de programa, mais o seguro
  saude (que ja e um valor total do programa)
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
    "Nome da Universidade": "UCLA",
    "Estado": "CA",
    "Cidade": "Los Angeles",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Tuition Total (Residente) USD": 75445,  # programa "self-support": mesma taxa para todos
    "GMAT/GRE Obrigatório": (
        "Obrigatório (GMAT ou GRE); média da Classe de 2026 no percentil "
        "91-93; perfil competitivo: GMAT 710+ / GRE Quant 167+"
    ),
    "Experiência Profissional": (
        "Mínimo 2 anos de experiência profissional full-time (exigido, "
        "não apenas típico)"
    ),  # CORRECAO: planilha tinha "~2 anos típico", mas e um requisito, nao so uma media
    "Deadline Aplicação": (
        "Round 1: 9 jan; Round 2: 6 mar; Round 3: 8 mai (datas do último "
        "ciclo encontrado; confirmar cronograma atualizado no site oficial)"
    ),
    "Rounds Aplicação": "3 rounds",
    "Documentos Necessários": "Transcripts oficiais, 3 ensaios",
    "Entrevista": "Por convite apenas",
    "GPA Mínimo": "Sem mínimo formal publicado",
    "Custo Vida Total (programa) USD": 60237,  # (2450+450+15+277+50+60+105+150) x 15 meses + 6882 seguro saúde
    "Ranking US News": 20,  # ranking geral do Anderson School of Management (não específico do MSBA)
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
        raise ValueError("Nenhuma linha encontrada para UCLA (Los Angeles, CA). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (UCLA, Los Angeles, CA)")

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
