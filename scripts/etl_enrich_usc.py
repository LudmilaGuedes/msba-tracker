"""
Data enrichment: atualiza APENAS a linha 'USC' (Marshall School of
Business, Los Angeles, CA) no msba_programs_database.xlsx.

Foco exclusivo: MS in Business Analytics, presencial, STEM.

ACHADO RELEVANTE: o programa NAO EXIGE NEM ACEITA MAIS GMAT/GRE
("no longer requires or accepts the GMAT/GRE") -- um resultado isolado
mencionava "GMAT 600+" como perfil tipico, mas isso parece informacao
desatualizada de antes da mudanca de politica; a fonte que fala
explicitamente do MSBA (nao de outro programa da Marshall) confirma a
remocao do teste, entao 'GMAT Medio' permanece "Verificar" (nao se aplica).

Fontes principais (pesquisadas manualmente, fora deste script):
- https://www.marshall.usc.edu/programs/graduate-programs/specialized-masters/ms-business-analytics/admissions
  (GMAT/GRE nao exigido, TOEFL/IELTS/PTE com notas minimas, sem cartas de
  recomendacao, taxa de inscricao $155)
- busca por duracao/creditos/turma (33 unidades, 3 semestres -- Fall/Spring/
  Fall, ~189 alunos/ano)
- busca por rankings (US News 18o nacional 2024; QS 14o mundial 2025,
  atualizando o valor antigo de 16)
- Custo Vida Total calculado a partir dos campos mensais que ja estavam
  preenchidos na planilha (aluguel+mercado+seguro residencial+utilities+
  celular+wifi+transporte+outros) x 12 meses -- seguro saude permanece
  desconhecido e NAO foi incluido na soma, entao o total e uma estimativa
  parcial (subestimada) e nao o custo de vida completo
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
    "Nome da Universidade": "USC",
    "Estado": "CA",
    "Cidade": "Los Angeles",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Duração (meses)": 12,  # ritmo padrão de 3 semestres (Fall/Spring/Fall); opção de estender a 4 semestres
    "Total Créditos": 33,
    "Tamanho Turma": 189,  # média anual
    "Tuition Total (Residente) USD": 72500,  # privada, mesma taxa que intl
    "GMAT/GRE Obrigatório": "Não aceito (removido do processo de admissão do MSBA)",
    "GPA Mínimo": "Sem mínimo formal; candidatos competitivos geralmente têm GPA ≥ 3.0",
    "Taxa Aceitação %": 14,
    "Ranking US News": 18,
    "Ranking QS": 14,  # atualização (planilha tinha 16)
    "Taxa Aplicação USD": 155,  # dispensa disponível para afiliados USC, veteranos militares e outros
    "Testes Inglês Aceitos": "TOEFL 100 total (mín. 20 por seção) / IELTS 7.0 (mín. 6 por seção) / PTE aceito",
    "Documentos Necessários": "Sem GMAT/GRE nem cartas de recomendação exigidas; transcript, resume e ensaios",
    "Custo Vida Total (programa) USD": 42684,  # (2450+450+15+277+50+60+105+150) x 12 meses; exclui seguro saúde (desconhecido)
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
        raise ValueError("Nenhuma linha encontrada para USC (Los Angeles, CA). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (USC, Los Angeles, CA)")

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
