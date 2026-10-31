"""
Data enrichment (round 2): atualiza APENAS a linha da 'Baruch CUNY'
(Zicklin School of Business, New York, NY) no msba_programs_database.xlsx.

Fonte: resposta por e-mail recebida diretamente do Office of Graduate
Admissions da Zicklin School of Business (Kanishka Kapoor), em resposta
a contato direto da aluna, em 2026-09-08. Fonte primária de alta
confiabilidade (não é research de site público, é confirmação direta
do programa) -- sobrepõe estimativas anteriores quando há conflito.

MUDANÇAS:

1. Tuition/crédito ($855, categoria "Masters Degree" != $1.110 da MBA):
   e-mail CONFIRMA explicitamente o valor de $855/crédito já corrigido
   no round 1 (etl_enrich_baruch_cuny.py), incluindo a mesma distinção
   MBA vs. MSBA. Nenhuma mudança necessária nesse campo.

2. Total de créditos: e-mail informa "approximately 30-33 credits,
   depending on course waivers" -- DIVERGE do valor anterior na planilha
   (36, de origem não confirmada). Fonte oficial prevalece. Atualizado
   'Total Créditos' para '30-33' (formato de faixa, mesmo padrão já
   usado em outras linhas do schema, ex. '111-147').
   Consequentemente, 'Tuition Total (Intl)' e 'Tuition Total
   (Residente)' foram recalculados como faixa (30-33 créditos x
   $855 / $470), substituindo os valores fixos baseados em 36 créditos.

3. Graduate assistantships: e-mail confirma que existem, são pagas por
   hora, SEM isenção de tuition (diferente do modelo comum de GA/TA em
   outras escolas, que costuma incluir waiver parcial/total). Preenchido
   'Graduate Assistantships' e 'TA/RA Disponível'.

4. Bolsas para internacionais: e-mail confirma consideração automática
   na admissão, mas funding limitado, com recomendação explícita de
   também buscar bolsas externas listadas no site da Zicklin. Preenchido
   'Bolsas Disponíveis' (Sim) e 'Tipos de Bolsa'.

5. Deadline de aplicação: e-mail confirma que o programa presencial
   (on-campus) está com aplicações abertas para Spring 2027, primeiro
   deadline 1º de outubro. Preenchido 'Deadline Aplicação'. Também
   atualizado 'Início' de 'Fall' para 'Fall, Spring' -- e-mail confirma
   existência de intake de Spring (não estava documentado antes); não
   há indicação de que o intake de Fall tenha sido descontinuado, então
   Fall foi mantido.

'Fees Adicionais' e demais campos ainda 'Verificar' não mencionados no
e-mail (o próprio remetente direciona para o International Student
Service Center e a Tuition Calculator da própria escola para esses
detalhes) permanecem inalterados nesta rodada.

Fonte:
- E-mail direto de Kanishka Kapoor, Office of Graduate Admissions,
  Zicklin School of Business, Baruch College (CUNY), recebido em
  2026-09-08, em resposta a contato da aluna.
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
    "Nome da Universidade": "Baruch CUNY",
    "Estado": "NY",
    "Cidade": "New York",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Total Créditos": "30-33",  # e-mail oficial: "approximately 30-33 credits, depending on course waivers" (era 36)
    "Tuition Total (Intl) USD": "25650-28215",  # recalculado: 30-33 x $855 (era 30780, baseado em 36 creditos)
    "Tuition Total (Residente) USD": "14100-15510",  # recalculado: 30-33 x $470 (era 16920, baseado em 36 creditos)
    "Graduate Assistantships": "Disponíveis; pagas por hora; SEM isenção de tuition (confirmado por e-mail do programa)",
    "TA/RA Disponível": "Sim, na forma de graduate assistantships pagas por hora (sem waiver de tuition)",
    "Bolsas Disponíveis": "Sim",
    "Tipos de Bolsa": (
        "Internacionais são considerados automaticamente para bolsas na "
        "admissão; funding limitado. Recomenda-se também buscar bolsas "
        "externas listadas no site da Zicklin (Scholarships & Financial Aid)."
    ),
    "Deadline Aplicação": "Spring 2027 (presencial): 1º deadline em 1º de outubro",
    "Início": "Fall, Spring",  # e-mail confirma intake de Spring 2027; Fall mantido (nao ha indicacao de descontinuacao)
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
        raise ValueError("Nenhuma linha encontrada para Baruch CUNY (New York, NY). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Baruch CUNY, New York, NY)")

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
