"""
Data enrichment: atualiza APENAS a linha da 'American University' (Kogod
School of Business, Washington, DC) no msba_programs_database.xlsx.

Foco exclusivo: MS in Analytics (MSAn), formato ON-CAMPUS (presencial --
distinto da versão online/híbrida separada da mesma escola), STEM, 33
créditos, 12 meses. Deep research em: custo total real de tuition para
internacionais, bolsas/assistantships, requisitos de GPA/GMAT/GRE/
inglês, e custo de vida em Washington, DC.

TUITION: CORRIGIDO. 'Tuition Total (Intl) USD' continha um texto de
faixa inválido para um campo numérico ("59.800–63.426 (2026-27)"), sem
origem rastreável clara. A página oficial de Tuition and Fees do MS in
Analytics confirma o valor FIXO e atual (2026-27): 33 créditos x
$2.163/crédito (já presente corretamente na planilha) = $71.379 TOTAL,
sem faixa. CORRIGIDO 'Tuition Total (Intl)' para $71.379 e PREENCHIDO
'Tuition Total (Residente)' com o mesmo valor -- a mesma página confirma
explicitamente que a tarifa é uniforme para TODOS os alunos, incluindo
internacionais, sem distinção de residência.

'Fees Adicionais USD' preenchido com $2.880 -- soma das taxas
obrigatórias oficiais por semestre para alunos em tempo integral
(Graduate Activity $50 + Support Services $450 + Metro U-Pass $136 +
Digital Courseware/Eagle Access $324 = $960/semestre) x 3 semestres
(estimativa de tempo integral ao longo de 12 meses). 'Transporte/mês'
preenchido com $0 -- a Metro U-Pass (acesso ilimitado ao metrô de DC)
já está incluída nessas fees obrigatórias, então contar um custo mensal
de transporte separado duplicaria a cobrança.

'GMAT/GRE Obrigatório' CORRIGIDO: a planilha tinha uma descrição
ambígua ("Exigido OU essay suplementar"); a página oficial confirma que
GMAT/GRE NÃO é exigido -- é totalmente opcional, podendo ser enviado
para reforçar a candidatura ou consideração de bolsa adicional.

'Seguro Saúde (programa) USD' preenchido com $2.759/ano -- prêmio
oficial 2026-27 do plano Aetna, obrigatório para todo aluno em tempo
integral e para internacionais com visto F-1/J-1.

CUIDADO APLICADO: um número de salário médio ($105.000, "MSBA
graduates") e uma taxa de emprego (93%) encontrados em buscas NÃO foram
usados -- vieram de uma fonte que se refere a "MSBA graduates" (nome
diferente do programa atual, "MS in Analytics"/MSAn), e a página
oficial de outcomes do Kogod ("Student Outcomes Hub") não confirma
esses números especificamente para o MSAn, trazendo apenas estatísticas
agregadas de TODOS os programas de pós-graduação do Kogod (exceto MBA
full-time). Por não conseguir confirmar a especificidade ao MSAn, esses
campos permanecem 'Verificar'. Pelo mesmo motivo, 'Tamanho Turma'
(um número de "10 alunos, 4.6 anos de experiência" encontrado na mesma
leva de busca não confiável) também permanece 'Verificar'.

Fontes principais (pesquisadas via busca web, ago/2026):
- https://kogod.american.edu/programs-admissions/masters/analytics/
  tuition-fees (tuition oficial 2026-27, breakdown completo de fees,
  confirmação de tarifa uniforme sem distinção de residência)
- https://kogod.american.edu/kogod-school-of-business-student-outcomes-hub
  (outcomes oficiais -- confirma que não há dados específicos do MSAn
  publicados nesta página, só agregados do Kogod em geral)
- Busca web: GMAT/GRE não exigido (opcional), TOEFL/IELTS/Duolingo
  aceitos (mínimos não divulgados publicamente)
- Busca web: seguro saúde $2.759/ano (oficial 2026-27, plano Aetna)
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
    "Nome da Universidade": "American University",
    "Estado": "DC",
    "Cidade": "Washington",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Tuition Total (Intl) USD": 71379,  # CORREÇÃO: planilha tinha uma faixa inválida ("59.800-63.426"); oficial 2026-27 = 33 x $2.163
    "Tuition Total (Residente) USD": 71379,  # tarifa uniforme, sem distinção de residência (confirmado oficialmente)
    "Fees Adicionais USD": 2880,  # (Graduate Activity + Support Services + Metro U-Pass + Digital Courseware) x 3 semestres
    "Transporte/mês USD": 0,  # Metro U-Pass já incluída nas fees obrigatórias (evitar dupla contagem)
    "Seguro Saúde (programa) USD": 2759,  # oficial 2026-27, plano Aetna
    "Caminho Tuition Residente": "N/A — tuition uniforme para todos, incluindo internacionais (confirmado oficialmente, sem distinção de residência)",
    "GMAT/GRE Obrigatório": (
        "CORREÇÃO: não exigido (totalmente opcional); pode ser "
        "enviado para reforçar a candidatura ou consideração de bolsa "
        "adicional"
    ),
    "Testes Inglês Aceitos": "TOEFL/IELTS/Duolingo aceitos (scores mínimos não divulgados publicamente)",
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
        raise ValueError("Nenhuma linha encontrada para American University (Washington, DC). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (American University, Washington, DC)")

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
