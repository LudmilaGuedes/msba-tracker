"""
Data enrichment: atualiza APENAS a linha da 'U of Florida' (Warrington
College of Business, Gainesville, FL) no msba_programs_database.xlsx.

Foco exclusivo: MS in Business Analytics, formato on-campus (presencial),
STEM, 10 meses, modulos de 7 semanas, 36 creditos. Deep research em:
custo total real de tuition para internacionais, bolsas/assistantships,
requisitos de GPA/GMAT/GRE/ingles, e custo de vida em Gainesville.

ACHADO CRITICO DA AUDITORIA: 'Tuition Total (Intl) USD' ($30.130) e
'Tuition Total (Residente) USD' ($30.134) estavam quase IDENTICOS entre
si -- um padrao suspeito, ja que programas com "market-rate"/preco
uniforme costumam ser assim, mas o MSBA da UF nao e um desses casos.
A pagina oficial do programa (warrington.ufl.edu/graduate/masters-
business-analytics/ e warrington.ufl.edu/master-business-analytics/,
ambas convergindo para os MESMOS 2 numeros) mostra um gap grande e
consistente entre residente e nao-residente: $19.171 (FL resident) vs
$52.174 (out-of-state/internacional) -- muito diferente dos $30.130/
$30.134 da planilha, que pareciam ser de uma fonte desatualizada ou
incorreta (nem batem com o valor historico mais antigo encontrado via
Poets&Quants, $19.105/$45.195, sugerindo que os $30k da planilha vieram
de outra fonte/programa por engano). CORRIGIDO para os valores oficiais
correntes (ciclo 2026-27): $19.171 residente / $52.174 internacional,
divididos por 36 creditos para 'Custo/Credito'.

'Fees Adicionais USD' estava 'Verificar'; a mesma pagina oficial de
custos (warrington.ufl.edu/admissions-aid/cost-aid/graduate/) lista uma
"professional development fee" obrigatoria de $350 (domestico) / $650
(internacional) especificamente associada aos mesmos valores de tuition
do MSBA ($19.171/$52.174) -- CORRIGIDO/PREENCHIDO com $650 (cenario
internacional, foco do research).

'Transporte/mes USD' (ja preenchido com $75) NAO foi alterado apesar de
RTS (onibus de Gainesville) ser gratuito para alunos da UF com Gator1 ID
-- o valor pode refletir custos de carro/estacionamento ocasionais nao
cobertos pelo onibus gratis, e nao ha evidencia forte o suficiente para
sobrescrever um dado ja preenchido (mesmo criterio de cautela usado em
outros enriquecimentos: so substituir dado existente quando ha
contradicao clara).

'Ranking US News' e 'Ranking QS' permanecem 'Verificar': nao foi
encontrado um ranking especifico de US News para o programa de Business
Analytics da UF (apenas o ranking geral do Warrington College/MBA, que
nao e o mesmo programa e seria enganoso usar aqui); o ranking QS
encontrado (#27 nacional) e da University of South Florida (USF), NAO da
University of Florida -- descartado por ser de instituicao errada.

Fontes principais (pesquisadas via busca web, ago/2026):
- https://warrington.ufl.edu/graduate/masters-business-analytics/ e
  https://warrington.ufl.edu/master-business-analytics/ (tuition oficial,
  STEM, admissao, deadlines, salario medio $82.166 + bonus $5.600,
  principais empregadores, GMAT/GRE nao exigido, GPA minimo 3.0)
- https://warrington.ufl.edu/admissions-aid/cost-aid/graduate/
  (tabela de tuition + professional development fee $350/$650)
- Busca web: RTS onibus gratuito para alunos UF com Gator1 ID
  (amberstudent.com, swamprentals.com)
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
    "Nome da Universidade": "U of Florida",
    "Estado": "FL",
    "Cidade": "Gainesville",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Intl) USD": 1449.28,  # CORRECAO: oficial 2026-27 $52.174 / 36 creditos
    "Custo/Crédito (Residente) USD": 532.53,  # CORRECAO: oficial 2026-27 $19.171 / 36 creditos
    "Tuition Total (Intl) USD": 52174,  # CORRECAO: planilha tinha $30.130 (quase igual ao residente, incoerente); oficial = $52.174
    "Tuition Total (Residente) USD": 19171,  # CORRECAO: planilha tinha $30.134; oficial = $19.171
    "Fees Adicionais USD": 650,  # Professional development fee obrigatória (cenário internacional); $350 para domésticos
    "Salário Médio Pós USD": 82166,  # oficial, turma recente; + $5.600 de signing bonus mediano (não capturado em campo separado do schema)
    "Principais Empregadores": "Accenture, Bank of America, Deloitte, EY, KPMG, PwC, NextEra Energy, Protiviti",
    "Matérias Pré-Requisito": (
        "Não há pré-requisitos formais específicos além do bacharelado "
        "(ou equivalente internacional de 4 anos); GPA mínimo de 3.0 "
        "(upper-division, para diplomas dos EUA)"
    ),
    "Rounds Aplicação": (
        "Sem rounds numerados formais; Intl prioridade 1º jan, Intl final "
        "15 fev, Doméstico 1º mai"
    ),
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
        raise ValueError("Nenhuma linha encontrada para U of Florida (Gainesville, FL). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (U of Florida, Gainesville, FL)")

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
