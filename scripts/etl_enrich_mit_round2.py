"""
Data enrichment (RODADA 2): atualiza APENAS a linha 'MIT' (Sloan
School of Management, Cambridge, MA) no msba_programs_database.xlsx.

Esta linha já passou por uma rodada de enrichment muito completa
(script etl_enrich_mit.py), restando 10/54 campos 'Verificar'. Usuário
confirmou via AskUserQuestion (ago/2026) para prosseguir com uma
Rodada 2.

'Total Créditos' preenchido: "111-147" -- a página oficial de
currículo confirma que o programa exige entre 111 e 147 unidades MIT
(fall + IAP + spring + summer), a variação vindo principalmente da
faixa de eletivas de analytics ("27 to 48 units"). Mantido como faixa
em vez de um único número, por ser a informação oficial real (não uma
imprecisão a resolver).

'Custo/Crédito' preenchido por ESTIMATIVA: $96.884 ÷ 129 (ponto médio
da faixa 111-147) ≈ $751,03 -- mesmo valor para Intl e Residente (MIT
Sloan é privada, tuition flat, já documentado). RESSALVA: como a
faixa de unidades é larga, o custo por unidade real pode variar entre
~$659 e ~$873 dependendo da carga exata de eletivas escolhida pelo
aluno -- o valor preenchido é um ponto médio aproximado, não uma
tarifa fixa oficial por unidade.

'Fees Adicionais USD' preenchido: $420 -- taxa GERAL de 9 meses da
MIT (não específica do Sloan), encontrada na página oficial de "Cost
of Attendance" da SFS (Student Financial Services). CONFIANÇA
MODERADA-ALTA: embora a página avise que "Graduate programs at Sloan
have their own tuition" (o que já era conhecido -- daí a tuition de
$96.884 já presente, diferente da tuition geral do MIT), o valor de
SEGURO SAÚDE dessa mesma tabela geral ($5.148) bate EXATAMENTE com o
valor já presente e confirmado na planilha, indicando que os demais
itens não-tuition dessa tabela (incluindo a fee de $420) também se
aplicam aos alunos de Sloan.

'Tipos de Bolsa' REFINADO com achado IMPORTANTE: a página oficial de
currículo revela que o Analytics Capstone (parte obrigatória do
currículo, no verão, com organizações parceiras) oferece "a tuition
offset of more than $22.000" -- ou seja, além da bolsa por mérito já
documentada (limitada, baixa realista para internacionais), TODOS os
alunos do MBAn se beneficiam desse abatimento estrutural embutido no
programa via o capstone patrocinado, reduzindo o custo líquido real
de forma significativa e mais garantida do que a bolsa por mérito.

CUIDADO APLICADO (dados NÃO encontrados mesmo após tentativas
adicionais nesta rodada, permanecem 'Verificar'): 'GPA Mínimo' e 'GPA
Médio Admitidos' -- a página de admissões só menciona qualitativamente
que candidatos fortes devem "demonstrate academic excellence... and
strong quantitative ability", sem nenhum número ou estatística de
GPA. 'Seguro Residencial/mês', 'Utilities/mês', 'Celular/mês', 'WiFi/
mês' permanecem 'Verificar' -- não encontrados em nenhuma fonte
consultada.

Fontes principais (pesquisadas via WebFetch, ago/2026):
- https://mitsloan.mit.edu/master-of-business-analytics/curriculum
  (unidades totais 111-147, estrutura curricular completa, achado do
  tuition offset do capstone >$22k)
- https://mitsloan.mit.edu/master-of-business-analytics/admissions
  (sem GPA numérico, apenas critérios qualitativos)
- https://sfs.mit.edu/graduate-students/cost-of-attendance/ (tabela
  geral de custo de atendimento do MIT 2026-27, fee de $420
  confirmada por validação cruzada com o seguro saúde já presente)
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
    "Nome da Universidade": "MIT",
    "Estado": "MA",
    "Cidade": "Cambridge",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Total Créditos": "111-147",  # oficial, faixa de unidades MIT (fall+IAP+spring+summer)
    "Custo/Crédito (Intl) USD": 751.03,  # 96.884 / 129 (ponto médio da faixa), estimativa
    "Custo/Crédito (Residente) USD": 751.03,  # privada, tarifa flat (igual ao Intl)
    "Fees Adicionais USD": 420,  # tabela geral de custo de atendimento MIT, validada via seguro saúde já confirmado
    "Tipos de Bolsa": (
        "Merit-based (limitado, baixa realista para internacionais) "
        "+ tuition offset estrutural de mais de $22.000 via Analytics "
        "Capstone (verão, com organizações parceiras) — aplicável a "
        "TODOS os alunos, não apenas quem recebe bolsa de mérito"
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
        raise ValueError("Nenhuma linha encontrada para MIT (Cambridge, MA). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (MIT, Cambridge, MA)")

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
