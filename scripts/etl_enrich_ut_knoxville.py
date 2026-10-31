"""
Data enrichment: atualiza APENAS a linha da 'UT Knoxville' (Haslam
College of Business, Knoxville, TN) no msba_programs_database.xlsx.

Foco exclusivo: MS in Business Analytics & AI (MSBAi), presencial, STEM,
36 créditos, 10 meses (3 semestres). Deep research em: custo total real
de tuition para internacionais, bolsas/assistantships, requisitos de GPA/
GMAT/GRE/inglês, e custo de vida em Knoxville, TN.

TUITION TOTAL (ja estava correto, apenas verificado): a pagina oficial
(haslam.utk.edu) confirma $26.606 (residente) / $54.371 (internacional)
para o programa completo -- batendo exatamente com os valores ja
presentes na planilha. 'Custo/Credito' estava 'Verificar'; PREENCHIDO
dividindo os totais oficiais por 36 creditos ($739,06 residente /
$1.510,31 internacional).

'Caminho Tuition Residente' preenchido com nota sobre o Academic Common
Market (SREB) -- programa regional que permite a alunos de certos
estados do Sudeste dos EUA pagar tuition de residente no Tennessee para
programas nao oferecidos no proprio estado; NAO se aplica a alunos
internacionais.

Custo de vida: 'Transporte/mes' preenchido com $0 -- UT Knoxville oferece
transporte gratuito via T (onibus do campus) + parceria com Knoxville
Area Transit (KAT), ambos gratuitos para alunos com o VolCard. 'Seguro
Saude (programa)' preenchido com $2.520 -- unico valor oficial
encontrado (Grad Estimate of Expenses do escritorio internacional,
ciclo 2022-23); pode estar desatualizado (o custo real pode ser maior
hoje), mas e o melhor dado disponivel. Com os 9 componentes mensais
completos, 'Custo Vida Total (programa) USD' foi calculado: (975 aluguel
+ 400 mercado + 15 seguro residencial + 155 utilities + 50 celular + 60
wifi + 0 transporte + 150 outros) x 10 meses + 2.520 seguro saude =
$20.570.

'Fees Adicionais USD', 'Ranking US News', 'Ranking QS', 'Principais
Empregadores', 'GMAT Medio', 'Taxa Aceitacao %' e 'Taxa Aplicacao USD'
permanecem 'Verificar' -- nenhuma fonte oficial encontrada com esses
numeros especificos para o MSBAi (os $26.606/$54.371 ja parecem
representar tuition + fees combinados, sem um "fees adicionais"
separado divulgado; e as paginas oficiais citam apenas cargos/carreiras
tipicas, nao empregadores nomeados).

Fontes principais (pesquisadas via busca web, ago/2026):
- https://haslam.utk.edu/business-analytics-statistics/msbai/ e
  https://haslam.utk.edu/haslam-msbai/ (tuition oficial, GPA 3.0 min/
  3.64 media, GRE/GMAT com waiver via Canvas Exam $9, TOEFL 80/IELTS
  6.5, entrevista de 30min, taxa de emprego ~90%, salario mediano
  $87.563 + bonus medio $9.950, Academic Common Market, deadlines)
- Busca web: seguro saude internacional $2.520/ano (2022-23, Grad
  Estimate of Expenses, international.utk.edu)
- Busca web: T bus + parceria KAT gratuitos para alunos UT Knoxville
  (VolCard)
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
    "Nome da Universidade": "UT Knoxville",
    "Estado": "TN",
    "Cidade": "Knoxville",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Intl) USD": 1510.31,  # $54.371 / 36 créditos (tuition total oficial já confirmado)
    "Custo/Crédito (Residente) USD": 739.06,  # $26.606 / 36 créditos (tuition total oficial já confirmado)
    "Seguro Saúde (programa) USD": 2520,  # oficial (ciclo 2022-23, international.utk.edu); pode estar desatualizado
    "Transporte/mês USD": 0,  # T bus (campus) + parceria KAT gratuitos para alunos UT Knoxville (VolCard)
    "Custo Vida Total (programa) USD": 20570,  # (975+400+15+155+50+60+0+150) x 10 meses + 2.520 seguro saúde
    "Caminho Tuition Residente": (
        "Academic Common Market (SREB) — alunos de certos estados do "
        "Sudeste dos EUA podem pagar tuition de residente para "
        "programas não oferecidos no próprio estado; não se aplica a "
        "internacionais"
    ),
    "Matérias Pré-Requisito": (
        "Nenhum pré-requisito formal exigido; bagagem quantitativa "
        "(estatística, cálculo, programação, engenharia) é fortemente "
        "recomendada e avaliada via coursework prévio ou Canvas Exam/"
        "GRE/GMAT"
    ),
    "Documentos Necessários": (
        "Resume, respostas a essay questions, histórico acadêmico "
        "oficial, 2 cartas de recomendação (preferencialmente 1 "
        "acadêmica em área quantitativa + 1 profissional), GRE/GMAT ou "
        "Canvas Exam ($9, alternativa interna), TOEFL/IELTS "
        "(internacionais)"
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
        raise ValueError("Nenhuma linha encontrada para UT Knoxville (Knoxville, TN). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (UT Knoxville, Knoxville, TN)")

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
