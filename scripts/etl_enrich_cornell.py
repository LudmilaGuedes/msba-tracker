"""
Data enrichment: atualiza APENAS a linha da 'Cornell University (SC
Johnson)' (New York City, NY) no msba_programs_database.xlsx.

Foco exclusivo: Master of Science in Business Analytics (MSBA), formato
NYC (presencial, full-time, 10 meses) -- distinto do formato Online+
(part-time, 16 meses, mesma tuition mas custo de vida diferente). STEM
(CIP 52.1302, já confirmado na planilha), 30 créditos. Deep research em:
custo total real de tuition para internacionais, bolsas/assistantships,
requisitos de GPA/GMAT/GRE/inglês, e custo de vida em Nova York (NYC).

A linha estava quase inteiramente 'Verificar' (nenhum enrichment prévio
encontrado -- confirmado via Glob antes de iniciar).

TUITION: PREENCHIDO. A página oficial de tuition do MSBA confirma o
valor 2026-27: $89.864 TOTAL, cobrado em 2 parcelas iguais (outono e
primavera) para o formato NYC -- mesma tarifa para Online+ e para NYC,
e SEM distinção de residência (Cornell é privada; tarifa única para
domésticos e internacionais). PREENCHIDO 'Tuition Total' (Intl e
Residente) com $89.864, e 'Custo/Crédito' com $2.995,47 (89.864 ÷ 30
créditos).

CUSTO DE VIDA: a mesma página oficial fornece um breakdown OFICIAL e
ITEMIZADO do "Estimated Cost of Attendance" para o formato NYC
(diferente de um total bruto sem decomposição -- aqui os componentes já
vêm separados pela própria Cornell, então não há necessidade de estimar
uma proporção): Housing $25.856, Meals $6.404, Books & supplies $2.464,
Transportation $1.364, Miscellaneous $4.658, Student activity/health
fee $466 -- soma exata = $41.212, que bate com o total de COA
divulgado ($131.076) menos a tuition ($89.864). Confirmado por
soma-verificação antes do uso.

Dividido por 10 meses (duração do formato NYC) e mapeado aos campos do
schema: 'Aluguel Privativo/mês' = $2.586, 'Mercado/mês' = $640,
'Transporte/mês' = $136, 'Outros Gastos/mês' = $712 (books + misc
combinados, já que o schema não tem campo específico para material
didático). 'Custo Vida Total (programa) USD' preenchido com o total
oficial de $41.212 (todos os componentes de custo de vida, excluindo
tuition).

'Fees Adicionais USD' preenchido com $466 -- o "student activity/health
fee" oficial do formato NYC. CUIDADO: uma busca separada encontrou um
prêmio de seguro saúde de $3.944-$4.058/ano para "Cornell Tech"
genericamente (fonte agregadora, confiança moderada), mas esse valor
NÃO é confirmado como específico do MSBA -- pode se referir a outro
programa do campus com mandato de seguro diferente, e o breakdown
oficial do MSBA já soma exatamente ao total divulgado sem deixar espaço
para esse valor adicional. Por isso, 'Seguro Saúde (programa) USD'
permanece 'Verificar' em vez de usar a estimativa não confirmada, e o
fee oficial $466 (rotulado pela própria Cornell como "activity/health"
combinado) foi colocado em 'Fees Adicionais' em vez de dividido
arbitrariamente entre os dois campos.

'GMAT/GRE Obrigatório' PREENCHIDO: confirmado que NÃO é exigido para o
MSBA. 'Testes Inglês Aceitos' preenchido com os mínimos e política
oficial (Duolingo mín. 130, TOEFL/IELTS conforme Cornell Graduate
School, "My Best Score" do TOEFL não aceito, validade de 2 anos).
'Matérias Pré-Requisito' preenchido com o requisito oficial de
proficiência quantitativa (cálculo ou estatística em nível
universitário); programação é opcional, não obrigatória.

'Deadline Aplicação', 'Rounds Aplicação', 'Documentos Necessários',
'Entrevista' e 'Taxa Aplicação USD' preenchidos com dados oficiais do
guia de aplicação (foco no track NYC/presencial: janela de 1 dez a 15
abr; video interview assíncrono obrigatório pós-submissão; taxa $105;
depósito de matrícula não-reembolsável de $1.500 após aceite).

'Bolsas Disponíveis' e 'Tipos de Bolsa' preenchidos: bolsas de mérito
com consideração automática para TODOS os candidatos (domésticos e
internacionais), fundos limitados -- recomenda-se aplicar em rounds
iniciais.

CUIDADO APLICADO (dados NÃO usados): buscas por GPA médio, GMAT médio,
taxa de aceitação, taxa de empregabilidade, salário médio pós-programa,
principais empregadores e rankings (US News/QS) só retornaram dados do
programa de MBA (Two-Year MBA) da Cornell Johnson -- um programa
DIFERENTE, com processo seletivo, corpo discente e outcomes distintos
do MSBA. Nenhum desses números foi usado; todos os campos correspondentes
permanecem 'Verificar'. Pelo mesmo motivo de ausência de dado
específico ao MSBA, 'TA/RA Disponível', 'Processo TA/RA' e 'Graduate
Assistantships' também permanecem 'Verificar' (comum em masters
profissionais de 1 ano, mas não confirmado nem negado explicitamente
para este programa).

Fontes principais (pesquisadas via busca web, ago/2026):
- https://business.cornell.edu/admissions/graduate/msba/tuition/
  (tuition oficial 2026-27, breakdown itemizado de cost of attendance
  para NYC e Online+, bolsas de mérito, application fee, depósito)
- https://business.cornell.edu/admissions/graduate/msba/application-guide/
  (deadlines por track, documentos exigidos, política de entrevista,
  requisito de proficiência quantitativa, mínimos de inglês)
- Busca web: GMAT/GRE não exigido para MSBA; seguro saúde Cornell Tech
  $3.944-4.058/ano (não confirmado como específico do MSBA, não usado)
- Busca web: dados de GPA/GMAT/salário/emprego encontrados eram todos
  do programa de MBA, não do MSBA -- não usados
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
    "Nome da Universidade": "Cornell University (SC Johnson)",
    "Estado": "NY",
    "Cidade": "New York City (NYC)",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Intl) USD": 2995.47,  # 89.864 / 30 créditos, oficial
    "Custo/Crédito (Residente) USD": 2995.47,  # tarifa flat, sem distinção de residência (privada)
    "Tuition Total (Intl) USD": 89864,  # oficial 2026-27, formato NYC
    "Tuition Total (Residente) USD": 89864,  # tarifa flat
    "Caminho Tuition Residente": "N/A — tuition uniforme para todos (Cornell é privada, sem distinção de residência)",
    "Fees Adicionais USD": 466,  # "student activity/health fee" oficial, formato NYC
    "Aluguel Privativo/mês USD": 2586,  # $25.856 (COA oficial, formato NYC) ÷ 10 meses
    "Mercado/mês USD": 640,  # $6.404 (COA oficial) ÷ 10 meses
    "Transporte/mês USD": 136,  # $1.364 (COA oficial) ÷ 10 meses
    "Outros Gastos/mês USD": 712,  # (books $2.464 + misc $4.658) ÷ 10 meses
    "Custo Vida Total (programa) USD": 41212,  # total oficial de custo de vida (COA - tuition), formato NYC
    "GMAT/GRE Obrigatório": "Não exigido para o MSBA",
    "Testes Inglês Aceitos": (
        "TOEFL, IELTS Academic ou Duolingo (mín. Duolingo 130; "
        "mínimos de TOEFL/IELTS conforme Cornell Graduate School); "
        "\"My Best Score\" do TOEFL não aceito; teste válido até 2 "
        "anos antes do deadline"
    ),
    "Matérias Pré-Requisito": (
        "Proficiência quantitativa em nível universitário (cálculo ou "
        "estatística); conhecimento de programação é opcional, não "
        "obrigatório"
    ),
    "Deadline Aplicação": "Formato NYC (presencial): janela de 1 dez a 15 abr, admissão contínua por rounds",
    "Rounds Aplicação": "Admissão contínua (rolling) de 1 dez a 15 abr; recomendado aplicar em rounds iniciais devido a bolsas de mérito limitadas",
    "Documentos Necessários": (
        "Application Cornell Graduate School + taxa $105; histórico "
        "acadêmico de todas as instituições; resume/CV; 2 cartas de "
        "recomendação (mín. 1 de supervisor); academic statement of "
        "purpose (até 500 palavras); personal statement (300-500 "
        "palavras); teste de inglês; video interview pós-submissão; "
        "depósito de matrícula não-reembolsável de $1.500 após aceite"
    ),
    "Entrevista": "Sim — video interview assíncrono obrigatório (respostas gravadas a perguntas pré-definidas), após submissão da aplicação",
    "Taxa Aplicação USD": 105,  # oficial (correção: antes só indicava "Aplicável")
    "Bolsas Disponíveis": "Sim — bolsas de mérito, consideração automática para todos os candidatos (domésticos e internacionais); fundos limitados",
    "Tipos de Bolsa": "Bolsa por mérito (parcial, consideração automática); Yellow Ribbon Program para veteranos elegíveis dos EUA",
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
        raise ValueError("Nenhuma linha encontrada para Cornell University (SC Johnson) (New York City, NY). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Cornell University (SC Johnson), New York City (NYC), NY)")

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
