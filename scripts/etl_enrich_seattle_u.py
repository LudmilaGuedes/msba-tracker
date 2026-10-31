"""
Data enrichment: atualiza APENAS a linha da 'Seattle U' (Albers School of
Business & Economics, Seattle, WA) no msba_programs_database.xlsx.

Foco exclusivo: MS in Business Analytics, formato presencial (First Hill
campus), STEM, 45 créditos trimestrais, 15-24 meses. Deep research em:
custo total real de tuition para internacionais, bolsas/assistantships,
requisitos de GPA/GMAT/GRE/inglês, e custo de vida em Seattle, WA.

ACHADO DA AUDITORIA: os campos 'Custo/Crédito (Intl)' ($1.042) e 'Tuition
Total (Intl)' ($37.512) eram INTERNAMENTE INCONSISTENTES entre si --
1.042 x 45 créditos = $46.890, não $37.512 (na verdade, 1.042 x 36 =
37.512 exatamente, sugerindo que o total foi calculado com o número de
créditos errado, talvez confundido com outro programa de 36 créditos).
A página oficial de Cost of Attendance (seattleu.edu, ciclo 2026-27)
confirma a taxa atual: $1.073/crédito para o MSBA. CORRIGIDO: Custo/
Crédito para $1.073 e Tuition Total para $48.285 (45 x 1.073). Como é
universidade PRIVADA, sem distinção residente/não-residente confirmada,
'Custo/Crédito (Residente)' e 'Tuition Total (Residente)' foram
preenchidos com os MESMOS valores (mesmo critério já usado para
University of Dallas, outra privada desta planilha).

'Fees Adicionais USD': calculado a partir das taxas obrigatórias oficiais
por trimestre (Wellness $172 + Technology $214 + Graduate Activity $20,
esta última dispensada no trimestre de verão) mais a Matriculation Fee de
$175 (cobrada uma única vez, no primeiro trimestre) -- projetado para um
percurso típico de 15 meses (5 trimestres, 1 deles de verão): PREENCHIDO
com $2.185. A tabela oficial de orçamento (Cost of Attendance) do próprio
site cruza e confirma os valores trimestrais de Wellness/Technology/
Activity usados neste cálculo.

Custo de vida: 'Aluguel Privativo/mês' preenchido com $950 (quarto
privativo com banheiro privativo, região de First Hill/Capitol Hill perto
do campus -- múltiplos anúncios convergentes). 'Mercado/mês' ($714),
'Transporte/mês' ($191, SEM evidência de transporte gratuito para alunos
da Seattle U -- ao contrário de outras universidades já enriquecidas
nesta planilha) e 'Outros Gastos/mês' ($219) vieram do próprio orçamento
oficial de Cost of Attendance da universidade (dividido por 9 meses).
'Seguro Saúde (programa)' preenchido com $1.823 -- único valor
encontrado, do ciclo de apólice 2022-23; pode estar desatualizado.
'Utilities/mês', 'Seguro Residencial/mês', 'Celular/mês' e 'WiFi/mês'
permanecem 'Verificar' -- sem fonte isolada confiável -- e por isso
'Custo Vida Total (programa) USD' também permanece 'Verificar' (faltam
4 dos 8 componentes mensais para uma soma completa).

Fontes principais (pesquisadas via busca web, ago/2026):
- https://www.seattleu.edu/admissions-aid/tuition-and-fees/
  graduate-cost-of-attendance/ (tuition oficial $1.073/crédito 2026-27,
  taxas trimestrais, orçamento detalhado de custo de vida)
- https://www.seattleu.edu/academics/all-programs/business-analytics-ms/
  admissions-process/ (documentos, TOEFL 86/IELTS 6.5/PTE 63/Duolingo
  120, pré-requisitos de Statistics/Calculus/Programming, isenção de
  taxa de aplicação, deadlines Early Review 20 jan / Final 20 mar)
- https://www.seattleu.edu/academics/all-programs/business-analytics-ms/
  (ranking #19 US News 2026, empregadores, formato, duração)
- https://www.seattleu.edu/business/graduate/
  graduate-assistantships-internships-fellowships/ (Genevieve Albers
  Graduate Fellowship: $15k tuition + $7k stipend, 12,5h/semana, 40%
  dos candidatos internacionais, 1 bolsa concedida entre 40 candidatos)
- Busca web: seguro saúde internacional $1.823/ano (2022-23,
  seattleu.edu/isc)
- Busca web: aluguel de quarto privativo com banheiro privativo em
  First Hill/Capitol Hill, Seattle WA (múltiplos anúncios convergentes)
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
    "Nome da Universidade": "Seattle U",
    "Estado": "WA",
    "Cidade": "Seattle",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Intl) USD": 1073,  # CORREÇÃO: oficial 2026-27 $1.073/crédito (planilha tinha $1.042, inconsistente com o Tuition Total)
    "Custo/Crédito (Residente) USD": 1073,  # privada, mesma taxa (sem distinção de residência confirmada)
    "Tuition Total (Intl) USD": 48285,  # CORREÇÃO: 45 créditos x $1.073 (planilha tinha $37.512, calculado com número de créditos errado)
    "Tuition Total (Residente) USD": 48285,  # privada, mesma taxa
    "Fees Adicionais USD": 2185,  # taxas obrigatórias por trimestre (Wellness+Technology+Activity) x ~5 trimestres (15 meses) + Matriculation Fee $175
    "Aluguel Privativo/mês USD": 950,  # quarto privativo com banheiro privativo, First Hill/Capitol Hill (perto do campus)
    "Mercado/mês USD": 714,  # orçamento oficial de Cost of Attendance ($6.426 / 9 meses)
    "Seguro Saúde (programa) USD": 1823,  # oficial (ciclo de apólice 2022-23); pode estar desatualizado
    "Transporte/mês USD": 191,  # orçamento oficial de Cost of Attendance ($1.716 / 9 meses); sem evidência de transporte gratuito para alunos
    "Outros Gastos/mês USD": 219,  # orçamento oficial de Cost of Attendance, "Personal Expenses" ($1.968 / 9 meses)
    "Bolsas Disponíveis": "Sim",
    "Tipos de Bolsa": (
        "Albers Incoming Student Graduate Scholarship; Redapt "
        "Scholarship (específica do MSBA, candidatura até 20 mai); taxa "
        "de aplicação sempre isenta para todos os candidatos"
    ),
    "Quão Realista (Intl)": (
        "Baixa a Moderada — a Genevieve Albers Graduate Fellowship "
        "teve apenas 1 bolsa concedida entre 40 candidatos no último "
        "ciclo divulgado (40% internacionais); demais bolsas nomeadas "
        "existem mas são competitivas"
    ),
    "TA/RA Disponível": "Sim (Graduate Assistantships disponíveis para todos os programas de pós-graduação do Albers, incluindo MSBA)",
    "Processo TA/RA": (
        "Candidatura própria por bolsa/fellowship específica (ex.: "
        "Genevieve Albers Graduate Fellowship, deadline 30 abr); sem "
        "processo centralizado único para todas as vagas de "
        "assistantship"
    ),
    "Graduate Assistantships": (
        "Genevieve Albers Graduate Fellowship: $15.000 de bolsa de "
        "tuition + $7.000 de stipend, 12,5h/semana por 3 trimestres "
        "(350h/ano); aberta a domésticos e internacionais"
    ),
    "Ranking US News": 19,  # Best Master's in Business Analytics, U.S. News & World Report 2026
    "Matérias Pré-Requisito": (
        "Statistics, Calculus e Programming — cursos online gratuitos e "
        "self-paced desenhados pela própria Seattle University; "
        "Programming deve ser concluído antes do início do programa; "
        "Statistics/Calculus podem ser dispensados por teste"
    ),
    "Documentos Necessários": (
        "Histórico acadêmico oficial (últimos 5 anos), resume, ensaio "
        "de 300 palavras (objetivos de carreira ou resolução analítica "
        "de problemas), comprovação de conclusão dos pré-requisitos "
        "(Estatística, Cálculo, Programação)"
    ),
    "Rounds Aplicação": "Early Review 20 jan (com bolsa); Final 20 mar; somente admissão Fall",
    "Deadline Aplicação": "Early Review 20 jan (bolsa); Final 20 mar",  # CORREÇÃO: planilha tinha "Mar 30", oficial é 20 mar
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
        raise ValueError("Nenhuma linha encontrada para Seattle U (Seattle, WA). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Seattle U, Seattle, WA)")

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
