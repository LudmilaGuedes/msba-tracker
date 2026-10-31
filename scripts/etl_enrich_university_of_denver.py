"""
Data enrichment: atualiza APENAS a linha da 'University of Denver'
(Daniels College of Business, Denver, CO) no msba_programs_database.xlsx.

Foco exclusivo: MS in Business Analytics, presencial, STEM (OPT +24
meses), 58 créditos, 12 meses. Deep research em: custo total real de
tuition para internacionais, bolsas/assistantships, requisitos de GPA/
GMAT/GRE/inglês, e custo de vida em Denver, CO.

TUITION: PREENCHIDO (estava 'Verificar' em ambos os campos). A página
oficial de Tuition and Aid do Daniels College confirma explicitamente
que a tarifa é a MESMA para todos os alunos, "inside and out of
Colorado" (sem diferenciação por residência) -- $72.500 total para 58
créditos, $1.250/crédito exato (72.500 ÷ 58 = 1.250,00). PREENCHIDO
'Custo/Crédito' e 'Tuition Total' com esses valores (Intl e Residente),
e 'Total Créditos' com 58. 'Caminho Tuition Residente' documentado como
não aplicável (tuition já é flat).

'TA/RA Disponível' e 'Graduate Assistantships' preenchidos com CAUTELA:
a página geral de assistantships da DU confirma que GTAs/GRAs recebem
stipend mensal + tuition waiver + possível bolsa de seguro saúde -- um
pacote bastante generoso -- MAS a mesma página não confirma
explicitamente se alunos do MSBA/Daniels College têm acesso a essas
vagas (recomenda contato direto com o departamento). Não foi assumida
elegibilidade sem confirmação.

'Taxa Empregabilidade %' e 'Salário Médio Pós USD' permanecem
'Verificar': os únicos números encontrados (91% empregados em 6 meses,
salário médio $79.505) são estatísticas AGREGADAS de TODOS os
mestrados/profissionais da DU, não específicas do MSBA -- não usados
aqui para evitar misturar outcomes de programas diferentes (mesmo
critério de cautela já aplicado a outras universidades desta planilha).

Custo de vida: 'Transporte/mês' preenchido com $0 -- o RTD CollegePass
dá acesso ILIMITADO e gratuito ao sistema de trem/ônibus de Denver para
alunos da DU. 'Seguro Saúde (programa) USD' preenchido com $4.400/ano
-- prêmio oficial 2026-27 do DU Student Health Insurance Plan (SHIP).
'Aluguel Privativo/mês' preenchido com $700 (quarto privativo, faixa
$500-900 encontrada para Denver, ponderando bairros populares entre
alunos).

Fontes principais (pesquisadas via busca web, ago/2026):
- https://daniels.du.edu/graduate-admissions/tuition-and-aid/ (tuition
  oficial 2026-27, confirmação de tarifa flat sem distinção de
  residência, consideração automática para bolsas de mérito e GA)
- Busca web cruzada: GPA mínimo 2.5 (bacharelado) ou 3.0 (pós-graduação
  prévia), GMAT/GRE opcional, TOEFL 80/IELTS 7.0/Duolingo 115/C1
  Advanced aceito, deadlines Fall Priority 1 (15 out) e Spring Priority
  2 (5 dez)
- https://www.du.edu/graduate-education/current-student-resources/
  gta-information (benefícios de GTA/GRA: stipend + tuition waiver +
  possível seguro saúde; elegibilidade para Daniels não confirmada)
- Busca web: RTD CollegePass gratuito para alunos DU; seguro saúde
  $4.400/ano (SHIP oficial 2026-27); aluguel de quarto privativo em
  Denver, CO
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
    "Nome da Universidade": "University of Denver",
    "Estado": "CO",
    "Cidade": "Denver",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Total Créditos": 58,  # oficial
    "Custo/Crédito (Intl) USD": 1250,  # oficial, flat (mesma taxa dentro/fora do Colorado)
    "Custo/Crédito (Residente) USD": 1250,  # oficial, flat
    "Tuition Total (Intl) USD": 72500,  # 58 créditos x $1.250, flat
    "Tuition Total (Residente) USD": 72500,  # flat, mesma taxa
    "Caminho Tuition Residente": "N/A — tuition já é flat (mesma taxa dentro e fora do Colorado, confirmado oficialmente)",
    "Seguro Saúde (programa) USD": 4400,  # SHIP oficial 2026-27
    "Transporte/mês USD": 0,  # RTD CollegePass gratuito para alunos DU
    "Aluguel Privativo/mês USD": 700,  # quarto privativo, Denver CO (faixa $500-900 encontrada)
    "GPA Mínimo": "2.5 no bacharelado (ou 3.0 em pós-graduação prévia, se aplicável)",
    "Testes Inglês Aceitos": "TOEFL mín. 80 / IELTS mín. 7.0 / Duolingo mín. 115 / C1 Advanced aceito",
    "Deadline Aplicação": "Fall Priority 1: 15 out; Spring Priority 2: 5 dez (admissão contínua após os prazos de prioridade)",
    "Rounds Aplicação": "Priority rounds (Fall/Spring); admissão contínua (rolling) após os deadlines de prioridade",
    "Documentos Necessários": (
        "Application online, taxa de aplicação, histórico acadêmico, "
        "2 essays obrigatórios + 1 opcional, resume, teste de inglês "
        "(internacionais)"
    ),
    "TA/RA Disponível": (
        "Incerto para o MSBA especificamente — o programa geral de "
        "GTA/GRA da DU oferece stipend mensal + tuition waiver + "
        "possível bolsa de seguro saúde, mas elegibilidade para "
        "Daniels College/MSBA não confirmada (recomendado contato "
        "direto com o departamento)"
    ),
    "Graduate Assistantships": (
        "Consideração automática mencionada para admitidos, mas "
        "elegibilidade específica do MSBA não confirmada; benefício "
        "padrão da DU (quando disponível) inclui stipend + tuition "
        "waiver + possível seguro saúde"
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
        raise ValueError("Nenhuma linha encontrada para University of Denver (Denver, CO). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (University of Denver, Denver, CO)")

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
