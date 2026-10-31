"""
Data enrichment: atualiza APENAS a linha da 'Texas State' (McCoy College of
Business, San Marcos, TX) no msba_programs_database.xlsx.

Foco exclusivo: M.S. in Data Analytics and Information Systems (MSDAIS) --
o programa STEM presencial de analytics do McCoy College, 30 creditos,
3 semestres (~16 meses). Deep research em: custo total real de tuition
para internacionais, bolsas/assistantships, requisitos de GPA/GMAT/GRE/
ingles, e custo de vida em San Marcos, TX.

'Nome Oficial do Programa' estava abreviado ('MS in Data Analytics & IS');
CORRIGIDO para o nome oficial completo usado pela propria Texas State:
'Master of Science in Data Analytics and Information Systems (M.S.D.A.I.S.)'.
Confirmado que este e de fato o programa STEM presencial de business
analytics da McCoy College (nao ha um programa separado chamado "MS in
Business Analytics" nesta universidade).

TUITION -- processo de calculo (varias paginas oficiais, ago/2026):
- Tabela oficial "Fall 2026 Graduate Tuition and Fees" (sbs.txst.edu)
  mostra $4.514,49 (residente) / $8.363,49 (nao-residente) para 9 creditos
  e $5.875,32 / $11.008,32 para 12 creditos. A taxa marginal por credito
  (excluindo taxas fixas por semestre) fica em ~$453,61/credito
  (residente) e ~$881,61/credito (nao-residente) -- validada de forma
  cruzada contra a tabela de Summer 2026 do mesmo site, que da uma taxa
  marginal quase identica (~$452,96 / ~$858,00). CORRIGIDO/PREENCHIDO
  'Custo/Credito' com estes valores e 'Tuition Total' multiplicando por
  30 creditos (mesmo criterio usado em todo o projeto).
- 'Fees Adicionais USD': a pagina oficial de Estimated Cost of Attendance
  do escritorio de alunos internacionais (international.txst.edu) informa
  que "McCoy College of Business graduate students must show an
  Additional $1.800 on their I-20" -- e uma busca separada indicou uma
  differential tuition fee de $150/credito para programas de pos-graduacao
  do McCoy College. Os dois numeros SE VALIDAM MUTUAMENTE: $150 x 12
  creditos = exatamente $1.800, a base usada pelo ISSS para o calculo do
  I-20. CORRIGIDO/PREENCHIDO com $150/credito x 30 creditos = $4.500
  (fee especifica do McCoy, adicional a tuition institucional).

Custo de vida: a mesma pagina de Cost of Attendance da ISSS (2026-27, base
para I-20 de um aluno de mestrado nao-residente) fornece um breakdown
oficial de "Living Expenses (9 meses)": Room & Board $13.230 (combinado,
sem split individual confiavel entre aluguel e mercado -- por isso esses
2 campos permanecem 'Verificar'), Transporte $1.660/9m (~$184/mes,
PREENCHIDO), Personal/Misc $2.750/9m (~$306/mes, PREENCHIDO em 'Outros
Gastos/mes'), Seguro Saude $2.333 (PREENCHIDO em 'Seguro Saude (programa)'
-- valor e para o ano academico padrao Fall+Spring; o programa dura ~16
meses, entao o total real pode ser um pouco maior). 'Custo Vida Total
(programa) USD' permanece 'Verificar' pois faltam os componentes de
aluguel, mercado, utilities, seguro residencial, celular e wifi
individualmente -- mesmo criterio de rigor ja aplicado em outros scripts
deste projeto (nao somar quando faltam componentes).

Bolsas/Assistantships: 'Caminho Tuition Residente' confirmado -- GTAs/GRAs
recebem waiver automatico (processado pelo Student Business Services
assim que o departamento contratante submete os documentos de contratacao)
que reduz a tuition de nao-residente para residente. Bolsas competitivas/
merit do Graduate College de pelo menos $1.000 tambem dao direito a um
"nonresident tuition competitive scholarship waiver".

Fontes principais (pesquisadas via busca web, ago/2026):
- https://www.mccoy.txst.edu/graduate/data-analytics.html (formato, STEM,
  curriculo, salario medio $98.333 em 3 meses pos-formatura, principais
  empregadores)
- https://www.gradcollege.txst.edu/programs/msdais.html (GPA/GMAT/GRE
  competitivos, TOEFL/IELTS/PTE/Duolingo, taxa de aplicacao, deadlines,
  documentos)
- https://www.sbs.txst.edu/sbs-policies/tuition-breakdown/graduate-rates/
  fall-2026-graduate-rates.html e a tabela de Summer 2026 (tuition oficial)
- https://www.international.txst.edu/.../estimated-cost-of-attendance-
  2026-2027.html (COA oficial p/ alunos internacionais, incl. fee
  adicional do McCoy e breakdown de custo de vida)
- https://www.gradcollege.txst.edu/funding/assistantships.html e
  https://onestop.txst.edu/financial-aid/waivers-and-exemptions.html
  (waiver de tuition nao-residente para GTA/GRA e bolsistas)
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
    "Nome da Universidade": "Texas State",
    "Estado": "TX",
    "Cidade": "San Marcos",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Nome Oficial do Programa": "Master of Science in Data Analytics and Information Systems (M.S.D.A.I.S.)",  # nome oficial completo
    "Destaques Currículo": (
        "30 créditos, 3 semestres; opção thesis/non-thesis (internship "
        "como eletiva na opção non-thesis); concentrações em Artificial "
        "Intelligence ou Supply Chain Analytics"
    ),
    "Custo/Crédito (Intl) USD": 881.61,  # oficial Fall 2026, taxa marginal por crédito (não-residente), validada contra tabela de Summer 2026
    "Custo/Crédito (Residente) USD": 453.61,  # oficial Fall 2026, taxa marginal por crédito (residente)
    "Tuition Total (Intl) USD": 26448,  # 30 créditos x $881,61
    "Tuition Total (Residente) USD": 13608,  # 30 créditos x $453,61
    "Fees Adicionais USD": 4500,  # McCoy College differential fee: $150/crédito x 30 créditos (validado: $150x12=$1.800, valor citado pela ISSS para o I-20)
    "Transporte/mês USD": 184,  # ISSS Cost of Attendance 2026-27: $1.660 / 9 meses
    "Outros Gastos/mês USD": 306,  # ISSS COA: "Personal and Miscellaneous" $2.750 / 9 meses
    "Seguro Saúde (programa) USD": 2333,  # ISSS COA 2026-27, ano acadêmico padrão (Fall+Spring); programa dura ~16 meses, total real pode ser um pouco maior
    "Bolsas Disponíveis": "Sim",
    "Tipos de Bolsa": (
        "Bolsas competitivas/merit do Graduate College; bolsas de pelo "
        "menos $1.000 dão direito a um nonresident tuition competitive "
        "scholarship waiver"
    ),
    "Quão Realista (Intl)": (
        "Moderada — waivers de tuition não-residente são bem "
        "estabelecidos para GTA/GRA e para bolsistas competitivos "
        "(≥$1.000), mas não há bolsa integral garantida"
    ),
    "TA/RA Disponível": "Sim (GTA — Graduate Teaching Assistant / GRA — Graduate Research Assistant)",
    "Processo TA/RA": (
        "Contratação por departamento/professor; waiver automático de "
        "tuition não-residente para taxa de residente é processado pelo "
        "Student Business Services assim que o departamento submete os "
        "documentos de contratação"
    ),
    "Caminho Tuition Residente": (
        "Sim — GTA/GRA recebem waiver automático que reduz a tuition de "
        "não-residente para a taxa de residente; bolsistas competitivos "
        "do Graduate College (≥$1.000) também podem receber waiver "
        "similar"
    ),
    "Graduate Assistantships": (
        "GTA/GRA disponíveis via departamentos, com waiver automático de "
        "tuition não-residente → residente durante o período de contrato"
    ),
    "GMAT/GRE Obrigatório": "Obrigatório em geral (GMAT ou GRE, pontuação competitiva), mas com possibilidade de waiver — critérios no McCoy College",
    "Testes Inglês Aceitos": (
        "TOEFL iBT mín. 78 (L19/R19/S19/W18) / IELTS Academic mín. 6.5 "
        "(mín. 6.0 por módulo) / PTE mín. 52 / Duolingo mín. 110 / TOEFL "
        "Essentials mín. 8.5"
    ),
    "Deadline Aplicação": (
        "Fall Priority 1º fev; Fall Standard 1º ago (EUA) / 1º jun "
        "(Intl); Spring Priority 1º set; Spring Standard 1º nov (EUA) / "
        "1º out (Intl)"
    ),
    "Taxa Aplicação USD": 90,  # internacional ($55 doméstico)
    "Rounds Aplicação": "Prioridade e Standard (sem rounds numerados); ver Deadline Aplicação",
    "Documentos Necessários": (
        "Application online, histórico acadêmico oficial de todas as "
        "instituições cursadas, 2 ensaios obrigatórios (motivação/"
        "objetivos + 3 conquistas significativas) + personal statement "
        "opcional, resume/CV, 3 cartas de recomendação"
    ),
    "Salário Médio Pós USD": 98333,  # oficial, salário médio até 3 meses após formatura
    "Principais Empregadores": "Amazon, Google, Meta, IBM, Nike, Walmart, entre outras empresas de tecnologia e consultoria",
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
        raise ValueError("Nenhuma linha encontrada para Texas State (San Marcos, TX). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Texas State, San Marcos, TX)")

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
