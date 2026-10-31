"""
Data enrichment: atualiza APENAS a linha da 'UMass Amherst' (Isenberg
School of Management, Amherst, MA) no msba_programs_database.xlsx.

Foco exclusivo: MS in Business Analytics, formato on-campus (presencial),
STEM, 31 créditos (11 disciplinas). Deep research em: custo total real
de tuition para internacionais, bolsas/assistantships, requisitos de
GPA/GMAT/GRE/inglês, e custo de vida em Amherst, MA.

TUITION: a página oficial de tuition do formato on-campus confirma
$925/crédito, SEM distinção entre residentes e não-residentes ("no
distinction... between Massachusetts residents and non-residents") --
ou seja, é uma tarifa FLAT, igual para todos. PREENCHIDOS 'Custo/
Crédito (Residente)' e 'Tuition Total (Residente)' (estavam
'Verificar') com os mesmos valores já presentes para Intl ($925/
crédito, $28.675 total). 'Fees Adicionais' preenchido somando as taxas
obrigatórias oficiais: Graduate entering fee $507 (única vez) +
Registration fee $85 x 3 semestres + Student health fee $438,50 x 2
semestres (Fall+Spring) + Graduate Senate Tax $86 = $1.725 (não inclui
a fee de Field Experience Concentration, que é opcional/especifica de
uma trilha eletiva, não uma fee geral do programa).

ACHADO CRITICO: a mesma pagina oficial afirma explicitamente "Graduate
assistantships are not available to MSBA students" -- ou seja, NÃO HÁ
assistantships para este programa (mesmo padrão já encontrado em
Binghamton SUNY nesta planilha). CORRIGIDO/PREENCHIDO 'TA/RA
Disponível', 'Processo TA/RA' e 'Graduate Assistantships' refletindo
isso. Também não foi encontrada nenhuma bolsa especifica do MSBA (só
desconto de 10% para ex-alunos e patrocinio por empregador terceiro) --
'Bolsas Disponíveis' e 'Quão Realista (Intl)' preenchidos de forma
correspondente (Baixa).

Custo de vida: 'Transporte/mês' preenchido com $0 -- o PVTA (sistema de
onibus da regiao) e gratuito para alunos da UMass Amherst com UCard, e
esta oferecendo tarifa zero para todos os passageiros ate 30/jun/2026.
'Seguro Saúde (programa)' preenchido com $2.480 -- taxa oficial do
Student Health Benefit Plan (SHBP), $1.240/semestre x 2 (Fall+Spring),
obrigatoria para todos os internacionais. 'Aluguel Privativo/mês'
preenchido com $800 (quarto privativo mobiliado, multiplos anuncios
convergentes, muitos ja incluindo utilities). 'Utilities/mês' ($96) e
'WiFi/mês' ($32) estimados a partir das faixas anuais encontradas
($800-1.500/ano utilities; $300-480/ano internet). 'Mercado/mês',
'Seguro Residencial/mês', 'Celular/mês' e 'Outros Gastos/mês'
permanecem 'Verificar' -- sem fonte especifica -- e por isso 'Custo
Vida Total (programa) USD' também permanece 'Verificar'.

Fontes principais (pesquisadas via busca web, ago/2026):
- https://www.isenberg.umass.edu/programs/masters/ms/business-analytics/
  on-campus/tuition-aid (tuition oficial flat $925/credito, fees
  obrigatorias, confirmacao de que assistantships NAO estao
  disponiveis, desconto de alumni, patrocinio por empregador)
- https://www.isenberg.umass.edu/programs/masters/ms/business-analytics/
  on-campus/admissions (deadlines Early Action/Regular/Final por
  semestre de Fall/Spring/Summer, documentos exigidos, GMAT/GRE waiver
  com GPA>=3.0, taxa de aplicacao $90 reembolsada para Early Action)
- Busca web: GPA minimo 3.0, GMAT medio 620, TOEFL 100/IELTS 7.0/
  Duolingo 120/PTE 68 (paginas de admissao online, mesmos criterios
  aplicaveis ao on-campus)
- Busca web: PVTA gratuito para alunos UMass Amherst (UCard); tarifa
  zero geral ate 30/jun/2026
- Busca web: Student Health Benefit Plan (SHBP) $1.240/semestre para
  internacionais (umass.edu/uhs)
- Busca web: aluguel de quarto privativo mobiliado em Amherst, MA
  (multiplos anuncios convergentes em torno de $700-850)
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
    "Nome da Universidade": "UMass Amherst",
    "Estado": "MA",
    "Cidade": "Amherst",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Intl) USD": 925,  # confirma valor oficial já implícito no Tuition Total
    "Custo/Crédito (Residente) USD": 925,  # flat — sem distinção residente/não-residente (oficial)
    "Tuition Total (Residente) USD": 28675,  # flat — mesmo valor do Intl (oficial)
    "Fees Adicionais USD": 1725,  # Graduate entering fee + Registration + Student health fee + Senate Tax
    "Aluguel Privativo/mês USD": 800,  # quarto privativo mobiliado, Amherst MA (múltiplos anúncios convergentes)
    "Utilities/mês USD": 96,  # faixa anual oficial $800-1.500 ÷ 12, ponto médio
    "WiFi/mês USD": 32,  # faixa anual $300-480 ÷ 12
    "Transporte/mês USD": 0,  # PVTA gratuito para alunos UMass Amherst (UCard); tarifa zero geral até 30/jun/2026
    "Seguro Saúde (programa) USD": 2480,  # Student Health Benefit Plan (SHBP): $1.240/semestre x 2 (Fall+Spring)
    "TA/RA Disponível": "Não — assistantships de pós-graduação não estão disponíveis para alunos do MSBA (confirmado oficialmente)",
    "Processo TA/RA": "N/A — sem assistantships neste programa",
    "Graduate Assistantships": (
        "Não disponíveis para MSBA (confirmado oficialmente); "
        "patrocínio por empregador terceiro e desconto de 10% para "
        "ex-alunos são as alternativas de redução de custo"
    ),
    "Caminho Tuition Residente": "N/A — tuition já é flat (mesma taxa para residentes e não-residentes)",
    "Bolsas Disponíveis": "Não identificadas bolsas específicas do MSBA",
    "Tipos de Bolsa": (
        "Nenhuma bolsa específica do MSBA encontrada; desconto de 10% "
        "para ex-alunos da UMass; patrocínio por empregador terceiro; "
        "benefícios para militares"
    ),
    "Quão Realista (Intl)": "Baixa — sem assistantships nem bolsas específicas identificadas para internacionais",
    "GPA Mínimo": 3.0,
    "GMAT/GRE Obrigatório": (
        "Waiver automático com GPA de graduação ≥ 3.0 (deixar a seção "
        "de teste em branco na aplicação); pode ser exigido caso "
        "contrário, com possibilidade de waiver por experiência "
        "profissional"
    ),
    "GMAT Médio": 620,
    "Testes Inglês Aceitos": "TOEFL mín. 100 / IELTS Academic mín. 7.0 / Duolingo mín. 120 / PTE Academic mín. 68",
    "Matérias Pré-Requisito": "Cálculo, estatística ou probabilidade recomendados, mas não obrigatórios",
    "Deadline Aplicação": (
        "Fall: Early Action 1º fev, Regular 1º abr, Final 31 jul; "
        "Spring: Early Action 1º ago, Regular 1º out, Extended 15 nov; "
        "Summer: Early Action 1º nov, Regular 1º jan, Extended 1º abr "
        "(doméstico)"
    ),
    "Taxa Aplicação USD": 90,  # reembolsada na primeira fatura de tuition para candidatos Early Action
    "Rounds Aplicação": "Early Action, Regular, Final/Extended (ver Deadline Aplicação)",
    "Documentos Necessários": (
        "Application online, personal statement (1-2 páginas), resume/"
        "CV, histórico acadêmico oficial, 2 cartas de recomendação, "
        "GMAT/GRE (waiver possível), teste de inglês (internacionais)"
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
        raise ValueError("Nenhuma linha encontrada para UMass Amherst (Amherst, MA). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (UMass Amherst, Amherst, MA)")

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
