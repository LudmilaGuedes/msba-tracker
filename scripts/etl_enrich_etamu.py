"""
Data enrichment: atualiza APENAS a linha da 'East Texas A&M University
(ETAMU)' (College of Business, Commerce, TX) no msba_programs_database.xlsx.

Foco exclusivo: Master of Science in Business Analytics, presencial, STEM,
30 creditos, 18 meses. Deep research em: custo total real de tuition para
internacionais, bolsas/assistantships, requisitos de GPA/GMAT/GRE/ingles,
e custo de vida em Commerce, TX. Nota: a universidade foi renomeada de
"Texas A&M University-Commerce" para "East Texas A&M University" em 2025;
o nome ja estava correto na planilha.

AUDITORIA DE TUITION (ja estava correto, apenas verificado): a pagina
oficial (etamu.edu/programs/business-analytics-ms/) lista "Per Semester
(9 credit hours): Texas residents $3.851, Non-residents $7.703" -- dividido
por 9 creditos da $427,89/credito (residente) e $855,89/credito
(internacional), batendo EXATAMENTE com os valores ja presentes na
planilha ($427.89 / $855.89, $12.837 / $25.677 para 30 creditos). Nenhuma
correcao necessaria nesses campos.

ACHADO RELEVANTE: 'Caminho Tuition Residente' estava 'Verificar'. A pagina
de "Graduate Assistant Tuition Remission Program" (inside.tamuc.edu)
confirma que alunos contratados como GAT/GAR (Graduate Assistant Teaching/
Research) recebem remissao de tuition calculada pela taxa de RESIDENTE
("resident, designated graduate tuition"), independente da residencia
real do aluno -- e o mecanismo classico do Texas Education Code para
TAs/RAs em universidades publicas do Texas. A remissao cobre ate 6h de
matricula por semestre; fees NAO estao inclusas. CORRIGIDO/PREENCHIDO.

Custo de vida: 'Aluguel Privativo/mes' preenchido com $695 (quarto
privativo mobiliado listado em Commerce, TX -- fonte: agregadores de
aluguel estudantil pesquisados). Demais campos de custo de vida mensal
(Mercado, Seguro Saude, Seguro Residencial, Utilities, Celular, WiFi,
Transporte, Outros Gastos) permanecem 'Verificar' -- as unicas fontes
encontradas eram medias amplas (estaduais do Texas, ou de outros campi
do sistema TAMUS) sem numero especifico e confiavel para Commerce/ETAMU,
entao 'Custo Vida Total (programa) USD' tambem permanece 'Verificar'.

'Salario Medio Pos USD' e 'Principais Empregadores' permanecem
'Verificar': o unico dado salarial encontrado ($93.138) e do programa de
MBA (que tem uma concentracao em Business Analytics, mas e um diploma
diferente do MS in Business Analytics-alvo), entao nao foi usado para
nao misturar outcomes de programas distintos.

Fontes principais (pesquisadas via busca web, ago/2026):
- https://www.etamu.edu/programs/business-analytics-ms/ (programa, STEM,
  admissao, TOEFL/IELTS/PTE/Duolingo, tuition por semestre, taxa de
  aplicacao, Quick Start Grant, GAT/GAR/GANT, deadlines)
- https://www.etamu.edu/waivers-and-exemptions/ e
  https://inside.tamuc.edu/academics/graduateschool/funding/assistantships/
  tuitionRemission.aspx (remissao de tuition para GAT/GAR a taxa de
  residente)
- Busca web: GPA minimo institucional de 3.0 para admissao plena em
  pos-graduacao (graduateFaq/graduate school pages)
- Busca web: aluguel de quarto privativo mobiliado em Commerce, TX (~$695/
  mes) e faixas gerais de aluguel/utilities da cidade (agregadores)
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
    "Nome da Universidade": "East Texas A&M University (ETAMU)",
    "Estado": "TX",
    "Cidade": "Commerce",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Aluguel Privativo/mês USD": 695,  # quarto privativo mobiliado, Commerce TX (fonte: agregadores de aluguel)
    "Bolsas Disponíveis": "Sim",
    "Tipos de Bolsa": (
        "Quick Start Grant ($250–$500, renovável, para novos alunos de "
        "mestrado); bolsas adicionais possíveis após admissão, avaliadas "
        "caso a caso"
    ),
    "Quão Realista (Intl)": (
        "Baixa a Moderada — bolsas identificadas são modestas "
        "($250–$500); nenhuma bolsa integral encontrada para "
        "internacionais"
    ),
    "TA/RA Disponível": "Sim (Graduate Assistantships: GAT — ensino, GAR — pesquisa, GANT — não-ensino)",
    "Processo TA/RA": (
        "Sem inscrição centralizada separada conhecida; contratação por "
        "departamentos/professores; para GAT, alunos com inglês como "
        "língua não-nativa devem comprovar proficiência oral e escrita"
    ),
    "Caminho Tuition Residente": (
        "Sim — alunos contratados como GAT/GAR recebem remissão de "
        "tuition calculada à taxa de residente (Graduate Assistant "
        "Tuition Remission Program), válida apenas no semestre do "
        "contrato, cobrindo até 6h de matrícula; fees não inclusas"
    ),
    "Graduate Assistantships": (
        "GAT/GAR/GANT disponíveis via departamentos; remissão de tuition "
        "à taxa de residente para até 6h/semestre durante o período de "
        "contrato (fees não cobertas)"
    ),
    "GPA Mínimo": 3.0,  # padrão institucional para admissão plena em pós-graduação
    "Testes Inglês Aceitos": "TOEFL iBT mín. 79 / IELTS mín. 6.0 / PTE mín. 53 / Duolingo mín. 105",
    "Taxa Aplicação USD": 75,  # internacional ($50 doméstico)
    "Rounds Aplicação": "Admissão contínua por prazos semestrais (sem rounds numerados); ver Deadline Aplicação",
    "Documentos Necessários": (
        "Application via ApplyTexas, histórico acadêmico oficial, "
        "personal statement (até 500 palavras), resume/CV, 1 carta de "
        "recomendação, comprovante financeiro do patrocinador (saldo "
        "bancário mín. ~$33.601 USD, internacional), scores de "
        "proficiência em inglês (internacional)"
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
        raise ValueError("Nenhuma linha encontrada para East Texas A&M University (Commerce, TX). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (East Texas A&M University, Commerce, TX)")

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
