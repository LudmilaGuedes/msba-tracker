"""
Data enrichment: atualiza APENAS a linha da 'University of Illinois
Chicago (UIC)' (UIC Business / Liautaud Graduate School, Chicago, IL)
no msba_programs_database.xlsx.

Foco: Master of Science in Business Analytics, presencial, 32 créditos
(já confirmado), STEM. Formato flexível full-time (1 ano) ou part-time
(até 2 anos).

'Custo/Crédito (Intl) USD' preenchido por DERIVAÇÃO do total Intl já
presente e não alterado ($42.532 ÷ 32 créditos = $1.329,13) -- os
documentos oficiais de tuition (hospedados em Box.com, PDFs Fall 2026/
Spring 2027) estavam atrás de um visualizador JavaScript que o
WebFetch não conseguiu processar em nenhuma tentativa, impedindo
confirmar o valor residente ou refinar/corrigir o Intl com uma fonte
primária nesta rodada.

'GMAT/GRE Obrigatório' preenchido: test-optional (confirmado
oficialmente). 'GPA Mínimo' preenchido: 3.0/4.0 recomendado para
domésticos (não necessariamente um piso rígido de corte).

'Taxa Aplicação USD' preenchido: $80 (Fall 2026 em diante; era $70
para Summer 2026 ou anterior) -- dispensada integralmente para alunos/
ex-alunos da UIC. Taxa adicional de $100 para avaliação de credencial
internacional, se aplicável (não incluída neste campo por ser
condicional, não universal).

'Documentos Necessários' preenchido: aplicação via UIC Apply;
transcripts acadêmicos; 2 cartas de recomendação profissionais;
personal statement sobre objetivos de carreira; resume/CV; GMAT/GRE
(requisitos variam por programa -- para o MSBA, opcional); taxa de
avaliação de credencial internacional de $100 (se aplicável).

'Entrevista' preenchido com achado IMPORTANTE: a página oficial de
admissões declara explicitamente que entrevista é exigida "MBA
applicants only" -- ou seja, NÃO se aplica ao MSBA, confirmado como
negativa oficial, não uma lacuna de pesquisa.

'Matérias Pré-Requisito' preenchido: a própria UIC Business confirma
que "a business-related undergraduate degree is not required" --
diploma de graduação em business NÃO é exigido, e não há
pré-requisitos formais de disciplinas específicas mencionados.

CUIDADO APLICADO: a página oficial do programa lista um ciclo de
deadlines diferente do já presente na planilha (Spring 2027
Internacional: 1 set 2026; Spring 2027 Doméstico: 1 nov 2026; Summer
2027 Doméstico: 15 mar 2027 -- sem um deadline de "Fall" explícito
nesta fonte específica). Como o campo 'Deadline Aplicação' já estava
preenchido (não 'Verificar') com "Fall 1 abr 2026; Spring 1 set 2026",
e as duas fontes não são diretamente contraditórias (podem
simplesmente cobrir ciclos/anos diferentes de admissão), o valor
existente foi MANTIDO sem alteração, e essa nova informação fica
registrada aqui para referência futura.

CUIDADO APLICADO (dados NÃO encontrados, permanecem 'Verificar'):
'Custo/Crédito (Residente)', 'Tuition Total (Residente)', 'Fees
Adicionais USD', 'Seguro Saúde (programa) USD' (documentos oficiais de
tuition inacessíveis via Box.com nesta rodada), 'Tamanho Turma', 'GPA
Médio Admitidos', 'GMAT Médio', 'Taxa Aceitação %', 'Ranking US News',
'Ranking QS', 'Taxa Empregabilidade %', 'Salário Médio Pós USD',
'Principais Empregadores', 'Testes Inglês Aceitos' JÁ ESTAVA
preenchido com uma nota já hedged ("médias admitidos — Verificar
mínimos oficiais"), não Verificar, não tocado por não haver mínimos
oficiais numéricos novos para substituí-la. 'Caminho Tuition
Residente', 'Processo TA/RA', 'TA/RA Disponível', 'Bolsas Disponíveis'/
'Tipos de Bolsa' (a página do programa menciona genericamente
"scholarships and partial tuition waivers" sem detalhamento), e todos
os campos mensais de custo de vida em Chicago, IL.

Fontes principais (pesquisadas via WebFetch, ago/2026):
- https://business.uic.edu/graduate/degrees/master-science-business-analytics/
  (formato, duração, 32 créditos confirmados, STEM, currículo,
  GMAT/GRE test-optional, deadlines de um ciclo específico)
- https://business.uic.edu/graduate/admissions/ (GPA recomendado,
  proficiência em inglês genérica, taxa de aplicação $80/$70, taxa de
  avaliação de credencial internacional, documentos, entrevista
  restrita a MBA, ausência de pré-requisito de diploma em business)
- https://registrar.uic.edu/financial-matters/graduate-tuition-and-fees/
  e os PDFs vinculados no Box.com -- inacessíveis via WebFetch nesta
  rodada (visualizador JavaScript sem conteúdo extraível)
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
    "Nome da Universidade": "University of Illinois Chicago (UIC)",
    "Estado": "IL",
    "Cidade": "Chicago",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Intl) USD": 1329.13,  # 42.532 / 32 (total já presente, não alterado)
    "GMAT/GRE Obrigatório": "Test-optional (confirmado oficialmente)",
    "GPA Mínimo": "3.0/4.0 recomendado para domésticos (não necessariamente um piso rígido)",
    "Taxa Aplicação USD": 80,  # Fall 2026+ (era $70 para Summer 2026 ou anterior); dispensada para alunos/ex-alunos UIC
    "Documentos Necessários": (
        "Aplicação via UIC Apply; transcripts acadêmicos; 2 cartas de "
        "recomendação profissionais; personal statement sobre "
        "objetivos de carreira; resume/CV; GMAT/GRE (opcional para o "
        "MSBA); taxa de avaliação de credencial internacional $100 "
        "(se aplicável)"
    ),
    "Entrevista": "Exigida apenas para candidatos de MBA — NÃO se aplica ao MSBA (confirmado oficialmente)",
    "Matérias Pré-Requisito": "Diploma de graduação em business NÃO é exigido (confirmado oficialmente); sem pré-requisitos formais de disciplinas específicas",
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
        raise ValueError("Nenhuma linha encontrada para University of Illinois Chicago (UIC) (Chicago, IL). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (University of Illinois Chicago (UIC), Chicago, IL)")

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
