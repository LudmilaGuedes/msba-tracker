# -*- coding: utf-8 -*-
"""
Data enrichment (round 2): atualiza APENAS a linha 'University of Dallas'
(Satish & Yasmin Gupta College of Business, Irving, TX) no
msba_programs_database.xlsx.

CONTEXTO: rodada 1 (scripts/etl_enrich_university_of_dallas.py) ja
confirmou muito bem tuition, fees, bolsas, TA/RA, GMAT/GRE, testes de
ingles, taxa de aplicacao, deadlines e documentos. Esta rodada tentou
preencher os campos ainda "Verificar" (duracao, tamanho de turma,
rankings, empregabilidade/salario, GPA/GMAT medio, taxa de aceitacao,
pre-requisitos, entrevista), mas o resultado foi MAJORITARIAMENTE
NEGATIVO -- o site institucional simplesmente nao divulga a maioria
desses dados por programa individual.

TENTADO E NAO CONFIRMADO (mantido "Verificar", nao por falha de busca mas
por nao-divulgacao institucional confirmada apos 3 paginas oficiais
tentadas):
- Duração (meses): so confirmado oficialmente "programa de 30 creditos",
  sem duracao em meses/trimestres para tempo integral divulgada em
  nenhuma pagina oficial. Um agregador terceiro (mastersportal) sugeriu
  "1-2 anos part-time profissional", mas e estimativa de terceiro, nao
  fonte oficial, e nao cobre tempo integral -- nao usado.
- Tamanho Turma: so achada mencao generica a "small cohort structure" do
  college inteiro, sem numero especifico do MSBA.
- Ranking US News/QS: nao encontrado especifico do MSBA.
- Taxa Empregabilidade %, Salário Médio Pós, Principais Empregadores:
  ACHADO EXISTE mas NAO USADO -- a pagina oficial "Graduate Students By
  The Numbers" tem dados reais (salario medio $90.501, mediana $80.000),
  mas sao dados AGREGADOS de todo MBA/MS do Gupta College JUNTOS, sem
  separacao por programa e sem mencao especifica a "Business Analytics"
  -- usar misturaria dados de programas diferentes, o que o projeto
  proibe.
- GMAT Médio, Taxa de Aceitação %, Entrevista: nao encontrados.

ATUALIZADO -- Matérias Pré-Requisito: estava "Verificar". A pagina
oficial de admissoes do college (udallas.edu/admissions-aid/graduate/
college-of-business/) lista pre-requisitos de disciplinas especificos
APENAS para MS Accounting, MS Finance e MS Financial Planning -- o MS
Business Analytics NAO aparece nessa lista. Isso e uma INFERENCIA
razoavel (ausencia de mencao, nao confirmacao textual direta do tipo
"nao ha pre-requisito"), mas suficientemente indicativa para registrar
em vez de deixar puramente "Verificar", com essa ressalva explicita.

NOTA (nao usada para atualizar campo, apenas registrada aqui): STEM
Designated reconfirmado com elegibilidade a extensao de OPT STEM padrao
de 24 meses -- isso e o beneficio STEM padrao (ja implicito em "STEM
Designated: Sim"), entao nao foi adicionado como novidade a nenhum campo.

RISCO EVITADO: varias buscas retornaram resultados da UT DALLAS (Jindal
School) -- uma universidade DIFERENTE, apesar do nome parecido. Nenhum
dado da UT Dallas foi usado nesta linha (University of Dallas, privada,
Irving, TX -- nao confundir com University of Texas at Dallas, publica,
Richardson, TX, ja documentada em linha separada desta base).

Fontes principais (pesquisadas via WebSearch/WebFetch, set/2026):
- https://udallas.edu/academics/programs/business-analytics/index.php (oficial)
- https://udallas.edu/academics/programs/business-analytics/more-infornation.php (oficial)
- https://udallas.edu/admissions-aid/graduate/college-of-business/index.php
  (oficial -- lista de pré-requisitos por programa, MSBA ausente)
- https://udallas.edu/life-at-ud/student-support-resources/personal-career-development/success-rate/graduate-outcomes.php
  (oficial -- dados agregados MBA/MS, não usados por não serem específicos do MSBA)
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "University of Dallas",
    "Estado": "TX",
    "Cidade": "Irving",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Matérias Pré-Requisito": (
        "Provavelmente nenhum pré-requisito formal de disciplinas -- a "
        "página oficial de admissões do Gupta College lista "
        "pré-requisitos específicos apenas para MS Accounting, MS "
        "Finance e MS Financial Planning; o MS Business Analytics NÃO "
        "aparece nessa lista (inferência por ausência, não confirmação "
        "textual direta)"
    ),
}


def main():
    rotate_backup()
    print("Backup rotacionado (política fixa) em: data/msba_programs_database_backup.xlsx")

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
        raise ValueError("Nenhuma linha encontrada para University of Dallas (Irving, TX). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (University of Dallas, Irving, TX)")

    no_fill = PatternFill(fill_type=None)

    before = {}
    after = {}
    for header, new_value in UPDATES.items():
        col_idx = header_to_col[header]
        cell = ws.cell(row=row_idx, column=col_idx)
        before[header] = cell.value
        cell.value = new_value
        cell.fill = no_fill  # nenhum novo valor aqui e "Verificar"
        after[header] = new_value

    wb.save(DB_PATH)

    print("\nCampos atualizados (Como Estava -> Como Ficou):")
    for header in UPDATES:
        print(f"  - {header}:")
        print(f"      Antes: {before[header]!r}")
        print(f"      Depois: {after[header]!r}")


if __name__ == "__main__":
    main()
