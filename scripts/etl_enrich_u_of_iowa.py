"""
Data enrichment: atualiza APENAS a linha da 'U of Iowa' (Tippie
College of Business, Iowa City, IA) no msba_programs_database.xlsx.

Foco: Master of Science in Business Analytics (MSBA), trilha FULL-TIME
on-campus (distinta de outras trilhas do Tippie), 2 semestres +
bootcamp de nivelamento pré-Fall, STEM (confirmado pelo Department of
Homeland Security).

'Total Créditos' CORRIGIDO de 42 para 33 -- a página oficial de
currículo lista o breakdown completo: 7 disciplinas core de 3 créditos
+ 2 de comunicação (3 créditos) + Professional Development (1 crédito
por semestre, 2 no total) + trilha eletiva de 9 créditos (uma entre
3: AI/ML, AI Technology Management, ou Finance Analytics) = 33
créditos totais, não 42.

TUITION: valores ATUALIZADOS -- a página oficial de tuition traz
$28.069 (residente Iowa, 2 semestres) e $40.401 (não-residente/
internacional, 2 semestres), pequena diferença em relação aos valores
já presentes na planilha ($27.250 / $39.777), provavelmente refletindo
um reajuste anual de tuition entre a rodada de pesquisa anterior e
esta. 'Custo/Crédito' preenchido dividindo pelos 33 créditos agora
confirmados: $850,58 (residente) / $1.224,27 (internacional).

'GPA Mínimo' preenchido: 3.0 cumulativo (ou equivalente estrangeiro).
'Testes Inglês Aceitos' refinado com os dois patamares oficiais de
TOEFL (100 antes de jan/2026; 5.0 no novo formato depois), IELTS 7.0
(nenhum sub-score abaixo de 6.0), Duolingo 120+, com dispensa
disponível para quem qualificar (critério não detalhado).

'Taxa Aplicação USD' preenchido com $100 (internacional; doméstico é
$60). 'Deadline Aplicação' e 'Rounds Aplicação' refinados: aplicações
abrem 15 ago; deadline de prioridade (bolsas de mérito) 15 dez;
deadline final internacional 15 mar; deadline final doméstico 15 jun;
entrada apenas no Fall.

'Tipos de Bolsa' refinado com achado IMPORTANTE: o programa declara
EXPLICITAMENTE que "does not offer full tuition scholarships" --
apenas um número limitado de pacotes de bolsa parcial para alunos de
alto mérito, com preferência para quem tem 12+ meses de experiência
pós-graduação (já registrado em 'Experiência Profissional'), GPA
cumulativo ≥3.50, e/ou GMAT ≥600.

'Matérias Pré-Requisito' preenchido: bacharelado em QUALQUER área é
aceito -- diploma em business explicitamente NÃO é exigido.

'Taxa Empregabilidade %' e 'Salário Médio Pós USD' preenchidos com
dados oficiais e específicos do programa, mas da coorte 2023-2024
(mais antiga que o ideal -- documentado como ressalva): 78% empregados
em até 6 meses; salário médio $71.273 (máximo registrado: $107.000).

CUIDADO APLICADO (dados NÃO encontrados, permanecem 'Verificar'):
'Tamanho Turma' (só descrito qualitativamente como "small cohorts"),
'GPA Médio Admitidos', 'GMAT Médio', 'Taxa Aceitação %', 'Principais
Empregadores' (a página de careers não lista nomes específicos, só
estatísticas agregadas sobre recrutadores Fortune Global 100),
'Ranking US News', 'Ranking QS' (a única classificação encontrada foi
"#15 Most Affordable" do TechGuide 2025, uma fonte não-tradicional que
não corresponde ao formato esperado desses dois campos -- não usada),
'Entrevista', 'Documentos Necessários' (a página só confirma
bacharelado + GPA + proficiência em inglês como base, sem lista
detalhada de documentos adicionais), 'Seguro Saúde (programa) USD',
'TA/RA Disponível'/'Processo TA/RA'/'Graduate Assistantships' (não
mencionados nas páginas consultadas), e todos os campos mensais de
custo de vida.

Fontes principais (pesquisadas via WebFetch, ago/2026):
- https://tippie.uiowa.edu/masters-business-analytics/full-time
  (nome oficial confirmado, formato, STEM, tuition, outcomes 2023-24)
- .../full-time/admissions (GPA, GMAT/GRE opcional, mínimos de
  inglês, deadlines, taxas de aplicação, "business degree not
  required")
- .../full-time/tuition-aid (tuition detalhada, bolsas parciais
  apenas — sem bolsa integral —, critérios de mérito)
- .../full-time/curriculum (créditos totais corretos: 33, breakdown
  completo por trilha)
- .../full-time/careers (78% empregabilidade, $71.273 salário médio,
  sem lista de empregadores nomeados)
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
    "Nome da Universidade": "U of Iowa",
    "Estado": "IA",
    "Cidade": "Iowa City",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Total Créditos": 33,  # CORREÇÃO: oficial (breakdown completo confere), não 42
    "Custo/Crédito (Residente) USD": 850.58,  # 28.069 / 33
    "Custo/Crédito (Intl) USD": 1224.27,  # 40.401 / 33
    "Tuition Total (Residente) USD": 28069,  # ATUALIZADO: oficial, era 27.250
    "Tuition Total (Intl) USD": 40401,  # ATUALIZADO: oficial, era 39.777
    "GPA Mínimo": "3.0 cumulativo (ou equivalente estrangeiro)",
    "Testes Inglês Aceitos": (
        "TOEFL iBT mín. 100 (antes de jan/2026) ou 5.0 (novo "
        "formato, depois); IELTS mín. 7.0 (nenhum sub-score abaixo "
        "de 6.0); Duolingo mín. 120; dispensas disponíveis para quem "
        "qualificar"
    ),
    "Taxa Aplicação USD": 100,  # internacional (doméstico: $60)
    "Deadline Aplicação": "Aplicações abrem 15 ago; prioridade (bolsas de mérito) 15 dez; final internacional 15 mar; final doméstico 15 jun; entrada apenas Fall",
    "Rounds Aplicação": "Prioridade (15 dez, para bolsas de mérito) + deadline final por status (15 mar intl / 15 jun doméstico)",
    "Tipos de Bolsa": (
        "NÃO oferece bolsa integral de tuition (confirmado "
        "oficialmente) — apenas pacotes parciais limitados para "
        "alto mérito, com preferência para 12+ meses de experiência "
        "pós-graduação, GPA ≥3.50, e/ou GMAT ≥600"
    ),
    "Matérias Pré-Requisito": "Bacharelado em qualquer área aceito; diploma em business explicitamente NÃO exigido",
    "Taxa Empregabilidade %": 78,  # coorte 2023-2024, oficial e específico do programa (dado potencialmente desatualizado)
    "Salário Médio Pós USD": 71273,  # coorte 2023-2024, oficial (máximo registrado: $107.000)
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
        raise ValueError("Nenhuma linha encontrada para U of Iowa (Iowa City, IA). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (U of Iowa, Iowa City, IA)")

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
