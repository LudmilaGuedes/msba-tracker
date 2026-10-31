"""
Data enrichment: atualiza APENAS a linha da 'U of Minnesota' (Carlson School
of Management, Minneapolis, MN) no msba_programs_database.xlsx.

Foco exclusivo: MS in Business Analytics (presencial, STEM). Deep research em:
custo total real de tuition para internacionais, bolsas/assistantships,
requisitos de GPA/GMAT/GRE/inglês, e custo de vida local.

ACHADO DA AUDITORIA: o programa foi renomeado para "MS in Business Analytics
& Artificial Intelligence" (MSBAi), efetivo para a turma que inicia Fall
2027 -- mesmo programa, mesmo departamento (Carlson), currículo com maior
ênfase em IA. Nome oficial atualizado; dados de custo/admissão abaixo já são
os publicados sob o novo nome.

Tuition Total (Intl/Residente) já estava correta ($56.250 / $45.000,
2026-2027) -- confirmada, sem alteração. Novo dado: custo por crédito
($1.250 intl / $1.000 residente x 45 créditos = bate exatamente com os
totais já cadastrados).

Custo de vida (Aluguel, Mercado, Seguro Saúde, Utilities, etc.) mantido como
"Verificar": a página oficial de cost-of-attendance da UMN (onestop.umn.edu)
e a página de tuition-aid da Carlson (carlsonschool.umn.edu) bloquearam
acesso automatizado (HTTP 403) para os valores discriminados por categoria.
O único dado de aluguel encontrado via terceiros é uma média genérica de
mercado para Minneapolis/Dinkytown (não específica de "quarto individual +
banheiro privativo" para pós-graduandos), então não foi usado para não
inventar precisão que a fonte não tem -- mesmo critério já aplicado em
etl_enrich_uc_davis.py.

Fontes principais (pesquisadas manualmente, fora deste script, ago/2026):
- https://carlsonschool.umn.edu/graduate/masters/business-analytics-ai/tuition-aid
  (tuition 2026-27, per-credit rate, fees ~$9.100 intl, depósito $1.500)
- https://carlsonschool.umn.edu/graduate/masters/business-analytics-ai
  (renomeação do programa para MSBAi, efetivo Fall 2027)
- https://carlsonschool.umn.edu/graduate/masters/business-analytics/admissions/requirements-deadlines
  (GPA mínimo 3.0 p/ dispensa de GMAT/GRE, pré-requisitos, testes de inglês,
  documentos, rounds de aplicação, taxa de inscrição $75)
- https://carlsonschool.umn.edu/graduate/masters/business-analytics-ai/faq
  (GMAT médio 691.9 / GRE médio 320.1 / experiência média 3.3 anos --
  class profile do MSBAi)
- https://carlsonschool.umn.edu/graduate/masters/business-analytics-ai/tuition-aid/scholarships
  (bolsa de mérito automática; bolsa garantida p/ quem se inscreve até
  Round 1; TA/RA incomum em programa full-time de 12 meses)
- Kira Talent confirmado como plataforma de vídeo-entrevista assíncrona
  usada pela Carlson no processo de candidatura.
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
    "Nome da Universidade": "U of Minnesota",
    "Estado": "MN",
    "Cidade": "Minneapolis",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Nome Oficial do Programa": "MS in Business Analytics & Artificial Intelligence (MSBAi)",
    # CORRECAO: planilha tinha "MS in Business Analytics" -- programa renomeado,
    # efetivo para a turma que inicia Fall 2027 (mesmo depto/Carlson, mais enfase em IA).
    "Destaques Currículo": (
        "Python, R, ML, AI track, Capstone; renomeado para MSBAi (Business "
        "Analytics & AI) a partir da turma de Fall 2027, com currículo com "
        "maior ênfase em inteligência artificial"
    ),
    "Custo/Crédito (Intl) USD": 1250,  # confirma Tuition Total Intl: 45 x 1250 = 56.250 (já cadastrado, correto)
    "Custo/Crédito (Residente) USD": 1000,  # confirma Tuition Total Residente: 45 x 1000 = 45.000 (já cadastrado, correto)
    "Fees Adicionais USD": 9100,  # estimativa oficial de fees p/ estudante internacional, além da tuition
    "Caminho Tuition Residente": (
        "Não (exceto reciprocidade regional p/ residentes de ND, WI e "
        "Manitoba -- não aplicável a candidatos internacionais)"
    ),
    "GPA Mínimo": 3.0,
    "GMAT/GRE Obrigatório": (
        "Opcional se GPA de graduação ≥ 3.0 (escala 4.0) ou equivalente; "
        "obrigatório enviar GMAT ou GRE Geral se o GPA ficar abaixo desse "
        "mínimo"
    ),
    "GMAT Médio": 691.9,  # class profile MSBAi; GRE médio equivalente = 320.1 (quando enviado)
    "Experiência Profissional": (
        "Não obrigatória (mas média da turma admitida é de ~3,3 anos de "
        "experiência profissional)"
    ),
    "Matérias Pré-Requisito": (
        "Pelo menos 1 curso quantitativo de nível universitário (cálculo "
        "ou estatística) já na inscrição; proficiência em programação "
        "(Python, C, C++, C#, Visual Basic, Rust, Java ou JavaScript) "
        "exigida até a Orientação, não para a admissão"
    ),
    "Testes Inglês Aceitos": (
        "TOEFL/IELTS/PTE -- obrigatório se o candidato não for falante "
        "nativo de inglês ou não tiver se formado em instituição de ensino "
        "em inglês"
    ),
    "Deadline Aplicação": (
        "Round 1: 1/nov; Round 2: 4/jan; Round 3: 1/fev (deadline final "
        "para candidatos internacionais); Round 4: 15/mar"
    ),
    "Taxa Aplicação USD": 75,
    "Rounds Aplicação": "4 rounds (ver Deadline de Aplicação)",
    "Documentos Necessários": (
        "Transcripts, applicant statement (750 palavras / 2 páginas), "
        "resume, 2 cartas de recomendação, GMAT/GRE (dispensável com GPA "
        "≥ 3.0), TOEFL/IELTS/PTE (se aplicável), taxa de inscrição de $75, "
        "vídeo-ensaio via Kira Talent"
    ),
    "Entrevista": (
        "Vídeo-entrevista assíncrona via plataforma Kira Talent, integrada "
        "à candidatura (sem entrevista presencial/Zoom separada confirmada)"
    ),
    "TA/RA Disponível": (
        "Raro -- programa full-time de 12 meses; maioria financia via "
        "bolsa de mérito, ajuda federal, empréstimo privado ou poupança "
        "própria. Vagas de assistente de pós-graduação existem na "
        "universidade, mas exigem busca ativa (não são o caminho típico)"
    ),
    "Processo TA/RA": (
        "Sem processo formal específico do MSBAi; candidatos interessados "
        "devem perguntar ao orientador acadêmico na orientação ou buscar "
        "vagas abertas de graduate assistant no portal geral da "
        "universidade"
    ),
    "Graduate Assistantships": (
        "Incomuns para este programa de 12 meses; vagas gerais da "
        "universidade existem, mas não são o caminho típico de "
        "financiamento do MSBAi"
    ),
    "Tipos de Bolsa": (
        "Merit (automática, sem inscrição separada); candidatos que "
        "completam a inscrição até o Round 1 (1/nov) e são admitidos "
        "recebem bolsa garantida"
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
        raise ValueError("Nenhuma linha encontrada para U of Minnesota (Minneapolis, MN). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (U of Minnesota, Minneapolis, MN)")

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
