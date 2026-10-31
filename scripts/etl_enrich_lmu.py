"""
Data enrichment: atualiza APENAS a linha da 'Loyola Marymount
University (LMU)' (LMU College of Business Administration, Los
Angeles, CA) no msba_programs_database.xlsx.

Foco: MS in Business Analytics (MSBA), formato presencial full-time
(12 meses, distinto da trilha part-time de 24 meses), 30-36 unidades
dependendo de coursework prévio de business (já confirmado na
planilha), STEM.

'Início' preenchido: 'Fall' -- o ciclo de deadlines encontrado é
explicitamente rotulado "Fall 2027", confirmando entrada no outono.

'Tamanho Turma' preenchido: 13 (tamanho médio de turma, oficial),
proporção aluno-professor de 11:1 (mencionada no docstring, sem campo
próprio no schema).

'Destaques Currículo' preenchido: foco em data management, big data
modeling & analytics, e aplicações de negócio; eletivas em Data
Visualization/GIS, Text Mining/Social Media, Emerging Technologies,
Marketing, Healthcare.

'GMAT/GRE Obrigatório' refinado: NÃO exigido, mas recomendado para
candidatos com GPA abaixo de 3.3 para demonstrar habilidades
quantitativas (códigos institucionais: GMAT MSX-XN-89, GRE 4403).

'Matérias Pré-Requisito' preenchido: curso de estatística em nível
universitário concluído nos últimos 6 anos, nota B ou superior;
candidatos sem esse pré-requisito são orientados a cursar estatística
antes de completar a aplicação.

'Taxa Aplicação USD' preenchido: $50. 'Documentos Necessários'
preenchido: resume; personal statement; 2 cartas de recomendação;
transcripts de todas as instituições; entrevista por convite.
'Entrevista' preenchido: por convite apenas (confirmado oficialmente,
não é universal para todos os candidatos).

'Deadline Aplicação' e 'Rounds Aplicação' REFINADOS com a estrutura
oficial de 3 rounds para o ciclo Fall 2027 (mais precisa que o valor
anterior "Rolling", mantendo a nuance de que a admissão continua
sendo aceita após o último round até a capacidade ser atingida):
Round 1 (10 jan 2027); Round 2 (10 mar 2027); Round 3 -- internacional
(15 mai 2027) / doméstico (31 mai 2027); aceitação contínua depois
disso até completar as vagas.

'Tipos de Bolsa' refinado com detalhe oficial: 80% dos alunos de
graduate business recebem alguma bolsa; consideração automática para
todos que aplicam até o deadline final (sem aplicação de bolsa
separada); critérios incluem histórico acadêmico, test scores (se
enviados) e experiência profissional; elegibilidade CONFIRMADA tanto
para domésticos quanto internacionais.

'Principais Empregadores' preenchido com uma amostra da lista oficial
de empregadores de egressos: Activision Blizzard, Amazon, Creative
Artists Agency, Deloitte, Disney, Intuit, ServiceNow, Target.

CUIDADO APLICADO (dados NÃO encontrados, permanecem 'Verificar'):
'Testes Inglês Aceitos' JÁ ESTAVA preenchido com "TOEFL/IELTS exigido"
(genérico, não Verificar) -- nenhuma fonte nova trouxe mínimos
numéricos específicos para substituir essa nota genérica, então não
foi alterado. 'GPA Médio Admitidos', 'GMAT Médio', 'Taxa Aceitação %',
'Ranking US News', 'Ranking QS', 'Taxa Empregabilidade %', 'Salário
Médio Pós USD' (múltiplas páginas de career/outcomes consultadas não
trazem nenhuma estatística quantitativa de emprego ou salário, apenas
qualitativas), 'Fees Adicionais USD', 'Seguro Saúde (programa) USD',
'Processo TA/RA', 'TA/RA Disponível', 'Caminho Tuition Residente'
(a linha já indica tuition idêntica para Intl/Residente via valores
já preenchidos, mas o campo de texto explicativo em si não foi
confirmado por fonte nova), 'Experiência Profissional', e todos os
campos mensais de custo de vida em Los Angeles, CA.

Fontes principais (pesquisadas via WebFetch, ago/2026):
- https://cba.lmu.edu/academics/grad/analytics/ (formato, duração,
  créditos, tamanho de turma, currículo, deadlines dos 3 rounds, taxa
  de aplicação, documentos, entrevista por convite, bolsas 80%)
- .../admission/ (GMAT/GRE não exigido com recomendação condicional,
  pré-requisito de estatística detalhado)
- .../career/ e .../career/careerpath/ (sem estatísticas
  quantitativas de emprego/salário; lista de empregadores de egressos)
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
    "Nome da Universidade": "Loyola Marymount University (LMU)",
    "Estado": "CA",
    "Cidade": "Los Angeles",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Início": "Fall",  # oficial, ciclo "Fall 2027" confirmado
    "Tamanho Turma": 13,  # oficial, tamanho médio de turma
    "Destaques Currículo": (
        "Data management, big data modeling & analytics, aplicações "
        "de negócio; eletivas em Data Visualization/GIS, Text Mining/"
        "Social Media, Emerging Technologies, Marketing, Healthcare"
    ),
    "GMAT/GRE Obrigatório": "Não exigido; recomendado para GPA <3.3 (códigos: GMAT MSX-XN-89, GRE 4403)",
    "Matérias Pré-Requisito": (
        "Curso de estatística em nível universitário nos últimos 6 "
        "anos, nota B ou superior; se ausente, recomenda-se cursar "
        "antes de aplicar"
    ),
    "Taxa Aplicação USD": 50,
    "Documentos Necessários": "Resume; personal statement; 2 cartas de recomendação; transcripts de todas as instituições; entrevista por convite",
    "Entrevista": "Por convite apenas (confirmado oficialmente, não universal)",
    "Deadline Aplicação": "Fall 2027: Round 1 (10 jan); Round 2 (10 mar); Round 3 — internacional (15 mai) / doméstico (31 mai); aceitação contínua depois até completar vagas",
    "Rounds Aplicação": "3 rounds (10 jan, 10 mar, 15/31 mai por status) + aceitação contínua até completar vagas",
    "Tipos de Bolsa": (
        "80% dos alunos de graduate business recebem bolsa; "
        "consideração automática ao aplicar até o deadline final "
        "(sem aplicação separada); critérios: histórico acadêmico, "
        "test scores (se enviados), experiência profissional; "
        "elegível para domésticos E internacionais"
    ),
    "Principais Empregadores": "Activision Blizzard, Amazon, Creative Artists Agency, Deloitte, Disney, Intuit, ServiceNow, Target",
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
        raise ValueError("Nenhuma linha encontrada para Loyola Marymount University (LMU) (Los Angeles, CA). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Loyola Marymount University (LMU), Los Angeles, CA)")

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
