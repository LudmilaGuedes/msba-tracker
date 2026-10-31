"""
Data enrichment: atualiza APENAS a linha da 'Temple' (Fox School of
Business, Philadelphia, PA) no msba_programs_database.xlsx.

Foco exclusivo: MS in Business Analytics, presencial, STEM, 30 créditos
(10 disciplinas), formato full-time ou part-time (até 5 anos para
concluir). Deep research em: custo total real de tuition para
internacionais, bolsas/assistantships, requisitos de GPA/GMAT/GRE/
inglês, e custo de vida em Philadelphia, PA.

TUITION: já estava correto e verificado (tarifa flat $1.165/crédito,
$34.950 total, igual para residentes e não-residentes -- confirmado
oficialmente). Nenhuma correção necessária.

'Seguro Saúde (programa)' preenchido com $2.760 -- prêmio anual oficial
do plano padrão de seguro saúde de Temple (ciclo 1º ago/2025 - 31 jul/
2026). Outros valores encontrados (até $6.118/ano) pareciam referir-se
a coberturas mais amplas ou de fontes menos diretas; o valor de $2.760
é o citado como prêmio oficial do plano padrão.

'Tipos de Bolsa' preenchido com a única bolsa nomeada encontrada: 5% de
desconto na tuition para EX-ALUNOS de Temple admitidos ao programa
(fundos limitados, baseados em mérito) -- não é uma bolsa geral aberta
a todo candidato internacional, por isso 'Quão Realista (Intl)' foi
preenchido como Baixa.

'TA/RA Disponível' preenchido com CAUTELA: existem teaching
assistantships na universidade (20h/semana, stipend de 9 meses, seguro
saúde básico, 9 créditos de remissão de tuition por termo), mas as
fontes encontradas referem-se majoritariamente a programas de
doutorado -- não foi confirmado se alunos do MSBA (mestrado) têm acesso
à mesma via, então o campo foi preenchido como "Incerto" em vez de
assumir disponibilidade.

'Principais Empregadores' preenchido com a lista de parceiros/relações
do Fox School (Boeing, CitiBank, Comcast, Domino's, Merck, Sallie Mae),
encontrada na própria página de carreiras do MSBA -- mas com a ressalva
de que não é uma lista confirmada de colocação de egressos do MSBA
especificamente, e sim de relações institucionais do Fox School como um
todo.

Fontes principais (pesquisadas via busca web, ago/2026):
- Busca web cruzada (temple.edu, bulletin.temple.edu -- fetch direto a
  www.temple.edu foi bloqueado por política de rede/segurança):
  GPA mínimo 3.0, GMAT/GRE opcional, TOEFL/IELTS/PTE/Duolingo exigido
  sem diploma dos EUA, taxa de aplicação $60, documentos (statement of
  goals, resume, transcripts, recomendadores via formulário),
  admissão rolling com prioridade 15 dez e deadline de bolsa/
  internacional 1 mar
- https://www.temple.edu/academics/degree-programs/graduate-programs/
  business-analytics-ms-busa/careers-internships-opportunities
  (empregadores parceiros do Fox School, estágio incentivado mas não
  obrigatório)
- Busca web: seguro saúde $2.760/ano (oficial, ciclo 2025-26)
- Busca web: custo de vida em Philadelphia (aluguel, SEPTA) -- não
  usado para sobrescrever campos já preenchidos e não claramente
  incorretos (Aluguel $1.650 e Transporte $96 já presentes)
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
    "Nome da Universidade": "Temple",
    "Estado": "PA",
    "Cidade": "Philadelphia",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Seguro Saúde (programa) USD": 2760,  # oficial, ciclo 2025-26
    "Bolsas Disponíveis": "Sim",
    "Tipos de Bolsa": "5% de desconto na tuition para ex-alunos (alumni) de Temple admitidos ao programa; fundos limitados, baseados em mérito",
    "Quão Realista (Intl)": "Baixa — a única bolsa nomeada encontrada é um desconto de 5% exclusivo para ex-alunos de Temple, não disponível a candidatos internacionais em geral",
    "TA/RA Disponível": (
        "Incerto — teaching assistantships existem na universidade "
        "(20h/semana, stipend de 9 meses, seguro saúde básico, 9 "
        "créditos de remissão de tuition/termo), mas as fontes "
        "encontradas referem-se majoritariamente a programas de "
        "doutorado; não confirmado para alunos do MSBA"
    ),
    "GPA Mínimo": 3.0,
    "Matérias Pré-Requisito": "Bagagem quantitativa (estatística) e familiaridade com Excel avaliadas na admissão; sem pré-requisito formal de curso específico",
    "Testes Inglês Aceitos": (
        "TOEFL/IELTS/PTE Academic/Duolingo aceitos — exigido para quem "
        "não tem diploma de instituição dos EUA; score dentro de 2 "
        "anos do teste"
    ),
    "Deadline Aplicação": "Admissão contínua (rolling); prioridade 15 dez; deadline de bolsa e internacional 1º mar",
    "Taxa Aplicação USD": 60,
    "Rounds Aplicação": "Rolling; sem rounds numerados formais (deadline de prioridade 15 dez)",
    "Documentos Necessários": (
        "Statement of goals, resume, histórico acadêmico (transcripts), "
        "detalhes de recomendadores inseridos via formulário online"
    ),
    "Principais Empregadores": (
        "Boeing, CitiBank, Comcast, Domino's, Merck, Sallie Mae "
        "(parcerias/relações institucionais do Fox School; não "
        "confirmado como colocação específica de egressos do MSBA)"
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
        raise ValueError("Nenhuma linha encontrada para Temple (Philadelphia, PA). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Temple, Philadelphia, PA)")

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
