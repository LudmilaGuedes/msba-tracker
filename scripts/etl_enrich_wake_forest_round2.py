"""
Data enrichment (RODADA 2): atualiza APENAS a linha da 'Wake Forest
University' (School of Business, Winston-Salem, NC) no
msba_programs_database.xlsx.

Esta linha já havia passado por um enrichment anterior (script
etl_enrich_wake_forest.py, de sessão prévia a esta conversa), que
preencheu com boa qualidade: perfil de turma (104 alunos, GPA médio
3.5, GMAT médio 685/GRE 320 -- CONFIRMADOS novamente aqui via página
oficial de Class Profile, sem alterações), taxa de emprego, salário,
empregadores, GMAT/GRE, testes de inglês, deadlines e documentos. Esta
rodada 2 foca em fechar as lacunas que restaram.

Foco exclusivo: MS in Business Analytics (MSBA), presencial, STEM/OPT,
37 créditos, 11 meses (3 semestres, início em julho). Deep research em:
custo total real de tuition para internacionais, bolsas/assistantships,
requisitos de GPA/GMAT/GRE/inglês, e custo de vida em Winston-Salem, NC.

TUITION: CORRIGIDO para o ciclo 2026-2027 oficial. A página oficial de
Tuition & Financial Aid (business.wfu.edu) mostra: Tuition $55.007 +
Required Activity/Health Services Fee $1.224 = Total $56.231. A
planilha tinha $51.084 (provavelmente de um ciclo anterior).
CORRIGIDO 'Tuition Total' (Intl e Residente, universidade PRIVADA) para
$55.007 (tuition pura) e PREENCHIDO 'Fees Adicionais' com $1.224 (fee
obrigatória, mantida separada por ser uma categoria distinta no
schema). 'Custo/Crédito' (estava 'Verificar' em ambos) PREENCHIDO
dividindo por 37 créditos = $1.486,68/crédito.

'Seguro Saúde (programa) USD' preenchido com $3.210/ano -- prêmio
oficial 2026-27 do WFU Student Health Insurance Plan (SHIP),
obrigatório para todo aluno matriculado em ao menos 1 crédito
presencial (incl. internacionais com visto F/J).

'Tipos de Bolsa' (já tinha "Mérito (sem aplicação separada)")
enriquecido com a lista completa de fundos de bolsa nomeados do School
of Business (ex.: Alm Family Scholarship, Bowman Family Scholarship,
Mary Reynolds Babcock Foundation Scholarships, entre ~19 fundos) --
NENHUM valor individual em dólares nem elegibilidade especifica para
internacionais foi divulgado publicamente, então 'Quão Realista (Intl)'
permanece 'Verificar' por falta desses detalhes.

Custo de vida: 'Utilities/mês' preenchido com $200 (faixa $180-220
encontrada para Winston-Salem, NC). 'Transporte/mês' permanece
'Verificar' -- o Wake Line (shuttle gratuito do campus) cobre rotas
para áreas de moradia fora do campus e alguns pontos comerciais, mas
não substitui integralmente um orçamento de transporte pessoal, e não
foi encontrada confirmação de tarifa zero da WSTA (transporte público
da cidade) especificamente para alunos.

Fontes principais (pesquisadas via busca web, ago/2026):
- https://business.wfu.edu/tuition-financial-aid/ (tuition oficial
  2026-27, fee obrigatória, lista de fundos de bolsa nomeados)
- https://business.wfu.edu/masters-in-business-analytics/class-profile/
  (perfil de turma, confirma dados já presentes)
- Busca web cruzada: GPA mínimo 3.0, pré-requisitos (Estatística + 1
  disciplina quantitativa adicional, B ou superior recomendado),
  "all majors welcome to apply" (experiência profissional não exigida)
- Busca web: seguro saúde $3.210/ano (SHIP oficial 2026-27); custo de
  vida em Winston-Salem, NC (aluguel já presente na planilha,
  utilities $180-220/mês)
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
    "Nome da Universidade": "Wake Forest University",
    "Estado": "NC",
    "Cidade": "Winston-Salem",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Intl) USD": 1486.68,  # $55.007 (oficial 2026-27) / 37 créditos
    "Custo/Crédito (Residente) USD": 1486.68,  # privada, mesma taxa do Intl
    "Tuition Total (Intl) USD": 55007,  # CORREÇÃO: planilha tinha $51.084 (ciclo anterior); oficial 2026-27 = tuition pura
    "Tuition Total (Residente) USD": 55007,  # privada, mesma taxa do Intl
    "Fees Adicionais USD": 1224,  # Required Activity/Health Services Fee, oficial 2026-27
    "Seguro Saúde (programa) USD": 3210,  # SHIP oficial 2026-27, obrigatório
    "Utilities/mês USD": 200,  # faixa $180-220 encontrada para Winston-Salem, NC
    "Caminho Tuition Residente": "N/A (privada)",
    "Tipos de Bolsa": (
        "Mérito (sem aplicação separada); ~19 fundos de bolsa nomeados "
        "do School of Business (ex.: Alm Family Scholarship, Bowman "
        "Family Scholarship, Mary Reynolds Babcock Foundation "
        "Scholarships, entre outros) — valores individuais e "
        "elegibilidade para internacionais não divulgados publicamente"
    ),
    "GPA Mínimo": 3.0,
    "Experiência Profissional": "Não obrigatória; candidatos de todas as áreas de graduação são bem-vindos",
    "Matérias Pré-Requisito": (
        "Estatística + 1 disciplina quantitativa adicional (ex.: "
        "regressão, probabilidade, cálculo discreto, álgebra linear), "
        "idealmente com nota B ou superior"
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
        raise ValueError("Nenhuma linha encontrada para Wake Forest University (Winston-Salem, NC). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Wake Forest University, Winston-Salem, NC)")

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
