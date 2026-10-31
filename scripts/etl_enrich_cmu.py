"""
Data enrichment: atualiza APENAS a linha da 'CMU' (Carnegie Mellon
University, Tepper School of Business, Pittsburgh, PA) no
msba_programs_database.xlsx.

Foco exclusivo: MS in Business Analytics, formato full-time presencial
(9 meses), STEM, 108 unidades/créditos (sistema de "units" da CMU).
Deep research em: custo total real de tuition para internacionais,
bolsas/assistantships, requisitos de GPA/GMAT/GRE/inglês, e custo de
vida em Pittsburgh, PA.

TUITION: CORRIGIDO. A planilha tinha $718/unidade e $69.412 total, sem
origem claramente rastreável ao ciclo corrente. A página oficial de
Student Financial Services (cmu.edu/sfs, ciclo 2026-2027) confirma
$743/unidade e 108 unidades = $80.244 de tuition pura. CORRIGIDO:
'Custo/Crédito' para $743 e 'Tuition Total' para $80.244 (ambos para
Intl e Residente, ja que e universidade PRIVADA e o campo 'Residente'
estava 'Verificar').

'Total Créditos' estava 'Verificar'; PREENCHIDO com 108 -- CMU usa um
sistema de "units" (nao o credito-hora semestral tradicional de outras
universidades americanas), confirmado pela mesma pagina oficial.

'Fees Adicionais USD' preenchido com $1.040 = Activity Fee $248 +
Technology Fee $490 + Transportation Fee $302 (taxas obrigatorias por
ano letivo). 'Seguro Saude (programa) USD' preenchido separadamente com
$3.352 (aluno solteiro; a mesma fonte oficial lista essa fee de forma
distinta das demais). A fee de emprestimo federal ($1.500) NAO foi
incluida em nenhum dos dois campos por ser condicional (so se aplica a
quem usa Direct Loans), nao uma cobranca obrigatoria universal.

ACHADO SOBRE TRANSPORTE: a Transportation Fee de $302/ano ja cobre
tarifas do Pittsburgh Regional Transit (PRT) -- CMU tem acordo que
torna onibus, light rail e inclines GRATUITOS no ponto de uso para
alunos (via app Ready2Ride). Como esse custo ja esta contabilizado em
'Fees Adicionais', 'Transporte/mes' foi preenchido com $0 para evitar
contagem dupla.

Ranking US News: a pagina oficial lista "#2 Business Analytics -- U.S.
News & World Report (2026)" ao lado de rankings explicitamente
rotulados como do programa ONLINE ("#1 Best Online Master's", "#4
Online MSBA Program"). Apenas o ranking SEM o qualificador "online" foi
usado aqui (nosso alvo e o programa presencial/on-campus), para nao
atribuir um ranking do programa online a linha do presencial.

Fontes principais (pesquisadas via busca web, ago/2026):
- https://www.cmu.edu/sfs/tuition/graduate/tepper/full-msba.html
  (tuition oficial 2026-27, breakdown completo de fees)
- https://www.cmu.edu/tepper/programs/master-business-analytics
  (formato, ranking, setores de emprego -- sem nomes especificos de
  empregadores nem taxa de emprego/salario divulgados)
- Busca web cruzada (yocket, mim-essay, f1gmat): GPA medio admitidos
  3.56, GMAT faixa 640-700 (80% central, sem media unica divulgada),
  GMAT/GRE com waiver possivel (nao afeta elegibilidade a bolsa),
  TOEFL recomendado 25/secao, IELTS recomendado 7.0 geral e por banda,
  documentos exigidos (1 ensaio escrito + 1 video essay + 1 carta
  profissional), deadlines Round 1-4
- https://www.cmu.edu/transportation/transport/prt.html e busca web:
  PRT gratuito para alunos CMU via Transportation Fee
- Busca web: aluguel de quarto privativo em Oakland/Squirrel Hill
  (bairros populares entre alunos da CMU), Pittsburgh PA
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
    "Nome da Universidade": "CMU",
    "Estado": "PA",
    "Cidade": "Pittsburgh",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Total Créditos": 108,  # sistema de "units" da CMU (51 no Fall + 57 na Spring)
    "Custo/Crédito (Intl) USD": 743,  # CORREÇÃO: planilha tinha $718; oficial 2026-27 = $743/unidade
    "Custo/Crédito (Residente) USD": 743,  # privada, mesma taxa do Intl
    "Tuition Total (Intl) USD": 80244,  # CORREÇÃO: planilha tinha $69.412; oficial 2026-27 = 108 x $743
    "Tuition Total (Residente) USD": 80244,  # privada, mesma taxa do Intl
    "Fees Adicionais USD": 1040,  # Activity Fee $248 + Technology Fee $490 + Transportation Fee $302 (ano letivo)
    "Seguro Saúde (programa) USD": 3352,  # oficial 2026-27, aluno solteiro
    "Transporte/mês USD": 0,  # PRT gratuito para alunos CMU (já coberto pela Transportation Fee, evitar dupla contagem)
    "Aluguel Privativo/mês USD": 850,  # quarto privativo, Oakland/Squirrel Hill (perto do campus)
    "GPA Médio Admitidos": 3.56,
    "Testes Inglês Aceitos": (
        "TOEFL recomendado 25 em cada seção (iBT) / IELTS recomendado "
        "7.0 geral e 7.0 em cada banda / Duolingo aceito — dispensa "
        "possível com diploma de instituição de instrução "
        "exclusivamente em inglês"
    ),
    "Deadline Aplicação": (
        "Round 1: 14 out 2026; Round 2: 13 jan 2027; Round 3: 3 mar "
        "2027; Round 4 Priority: 5 mai 2027; Round 4 Final: 6 jun 2027"
    ),
    "Rounds Aplicação": "4 rounds (Round 1 a Round 4, com etapa Priority/Final no Round 4); ver Deadline Aplicação",
    "Documentos Necessários": (
        "Application online, resume, 1 ensaio escrito + 1 video essay, "
        "1 carta de recomendação profissional, histórico acadêmico "
        "não-oficial, GMAT/GRE oficial (ou waiver), TOEFL/IELTS/"
        "Duolingo (se aplicável)"
    ),
    "Ranking US News": 2,  # "Business Analytics" (2026) — versão presencial; rankings rotulados "Online" não usados aqui
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
        raise ValueError("Nenhuma linha encontrada para CMU (Pittsburgh, PA). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (CMU, Pittsburgh, PA)")

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
