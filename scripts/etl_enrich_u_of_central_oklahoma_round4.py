"""
Data enrichment (round 4): atualiza APENAS a linha 'U of Central
Oklahoma' (College of Business, Edmond, OK) no
msba_programs_database.xlsx.

Pedido da usuaria: cartas de recomendação são exigidas? Aceitam Duolingo
English Test?

CARTAS DE RECOMENDACAO: NAO aparecem em nenhuma lista de requisitos do
MSBA em MULTIPLAS buscas independentes (todas derivadas da mesma pagina
oficial e do catalogo): "bachelor's or advanced degree..., overall GPA
of 3.0, complete the application..., current resume, and a 500-word
personal statement." Nenhuma menciona carta de recomendacao. Ausencia
consistente em varias fontes que citam a mesma lista oficial e forte
indicio de que NAO e exigida -- mas nao e uma confirmacao textual
explicita ("nao exigimos carta"), entao registrado com essa ressalva.

DUOLINGO: CONFIRMADO que a UCO aceita Duolingo English Test para
admissao de GRADUATE, mas com uma restricao: "Duolingo English Test for
graduate admission is only for Fall term" -- ou seja, so vale se a
candidata pretende ingressar no semestre de Fall (o que ja bate com
'Início' = Fall, ja cadastrado nesta linha). Score minimo de Duolingo
NAO confirmado com confianca -- um blog citava "105", mas com a ressalva
de ter sido uma aprovacao TEMPORARIA da era COVID, entao esse numero
especifico NAO foi usado.

TESTES DE INGLES -- scores minimos encontrados (fonte: agregacao de
paginas oficiais da UCO indexadas por busca, aplicaveis a admissao
GRADUATE em geral -- podem ou nao ser especificos do MSBA, ja que
escolas de negocio as vezes tem padrao proprio mais alto que o minimo
geral da universidade; usado como melhor dado disponivel):
  TOEFL: 500 (PBT) ou 61 (iBT)
  IELTS: 6.0
  PTE: 46-52
  iTEP Academic: 3.7
  Duolingo: aceito, só para ingresso no Fall, score mínimo não confirmado
Alternativa para quem fica abaixo desses minimos: comprovar 12 semanas
de curso intensivo de ingles (ESL) em programa aprovado nos EUA.
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, ".")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = {
    "Nome da Universidade": "U of Central Oklahoma",
    "Estado": "OK",
    "Cidade": "Edmond",
}

UPDATES = {
    "Testes Inglês Aceitos": (
        "TOEFL 500 (PBT) / 61 (iBT); IELTS 6.0; PTE 46-52; iTEP "
        "Academic 3.7; Duolingo English Test ACEITO, mas só para "
        "ingresso no semestre Fall (score mínimo não confirmado com "
        "confiança). Alternativa: 12 semanas de curso intensivo de "
        "inglês (ESL) em programa aprovado nos EUA para quem fica "
        "abaixo dos mínimos"
    ),
    "Documentos Necessários": (
        "Application completa, resume, personal statement (500 "
        "palavras), teste de inglês (IELTS/TOEFL/PTE/Duolingo). "
        "Internacionais, adicionalmente: bank statement oficial em "
        "inglês (emitido há menos de 6 meses), Certificate of "
        "Financial Resources (p/ I-20/F-1 ou DS-2019/J-1), Statement of "
        "Financial Sponsorship, cópia do passaporte, transcripts "
        "oficiais (se não emitidos em inglês, exige documento original "
        "+ tradução oficial/certificada). CARTA DE RECOMENDAÇÃO: não "
        "aparece em nenhuma lista oficial encontrada -- provavelmente "
        "NÃO exigida para o MSBA (ausência consistente em múltiplas "
        "fontes que citam a mesma lista oficial), mas sem confirmação "
        "textual explícita do tipo 'não exigimos'"
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
        raise ValueError("Nenhuma linha encontrada para U of Central Oklahoma (Edmond, OK). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (U of Central Oklahoma, Edmond, OK)")

    no_fill = PatternFill(fill_type=None)

    before = {}
    after = {}
    for header, new_value in UPDATES.items():
        col_idx = header_to_col[header]
        cell = ws.cell(row=row_idx, column=col_idx)
        before[header] = cell.value
        cell.value = new_value
        cell.fill = no_fill
        after[header] = new_value

    wb.save(DB_PATH)

    print("\nCampos atualizados (Como Estava -> Como Ficou):")
    for header in UPDATES:
        print(f"  - {header}:")
        print(f"      Antes: {before[header]!r}")
        print(f"      Depois: {after[header]!r}")


if __name__ == "__main__":
    main()
