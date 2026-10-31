"""
Data enrichment (round 3): atualiza APENAS a linha 'U of Central
Oklahoma' (College of Business, Edmond, OK) no
msba_programs_database.xlsx.

Pedido da usuaria: quais documentos são exigidos na aplicação e quanto
custa.

TAXA DE APLICACAO -- resolvida uma contradicao entre duas buscas:
- 1a busca: "$40 international admission fee... required for your
  application... finalized by your admissions advisor" -- lida
  inicialmente como uma taxa ADICIONAL para internacionais, somada aos
  "$60" de taxa geral de graduate (total $100).
- 2a busca, mais explicita, desfaz a ambiguidade: "$40.00" e a taxa de
  UNDERGRADUATE; "$60.00" e a taxa de GRADUATE -- e o proprio texto
  confirma que um candidato internacional a programas de GRADUATE
  aplica com a taxa de $60 (nao ha um adicional de $40 por cima disso
  para pos-graduacao). Preenchido $60 (nao $100).
site.uco.edu e oga.uco.edu bloquearam fetch direto (HTTP 403) em todas
as tentativas -- taxa confirmada via busca (Google), nao fetch direto.

DOCUMENTOS -- lista oficial expandida (fonte: pagina "UCO: Admission
Documents", indexada por busca; fetch direto tambem bloqueado por 403,
mas o snippet trouxe trechos citados diretamente da pagina):
  - Bank statement oficial/certificado, em ingles, emitido ha menos de
    6 meses da data de submissao da aplicacao -- valor minimo exigido
    varia conforme o tipo de visto
  - Certificate of Financial Resources form -- obrigatorio para quem
    vai receber I-20 (F-1) ou DS-2019 (J-1)
  - Statement of Financial Sponsorship
  - Copia do passaporte
  - Transcripts academicos oficiais; se emitidos em idioma diferente do
    ingles, exige-se o documento original NO IDIOMA ORIGINAL + traducao
    oficial/certificada para o ingles
  - (ja cadastrados) resume, personal statement (500 palavras), teste
    de ingles (IELTS/TOEFL/PTE)
Tudo enviado digitalmente junto com a aplicacao eletronica -- documento
faltante atrasa o processamento.
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
    "Taxa Aplicação USD": 60,  # taxa de graduate (não $100 -- $40 é a taxa de undergraduate, não um adicional)
    "Documentos Necessários": (
        "Application completa, resume, personal statement (500 "
        "palavras), teste de inglês (IELTS/TOEFL/PTE). Internacionais, "
        "adicionalmente: bank statement oficial em inglês (emitido há "
        "menos de 6 meses), Certificate of Financial Resources (p/ "
        "I-20/F-1 ou DS-2019/J-1), Statement of Financial Sponsorship, "
        "cópia do passaporte, transcripts oficiais (se não emitidos em "
        "inglês, exige documento original + tradução oficial/"
        "certificada)"
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
