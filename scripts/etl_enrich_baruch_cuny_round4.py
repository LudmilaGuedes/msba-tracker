"""
Data enrichment (round 4): atualiza APENAS a linha da 'Baruch CUNY'
(Zicklin School of Business, New York, NY) no msba_programs_database.xlsx.

Objetivo desta rodada: taxa de aplicação e requisitos/documentos do
processo de aplicação (pedido do usuário).

FONTE: página oficial "Application Components" da Zicklin, aplicável a
candidatos de MBA & MS (inclui MS Business Analytics):
  https://zicklin.baruch.cuny.edu/academic-programs/graduate/grad-admissions/application-process/application-components/
(acessado via download direto em 2026-09-08; zicklin.baruch.cuny.edu
segue com erro de certificado SSL na ferramenta de fetch padrão --
mesma limitação já documentada nos rounds anteriores -- contornado com
download direto do HTML.)

ACHADOS relevantes (aplicam-se a "MBA & MS Applicants", texto da
própria página, portanto cobre MS Business Analytics):
- Application Fee: $75, não-reembolsável (isento para militares/veteranos
  dos EUA e cônjuges, mediante comprovação)
- Cartas de recomendação: 2 OBRIGATÓRIAS (a maioria: 2 referências
  profissionais; recém-formados: 1 profissional + 1 acadêmica; sem
  bacharelado ainda: 2 acadêmicas). CORREÇÃO: uma busca anterior (fonte
  agregadora não-oficial) havia sugerido que LOR não seriam
  obrigatórias -- a página oficial contradiz isso claramente, fonte
  oficial prevalece.
- Ensaios: 2 obrigatórios (Essay 1: 500 palavras, sobre motivação para
  o mestrado e plano de carreira; Essay 2: 250 palavras, escolha entre
  tópicos) + 1 opcional
- Transcripts: cópia não-oficial de TODAS as faculdades cursadas,
  exigida para decisão inicial; oficial exigido só após admissão
  (enviado diretamente pela instituição)
- Prova de proficiência em inglês: dispensada se o candidato tem
  diploma (bacharelado/mestrado/doutorado) dos EUA ou de país de língua
  inglesa listado, OU histórico documentado de 1º ao 12º ano em escola
  dos EUA. Caso não se enquadre, aceita: TOEFL iBT (mín. 100 na escala
  antiga / 5 na nova), Duolingo English Test (mín. 120), IELTS Academic
  (mín. 7), PTE Academic (mín. 70). Validade: TOEFL/IELTS por 2 anos.
- Teste padronizado (GMAT/GRE/EA): página confirma "Standardized
  testing is optional for applicants to MS and Advanced Certificate
  programs" -- ou seja, opcional para TODOS os candidatos de MS (não
  é uma dispensa condicional como no MBA, é opcional por padrão).

CAMPOS **NÃO** atualizados nesta rodada (não encontrados na fonte
oficial, permanecem 'Verificar'): Rounds Aplicação (não há menção
explícita de rounds vs. rolling), Entrevista, Taxa Aceitação %
(específica do MSBA), GPA Mínimo (só há menção de 3.30 como critério
de dispensa de teste padronizado no MBA, não é um mínimo de admissão
do MS).
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
    "Nome da Universidade": "Baruch CUNY",
    "Estado": "NY",
    "Cidade": "New York",
}

UPDATES = {
    "Taxa Aplicação USD": 75,  # não-reembolsável; isento p/ militares/veteranos dos EUA e cônjuges
    "Documentos Necessários": (
        "Transcripts não-oficiais de todas as faculdades cursadas "
        "(oficiais só após admissão), CV/currículo, 2 cartas de "
        "recomendação obrigatórias (2 profissionais; recém-formados: "
        "1 profissional + 1 acadêmica), 2 ensaios obrigatórios (500 e "
        "250 palavras) + 1 opcional, comprovante de proficiência em "
        "inglês (se aplicável, ver campo 'Testes Inglês Aceitos')"
    ),
    "Testes Inglês Aceitos": (
        "TOEFL iBT mín. 100 (escala antiga) / Duolingo DET mín. 120 / "
        "IELTS Academic mín. 7 / PTE Academic mín. 70 — dispensado com "
        "diploma dos EUA ou de país de língua inglesa listado, ou "
        "histórico de 1º-12º ano em escola dos EUA. TOEFL/IELTS válidos "
        "por 2 anos."
    ),
    "GMAT/GRE Obrigatório": "Opcional (GRE, GMAT ou Executive Assessment aceitos, nenhum obrigatório para candidatos de MS)",
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
        raise ValueError("Nenhuma linha encontrada para Baruch CUNY (New York, NY). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Baruch CUNY, New York, NY)")

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
