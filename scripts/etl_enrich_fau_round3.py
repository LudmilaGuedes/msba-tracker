"""
Data enrichment (RODADA 3): atualiza APENAS a linha 'Florida Atlantic
University (FAU)' (College of Business, Boca Raton, FL) no
msba_programs_database.xlsx.

Programa: Master of Science in Business Analytics (Traditional),
presencial, STEM.

FONTES OFICIAIS (acessadas em 2026-10-17):
  - https://business.fau.edu/masters-phd/business-analytics/traditional/admissions/
  - https://business.fau.edu/masters-phd/business-analytics/traditional/program-information/
  - https://business.fau.edu/masters-phd/business-analytics/traditional/career-outlook
  - https://www.fau.edu/graduate/graduate-programs/ms-business-analytics/
  - https://www.fau.edu/controllers-office/student-services/tuition-breakdown/
  - https://www.fau.edu/shs/insurance/international

CONFIRMADO (sem alteração): tabela 2026-27 por crédito (tuition + fees)
pós residente US$ 371,82 / não-residente US$ 1.125,06 (inclui a non-
resident fee de US$ 717,37) -> 30 créditos = US$ 11.155 / US$ 33.752;
12 meses full-time (12 + 12 + 6 créditos em Fall/Spring/Summer), aulas de
dia em Boca Raton; GPA 3.0 nos últimos 60 créditos; taxa US$ 30; Duolingo
110 (página da Graduate College; um PDF de outro programa citava 90).

CORREÇÕES:
- GMAT/GRE: a planilha dizia "waiver automático com GPA >= 3.3". Na
  Graduate College, a dispensa automática é só para quem se formou no
  departamento ITOM da própria FAU (GPA >= 3.25); a página do College of
  Business diz que quem tem GPA de admissão >= 3.3 "may be eligible" --
  ou seja, pode pedir, não é automático. Também dispensa quem já tem
  mestrado.
- Inglês: TOEFL agora é 4.5 (nova escala) na página da Graduate College;
  79 era o corte da escala antiga.
- Fees Adicionais: fora do valor por crédito há Transportation Access Fee
  (US$ 76,90 Fall/Spring, US$ 32,04 Summer), Owl Card US$ 10 e Graduate
  Orientation Fee US$ 35 -> 76,90 x 2 + 32,04 + 10 + 35 = US$ 230,84.
- Rounds: não é rolling -- prazos fixos por semestre e a aplicação
  precisa estar completa até o prazo.
- Seguro saúde: obrigatório para F-1; a FAU inscreve automaticamente no
  plano da universidade (dá para dispensar com plano equivalente). O
  preço não está na página oficial -- continua 'Verificar'.
- Admissão condicional: quem fica abaixo do GPA ou do GMAT/GRE pode ser
  aceito condicionalmente, caso a caso.
- Empregadores: a página do programa cita Microsoft, Citrix e Office
  Depot como colocações de formados "dos nossos programas" (ITOM), sem
  separar o MSBA.

NÃO ALTERADOS: Salário Médio Pós (o US$ 72.560 da página é média do
Payscale para o mercado, não dos formados), empregabilidade, rankings,
tamanho da turma, GPA médio, GMAT médio, taxa de aceitação, bolsas.
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
    "Nome da Universidade": "Florida Atlantic University (FAU)",
    "Estado": "FL",
    "Cidade": "Boca Raton",
}

UPDATES = {
    "Fees Adicionais USD": 231,  # Transportation Access 76,90 x 2 + 32,04 + Owl Card 10 + Orientation 35
    "Caminho Tuition Residente": (
        "Não aplicável para F-1 -- internacional paga a non-resident fee "
        "(US$ 717,37/crédito a mais que o residente)"
    ),
    "Principais Empregadores": (
        "Microsoft, Citrix, Office Depot (citados como colocações de "
        "formados dos programas do departamento ITOM, sem separar o MSBA)"
    ),
    "GMAT/GRE Obrigatório": (
        "Obrigatório: GMAT 495+ ou GRE (verbal + quant) 295+, de até 5 anos. "
        "Dispensa automática só para quem se formou no ITOM da própria FAU "
        "(GPA >= 3.25) ou já tem mestrado; com GPA de admissão >= 3.3 dá "
        "para pedir dispensa (não é automática)"
    ),
    "Testes Inglês Aceitos": (
        "TOEFL iBT 4.5 (nova escala; 79 na antiga) / IELTS 6.0 / Duolingo "
        "110 / alternativa: Level VI (Academic Track) do Intensive English "
        "Institute da FAU. PTE não consta na lista"
    ),
    "Deadline Aplicação": (
        "Round 1: 15 fev (Fall, internacional); Spring: 15 jul "
        "(internacional); Summer: 15 jan (internacional). Domésticos: 1 jul "
        "/ 1 nov / 1 abr. Para ingresso (Fall 2028): ~15 fev 2028"
    ),
    "Rounds Aplicação": (
        "Sem rounds e não é rolling: um prazo fixo por semestre, com a "
        "aplicação e todos os documentos completos até a data"
    ),
    "Documentos Necessários": (
        "Aplicação online na Graduate College (taxa US$ 30); histórico "
        "oficial ou avaliação do diploma; internacionais: avaliação "
        "course-by-course por membro NACES ou pela própria FAU (sem "
        "preferência); GMAT ou GRE (salvo dispensa); teste de inglês; "
        "comprovante de imunização antes da matrícula; documentos de "
        "residência para fins de tuition. Nenhuma carta de recomendação "
        "exigida"
    ),
    "Dificuldade Entrada": (
        "Moderada (aceita admissão condicional, caso a caso, para quem "
        "fica abaixo do GPA ou do GMAT/GRE)"
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
        raise ValueError("Nenhuma linha encontrada para FAU (Boca Raton, FL). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (FAU, Boca Raton, FL)")

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
