"""
Data enrichment (RODADA 2): atualiza APENAS a linha 'U of Tulsa'
(Collins College of Business, Tulsa, OK) no msba_programs_database.xlsx.

Esta linha já passou por uma rodada de enrichment (script
etl_enrich_u_of_tulsa.py), que teve acesso severamente limitado --
dezenas de tentativas de achar a página específica do programa
retornaram 404. Usuário confirmou via AskUserQuestion (ago/2026) para
tentar novamente com URLs diferentes.

ACHADO DESTA RODADA: a página correta do programa foi finalmente
localizada via o "Graduate Program Adviser" listado pela Graduate
School (https://utulsa.edu/academics/business/academics/graduate-business-programs/business-analytics/)
-- as tentativas anteriores usavam padrões de URL incorretos.

VALIDAÇÃO CRUZADA: a nova página confirma "96% placement rate" para
formandos -- bate EXATAMENTE com o valor já presente em 'Taxa
Empregabilidade %' (96, preenchido em rodada anterior por outra fonte),
confirmando que o dado já estava correto. Também confirma "GMAT: Not
required, but applicants submitting scores qualify for enhanced
financial aid" e bolsas "Merit scholarships covering up to 75% of
tuition" + "Graduate assistantships with full tuition coverage plus
stipends" -- ambos já consistentes com o que já estava preenchido em
'GMAT/GRE Obrigatório' e 'Tipos de Bolsa' (não Verificar, não
tocados).

'Documentos Necessários' preenchido com a lista oficial geral da
Graduate School (a página específica do MSBA não distingue requisitos
próprios): Graduate School application; taxa de aplicação $60 (já
confirmada); transcripts oficiais de todas as instituições; Statement
of Purpose (500-750 palavras); 2-3 cartas de recomendação; GRE
dispensado para a maioria dos programas de business, GMAT opcional;
TOEFL/IELTS/Duolingo (internacionais).

'Rounds Aplicação' preenchido: admissão contínua (rolling), mas
recomendado aplicar até 1 fev para consideração de graduate
assistantship, e até 3 fev para bolsas/fellowships competitivos.

CUIDADO APLICADO: a página do "Graduate School" cita "Top 100 private
research institution (U.S. News)" -- um ranking GERAL da universidade,
não específico do programa MSBA, e sem número exato de posição --
NÃO usado para preencher 'Ranking US News' (mesmo critério de cautela
já aplicado a outras universidades desta sessão). Uma duração de "12
months minimum" foi encontrada, mas pertence à trilha ONLINE do MBA em
Business Analytics, um produto DIFERENTE do MSBA presencial -- não
usada para não conflar dados de programas distintos.

CUIDADO APLICADO (dados NÃO encontrados mesmo após tentativas
adicionais, permanecem 'Verificar'): 'Duração (meses)', 'Total
Créditos' (a página do bulletin com o detalhamento curricular retornou
conteúdo vazio em múltiplas tentativas -- provável renderização via
JavaScript não capturável por este método), 'Tamanho Turma', 'Tuition
Total' (ambos os campos, pois dependem do total de créditos ainda não
confirmado), 'Fees Adicionais', todos os campos mensais de custo de
vida, 'Ranking QS', 'Salário Médio Pós USD', 'GPA Mínimo', 'GPA Médio
Admitidos', 'GMAT Médio', 'Taxa Aceitação %', 'Matérias
Pré-Requisito', 'Entrevista'.

Fontes principais (pesquisadas via WebFetch, ago/2026):
- https://utulsa.edu/admissions/graduate-admission/program-advisers/
  (localizou finalmente a URL correta do programa)
- https://utulsa.edu/academics/business/academics/graduate-business-programs/business-analytics/
  (formato, STEM, currículo, GMAT opcional, bolsas, 96% placement --
  validação cruzada com dado já existente)
- https://utulsa.edu/admissions/graduate-admission/admission-deadline-requirements/
  (taxa de aplicação $60 confirmada, documentos, deadlines gerais)
- https://utulsa.edu/tuition-aid/graduate-financial-assistance/
  (detalhamento de assistantship: stipend ~$14.349, cobertura de até 9
  créditos/semestre, GPA ≥3.0 para manter)
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
    "Nome da Universidade": "U of Tulsa",
    "Estado": "OK",
    "Cidade": "Tulsa",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Documentos Necessários": (
        "Graduate School application; taxa de aplicação $60; "
        "transcripts oficiais de todas as instituições; Statement "
        "of Purpose (500-750 palavras); 2-3 cartas de recomendação; "
        "GRE dispensado para a maioria dos programas de business "
        "(GMAT opcional); TOEFL/IELTS/Duolingo (internacionais)"
    ),
    "Rounds Aplicação": "Admissão contínua (rolling); recomendado até 1 fev para consideração de graduate assistantship, e até 3 fev para bolsas/fellowships competitivos",
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
        raise ValueError("Nenhuma linha encontrada para U of Tulsa (Tulsa, OK). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (U of Tulsa, Tulsa, OK)")

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
