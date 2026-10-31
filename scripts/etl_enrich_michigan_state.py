"""
Data enrichment: atualiza APENAS a linha da 'Michigan State' (Eli Broad
College of Business, East Lansing, MI) no msba_programs_database.xlsx.

Foco exclusivo: MS in Business Data Science and Analytics (MSBDSA),
presencial, STEM, programa de 30 créditos / 1 ano calendário (Fall+Spring+
Summer). Deep research em: custo total real de tuition para internacionais,
bolsas/assistantships, requisitos de GPA/GMAT/GRE/inglês, e custo de vida
local.

ACHADO CRITICO DA AUDITORIA: 'Tuition Total (Intl) USD' estava em $39.000 --
esse valor reflete a ESTRUTURA ANTIGA de tuition (diferenciada por
residência, corrigida via waiver de GA). A partir de Fall 2026 (fonte
oficial, mudança explicitamente descrita como "effective Fall 2026, a
change from the previous flat-fee structure that varied by residency"), a
tuition passou a ser tarifa ÚNICA de $1.200/crédito para TODOS os alunos,
independente de residência. Com os 30 créditos totais confirmados
(brochure oficial + busca cruzada), o total correto é 30 x $1.200 =
$36.000 -- CORRIGIDO (vários agregadores terceiros ainda mostram o
$39.000 desatualizado, não usados). Como a tarifa agora é igual para todos,
'Tuition Total (Residente) USD' recebeu o MESMO valor, e 'Caminho Tuition
Residente' (que dizia "Via GA (out-of-state waiver)", correto sob a
estrutura antiga) foi atualizado para refletir que essa distinção não
existe mais.

'Graduate Assistantships' e 'TA/RA Disponível' já cadastrados ("9cr F/S +
5cr Summer + waiver") foram CONFIRMADOS contra a política geral de GA da
MSU (isenção de 9 créditos Fall, 9 Spring, 5 Summer -- cobre virtualmente
os 30 créditos do programa) e mantidos sem alteração.

Resultados de emprego/salário (placement 100%, salário médio $89k) só
foram encontrados numa brochure sem data de publicação clara, referenciando
explicitamente "Class of 2020" e "Class of 2021" -- dados velhos demais
(sem uma classe mais recente confirmada) para representar o programa hoje
com confiança, então 'Taxa Empregabilidade %' e 'Salário Médio Pós USD'
foram MANTIDOS "Verificar" em vez de usar números de 5-6 anos atrás.

Fontes principais (pesquisadas manualmente, fora deste script, ago/2026):
- Busca por tuition: tarifa 2026-27 $1.200/crédito flat, efetiva a partir
  de Fall 2026, substituindo a estrutura antiga diferenciada por residência
- https://publications.broad.msu.edu/view/125441963/i (brochure oficial):
  30 créditos, programa de 12 meses, turma pequena de 48 alunos, "basic
  exposure to a programming language and statistics" como pré-requisito
- Busca por admissions: taxa de inscrição $85; TOEFL mín. 100 / IELTS
  mín. 7.0; GMAT ou GRE obrigatório; documentos: resume, personal
  statement, vídeo-ensaio, transcripts, cartas de recomendação
- Busca por assistantships: política geral de GA da MSU (isenção de 9+9+5
  créditos ao longo dos 3 semestres do calendário do programa)
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
    "Nome da Universidade": "Michigan State",
    "Estado": "MI",
    "Cidade": "East Lansing",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Total Créditos": 30,
    "Custo/Crédito (Intl) USD": 1200,  # tarifa única efetiva a partir de Fall 2026
    "Custo/Crédito (Residente) USD": 1200,  # mesma tarifa, sem diferenciação por residência
    "Tuition Total (Intl) USD": 36000,  # CORRECAO: 30 x $1.200 (nova estrutura Fall 2026); planilha tinha 39.000 (estrutura antiga)
    "Tuition Total (Residente) USD": 36000,  # mesma tarifa flat
    "Tamanho Turma": 48,  # brochure oficial: turma pequena
    "Caminho Tuition Residente": (
        "Não há mais distinção -- tarifa única de $1.200/crédito para "
        "todos os alunos independente de residência, efetiva a partir de "
        "Fall 2026 (antes havia diferenciação por residência, corrigida "
        "via waiver de GA)"
    ),  # CORRECAO: planilha tinha "Via GA (out-of-state waiver)", correto sob a estrutura antiga
    "Taxa Aplicação USD": 85,
    "Testes Inglês Aceitos": "TOEFL mín. 100 / IELTS mín. 7.0",  # CORRECAO/detalhamento: planilha tinha só "TOEFL/IELTS" vago
    "GMAT/GRE Obrigatório": "Obrigatório (GMAT ou GRE)",
    "Matérias Pré-Requisito": (
        "Exposição básica a uma linguagem de programação e a estatística, "
        "podendo vir da graduação ou de estudo autônomo"
    ),
    "Documentos Necessários": (
        "Resume, personal statement, vídeo-ensaio, transcripts, GMAT ou "
        "GRE, TOEFL/IELTS (internacionais), cartas de recomendação"
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
        raise ValueError("Nenhuma linha encontrada para Michigan State (East Lansing, MI). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Michigan State, East Lansing, MI)")

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
