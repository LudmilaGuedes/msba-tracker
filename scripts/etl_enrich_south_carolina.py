"""
Data enrichment: atualiza APENAS a linha da 'U of South Carolina' (Darla
Moore School of Business, Columbia, SC) no msba_programs_database.xlsx.

Foco exclusivo: MS in Business Analytics (MSBA), presencial, STEM, 30
créditos. Deep research em: custo total real de tuition para internacionais,
bolsas/assistantships, requisitos de GPA/GMAT/GRE/inglês, e custo de vida
local.

ACHADO CRITICO DA AUDITORIA: 'Taxa Aplicação USD' estava em 250 -- muito
acima do valor real. A partir de 1/ago/2026 (já vigente na data desta
coleta), a taxa é $108 ($50 de processamento da Graduate School + $58 de
taxa própria da Moore School). CORRIGIDO. 'Deadline Aplicação' também foi
ajustada: a página oficial de admission requirements lista Priority 1/dez,
Normal Decision 1/mar, Late Decision 15/mai (com rolling admission depois)
-- ligeiramente diferente das datas que estavam na planilha (5/dez, 25/fev,
15/mai), agora alinhadas à fonte oficial mais precisa.

Tuition: o Moore School NÃO oferece tarifa diferenciada por residência para
o MSBA -- $1.100/crédito vale igualmente para residentes da Carolina do Sul
e não-residentes (incluindo internacionais). Por isso 'Custo/Crédito' e
'Tuition Total' ficaram iguais para Intl e Residente (30 x $1.100 =
$33.000). Internacionais pagam fees adicionais específicas por cima dessa
mesma tuition (não uma tuition diferente) -- por isso 'Caminho Tuition
Residente' foi preenchido esclarecendo que não há tarifa premium a ser
dispensada (a tuition já é única).

'Fees Adicionais USD' foi calculado para o perfil internacional, assumindo
3 semestres (Fall+Spring+Summer, consistente com a duração de 12 meses já
cadastrada): matrícula $80 (único) + saúde $100/sem x3 + tecnologia
$200/sem x3 + taxa internacional $112 (único) + $200/sem x3 = $1.692.

Custo de vida: 'Seguro Saúde (programa)' foi preenchido com o prêmio anual
mais recente encontrado (2024-25: Fall $1.355,29 + Spring/Summer $1.849,71
= $3.205), com ressalva de que pode ter sido reajustado para 2026-27 (não
encontrada uma tabela oficial mais nova). 'Utilities/mês' e 'Transporte/mês'
permanecem "Verificar" -- a COA oficial da USC só publica uma categoria
agregada de "Personal & Miscellaneous Expenses" (livros, transporte,
utilities e lazer juntos), que não pode ser decomposta com confiança; por
isso 'Custo Vida Total (programa) USD' também permanece "Verificar" (2 dos
9 componentes ainda faltam).

Fontes principais (pesquisadas manualmente, fora deste script, ago/2026):
- https://sc.edu/study/colleges_schools/moore/study/management_science/degree_programs/masters_business_analytics/tuition_fees.php
  (tabela oficial de tuition/fees, sem diferenciação por residência;
  fellowships de mérito + GA no Moore School Data Lab)
- https://sc.edu/study/colleges_schools/moore/study/management_science/degree_programs/masters_business_analytics/admission_requirements/index.php
  (GPA mínimo 3.0; GMAT/GRE obrigatório com dispensa possível; TOEFL 100/
  IELTS 7.5/TOEIC/PTE; documentos: transcripts, 2 cartas, CV, SoP de 500
  palavras; entrevista para candidatos qualificados; deadlines oficiais)
- Busca por class profile: turma média de 23 alunos, GPA mediano 3.80,
  experiência profissional mediana de 15 meses
- Busca por taxa de inscrição: $108 efetivo a partir de 1/ago/2026 ($50 +
  $58), corrigindo o valor antigo de $250
- Busca por seguro-saúde internacional: obrigatório para todos os
  internacionais independente de carga de créditos; prêmio mais recente
  encontrado é de 2024-25
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
    "Nome da Universidade": "U of South Carolina",
    "Estado": "SC",
    "Cidade": "Columbia",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Custo/Crédito (Intl) USD": 1100,
    "Custo/Crédito (Residente) USD": 1100,  # mesma tarifa, sem diferenciação por residência
    "Tuition Total (Intl) USD": 33000,  # 30 x $1.100
    "Tuition Total (Residente) USD": 33000,  # mesma tarifa flat
    "Fees Adicionais USD": 1692,  # perfil internacional, 3 semestres: matrícula+saúde+tecnologia+taxa intl
    "Seguro Saúde (programa) USD": 3205,  # prêmio 2024-25 (mais recente encontrado); pode ter reajustado para 2026-27
    "Tamanho Turma": 23,  # turma média
    "GPA Mínimo": 3.0,
    "GPA Médio Admitidos": 3.80,  # mediana da turma
    "GMAT/GRE Obrigatório": (
        "Obrigatório GMAT ou GRE (validade de até 5 anos); dispensável "
        "para candidatos com desempenho de graduação superior e/ou "
        "experiência profissional relevante"
    ),  # detalhamento: planilha tinha só "Waivable"
    "Testes Inglês Aceitos": (
        "TOEFL mín. 100 (internet-based) / IELTS mín. 7.5 / TOEIC / PTE "
        "Academic; dispensável para quem se formou em programa "
        "ministrado em inglês"
    ),  # detalhamento: planilha tinha só "TOEFL/IELTS/TOEIC/PTE" sem notas mínimas
    "Experiência Profissional": (
        "Não obrigatória (mediana da turma: 15 meses de experiência "
        "profissional)"
    ),
    "Deadline Aplicação": (
        "Priority: 1/dez; Normal Decision: 1/mar; Late Decision: 15/mai "
        "(admissão rolling depois)"
    ),  # CORRECAO: planilha tinha 5/dez, 25/fev, 15/mai
    "Taxa Aplicação USD": 108,  # CORRECAO: planilha tinha 250; oficial a partir de 1/ago/2026 = $50 (Grad School) + $58 (Moore School)
    "Documentos Necessários": (
        "Transcripts oficiais, GRE/GMAT, 2 cartas de recomendação "
        "(preferencialmente profissionais), CV/Resume, Statement of "
        "Purpose (500 palavras)"
    ),  # detalhamento: planilha tinha versão resumida sem o detalhe do SoP
    "Caminho Tuition Residente": (
        "Não aplicável -- a tuition já é única ($1.100/crédito) para "
        "residentes e não-residentes; não há tarifa premium de fora do "
        "estado a ser dispensada"
    ),
    "Tipos de Bolsa": (
        "Fellowships de mérito (valores variáveis) + Graduate "
        "Assistantships no Moore School Data Lab"
    ),  # detalhamento: planilha tinha só "Merit fellowships + GA"
    "Graduate Assistantships": "GA no Moore School Data Lab",  # detalhamento: planilha tinha só "GA"
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
        raise ValueError("Nenhuma linha encontrada para U of South Carolina (Columbia, SC). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (U of South Carolina, Columbia, SC)")

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
