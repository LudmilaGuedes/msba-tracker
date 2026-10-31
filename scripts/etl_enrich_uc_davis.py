"""
Data enrichment: atualiza APENAS a linha da 'UC Davis' (Graduate School of
Management, Davis/San Francisco, CA) no msba_programs_database.xlsx.

Foco exclusivo: Master of Science in Business Analytics (MSBA), presencial,
STEM, campus fisico em San Francisco (Civic Center / UC Law SF).

Fontes principais (pesquisadas manualmente, fora deste script):
- https://gsm.ucdavis.edu/master-science-business-analytics/tuition-financial-aid
  (tarifa por unidade confirmada para a coorte de verao 2026: $1.560/unidade,
  corrigindo o valor antigo de $1.490/unidade, que era da coorte 2025)
- https://gsm.ucdavis.edu/msba-start-2026 (duracao, formato, campus, ranking)
- https://gsm.ucdavis.edu/master-science-business-analytics-msba/admissions/gmat-gre-options
  (GMAT/GRE opcional)
- https://gsm.ucdavis.edu/master-science-business-analytics-msba/admissions/application-process
  (documentos, taxa de inscricao, entrevista)
- Class of 2024 Placement Report (taxa de emprego e salario mais recentes encontrados)

Custo de vida (Aluguel, Mercado, Seguro Saude, etc.) foi mantido como
"Verificar": as paginas oficiais de cost-of-attendance da UC Davis
(financialaid.ucdavis.edu, gradstudies.ucdavis.edu) bloquearam o acesso
automatizado (HTTP 403), e os unicos numeros encontrados via terceiros eram
ou dados genericos de aluguel em San Francisco (nao especificos de aluno de
pos-graduacao) ou uma cifra de orcamento total desatualizada (ano letivo
2023-24) que nao pode ser decomposta com confianca nas colunas do schema.
"""

import shutil
import sys
from datetime import datetime

import openpyxl
from openpyxl.styles import PatternFill

sys.stdout.reconfigure(encoding="utf-8")

DB_PATH = "data/msba_programs_database.xlsx"
SHEET_DATA = "MSBA Programs"

VERIFICAR = "Verificar"

MATCH = {
    "Nome da Universidade": "UC Davis",
    "Estado": "CA",
    "Cidade": "Davis/San Francisco",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Duração (meses)": 12,
    "Destaques Currículo": (
        "Classes no campus de San Francisco (Civic Center); No. 1 globalmente "
        "em ROI e No. 4 em 'The Money' (QS 2024, 4 anos consecutivos no topo "
        "de ROI); practicum analítico ao longo de todo o ano com patrocinadores "
        "corporativos; currículo: liderança, gestão de ciclo de vida de dados, "
        "modelagem estatística, machine learning, pesquisa operacional"
    ),
    "Custo/Crédito (Intl) USD": 1560,  # CORRECAO: valor antigo (1490) era da coorte 2025; oficial p/ coorte verao 2026
    "Custo/Crédito (Residente) USD": 1560,
    "Tuition Total (Intl) USD": 68640,  # CORRECAO: 44 unidades x $1.560 (coorte 2026)
    "Tuition Total (Residente) USD": 68640,
    "GPA Mínimo": 3.0,
    "GMAT/GRE Obrigatório": (
        "Opcional (não enviar não prejudica a candidatura; o coursework "
        "quantitativo passa a ser a base de avaliação; dispensa também "
        "possível por GPA ou experiência profissional)"
    ),
    "Testes Inglês Aceitos": "TOEFL 100 (iBT) mínimo / IELTS 7.0 mínimo",  # CORRECAO: antes so "TOEFL/IELTS" vago
    "Matérias Pré-Requisito": "Pré-requisitos em computação, matemática e estatística",
    "Taxa Aplicação USD": 125,
    "Documentos Necessários": (
        "Transcripts (não oficiais na inscrição, oficiais após admissão), "
        "GRE/GMAT (opcional), resume (máx. 2 páginas), 2 ensaios (1 "
        "obrigatório + 1 opcional), 2 cartas de recomendação, TOEFL/IELTS "
        "(candidatos internacionais)"
    ),
    "Entrevista": (
        "Entrevista online assíncrona (one-way), podendo ser seguida de "
        "entrevista ao vivo por Zoom"
    ),
    "Taxa Empregabilidade %": 95,  # Class of 2024, dado mais recente encontrado
    "Salário Médio Pós USD": 131000,  # Class of 2024
    "Principais Empregadores": (
        "Amazon (maior contratante histórico da turma), Meta, BlackRock, e "
        "outras instituições financeiras"
    ),
    "Quão Realista (Intl)": (
        "Moderada para bolsa por mérito (automática, sem inscrição separada); "
        "baixa para fellowships por necessidade do Graduate Studies (muito "
        "limitadas para internacionais)"
    ),  # CORRECAO/refinamento: antes so "Moderada", sem distinguir tipo de bolsa
    "TA/RA Disponível": (
        "Comum em programas mais longos/PhD; não confirmado especificamente "
        "para o MSBA (1 ano)"
    ),
}


def main():
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    backup_path = f"data/msba_programs_database.backup-{timestamp}.xlsx"
    shutil.copy2(DB_PATH, backup_path)
    print(f"Backup criado em: {backup_path}")

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
        raise ValueError("Nenhuma linha encontrada para UC Davis. Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (UC Davis, Davis/San Francisco, CA)")

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
