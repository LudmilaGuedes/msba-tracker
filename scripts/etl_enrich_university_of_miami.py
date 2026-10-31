"""
Data enrichment: atualiza APENAS a linha da 'University of Miami' (Coral Gables, FL)
no msba_programs_database.xlsx, com base em pesquisa web focada no programa
MS in Business Analytics (Miami Herbert Business School).

ATENCAO CRITICA: nao confundir com 'Miami U Ohio' (Oxford, OH), que e uma
universidade diferente e ja existe como linha separada na planilha. O
matching abaixo exige Nome da Universidade == 'University of Miami' E
Estado == 'FL' E Cidade == 'Coral Gables' para blindar contra isso.

Fontes principais (pesquisadas manualmente, fora deste script):
- https://www.mbs.miami.edu/graduate/find-and-compare-programs/business-analytics/cost-of-attendance.html
- https://www.mbs.miami.edu/graduate/find-and-compare-programs/business-analytics/requirements.html
- https://bulletin.miami.edu/graduate-academic-programs/business/specialized-master-degrees/business-analytics-ms/
- https://go.miami.edu/analytics
- https://grad-prof.miami.edu/cost/graduate-costs/international-graduate-student-folder/ (orcamento
  oficial de custo de vida para pos-graduandos internacionais, 2026-2027, usado como proxy
  para os campos de custo de vida do MSBA, que nao tem orcamento proprio publicado)
"""

import shutil
from datetime import datetime

import openpyxl
from openpyxl.styles import PatternFill

DB_PATH = "data/msba_programs_database.xlsx"
SHEET_DATA = "MSBA Programs"

VERIFICAR = "Verificar"

MATCH = {
    "Nome da Universidade": "University of Miami",
    "Estado": "FL",
    "Cidade": "Coral Gables",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "STEM Designated": "Sim",
    "Início": "Fall, Spring",
    "Destaques Currículo": (
        "Core: Python, Applied Regression Analysis I/II, Machine Learning for "
        "Data Analytics I, ML Methods for AI, Capstone (2 créditos) com parceiros "
        "da indústria; eletivas em accounting/supply chain/marketing analytics, "
        "cloud computing, big data analytics; DIRPA institute top 10 em analytics/"
        "data science; turma com 40% de alunos nascidos fora dos EUA; trilha "
        "part-time de 2 anos também disponível"
    ),
    "Total Créditos": 33,  # CORRECAO: planilha tinha 30, oficial (bulletin + cost-of-attendance) confirma 33
    "Custo/Crédito (Intl) USD": 2561,
    "Custo/Crédito (Residente) USD": 2561,  # privada: mesma taxa para todos
    "Tuition Total (Intl) USD": 84523,  # CORRECAO: valor antigo (78773) vinha de agregador nao oficial
    "Tuition Total (Residente) USD": 84523,
    "Fees Adicionais USD": 1438,  # $714/semestre x 2
    "Aluguel Privativo/mês USD": 2181,  # orcamento oficial UM p/ pos-graduandos intl, 2026-27 (19636/9 meses)
    "Mercado/mês USD": 1030,  # 9276/9
    "Seguro Saúde (programa) USD": 4483,  # premio anual oficial 2026-27
    "Transporte/mês USD": 464,  # 4180/9
    "Outros Gastos/mês USD": 887,  # Personal (517) + Books (370), sem coluna propria p/ livros
    "Custo Vida Total (programa) USD": 45563,  # soma do orcamento oficial (moradia+comida+livros+pessoal+transporte+seguro)
    "Bolsas Disponíveis": "Sim",
    "Tipos de Bolsa": (
        "Merit-based automático (sem inscrição separada), baseado em histórico "
        "acadêmico, experiência profissional e potencial de liderança"
    ),
    "Quão Realista (Intl)": (
        "Alta (74% dos admitidos no Fall 2025 receberam bolsa por mérito; "
        "+$10 milhões concedidos aos mestrados/doutorados em 2025)"
    ),
    "Taxa Empregabilidade %": 94,
    "Principais Empregadores": (
        "McKinsey, BCG, Deloitte, Amazon, Cisco, Capital One, American Express, "
        "Citi, Visa, Mastercard, Carnival Cruise, Chicago Cubs, ChenMed, EY"
    ),
    "GMAT/GRE Obrigatório": "Opcional (recomendado para bolsa por mérito máxima)",
    "Experiência Profissional": "Não obrigatória",
    "Matérias Pré-Requisito": "Nenhum pré-requisito formal além do bacharelado",
    "Testes Inglês Aceitos": (
        "TOEFL 94 / IELTS 7.0 / Duolingo 125 (dispensa se diploma de instituição "
        "dos EUA ou de país de língua inglesa)"
    ),
    "Deadline Aplicação": "Fall: 1 jun; Spring: 1 fev – 15 mar",
    "Taxa Aplicação USD": 100,
    "Documentos Necessários": (
        "Application online, histórico acadêmico (avaliação WES para diplomas "
        "internacionais), resume, 2 ensaios, cartas de recomendação"
    ),
    "Entrevista": "Por convite apenas (virtual, ~30 min)",
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
        raise ValueError("Nenhuma linha encontrada para University of Miami (Coral Gables, FL). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (University of Miami, Coral Gables, FL)")

    verificar_fill = PatternFill(start_color="00FFF3CD", end_color="00FFF3CD", fill_type="solid")
    no_fill = PatternFill(fill_type=None)

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
