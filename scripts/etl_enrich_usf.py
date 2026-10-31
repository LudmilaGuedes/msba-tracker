"""
Data enrichment: atualiza APENAS a linha da 'USF' (University of South Florida,
Tampa, FL) no msba_programs_database.xlsx.

ATENCAO CRITICA DE DESAMBIGUACAO: NAO confundir com University of San
Francisco (usfca.edu, California). O matching abaixo exige Nome da
Universidade == 'USF' E Estado == 'FL' E Cidade == 'Tampa'. Durante a
pesquisa, um resultado de busca da usfca.edu (San Francisco) apareceu
misturado nos resultados e foi descartado manualmente sem ser usado.

ACHADO CRITICO DA AUDITORIA: o programa que constava na planilha ('MS in AI
& Business Analytics') esta sendo DESCONTINUADO/RENOMEADO neste exato Fall
2026 -- e o mais recente de uma linhagem de renomeacoes ('MS in Business
Analytics and Information Systems' -> 'MS in AI and Business Analytics' ->
'MS in AI in Business', a partir de Fall 2026). O novo programa tem duas
trilhas: 'Digital Transformation' (100% online) e 'Enterprise Integration'
(presencial, STEM, full-time) -- esta ultima e a sucessora direta do
programa presencial que ja estava documentado, e foi usada para atualizar
esta linha.

Custo de vida (Aluguel, Mercado, Transporte, Outros Gastos, Custo Vida
Total) foi mantido como "Verificar": os unicos numeros encontrados vinham
de agregadores terceiros com valores internamente inconsistentes (ex.:
"moradia no campus" custando mais que "moradia + alimentacao no campus"),
o que nao e confiavel o suficiente para popular a planilha, especialmente
dado o risco de contaminacao com dados da USF errada nesta pesquisa.

Fontes principais (pesquisadas manualmente, fora deste script):
- https://www.usf.edu/business/graduate/ms-ai-business-enterprise-integration/index.aspx
  (pagina oficial da nova trilha presencial, lancamento Fall 2026)
- https://www.usf.edu/business/graduate/ms-ai-business/index.aspx (pagina do programa novo,
  com as duas trilhas e tuition por credito)
- https://www.usf.edu/business/graduate/ms-ai-business-analytics/ (pagina do programa
  predecessor, usada para requisitos de admissao ainda nao republicados na pagina nova)
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
    "Nome da Universidade": "USF",
    "Estado": "FL",
    "Cidade": "Tampa",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Nome Oficial do Programa": (
        "Master of Science in Artificial Intelligence (AI) in Business – "
        "Enterprise Integration"
    ),  # CORRECAO: nome anterior ('MS in AI & Business Analytics') descontinuado a partir do Fall 2026
    "Duração (meses)": 14,  # fonte oficial: "12-16 months", usado o ponto medio
    "Início": (
        "Fall (lançamento do programa reformulado em 2026; demais intakes a confirmar)"
    ),
    "Destaques Currículo": (
        "Trilha presencial 'Enterprise Integration' (distinta da 'Digital "
        "Transformation', 100% online); sucessora do 'MS in AI and Business "
        "Analytics', que por sua vez sucedeu o 'MS in Business Analytics and "
        "Information Systems'; carreiras-alvo: GenAI analyst, AI engineer, "
        "AI/ML architect, data pipeline architect; campus Tampa (Muma College "
        "of Business)"
    ),
    "Total Créditos": 30,
    "Custo/Crédito (Intl) USD": 1030.20,
    "Custo/Crédito (Residente) USD": 466.34,
    "Tuition Total (Intl) USD": 30906,  # CORRECAO: valor antigo (17325) era do programa predecessor
    "Tuition Total (Residente) USD": 13990,
    "GMAT/GRE Obrigatório": (
        "Obrigatório, salvo dispensa (alumni da Muma College com GPA ≥ 3.5, "
        "ou 3+ anos de experiência técnica relevante)"
    ),
    "Testes Inglês Aceitos": "TOEFL 79 (iBT) mínimo; Duolingo e IELTS também aceitos",  # CORRECAO: antes so "TOEFL/IELTS" vago
    "Documentos Necessários": (
        "Histórico acadêmico, experiência profissional relevante, 2 cartas de "
        "recomendação (acadêmicas ou profissionais), statement of purpose"
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
        raise ValueError("Nenhuma linha encontrada para USF (Tampa, FL). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (USF, Tampa, FL) -- NAO University of San Francisco")

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
