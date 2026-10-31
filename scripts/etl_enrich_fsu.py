"""
Data enrichment: atualiza APENAS a linha da 'Florida State' (Herbert
Wertheim College of Business, Tallahassee, FL) no msba_programs_database.xlsx.

Foco exclusivo: Master of Science in Business Analytics (MS-BA), presencial,
STEM, programa lock-step de 3 semestres (Summer/Fall/Spring).

CORRECOES DA AUDITORIA:
- 'Business School' estava como 'College of Business' -- nome oficial
  atual e 'Herbert Wertheim College of Business'.
- 'Início' estava como 'Fall' -- o programa e lock-step de 3 semestres
  comecando no SUMMER (confirmado por 'Priority: January 15 (summer
  entry)' e pela estrutura Summer/Fall/Spring na pagina oficial).

Fontes principais (pesquisadas manualmente, fora deste script):
- https://wertheim.fsu.edu/msba (pagina oficial: creditos, curriculo, TOEFL/IELTS,
  documentos, deadlines)
- busca por tuition per credit hour 2026-27 (confirma $479,32/credito residente e
  $1.173,86/credito nao-residente, batendo exatamente com os totais encontrados)
- https://cge.fsu.edu (funding para internacionais: assistantships ~$29k/ano + waiver)
- busca por custo de atendimento de pos-graduacao da FSU (moradia/alimentacao no campus)

Custo de vida ficou parcialmente preenchido: so encontrei moradia+alimentacao
(fonte oficial-adjacente), sem dados confiaveis de transporte/pessoal/seguro
saude especificos para pos-graduandos da FSU -- por isso Custo Vida Total
(que exigiria somar todas as categorias) permanece "Verificar" em vez de ser
calculado de forma incompleta.
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
    "Nome da Universidade": "Florida State",
    "Estado": "FL",
    "Cidade": "Tallahassee",
}

# Apenas os campos que estao sendo corrigidos/enriquecidos.
# Campos nao listados aqui permanecem inalterados (incluindo os que continuam "Verificar").
UPDATES = {
    "Business School": "Herbert Wertheim College of Business",  # CORRECAO: nome oficial atual
    "Início": "Summer",  # CORRECAO: programa lock-step Summer/Fall/Spring, entrada e no Summer, nao Fall
    "Destaques Currículo": (
        "Programa lock-step de 3 semestres (Summer/Fall/Spring); 9 cursos "
        "centrais (27 créditos) + 1 eletiva (3 créditos) + desenvolvimento "
        "profissional (1 crédito/semestre); núcleo: otimização probabilística, "
        "programação para analytics, métodos quantitativos, mineração de "
        "dados, gestão e visualização de dados, capstone; eletivas em supply "
        "chain, healthcare, fintech, algoritmos"
    ),
    "Total Créditos": 33,
    "Custo/Crédito (Residente) USD": 479.32,
    "Custo/Crédito (Intl) USD": 1173.86,
    "Tuition Total (Residente) USD": 15818,  # 33 x 479,32
    "Tuition Total (Intl) USD": 38737,  # 33 x 1173,86
    "Aluguel Privativo/mês USD": 936,  # orçamento oficial-adjacente FSU (8420/9 meses)
    "Mercado/mês USD": 622,  # 5600/9 meses
    "GMAT/GRE Obrigatório": (
        "Opcional (dispensa automática com GPA geral ≥ 3.4 e GPA "
        "upper-division ≥ 3.2 em curso de negócios AACSB)"
    ),
    "Testes Inglês Aceitos": (
        "TOEFL 600 (PBT, antes de 21/jan/2026) ou 5.0 (a partir de "
        "21/jan/2026, conforme página oficial); IELTS 7.0; PTE 66; "
        "Duolingo 125"
    ),
    "Matérias Pré-Requisito": "Cálculo, probabilidade e estatística, programação (nível universitário)",
    "Documentos Necessários": (
        "Applicant statement, resume (com experiência profissional), 2 "
        "cartas de recomendação (3 preferível), histórico acadêmico, "
        "GMAT/GRE (opcional); internacionais: avaliação de credencial via NACES"
    ),
    "Taxa Aplicação USD": 30,
    "Deadline Aplicação": "Prioritária: 15 jan (também deadline para concorrer a assistantship); Final: 1 mar",
    "TA/RA Disponível": (
        "Possível (GA com stipend ~$29.000/ano + waiver de tuition + "
        "seguro saúde); não garantido, vagas limitadas por departamento"
    ),
    "Graduate Assistantships": (
        "Stipend ~$29.000/ano + waiver de tuition + seguro saúde; não "
        "garantido com a admissão, depende de disponibilidade de vaga"
    ),
    "Bolsas Disponíveis": "Sim",
    "Tipos de Bolsa": (
        "Bolsas por mérito do Herbert Wertheim College of Business; "
        "Graduate Assistantships com stipend + waiver"
    ),
    "Quão Realista (Intl)": (
        "Moderada (GAs disponíveis mas limitadas e não garantidas; "
        "concorrer exige aplicar até o deadline prioritário de 15 jan)"
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
        raise ValueError("Nenhuma linha encontrada para Florida State (Tallahassee, FL). Abortando.")
    if len(target_rows) > 1:
        raise ValueError(f"Mais de uma linha corresponde ao filtro (linhas {target_rows}). Abortando por seguranca.")

    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (Florida State, Tallahassee, FL)")

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
