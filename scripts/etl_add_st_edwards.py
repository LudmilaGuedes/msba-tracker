# -*- coding: utf-8 -*-
"""
Adiciona a St. Edward's University (Bill Munday School of Business, Austin,
TX) com o Master of Science in Business Analytics and Applied AI.

Contexto: em 20/out/2026 foi feita uma varredura para conferir se os
programas do Texas na base eram os unicos MSBA STEM presenciais do estado.
A St. Edward's foi o unico MSBA STEM com opcao presencial que faltava.

Fontes (out/2026):
- stedwards.edu/academics/majors-programs/master-science-business-analytics
  (STEM; 33 creditos; ~16 meses, ou 12-18; online, hibrido e presencial no
  campus de Austin; F-1 so pode cursar as opcoes presenciais, nao o 100%
  online; blocos de 7 semanas, uma materia por vez; 10 materias BANA;
  practicum com cliente real; requisito: bacharelado, so algebra basica)
- stedwards.edu/cost-aid/graduate-tuition-fees (Fall 2026/Spring 2027:
  US$1.242/credito; comprehensive fee US$145/periodo com menos de 12 horas)
- stedwards.edu/admission/graduate/deadlines-checklist (Fall programas
  online: prioridade 1 jun, final 15 jul; Spring todos: prioridade 1 out,
  final 15 nov; taxa US$50, dispensada no prazo de prioridade; GRE nao
  exigido; TOEFL/IELTS/Duolingo sem nota minima publicada; avaliacao de
  diploma estrangeiro WES/ECE)
- stedwards.edu/admission/international (prazo final de Fall para
  internacionais: 1 jun)

CALCULOS:
- Tuition total: 33 x 1.242 = 40.986 (privada: mesmo valor para residente).
- Fees: 145 x 4 periodos (~16 meses, part-time) = 580, estimativa.

INCERTEZAS REGISTRADAS:
- Creditos: a pagina diz 33, mas lista 10 materias de 3 creditos (30).
  Agregadores e uma landing page antiga dizem 30 creditos / 18 meses. Foi
  usado o valor oficial atual (33).
- Nao ha horario das aulas presenciais nem quantas materias sao oferecidas
  presencialmente a cada bloco: confirmar com o programa se da para cumprir
  a exigencia do F-1 (no maximo 1 materia online por periodo).
"""

import sys
from datetime import datetime

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"
SHEET_SUMMARY = "Resumo"
VERIFICAR = "Verificar"

NAME, STATE, CITY = "St. Edward's University", "TX", "Austin"

NEW_PROGRAM = {
    "Nome da Universidade": NAME,
    "Estado": STATE,
    "Cidade": CITY,
    "Business School": "Bill Munday School of Business",
    "Tipo (Pub/Priv)": "Privada",
    "Nome Oficial do Programa": "Master of Science in Business Analytics and Applied AI",
    "STEM Designated": "Sim",
    "Formato": (
        "Presencial/híbrido no campus de Austin (F-1 só pode cursar as opções "
        "presenciais); também há versão 100% online"
    ),
    "Duração (meses)": 16,
    "Início": "Fall, Spring",
    "Tamanho Turma": VERIFICAR,
    "Destaques Currículo": (
        "Blocos de 7 semanas, uma matéria por vez; R, Python, SQL, Power BI, big data, "
        "predictive analytics, simulação, otimização, IA/machine learning (GenAI, NLP, "
        "agentic AI); practicum final com projeto de consultoria para cliente real"
    ),
    "Total Créditos": 33,
    "Custo/Crédito (Intl) USD": 1242,
    "Custo/Crédito (Residente) USD": 1242,
    "Tuition Total (Intl) USD": 40986,
    "Tuition Total (Residente) USD": 40986,
    "Fees Adicionais USD": 580,
    "Aluguel Privativo/mês USD": VERIFICAR,
    "Mercado/mês USD": VERIFICAR,
    "Seguro Saúde (programa) USD": VERIFICAR,
    "Seguro Residencial/mês USD": VERIFICAR,
    "Utilities/mês USD": VERIFICAR,
    "Celular/mês USD": VERIFICAR,
    "WiFi/mês USD": VERIFICAR,
    "Transporte/mês USD": VERIFICAR,
    "Outros Gastos/mês USD": VERIFICAR,
    "Custo Vida Total (programa) USD": VERIFICAR,
    "Bolsas Disponíveis": VERIFICAR,
    "Tipos de Bolsa": VERIFICAR,
    "Quão Realista (Intl)": VERIFICAR,
    "TA/RA Disponível": VERIFICAR,
    "Processo TA/RA": VERIFICAR,
    "Caminho Tuition Residente": "Não se aplica (privada: mesmo valor para todos)",
    "Graduate Assistantships": VERIFICAR,
    "Ranking US News": VERIFICAR,
    "Ranking QS": VERIFICAR,
    "Taxa Empregabilidade %": VERIFICAR,
    "Salário Médio Pós USD": VERIFICAR,
    "Principais Empregadores": VERIFICAR,
    "GPA Mínimo": VERIFICAR,
    "GPA Médio Admitidos": VERIFICAR,
    "GMAT/GRE Obrigatório": "Não obrigatório",
    "GMAT Médio": VERIFICAR,
    "Taxa Aceitação %": VERIFICAR,
    "Experiência Profissional": VERIFICAR,
    "Matérias Pré-Requisito": "Nenhuma além de álgebra básica",
    "Testes Inglês Aceitos": "TOEFL, IELTS ou Duolingo (nota mínima não publicada)",
    "Deadline Aplicação": (
        "Fall: 1 jun (prioridade e prazo final para internacionais), 15 jul (final, "
        "domésticos); Spring: 1 out (prioridade), 15 nov (final)"
    ),
    "Taxa Aplicação USD": 50,
    "Rounds Aplicação": (
        "Prazo de prioridade e prazo final; a taxa de US$50 é dispensada para quem "
        "aplica até o prazo de prioridade"
    ),
    "Documentos Necessários": (
        "Personal statement; transcripts oficiais; resume; avaliação de diploma "
        "estrangeiro (WES ou ECE); teste de inglês"
    ),
    "Entrevista": VERIFICAR,
    "Dificuldade Entrada": VERIFICAR,
}


def main():
    rotate_backup()
    print("Backup rotacionado (política fixa) em: data/msba_programs_database_backup.xlsx")

    wb = openpyxl.load_workbook(DB_PATH)
    ws = wb[SHEET_DATA]
    headers = [cell.value for cell in ws[1]]
    col = {h: i + 1 for i, h in enumerate(headers)}

    missing = [h for h in headers if h not in NEW_PROGRAM]
    extra = [h for h in NEW_PROGRAM if h not in col]
    if missing or extra:
        raise ValueError(f"Schema divergente. Faltando: {missing}; sobrando: {extra}")

    for row in ws.iter_rows(min_row=2, values_only=True):
        if (row[col["Nome da Universidade"] - 1], row[col["Estado"] - 1], row[col["Cidade"] - 1]) == (
            NAME, STATE, CITY
        ):
            raise ValueError(f"{NAME} ({CITY}, {STATE}) já existe na base. Abortando.")

    ref_cell = ws.cell(row=2, column=1)
    base_font = Font(name=ref_cell.font.name, sz=ref_cell.font.sz)
    side = Side(style="thin", color="00D0D0D0")
    base_border = Border(left=side, right=side, top=side, bottom=side)
    base_alignment = Alignment(vertical="top", wrap_text=True)
    verificar_fill = PatternFill(start_color="00FFF3CD", end_color="00FFF3CD", fill_type="solid")

    new_row = ws.max_row + 1
    for header, col_idx in col.items():
        value = NEW_PROGRAM[header]
        cell = ws.cell(row=new_row, column=col_idx, value=value)
        cell.font = base_font
        cell.border = base_border
        cell.alignment = base_alignment
        if value == VERIFICAR:
            cell.fill = verificar_fill

    last_col_letter = ws.cell(row=1, column=len(headers)).column_letter
    ws.auto_filter.ref = f"A1:{last_col_letter}{ws.max_row}"

    total = ws.max_row - 1
    summary = wb[SHEET_SUMMARY]
    for row in summary.iter_rows(min_row=2):
        label = row[0].value
        if label in ("Total de Programas", "Programas Presenciais STEM"):
            print(f"Resumo '{label}': {row[1].value} -> {total}")
            row[1].value = total
        if label == "Data de Compilação":
            row[1].value = f"{row[1].value} + St. Edward's em {datetime.now().strftime('%d/%m/%Y')}"

    wb.save(DB_PATH)
    print(f"\n{NAME} adicionada na linha {new_row}.")
    filled = sum(1 for v in NEW_PROGRAM.values() if v != VERIFICAR)
    print(f"{filled} de {len(NEW_PROGRAM)} campos preenchidos; {len(NEW_PROGRAM) - filled} como 'Verificar'.")


if __name__ == "__main__":
    main()
