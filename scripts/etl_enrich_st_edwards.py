# -*- coding: utf-8 -*-
"""
Data enrichment (30/out/2026): atualiza APENAS a linha da St. Edward's
University (Bill Munday School of Business, Austin, TX) no
msba_programs_database.xlsx. Foco: MS in Business Analytics and Applied AI,
presencial (exigido para F-1), STEM.

CONFIRMADO SEM ALTERACAO (páginas oficiais, out/2026):
- STEM; 33 créditos; ~16 meses (faixa 12-18); blocos de 7 semanas, uma
  matéria por vez; 10 matérias BANA + practicum com cliente real.
- Tuition 2026-27: US$1.242/crédito (privada, mesmo valor para todos);
  total 33 x 1.242 = 40.986. Comprehensive fee US$145/período (<12 horas).
- Taxa US$50, dispensada para quem aplica até o prazo de prioridade.
- GRE/GMAT não exigido; sem cartas de recomendação; só bacharelado.

CORRIGIDO/PREENCHIDO:
- Formato: acrescentada a regra para F-1 (mínimo 9 créditos/semestre, pelo
  menos 6 em aulas presenciais; programa 100% online não aceita F-1).
- Deadline: a rodada anterior dizia que 1 jun era "prazo final para
  internacionais", mas a página de prazos da pós não tem prazo separado
  para internacionais (o "May 1" da página internacional é da graduação).
  Mantido o 1 jun como prioridade.
- Seguro Saúde: Verificar -> 3.962. Plano 2026-27 (Academic HealthPlans):
  US$1.320,50 Fall + US$1.320,50 Spring/Summer = 2.641/ano; 16 meses ~
  3 períodos = 3.961,50. Obrigatório para quem cursa 6+ créditos, salvo
  waiver. Há ainda uma taxa de US$50 do Student Health Center (não somada:
  o folheto não diz se é por período ou por ano).
- Custos mensais: Verificar -> mesmos valores da linha UT Austin (mesma
  cidade): aluguel 1.700, mercado 300, seguro residencial 15, utilities 185,
  celular 50, WiFi 60, transporte 204, outros 150. São estimativas para
  Austin, não números da St. Edward's.
- Custo Vida Total: Verificar -> 46.586, fórmula do dashboard
  (computeCustoVidaTotal): soma mensal 2.664 x 16 meses + seguro 3.962.
- Bolsas: a página de bolsas só lista bolsas de graduação e a página de
  tuition da pós só cita empréstimos federais e bolsas externas. Nenhuma
  assistantship de pós divulgada.
- Testes de inglês: sem nota mínima publicada para a pós. A única tabela
  oficial com notas é a do intercâmbio (IELTS 6.5 / TOEFL iBT 79 /
  Duolingo 105), registrada como referência.
- Principais Empregadores: lista da página do programa (NXP Semiconductor,
  Box, Ready.net, City of Austin, TrueNorth Companies, Texas DPS).
- Entrevista, Experiência Profissional, Dificuldade: preenchidos a partir
  dos requisitos (só bacharelado; entrevista só no MA Counseling).

MANTIDO COMO Verificar: GPA mínimo/médio, taxa de empregabilidade, salário
médio (não há relatório do MSBA; o salário inicial de US$63.712 da US News
mistura graduação e pós de business), taxa de aceitação, tamanho da turma,
GMAT médio, Ranking US News/QS (o "#12 nacional" citado pelo programa é
um ranking de engajamento de programas online, não de Business Analytics;
foi para Destaques Currículo).

INCERTEZA: com 9 créditos/semestre só em Fall/Spring, 33 créditos levam ~4
semestres (~21 meses). Os 16 meses dependem de cursar Summer presencial.
Confirmar com o programa.

Fontes (out/2026):
- stedwards.edu/academics/majors-programs/master-science-business-analytics
- stedwards.edu/admission/graduate/deadlines-checklist
- stedwards.edu/cost-aid/graduate-tuition-fees
- stedwards.edu/cost-aid/scholarships
- stedwards.edu/academics/graduate-professional-studies/international-students
- stedwards.edu/campus-services/health-and-counseling-center/health-services/student-health-insurance
- myahpcare.com/wp-content/uploads/St.-Edwards_PHF_26_27.pdf
- stedwards.edu/academics/student-success/international-students/exchange-students

Uso: python scripts/etl_enrich_st_edwards.py (a partir da raiz).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"
MATCH = ("St. Edward's University", "TX", "Austin")

UPDATES = {
    "Formato": (
        "Presencial/híbrido no campus de Austin; F-1 precisa de 9+ créditos por semestre, "
        "com pelo menos 6 em aulas presenciais, e não pode cursar a versão 100% online"
    ),
    "Destaques Currículo": (
        "Blocos de 7 semanas, uma matéria por vez; R, Python, SQL, Power BI, big data, "
        "predictive analytics, simulação, otimização, IA/machine learning (GenAI, NLP, "
        "agentic AI); practicum final com projeto de consultoria para cliente real; turmas "
        "pequenas; AACSB; US News 2024-25: #12 nacional e #1 no Texas em engajamento "
        "estudantil (ranking de programas online, não de Business Analytics)"
    ),
    "Aluguel Privativo/mês USD": 1700,
    "Mercado/mês USD": 300,
    "Seguro Saúde (programa) USD": 3962,
    "Seguro Residencial/mês USD": 15,
    "Utilities/mês USD": 185,
    "Celular/mês USD": 50,
    "WiFi/mês USD": 60,
    "Transporte/mês USD": 204,
    "Outros Gastos/mês USD": 150,
    # Mesma fórmula de computeCustoVidaTotal (derived.ts): soma mensal 2.664 x 16 meses + seguro 3.962.
    "Custo Vida Total (programa) USD": 46586,
    "Bolsas Disponíveis": "Não",
    "Tipos de Bolsa": (
        "Nenhuma bolsa institucional de pós divulgada (a página de bolsas só lista bolsas de "
        "graduação); a pós é financiada por empréstimos federais (só para americanos) e bolsas "
        "externas"
    ),
    "Quão Realista (Intl)": (
        "Baixa: sem bolsa de mérito nem assistantship publicada para a pós; perguntar ao "
        "Student Financial Services se há algum desconto para o MSBA"
    ),
    "TA/RA Disponível": "Não divulgado (nenhuma TA/RA de pós no site)",
    "Processo TA/RA": (
        "Não se aplica; só empregos on-campus gerais (F-1: até 20h/semana durante as aulas)"
    ),
    "Graduate Assistantships": "Nenhuma divulgada para o MSBA",
    "Principais Empregadores": (
        "NXP Semiconductor, Box, Ready.net, City of Austin, TrueNorth Companies, Texas "
        "Department of Public Safety (empregadores de alunos citados na página do programa); "
        "sem relatório de empregabilidade ou salário do MSBA"
    ),
    "Experiência Profissional": "Não obrigatória (só bacharelado)",
    "Testes Inglês Aceitos": (
        "TOEFL, IELTS ou Duolingo; sem nota mínima publicada para a pós (referência: o "
        "intercâmbio exige IELTS 6.5 / TOEFL iBT 79 / Duolingo 105)"
    ),
    "Deadline Aplicação": (
        "Fall: 1 jun (prioridade, isenta a taxa), 15 jul (final); Spring: 1 out (prioridade), "
        "15 nov (final). Não há prazo separado para internacionais"
    ),
    "Rounds Aplicação": (
        "Prazo de prioridade e prazo final por semestre; a taxa de US$50 é dispensada para quem "
        "aplica até o prazo de prioridade; F-1 deve ser admitido antes de entrar nos EUA, então "
        "convém aplicar no prazo de prioridade para ter tempo para o I-20"
    ),
    "Documentos Necessários": (
        "Personal statement (histórico, interesses, motivo do programa e objetivos); transcripts "
        "oficiais (enviados depois da aplicação); resume com histórico de trabalho detalhado; "
        "avaliação de diploma estrangeiro (WES, ECE ou similar); teste de inglês; para o I-20, "
        "modality attestation (declaração de modalidade presencial)"
    ),
    "Cartas de Recomendação": "0",
    "Entrevista": "Não exigida (só o MA Counseling tem entrevista)",
    "Dificuldade Entrada": "Acessível",
}


def main():
    rotate_backup()
    print("Backup rotacionado (política fixa) em: data/msba_programs_database_backup.xlsx")

    wb = openpyxl.load_workbook(DB_PATH)
    ws = wb[SHEET_DATA]

    headers = [cell.value for cell in ws[1]]
    col = {h: i + 1 for i, h in enumerate(headers)}

    unknown = [h for h in UPDATES if h not in col]
    if unknown:
        raise ValueError(f"Colunas inexistentes no schema: {unknown}")

    target_rows = [
        row[0].row
        for row in ws.iter_rows(min_row=2)
        if (
            row[col["Nome da Universidade"] - 1].value,
            row[col["Estado"] - 1].value,
            row[col["Cidade"] - 1].value,
        ) == MATCH
    ]
    if len(target_rows) != 1:
        raise ValueError(f"Esperava 1 linha para {MATCH}, achei {len(target_rows)}. Abortando.")
    row_idx = target_rows[0]
    print(f"Linha alvo confirmada: {row_idx} (St. Edward's University, Austin, TX)")

    no_fill = PatternFill(fill_type=None)
    changes = []
    for header, new_value in UPDATES.items():
        cell = ws.cell(row=row_idx, column=col[header])
        if cell.value != new_value:
            changes.append((header, cell.value, new_value))
        cell.value = new_value
        cell.fill = no_fill

    wb.save(DB_PATH)

    print("\nCampos atualizados (Antes -> Depois):")
    for header, old, new in changes:
        print(f"  - {header}:")
        print(f"      Antes:  {old!r}")
        print(f"      Depois: {new!r}")


if __name__ == "__main__":
    main()
