# -*- coding: utf-8 -*-
"""
Data enrichment (rodada 3, 20/out/2026): atualiza APENAS a linha da 'USC'
(University of Southern California, Marshall School of Business, Los
Angeles, CA) no msba_programs_database.xlsx. Foco: MS in Business Analytics,
presencial, STEM. (Não confundir com University of South Carolina.)

CONFIRMADO SEM ALTERACAO (páginas oficiais do MSBA, out/2026):
- Tuition 2025-26: US$2.541/unit x 33 units = US$83.853 (mesmo valor para
  intl e residente, privada); fees US$4.046 (3 semestres) / US$4.243 (4).
- Bolsas: mérito limitadas, avaliação holística, sem inscrição separada,
  internacionais elegíveis, prioridade nos Rounds 1-2.
- Rounds Fall 2027: R1 1/nov/2026, R2 15/jan/2027, R3 15/abr/2027;
  taxa US$175 (nenhuma isenção mencionada); GMAT/GRE não exigidos; sem
  cartas; KIRA por convite; pré-requisito de estatística; sem GPA mínimo.
- Class profile Fall 2025: 1.995 aplicações, turma 254, GPA médio 3.60;
  94% dos formados do Fall 2024 empregados; lista de empregadores.

CORRIGIDO/PREENCHIDO:
- Ranking QS: 14 -> 19. QS Business Analytics 2027 (divulgado set/2026):
  #19, score 82,9 (era #16 em 2026, como a própria Marshall cita).
- Ranking US News: 18 -> 16. O 18 era o ranking do MBA full-time (2024),
  não de analytics. A US News lista a Marshall em #16 (tie) em Business
  Analytics (ranking de especialidade; o MBA está em #25).
- Seguro Saúde: Verificar -> 5.100. USC SHIP on-campus 2026-27: Fall
  US$1.313, Spring/Summer US$2.474. O MSBA de 3 semestres (Fall-Spring-Fall,
  ~16 meses) cobre Fall + Spring/Summer + Fall = 1.313 + 2.474 + 1.313.
  Estimativa com os preços de 2026-27.
- Custo Vida Total: 42.684 -> 59.612. O valor antigo usava 12 meses e sem
  seguro; recalculado com a fórmula do dashboard (computeCustoVidaTotal):
  soma mensal 3.407 x 16 meses + seguro 5.100.
- Destaques Currículo: % de internacionais 54% -> 78% (class profile
  Fall 2025: 78% internacionais / 22% domésticos).
- Testes Inglês: a USC passou a usar a nova escala do TOEFL (exames a
  partir de 21/jan/2026: 4.5, com 4+ por seção; antes: iBT 90, 20+ por
  seção). Duolingo continua fora da lista.
- TA/RA Disponível / Processo TA/RA / Graduate Assistantships: estavam
  Verificar. Regras gerais da Graduate School (alunos de mestrado são
  elegíveis se full-time e com GPA >= 3.0, mas PhD tem prioridade); na
  Marshall só aparecem vagas de grader/instructional assistant (Data
  Sciences and Operations) e o pool de assistentes do Behavioral Research
  Lab (US$25/h, até 20h/semana, tarefas curtas).

MANTIDO: Salário Médio Pós 100.000. A Marshall não publica salário do
MSBA; agregadores citam ~US$100.863 (média da última turma), sem fonte
oficial. GMAT Médio fica Verificar (GMAT/GRE não exigidos).

Fontes (out/2026):
- marshall.usc.edu/programs/graduate-programs/specialized-masters/ms-business-analytics (+ /admissions, /tuition-and-fees, /class-profile)
- gradadm.usc.edu/prospective-international-students/english-proficiency/
- dworakpeck.usc.edu/admissions/tuition-and-financial-aid/tuition-fees (SHIP 2026-27)
- usnews.com/best-graduate-schools/top-business-schools/university-of-southern-california-01034
- planetegrandesecoles.com (tabela QS Business Analytics 2027, 05/out/2026)
- students.marshall.usc.edu (Behavioral Research Lab), marshall.usc.edu DSO student job board
- graduateschool.usc.edu TA/RA/AL Handbook

Uso: python scripts/etl_enrich_usc_round3.py (a partir da raiz).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = ("USC", "CA", "Los Angeles")

UPDATES = {
    "Ranking QS": 19,
    "Ranking US News": 16,
    "Seguro Saúde (programa) USD": 5100,
    # Mesma fórmula de computeCustoVidaTotal (derived.ts): soma mensal 3.407 x 16 meses + seguro 5.100.
    "Custo Vida Total (programa) USD": 59612,
    "Destaques Currículo": (
        "33 units: 16,5 de núcleo (decisão orientada a dados, computação estatística e "
        "visualização, Python, SQL/bancos de dados, estratégia) + estágio/experiência "
        "prática + 16,5 de eletivas (Deep Learning, Marketing/Operations Analytics, NLP, "
        "Product/Project Management, Consulting, Finance); opção de 3 ou 4 semestres; "
        "78% internacionais (Fall 2025)"
    ),
    "Testes Inglês Aceitos": (
        "TOEFL ou IELTS obrigatório para internacionais (página do MSBA); pisos da USC para "
        "mestrado (isentam do exame ISE): TOEFL 4.5 com 4+ por seção (nova escala, exames a "
        "partir de 21/jan/2026; antes: iBT 90 com 20+ por seção) / IELTS 6.5 com 6.0+ por "
        "banda / PTE 61 com 53+ por seção / Cambridge C1 176; USC não fixa mínimo de "
        "admissão (usados para placement); Duolingo não consta na lista de testes aceitos "
        "da pós-graduação da USC"
    ),
    "TA/RA Disponível": (
        "Limitado -- pela regra da Graduate School, alunos de mestrado são elegíveis "
        "(full-time, GPA ≥ 3.0), mas PhD tem prioridade; para o MSBA só aparecem vagas de "
        "grader/instructional assistant e de assistente de pesquisa por hora"
    ),
    "Processo TA/RA": (
        "Formulário de interesse do departamento de Data Sciences and Operations (grader, "
        "instructional assistant, student worker); pool de assistentes do Marshall "
        "Behavioral Research Lab (US$25/h, até 20h/semana, tarefas de 1-2 meses; inscrição "
        "anual, prazo do Fall 2026 foi 15/jun/2026)"
    ),
    "Graduate Assistantships": (
        "Nenhuma assistantship divulgada para o MSBA; só trabalho pago por hora no campus"
    ),
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
    print(f"Linha alvo confirmada: {row_idx} (USC, Los Angeles, CA)")

    no_fill = PatternFill(fill_type=None)
    changes = []
    for header, new_value in UPDATES.items():
        cell = ws.cell(row=row_idx, column=col[header])
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
