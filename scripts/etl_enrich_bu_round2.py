# -*- coding: utf-8 -*-
"""
Data enrichment (rodada 2, 20/out/2026): atualiza APENAS a linha da 'BU'
(Boston University, Questrom School of Business, Boston, MA) no
msba_programs_database.xlsx. Foco: MS in Business Analytics, presencial, STEM.

CONFIRMADO SEM ALTERACAO (páginas oficiais do MSBA, out/2026):
- Trilha de 12 meses, 2026-27: tuition US$86.716 (Fall 36.512 + Spring
  36.512 + Summer 13.692) e University Fees US$1.823.
- Turma de set/2025: 110 alunos (997 aplicações), 60% com bolsa, 84%
  internacionais, GMAT mediano 740, GPA mediano 3.4, TOEFL mediano 101.
- 96% colocados em até 6 meses (média de 3 anos); empregadores (McKinsey,
  Dell, BlackRock, Cigna, NBC Universal, Deloitte, Fidelity, AstraZeneca).
- Taxa US$125, isenta para quem participa de eventos de admissão
  selecionados; 3 video essays no Kira; sem cartas; GMAT/GRE opcional.
- Inglês: TOEFL 90 recomendado / IELTS 6.5 por banda / Duolingo 120
  recomendado / PTE aceito.
- Transporte 130 e Outros 382 já batiam com o orçamento oficial de 12 meses
  (transportation US$1.554, personal US$4.580).

CORRIGIDO/PREENCHIDO:
- Seguro Saúde: 4.054 -> 3.538 (Health Insurance 2026-27 no orçamento
  oficial da trilha de 12 meses).
- Ranking US News: Verificar -> 39 (Business Analytics, ranking de
  especialidade, tie; perfil da Questrom na US News).
- Deadline/Rounds: a página principal do MSBA agora lista 6 rounds para o
  Fall 2027 (Round 6: 20/abr/2027); a página de admissão ainda mostra 5.
- Utilities: Verificar -> 0. O aluguel de US$1.100 (rodada 1: quarto
  mobiliado em Allston/Brighton) já inclui utilities.
- Seguro Residencial 15, Celular 50, WiFi 60 (valores padrão das outras linhas).
- Custo Vida Total: Verificar -> 27.982, fórmula do dashboard: soma mensal
  2.037 x 12 meses + seguro 3.538.
- Destaques Currículo: acrescentado que a trilha de 9 meses custa ~US$60k
  de tuition (página do programa) e que 84% da turma é internacional.
- Principais Empregadores: acrescentado "remuneração total média no 1º ano
  US$98k+ (empregos nos EUA)" e "48% com 2+ ofertas na formatura".
  O Salário Médio (93.000, salário-base da rodada 1) foi mantido porque
  os US$98k são remuneração total, outra medida.

MANTIDO COMO Verificar: Ranking QS (a Questrom não aparece no top 21 da
tabela QS Business Analytics 2027), Taxa de Aceitação (a página só dá 997
aplicações e 110 matriculados), TA/RA (nenhuma assistantship divulgada).

Fontes (out/2026):
- bu.edu/questrom/graduate-programs/specialty-masters-programs/ms-in-business-analytics/ (+ /admissions/, /admissions/class-profile/)
- usnews.com/best-graduate-schools/top-business-schools/boston-university-01108

Uso: python scripts/etl_enrich_bu_round2.py (a partir da raiz).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = ("BU", "MA", "Boston")

UPDATES = {
    "Seguro Saúde (programa) USD": 3538,
    "Ranking US News": 39,
    "Deadline Aplicação": (
        "Round 1: 13 out 2026; Round 2: 17 nov 2026; Round 3: 7 jan 2027; Round 4: 16 fev "
        "2027; Round 5: 16 mar 2027; Round 6: 20 abr 2027"
    ),
    "Rounds Aplicação": (
        "6 rounds (a página de admissão ainda lista 5); revisão rolling, decisão em 6-8 "
        "semanas; bolsas de mérito automáticas e mais prováveis nos primeiros rounds; taxa de "
        "US$125 isenta para quem participa de eventos de admissão selecionados"
    ),
    "Seguro Residencial/mês USD": 15,
    "Utilities/mês USD": 0,
    "Celular/mês USD": 50,
    "WiFi/mês USD": 60,
    # Mesma fórmula de computeCustoVidaTotal (derived.ts): soma mensal 2.037 x 12 meses + seguro 3.538.
    "Custo Vida Total (programa) USD": 27982,
    "Destaques Currículo": (
        "Trilhas de 9 (~US$60k de tuition), 12 (com capstone de indústria, set–ago) ou 16 "
        "meses (com estágio de 3-6 meses elegível a OPT); concentrações: AI/Data & Methods, "
        "Healthcare Analytics (rara em MSBA), Marketing Analytics, eletivas customizadas; "
        "IA integrada aos cursos e projetos com clientes reais; ~100 alunos por coorte, 84% "
        "internacionais"
    ),
    "Principais Empregadores": (
        "McKinsey & Company, Dell, BlackRock, Cigna, NBC Universal, Deloitte, Fidelity, "
        "AstraZeneca, Amazon (lista oficial do MSBA); remuneração total média no 1º ano "
        "US$98k+ (empregos nos EUA); 48% com 2+ ofertas na formatura"
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
    print(f"Linha alvo confirmada: {row_idx} (BU, Boston, MA)")

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
