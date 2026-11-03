"""
Data enrichment (round 2, 2026-11-02): atualiza APENAS a linha da 'RIT'
(Saunders College of Business, Rochester, NY) -- MS in Business Analytics and AI.
Não confundir com 'U of Rochester' (Simon), que fica na mesma cidade.

Match por Nome + Estado + Cidade. Rodar da raiz do repositório:
    python scripts/etl_enrich_rit_round2.py [--dry-run]

FONTES OFICIAIS (2026-11-02):
- rit.edu/study/business-analytics-ms (requisitos, prazos, outcomes, inglês)
- rit.edu/careerservices/study/business-analytics-and-ai-ms (outcomes Class of 2025)
- rit.edu/admissions/graduate/estimate-education-costs (tuition e fees 2026-27)
- rit.edu/admissions/graduate/apply (taxa de aplicação)
- rit.edu/admissions/masterup (Master Up)
- go.gallagherstudent.com -- 2026-2027 RIT Student Health Insurance (Aetna)

CORRIGIDO:
- Deadline/Rounds: o Fall não é mais "prioridade 1 ago, rolling". A página atual
  traz 15 jan (prioridade) e 1 mar (padrão) para o Fall; Spring é rolling.
- Principais Empregadores: lista oficial atualizada (Class of 2025).
- Seguro Saúde: Verificar -> 3516 (plano Aetna 2026-27, student only, cobertura
  anual 15/ago/2026-31/jul/2027).

CONFIRMADO SEM MUDANÇA: tuition full-time é tarifa fixa de US$32.271 por
semestre (12-18 créditos); 30 créditos em 2 semestres = US$64.542 (o valor por
crédito de 2.151 na planilha é esse total dividido por 30; abaixo de 12
créditos cobra-se US$2.690/crédito). Fees: activity US$235 + health services
US$315 por semestre = US$1.100. Outcome rate 86,4% e salário médio US$84.000
(mediana US$67.000, Class of 2025, knowledge rate 73,3%). Taxa US$65 (isenta só
para alunos/ex-alunos da RIT). GMAT/GRE obrigatório para diploma estrangeiro.
Inglês: Duolingo 130, TOEFL 88/4.5, IELTS 6.5, PTE 60. GPA 3.0 recomendado.
Master Up (30%) exclui quem mora fora dos EUA e vale só para início até ago/2027.

MANTIDO COMO Verificar: Utilities e Custo Vida Total (sem fonte para utilities
em Rochester), Tamanho Turma, GPA/GMAT médios, Taxa Aceitação, rankings.
"""

import sys

import openpyxl

from db_retention import DB_PATH, rotate_backup

SHEET = "MSBA Programs"
MATCH = ("RIT", "NY", "Rochester")

UPDATES = {
    "Deadline Aplicação": (
        "Fall 2027: prioridade 15 jan 2027; padrão 1 mar 2027. Spring: rolling "
        "(bolsas de admissão por ordem de chegada das aplicações completas -- aplicar até a "
        "prioridade)"
    ),
    "Rounds Aplicação": (
        "Fall: 2 prazos (prioridade 15 jan, padrão 1 mar); Spring: rolling; bolsas de mérito "
        "decididas junto com a admissão, por ordem de chegada das aplicações completas"
    ),
    "Principais Empregadores": (
        "TikTok, Tata Consultancy Services, Ernst & Young (EY), Dow Jones & Co., CooperVision, "
        "Citibank, Calero-MDSL, DataCore Software, The Bonadio Group, Excellus BCBS, Mindex "
        "Technologies"
    ),
    "Seguro Saúde (programa) USD": 3516,
    "Taxa Aplicação USD": 65,
}


def main() -> None:
    dry = "--dry-run" in sys.argv
    if not dry:
        rotate_backup()
    wb = openpyxl.load_workbook(DB_PATH)
    ws = wb[SHEET]
    header = [c.value for c in ws[1]]
    col = {name: i + 1 for i, name in enumerate(header)}
    missing = [k for k in UPDATES if k not in col]
    if missing:
        raise SystemExit(f"colunas inexistentes no schema: {missing}")

    nome, estado, cidade = MATCH
    rows = [
        r for r in range(2, ws.max_row + 1)
        if ws.cell(r, col["Nome da Universidade"]).value == nome
        and ws.cell(r, col["Estado"]).value == estado
        and ws.cell(r, col["Cidade"]).value == cidade
    ]
    if len(rows) != 1:
        raise SystemExit(f"esperava 1 linha para {MATCH}, achei {len(rows)}")
    r = rows[0]

    changed = 0
    for field, new in UPDATES.items():
        cell = ws.cell(r, col[field])
        if cell.value == new:
            continue
        changed += 1
        print(f"== {field}\n   ANTES : {cell.value}\n   DEPOIS: {new}\n")
        cell.value = new
    print(f"{changed} campo(s) alterado(s) na linha {r} ({nome}).")

    if dry:
        print("dry-run: nada gravado")
        return
    wb.save(DB_PATH)
    print(f"gravado em {DB_PATH}")


if __name__ == "__main__":
    main()
