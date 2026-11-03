"""
Data enrichment (round 3, 2026-11-02): atualiza APENAS a linha da 'UW-Madison'
(Wisconsin School of Business, Madison, WI) -- MS in Business Analytics.
Não confundir com 'Dakota State' (Madison, SD).

Match por Nome + Estado + Cidade. Rodar da raiz do repositório:
    python scripts/etl_enrich_uw_madison_round3.py [--dry-run]

RECONFERIDO SEM MUDANÇA (business.wisc.edu, 2026-11-02): tuition US$1.500/crédito
x 30 = US$45.000 (mesma para residente e não residente), segregated fees
US$1.699,48; bolsa automática (74% da Class of 2026, média US$9.000); outcomes
Class of 2025 (89% com oferta em 6 meses, média US$75.629, empregadores);
Class of 2027 (43 alunos, GPA 3.33, GRE V158/Q167); taxa US$75 (+US$6 intl),
isenção só para militares; prazos 4 nov/13 jan/3 mar/29 abr (intl)/3 jun;
Duolingo 125, IELTS 7.0, TOEFL 92 ou 5.0; GMAT/GRE exigido com dispensa
seletiva; 3 essays + 3 video essays; carta opcional; sem entrevista; 10 meses
(bootcamp online + 2 semestres presenciais); QS 2025 #7 entre públicos.

PREENCHIDO:
- Processo TA/RA: studentjobs.wisc.edu/uw-graduate-assistantships/ e
  grad.wisc.edu/funding/graduate-assistantships/ -- a maioria das vagas é
  "owned" por programas e não é publicada; parte fica aberta a qualquer aluno
  de pós no Student Job Center, publicada pouco antes do início (fall: jul-ago);
  dá para conseguir depois do início das aulas.
- Graduate Assistantships: esclarecido que a regra geral da UW (nomeação de
  33,3%+ = tuition remission) NÃO se aplica ao MSBA -- a página de tuition da
  WSB diz que a nomeação "will not be eligible for tuition remission"; ficam
  stipend mensal e seguro saúde.

MANTIDO COMO Verificar (a escola não publica): GMAT Médio, Taxa Aceitação,
Ranking US News (o US News não ranqueia MSBA presencial).
"""

import sys

import openpyxl

from db_retention import DB_PATH, rotate_backup

SHEET = "MSBA Programs"
MATCH = ("UW-Madison", "WI", "Madison")

UPDATES = {
    "Processo TA/RA": (
        "Sem processo próprio da WSB: a maioria das assistantships da UW pertence a "
        "programas específicos e não é publicada (é preciso contatar departamentos, "
        "coordenadores de pós e professores, de preferência logo após a admissão); uma "
        "parte menor fica aberta a qualquer aluno de pós no Student Job Center "
        "(studentjobs.wisc.edu), publicada pouco antes do início (vagas de fall em "
        "jul-ago). Muitos alunos de mestrado conseguem TA em outros departamentos, às "
        "vezes já com as aulas em andamento"
    ),
    "Graduate Assistantships": (
        "Disponíveis via Student Job Center ou outros departamentos, mas no MSBA a "
        "nomeação NÃO dá tuition remission (a regra geral da UW de isenção com 33,3%+ não "
        "vale para este programa); dá stipend mensal e seguro saúde proporcionais à "
        "nomeação"
    ),
}


def main() -> None:
    dry = "--dry-run" in sys.argv
    if not dry:
        rotate_backup()
    wb = openpyxl.load_workbook(DB_PATH)
    ws = wb[SHEET]
    header = [c.value for c in ws[1]]
    col = {name: i + 1 for i, name in enumerate(header)}

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
