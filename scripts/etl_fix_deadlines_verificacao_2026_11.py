"""
Correção de 'Deadline Aplicação' / 'Rounds Aplicação' após a verificação
geral dos prazos de 2026-11-01 (conferência contra as páginas oficiais).

Só altera as linhas em que a data no banco estava comprovadamente errada
ou desatualizada. Match por Nome + Estado + Cidade (nunca por número de linha).

Rodar da raiz do repositório:
    python scripts/etl_fix_deadlines_verificacao_2026_11.py [--dry-run]
"""

import sys

import openpyxl

from db_retention import DB_PATH, rotate_backup

SHEET = "MSBA Programs"

FIXES = [
    {
        # rady.ucsd.edu/programs/ms-business-analytics/admissions/ (Summer 2027)
        "match": ("UC San Diego", "CA", "La Jolla"),
        "Deadline Aplicação": (
            "Round 1: 1 out 2026; Round 2: 15 nov 2026; Round 3: 15 jan 2027; "
            "Round 4: 15 mar 2027 (deadline de fellowship e internacional); "
            "Round 5: 15 mai 2027 (só domésticos)"
        ),
        "Rounds Aplicação": (
            "5 rounds (ingresso Summer 2027, único início do ano); Round 4 (15 mar 2027) "
            "é o último para internacionais e para fellowships; revisão rolling, aplicar "
            "cedo aumenta chance de vaga; taxa US$150"
        ),
    },
    {
        # infosystems.utdallas.edu FAQ "What are the application deadlines?" (mesmas datas p/ intl)
        "match": ("UT Dallas", "TX", "Richardson"),
        "Deadline Aplicação": (
            "Fall: prioritário 15 jan, regular 1 mai | Spring: prioritário 15 mai, "
            "regular 1 out | Summer: prioritário 15 out, regular 1 mar "
            "(mesmas datas para internacionais)"
        ),
        "Rounds Aplicação": (
            "Sem rounds competitivos -- admissão por semestre (Fall/Spring/Summer) com "
            "prazo prioritário + regular por termo; aplicar até o prioritário (15 jan para o Fall)"
        ),
    },
    {
        # mays.tamu.edu/graduate/masters/ms-programs/analytics/apply-to-ms-analytics
        "match": ("Texas A&M", "TX", "Houston"),
        "Deadline Aplicação": (
            "Round 1: 7 dez 2026; Round 2: 5 abr 2027 (consideração para bolsa); "
            "Round 3: 28 jun 2027 (final) -- ciclo Fall 2027; aplicações abrem em 1 set 2026"
        ),
        "Rounds Aplicação": (
            "3 rounds (7 dez, 5 abr com consideração para bolsa, 28 jun); taxa de US$114 "
            "para internacionais (US$89 domésticos)"
        ),
    },
    {
        # Página oficial bloqueada (403); datas do ciclo Fall 2027 via ApplyBuds
        "match": ("Johns Hopkins", "DC", "Washington"),
        "Deadline Aplicação": (
            "Round 1: 8 out 2026; Round 2: 14 jan 2027 (ciclo Fall 2027, fonte não "
            "oficial; Rounds 3-4 devem cair em ~mar e ~mai -- internacionais devem "
            "aplicar até o Round 2)"
        ),
        "Rounds Aplicação": (
            "4 rounds (padrão dos ciclos anteriores; Fall 2026 foi 8 out / 14 jan / 11 mar / "
            "20 mai); internacionais recomendados a aplicar até o Round 2 (visto); isenção "
            "de taxa só caso a caso, com aplicação completa"
        ),
    },
    {
        # scu.edu/business/ms -- 1 abr prioridade bolsa/internacional, 1 mai final
        "match": ("Santa Clara University", "CA", "Santa Clara"),
        "Deadline Aplicação": (
            "Fall: prioritário 1 abr (recomendado para internacionais/I-20 e bolsas); "
            "final 1 mai; Spring: prioritário 20 jan; aplicações abrem ~9 meses antes do prazo"
        ),
    },
    {
        # case.edu/weatherhead/.../master-business-analytics-and-intelligence/admission-information
        "match": ("Case Western Reserve University", "OH", "Cleveland"),
        "Deadline Aplicação": (
            "Early Action: 11 nov 2026; Early Decision: 20 jan 2027; Priority "
            "Scholarship: 31 mar 2027; Regular: 30 abr 2027; Final: 2 jun 2027; "
            "Extended: 30 jul 2027 (se houver vaga) -- ciclo Fall 2027"
        ),
        "Rounds Aplicação": (
            "6 prazos para o Fall 2027 (Early Action, Early Decision, Priority Scholarship, "
            "Regular, Final, Extended) + rolling no verão conforme vagas; bolsas de mérito "
            "favorecem quem aplica cedo; página não publica prazo internacional separado"
        ),
    },
    {
        # wsou.shu.edu/business/graduate-applying.html -- prazos internacionais separados
        "match": ("Seton Hall University", "NJ", "South Orange"),
        "Deadline Aplicação": (
            "Internacional: Fall 30 abr; Spring 30 set (sem Summer). EUA: Fall 31 mai "
            "(estendido para 15 jul); Spring 31 out; Summer 31 mar (admissão rolling; "
            "priority deadline vale para bolsas de mérito e GA)"
        ),
        "Rounds Aplicação": (
            "Rolling, com priority deadline por semestre -- internacional: Fall 30 abr, "
            "Spring 30 set; EUA: Fall 31 mai, Spring 31 out, Summer 31 mar; taxa US$75 "
            "(não reconfirmada)"
        ),
    },
    {
        # isenberg.umass.edu/content/business-analytics-fall-start (página oficial, sem ano)
        # Decisão da usuária em 2026-11-01: seguir a página oficial.
        "match": ("UMass Amherst", "MA", "Amherst"),
        "Deadline Aplicação": (
            "Fall: Early Action 1 fev; Regular 1 abr; Extended 15 jun (só domésticos e "
            "internacionais já nos EUA com outro visto). Spring: não confirmado na página oficial"
        ),
        "Rounds Aplicação": (
            "3 prazos para o Fall (Early Action 1 fev, Regular 1 abr, Extended 15 jun -- este "
            "último não vale para quem precisa de novo F-1); página oficial não informa o ano"
        ),
    },
    {
        # fau.edu/graduate/graduate-programs/ms-business-analytics/ -- prazo único, sem separar intl
        # Decisão da usuária em 2026-11-01: seguir a página oficial.
        "match": ("Florida Atlantic University (FAU)", "FL", "Boca Raton"),
        "Deadline Aplicação": (
            "Fall: 1 jul; Spring: 1 nov; Summer: 1 abr (prazo único da página do MSBA, "
            "sem prazo internacional separado -- internacionais devem aplicar bem antes "
            "por causa do I-20)"
        ),
        "Rounds Aplicação": (
            "Sem rounds e não é rolling: um prazo fixo por semestre (Fall 1 jul, Spring 1 nov, "
            "Summer 1 abr), com a aplicação e todos os documentos completos até a data"
        ),
    },
]


def main() -> None:
    dry = "--dry-run" in sys.argv
    if not dry:
        rotate_backup()
    wb = openpyxl.load_workbook(DB_PATH)
    ws = wb[SHEET]
    header = [c.value for c in ws[1]]
    col = {name: i + 1 for i, name in enumerate(header)}

    for fix in FIXES:
        nome, estado, cidade = fix["match"]
        rows = [
            r for r in range(2, ws.max_row + 1)
            if ws.cell(r, col["Nome da Universidade"]).value == nome
            and ws.cell(r, col["Estado"]).value == estado
            and ws.cell(r, col["Cidade"]).value == cidade
        ]
        if len(rows) != 1:
            raise SystemExit(f"{nome}: esperava 1 linha, achei {len(rows)}")
        r = rows[0]
        for field, value in fix.items():
            if field == "match":
                continue
            old = ws.cell(r, col[field]).value
            print(f"[{nome}] {field}\n  ANTES: {old}\n  DEPOIS: {value}\n")
            ws.cell(r, col[field]).value = value

    if dry:
        print("dry-run: nada gravado")
        return
    wb.save(DB_PATH)
    print(f"gravado em {DB_PATH}")


if __name__ == "__main__":
    main()
