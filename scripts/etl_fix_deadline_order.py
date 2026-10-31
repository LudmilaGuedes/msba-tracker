# -*- coding: utf-8 -*-
"""
Correcao em lote: reordena o campo "Deadline Aplicação" para que o texto
sempre COMECE pelo prazo mais cedo da entrada principal do programa.

Regra definida em set/2026: na classificacao por rounds, cada
faculdade deve ser considerada pela PRIMEIRA data de aplicacao possivel
(Round 1, Early Action, prioridade de bolsa...), porque aplicar cedo da
mais chance de admissao e de bolsa. A aba "Estrategia de Aplicacao" do
dashboard classifica cada programa pelo primeiro mes que aparece nesse
texto (strategyParsing.ts -> firstDeadlineMonth), entao a ordem importa.

Problemas encontrados na auditoria das 101 linhas:
- texto comecando pelo mes de inicio do curso ("Fall 2027 (agosto): ...");
- texto comecando pela data de abertura das inscricoes, que nao e prazo;
- prazo domestico listado antes de um prazo internacional mais cedo;
- prazo de Spring listado antes do prazo de Fall, em programas cuja
  entrada principal e Fall;
- prazo de bolsa (mais cedo) listado depois do prazo geral.

Nenhuma data e alterada, so a ordem e a redacao. Cada linha so e
atualizada se o valor atual for exatamente o esperado, para nao sobrescrever
uma edicao feita depois desta auditoria.

Uso: python scripts/etl_fix_deadline_order.py (a partir da raiz).
"""

import sys

import openpyxl

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"
COLUMN = "Deadline Aplicação"

# (Nome, Estado, Cidade) -> (valor atual esperado, novo valor)
FIXES = {
    ("ASU", "AZ", "Tempe"): (
        "Fall 2027 (agosto): Round I 1/dez/2026; Round II intl prioridade 1/fev/2027; Round III intl final 1/abr/2027; Round IV final 1/jul/2027. Spring 2027 (janeiro): 1/jul, 1/set (intl prioridade), 1/nov (intl final), 1/dez/2026",
        "Fall 2027: Round 1: 1/dez/2026 (Early Decision); Round 2: 1/fev/2027 (prioridade intl); Round 3: 1/abr/2027 (final intl); Round 4: 1/jul/2027 (final geral); Spring 2027: 1/jul/2026, 1/set/2026 (prioridade intl), 1/nov/2026 (final intl), 1/dez/2026",
    ),
    ("UAH", "AL", "Huntsville"): (
        "Doméstico: Fall 1º jun, Spring 30 nov, Summer 15 mar; Internacional (I-20): Fall 1º ago, Spring 1º dez. Prioridade para GA: 31/mar (Fall/Summer) e 31/out (Spring); bolsas competitivas: 15/dez",
        "Fall: bolsas competitivas da UAH 15/dez (ano anterior); prioridade para GA 31/mar; doméstico 1º jun; internacional (I-20) 1º ago. Spring: prioridade para GA 31/out; doméstico 30 nov; internacional 1º dez. Summer: doméstico 15 mar; prioridade para GA 31/mar",
    ),
    ("Stony Brook SUNY", "NY", "Long Island"): (
        "Fall 2026: Doméstico 15 jul, Internacional 1º mai; Spring 2026: Doméstico 15 jan, Internacional 1º dez; Summer 2026: Doméstico 15 mar",
        "Fall 2026: Internacional 1º mai, Doméstico 15 jul; Spring 2026: Internacional 1º dez, Doméstico 15 jan; Summer 2026: Doméstico 15 mar",
    ),
    ("U of Iowa", "IA", "Iowa City"): (
        "Aplicações abrem 15 ago; prioridade (bolsas de mérito) 15 dez; final internacional 15 mar; final doméstico 15 jun; entrada apenas Fall",
        "Prioridade (bolsas de mérito): 15 dez; final internacional 15 mar; final doméstico 15 jun; aplicações abrem 15 ago; entrada apenas Fall",
    ),
    ("Dakota State", "SD", "Madison"): (
        "Fall: 15 ago (doméstico) / 15 abr (intl); Spring: 30 nov (doméstico) / 1 nov (intl); Summer: 30 abr (só doméstico)",
        "Fall: 15 abr (intl) / 15 ago (doméstico); Spring: 1 nov (intl) / 30 nov (doméstico); Summer: 30 abr (só doméstico)",
    ),
    ("Cal Poly SLO", "CA", "San Luis Obispo"): (
        "Ciclo abre 1 out 2026; internacional 1 abr 2027; doméstico 1 jun 2027",
        "Internacional: 1 abr 2027; doméstico: 1 jun 2027; ciclo de aplicação abre 1 out 2026",
    ),
    ("Oklahoma State", "OK", "Stillwater"): (
        "Somente admissão Fall; inscrições abrem 1/out do ano anterior; prioritário 15/jan; final 1/mar; comitê começa avaliação em meados de janeiro",
        "Prioritário: 15/jan; final: 1/mar; inscrições abrem 1/out do ano anterior; comitê começa avaliação em meados de janeiro; somente admissão Fall",
    ),
    ("Florida Atlantic University (FAU)", "FL", "Boca Raton"): (
        "Fall: dom 1 jul / intl 15 fev; Spring: dom 1 nov / intl 15 jul; Summer: dom 1 abr / intl 15 jan",
        "Fall: intl 15 fev / dom 1 jul; Spring: intl 15 jul / dom 1 nov; Summer: intl 15 jan / dom 1 abr",
    ),
    ("University of North Texas (UNT)", "TX", "Denton"): (
        "Summer: 15 jun; Fall: 1 ago; Spring: 1 jan (candidatos F-1 devem se inscrever com 2 meses de antecedência para processar o I-20)",
        "Fall: 1 ago; Spring: 1 jan; Summer: 15 jun (candidatos F-1 devem se inscrever com 2 meses de antecedência para processar o I-20)",
    ),
    ("Rutgers University–Camden", "NJ", "Camden"): (
        "Internacionais: Fall (set) 1 jun; Spring (jan) 1 out. Domésticos: Fall 1 ago; Spring 1 dez; Summer 1 mai",
        "Fall: internacional 1 jun, doméstico 1 ago; Spring: internacional 1 out, doméstico 1 dez; Summer: doméstico 1 mai",
    ),
    ("George Mason University", "VA", "Fairfax"): (
        "Fall (início em agosto): 1º de maio",
        "Fall: 1º de maio (aulas começam em agosto)",
    ),
    ("Wayne State University", "MI", "Detroit"): (
        "Deadlines gerais da Graduate School — Fall: doméstico 1 jun / internacional 1 mai (programa pode ter prazo próprio, não encontrado especificamente para o MSDSBA)",
        "Deadlines gerais da Graduate School — Fall: internacional 1 mai / doméstico 1 jun (programa pode ter prazo próprio, não encontrado especificamente para o MSDSBA)",
    ),
    ("American University", "DC", "Washington"): (
        "Rolling 1 jul; rounds 15 nov / 15 jan / 15 mar / 15 maio",
        "Round 1: 15 nov; Round 2: 15 jan; Round 3: 15 mar; Round 4: 15 mai; rolling até 1 jul",
    ),
    ("West Texas A&M University (WTAMU)", "TX", "Canyon"): (
        "Admissão geral: Fall 1/ago; Spring 1/dez; Summer 1/mai. Prioridade de bolsas: 1/fev (para o próximo ano acadêmico)",
        "Prioridade de bolsas: 1/fev (para o próximo ano acadêmico); admissão geral: Fall 1/ago; Spring 1/dez; Summer 1/mai",
    ),
    ("Texas A&M", "TX", "Houston"): (
        "Abertura: 1 set; Round 1: 3 dez; Round 2 (consideração para bolsa): 1 abr; Final: 25 jun",
        "Round 1: 3 dez; Round 2: 1 abr (consideração para bolsa); Final: 25 jun; aplicações abrem 1 set",
    ),
    ("UW-Madison", "WI", "Madison"): (
        "Ciclo jul/2027: 4 nov 2026; 13 jan 2027; 3 mar 2027; 29 abr 2027 (final internacionais); 3 jun 2027 (final domésticos)",
        "Round 1: 4 nov 2026; Round 2: 13 jan 2027; Round 3: 3 mar 2027; Round 4: 29 abr 2027 (final internacionais); Round 5: 3 jun 2027 (final domésticos) -- turma com início em julho/2027",
    ),
    ("U of Houston", "TX", "Houston"): (
        "Intl: 1 out (Spring)/1 mai (Fall); Doméstico: 1 dez (Spring)/1 jul (Fall) — admissão contínua, aplicar cedo ajuda na bolsa",
        "Fall: intl 1 mai / doméstico 1 jul; Spring: intl 1 out / doméstico 1 dez — admissão contínua, aplicar cedo ajuda na bolsa",
    ),
}


def main():
    rotate_backup()
    print("Backup rotacionado (política fixa) em: data/msba_programs_database_backup.xlsx")

    wb = openpyxl.load_workbook(DB_PATH)
    ws = wb[SHEET_DATA]
    headers = [cell.value for cell in ws[1]]
    col = {h: i + 1 for i, h in enumerate(headers)}

    rows_by_key = {}
    for row in ws.iter_rows(min_row=2):
        key = (
            row[col["Nome da Universidade"] - 1].value,
            row[col["Estado"] - 1].value,
            row[col["Cidade"] - 1].value,
        )
        rows_by_key.setdefault(key, []).append(row[0].row)

    changed = 0
    for key, (expected, new_value) in FIXES.items():
        matches = rows_by_key.get(key, [])
        if len(matches) != 1:
            raise ValueError(f"{key}: esperava 1 linha, encontrou {len(matches)}. Abortando.")
        cell = ws.cell(row=matches[0], column=col[COLUMN])
        if cell.value != expected:
            print(f"\n[PULADO] {key[0]} (linha {matches[0]}): valor atual difere do esperado.\n  Atual: {cell.value!r}")
            continue
        cell.value = new_value
        changed += 1
        print(f"\n- {key[0]} (linha {matches[0]}):")
        print(f"    Antes:  {expected}")
        print(f"    Depois: {new_value}")

    wb.save(DB_PATH)
    print(f"\n{changed} de {len(FIXES)} linhas atualizadas na coluna '{COLUMN}'.")


if __name__ == "__main__":
    main()
