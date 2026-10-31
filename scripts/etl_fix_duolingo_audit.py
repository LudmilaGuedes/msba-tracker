# -*- coding: utf-8 -*-
"""
Auditoria em lote do campo "Testes Inglês Aceitos": confirma, nas fontes
oficiais, quais programas aceitam ou nao o Duolingo English Test (DET).

Pedido de set/2026: "confirme todas que nao aceitam o Duolingo". Foram
auditadas as 19 linhas que diziam "nao aceita" (3) ou que nao
mencionavam o Duolingo (15), mais a unica sem dado (SCSU). A Michigan ficou
de fora porque ja cita "DET recomendado 130+", ou seja, aceita.

U of Michigan: aceita Duolingo (michiganross.umich.edu, recomendado 130),
mas o texto dizia so "DET", e o dashboard (parseDuolingo) procura a
palavra "Duolingo" e marcava a Michigan como "nao aceita". O texto passou
a dizer "Duolingo (DET)".

Nao alteradas, pois ja estavam corretas e confirmadas: Georgia Tech
(esl.gatech.edu / grad.gatech.edu), UGA (grad.uga.edu) e U of Maryland
(gradschool.umd.edu). As tres NAO aceitam.

NAO ACEITAM (confirmado):
- Virginia Tech e VT Global BA: graduateschool.vt.edu/.../testing-requirements
  lista so TOEFL (4.5, ou 90 com >=20 por secao na escala antiga) e IELTS
  6.5. O Duolingo vale so para a graduacao.
- Georgetown: msb.georgetown.edu/msba/admissions-financing/application-process
  lista so TOEFL (100 ou 5) e IELTS (7.5).
- UC Berkeley: grad.berkeley.edu diz "We do not accept the Duolingo exam".
- UC Irvine: grad.uci.edu diz "approved English proficiency test (TOEFL or
  IELTS)".
- U of Arizona: grad.arizona.edu diz "The Graduate College does not accept
  Duolingo scores".
- William & Mary: mason.wm.edu/graduate/msba/residential/admissions lista
  so TOEFL (100) e IELTS (7).
- Wake Forest: business.wfu.edu/masters-in-business-analytics/application-process
  lista TOEFL, TOEFL Essentials e IELTS. Os minimos recomendados sao TOEFL
  5.0 e IELTS 7.5; a base tinha TOEFL 80 / IELTS 6.5, que foi corrigido. A
  Graduate School de artes e ciencias da Wake aceita Duolingo, mas a regra
  nao vale para o MSBA da School of Business.

ACEITAM (confirmado, minimo adicionado):
- GWU: Duolingo 115 (GW School of Business, via snippet; business.gwu.edu
  bloqueia fetch).
- Dakota State: dsu.edu/admissions/international/graduate, Duolingo 110.
- SMU: smu.edu/cox/ms-degrees/faqs; TOEFL, IELTS e Duolingo sao aceitos e
  opcionais.
- Illinois Tech: iit.edu, Duolingo 115.
- UNT: unt.edu/admissions/international/english-language-requirements,
  Duolingo 100 (pagina geral de admissao internacional).
- Oakland: oakland.edu/grad/how-to-apply/international-students, Duolingo
  aceito, sem minimo publicado na pagina.
- Florida Tech: fit.edu (demonstrating-english-proficiency), Duolingo 105.
- SCSU: catalog.southernct.edu (graduate), Duolingo 105.

Cada linha so e atualizada se o valor atual for exatamente o esperado.

Uso: python scripts/etl_fix_duolingo_audit.py (a partir da raiz).
"""

import sys

import openpyxl

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"
COLUMN = "Testes Inglês Aceitos"

VT_GRAD = (
    "TOEFL iBT 4.5 (nova escala) ou 90 com ≥20 em cada seção (escala antiga, antes de "
    "21/jan/2026) / IELTS 6.5 — padrão da Graduate School; Duolingo NÃO aceito (a "
    "Graduate School só aceita TOEFL/IELTS; Duolingo vale só para a graduação); "
    "dispensado se graduação em instituição de ensino em inglês"
)

# (Nome, Estado, Cidade) -> (valor atual esperado, novo valor)
FIXES = {
    ("Virginia Tech", "VA", "Blacksburg"): (
        "TOEFL (mín. 80, nenhuma sub-nota <16) ou IELTS (mín. 6.5, nenhuma sub-nota <6.5) — padrão institucional (Graduate School); dispensado se graduação em instituição de ensino em inglês",
        VT_GRAD,
    ),
    ("VT Global BA", "VA", "Blacksburg"): ("TOEFL/IELTS", VT_GRAD),
    ("Georgetown", "DC", "Washington"): (
        "TOEFL/IELTS",
        "TOEFL iBT 100 (escala antiga) ou 5 (nova escala) / IELTS Academic 7.5; Duolingo e "
        "PTE NÃO aceitos (a página oficial do MSBA lista só TOEFL e IELTS); dispensa para "
        "diploma dos EUA ou de universidade com ensino em inglês",
    ),
    ("UC Berkeley", "CA", "Berkeley"): (
        "TOEFL mín. 90 (iBT) / IELTS mín. 7.0 — exigido para não-cidadãos/não-residentes permanentes dos EUA",
        "TOEFL 90 (testes até 20/jan/2026) ou 5 na nova escala (mín. 4.5 em speaking e "
        "writing) / IELTS 7.0 (a partir de 21/jan/2026, mín. 6.5 em speaking e writing); "
        "Duolingo NÃO aceito (Graduate Division: 'We do not accept the Duolingo exam') — "
        "exigido para não-cidadãos/não-residentes permanentes dos EUA",
    ),
    ("UC Irvine", "CA", "Irvine"): (
        "TOEFL/IELTS exigido apenas para candidatos internacionais (mínimos específicos não divulgados publicamente)",
        "TOEFL ou IELTS, exigido apenas para candidatos internacionais (mínimos específicos "
        "não divulgados publicamente); Duolingo NÃO aceito (a Graduate Division da UCI só "
        "aceita TOEFL ou IELTS)",
    ),
    ("U of Arizona", "AZ", "Tucson"): (
        "TOEFL 85 iBT (antes de 21/jan/26) ou 4.5 (novo formato, depois); IELTS composto 7 (nenhuma sub-área <6); PTE não confirmado nesta fonte",
        "TOEFL 85 iBT (antes de 21/jan/26) ou 4.5 (novo formato, depois); IELTS composto 7 "
        "(nenhuma sub-área <6); PTE 60 (mínimo da Graduate College); Duolingo NÃO aceito "
        "('The Graduate College does not accept Duolingo scores')",
    ),
    ("William & Mary (Raymond A. Mason School of Business)", "VA", "Williamsburg"): (
        "TOEFL mín. 100 (iBT) / IELTS mín. 7.0 — dispensa para quem tem 4+ anos de graduação em país/instituição de língua inglesa (código TOEFL: B308)",
        "TOEFL mín. 100 (iBT) / IELTS mín. 7.0; Duolingo e PTE NÃO aceitos (a página oficial "
        "do MSBA lista só TOEFL e IELTS); dispensa para quem tem 4+ anos de graduação em "
        "país/instituição de língua inglesa (código TOEFL: B308); scores válidos por 2 anos",
    ),
    ("Wake Forest University", "NC", "Winston-Salem"): (
        "TOEFL iBT 80 (antes de 21/jan/2026) ou 4.5 (a partir de 21/jan/2026, conforme fonte oficial); IELTS 6.5; dispensa possível para quem residiu/trabalhou nos EUA por período extenso",
        "TOEFL 5.0 (após 21/jan/2026) ou IELTS 7.5 recomendados (diretrizes, não cortes "
        "rígidos); TOEFL Essentials também aceito; Duolingo NÃO aceito pelo MSBA (a página "
        "oficial lista só TOEFL/IELTS); dispensa possível para quem residiu/trabalhou nos "
        "EUA por período extenso",
    ),
    ("GWU", "DC", "Washington"): (
        "TOEFL/IELTS",
        "TOEFL, IELTS, PTE Academic ou Duolingo; Duolingo mín. 115 (GW School of Business); "
        "scores de até 2 anos",
    ),
    ("Dakota State", "SD", "Madison"): (
        "TOEFL iBT 76 / TOEFL PBT 550 / IELTS 6.0 (PTE também aceito)",
        "TOEFL iBT 76 (antes de 21/jan/2026) ou 4.0 na nova escala (nenhuma sub-nota <4.0) / "
        "TOEFL PBT 550 / IELTS 6.0 (nenhuma sub-nota <5.5) / PTE 60 / Duolingo 110 (nenhuma "
        "sub-nota integrada <105) — válidos por 2 anos",
    ),
    ("SMU", "TX", "Dallas"): (
        "TOEFL/IELTS (opcional, não obrigatório para internacionais; validade de 2 anos)",
        "TOEFL, IELTS ou Duolingo aceitos, mas OPCIONAIS (não obrigatório para "
        "internacionais; sem nota mínima publicada); validade de 2 anos",
    ),
    ("Illinois Tech", "IL", "Chicago"): (
        "TOEFL/IELTS",
        "TOEFL iBT 80 / IELTS 6.5 / Duolingo 115 (mínimos da Graduate Admission da Illinois Tech)",
    ),
    ("University of North Texas (UNT)", "TX", "Denton"): (
        "TOEFL 79 (iBT) / IELTS 6.5",
        "TOEFL 79 (iBT, antes de 21/jan/2026) ou 4 (nova escala) / IELTS 6.5 (a página geral "
        "da UNT cita 6.0) / Duolingo 100 / PTE 53 — mínimos da página geral de admissão "
        "internacional; confirmar se o programa exige mais",
    ),
    ("Oakland University", "MI", "Rochester (MI)"): (
        "TOEFL exigido para não-nativos, salvo diploma de bacharelado dos EUA ou 24+ créditos cursados em instituição dos EUA",
        "TOEFL, IELTS, IELTS Indicator, MELAB ou Duolingo aceitos (Graduate School; mínimos "
        "não publicados na página) — exigido para não-nativos, salvo diploma de bacharelado "
        "dos EUA ou 24+ créditos cursados em instituição dos EUA",
    ),
    ("Florida Institute of Technology (Florida Tech)", "FL", "Melbourne"): (
        "TOEFL iBT 79-80 (testes antes de 1/jan/2026) ou 4.0 (a partir de 1/jan/2026, conforme fontes oficiais-adjacentes); IELTS 6.5",
        "TOEFL iBT 80 (antes de 1/jan/2026) ou 4.0 (a partir de 1/jan/2026) / TOEFL "
        "Essentials 9 / IELTS 6.5 / PTE 58 / Cambridge Advanced 180 / Duolingo 105",
    ),
    ("U of Michigan", "MI", "Ann Arbor"): (
        "TOEFL recomendado 100+ / IELTS recomendado 7.0+ / DET recomendado 130+ — sem nota mínima obrigatória; dispensado com graduação integralmente em instituição de língua inglesa",
        "TOEFL recomendado 100+ / IELTS recomendado 7.0+ / Duolingo (DET) recomendado 130+ — "
        "sem nota mínima obrigatória; teste feito até 2 anos antes do deadline; dispensado "
        "com graduação integralmente em instituição de língua inglesa",
    ),
    ("Southern Connecticut State University (SCSU)", "CT", "New Haven"): (
        "Verificar",
        "TOEFL iBT 79 (antes de 2026) ou 4.5 (a partir de 1/jan/2026) / IELTS 6.5 / "
        "Duolingo 105 / PTE 53 (regra geral de pós-graduação da SCSU)",
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
