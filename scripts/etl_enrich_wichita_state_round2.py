# -*- coding: utf-8 -*-
"""
Data enrichment (rodada 2, 20/out/2026): atualiza APENAS a linha da
'Wichita State' (W. Frank Barton School of Business, Wichita, KS) no
msba_programs_database.xlsx. Foco: MS in Business Analytics, presencial, STEM.

CONFIRMADO SEM ALTERACAO: 30 créditos (+ até 15 de cursos preparatórios se
o histórico pedir); duas trilhas (Management e Data Science); GMAT/GRE não
exigido; sem cartas; personal statement + resume; inglês para analytics:
TOEFL 79 / IELTS 6.5 / PTE 58 / Duolingo 115; sem admissão no Summer;
bolsas Barton (1 fev) e fellowships de US$750-8.500 (1 jul); Global Select
automática para GPA ≥ 3.30, não combinável com GA.

CORRIGIDO/PREENCHIDO:
- Formato: o FAQ do MSBA diz "offered in person"; o catálogo fala em mix de
  aulas de 16 semanas e híbridas de 8 semanas.
- Duração: "18-24" (texto, o dashboard não calculava o custo de vida) -> 16.
  O FAQ sugere até 12 créditos por semestre para quem não trabalha: 30
  créditos = 3 semestres (Fall/Spring/Fall, ~16 meses), sem cursos
  preparatórios.
- Tuition (Comprehensive Fee Schedule, Fall 2026): base não-residente
  US$903,59/crédito (era 1.001,92) e residente US$367,92 (era 466,25).
  Totais: US$27.108 (intl) e US$11.038 (residente). Com a Global Select a
  tarifa cai para US$551,89/crédito (US$16.557 no programa, ~38% de
  desconto) -- registrado em Tipos de Bolsa, não na tuition cheia.
- Fees Adicionais: Verificar -> 4.930. Barton School Course Fee US$76/
  crédito x 30 = 2.280; Campus Infrastructure US$20,58 + Technology US$1 +
  Transportation US$0,75 por crédito x 30 = 670; Student Support Services
  Fee US$742,35 em 2 semestres com 9+ créditos e US$494,90 no último (6
  créditos) = 1.980.
- Seguro Saúde: Verificar -> 3.663. Plano obrigatório para internacionais
  (UnitedHealthcare, 2026-27): US$2.747/ano, proporcional a 16 meses.
- Custo de vida: Aluguel 725 e Mercado 300 mantidos (rodada 1); Transporte
  153 e Outros 270 (custo de attendance oficial da WSU 2026-27 para a pós:
  US$1.380 e US$2.430 por 9 meses); Utilities 206 (RentCafe, Wichita);
  Seguro Residencial 15, Celular 50, WiFi 60 (valores padrão das outras linhas).
- Custo Vida Total: Verificar -> 32.127, fórmula do dashboard: soma mensal
  1.779 x 16 meses + seguro 3.663.
- Taxa/Rounds/Deadline: taxa internacional em 3 faixas -- US$75 até o prazo
  prioritário (Fall 1 mar), US$135 até o secundário (30 abr) e US$175
  depois; "WSU does not provide application fee waivers". Spring: prioritário
  1 ago, secundário 30 set (catálogo). Aplicações processadas por ordem de
  chegada até lotar.
- Graduate Assistantships: acrescentado que o GTA isenta até 100% da tuition
  de até 12 créditos/semestre (não as fees); não-residente com 16+ h/semana
  paga tarifa de residente; valores variam por departamento.

MANTIDO COMO Verificar: Taxa Empregabilidade, Salário Médio, Principais
Empregadores (o FAQ não traz outcomes), Rankings, Tamanho Turma, GPA
Médio, Taxa de Aceitação.

Fontes (out/2026):
- wichita.edu/services/finance/ComprehensiveFeeSchedule.php (Fall 2026)
- wichita.edu/services/tuitionfees/collegefees.php (Barton course fee FY2027)
- wichita.edu/academics/business/msba/FAQ_MSBA.php
- wichita.edu/academics/business/msba/tuition-scholarships-aid.php
- wichita.edu/globalselect
- catalog.wichita.edu (admission application; MS Business Analytics)
- wichita.edu/admissions/graduate/checklist-international.php
- wichita.edu/administration/financial_aid/coa_grad.php (2026-27)
- wichita.edu/services/studenthealth/HealthInsuranceInternational.php
- wichita.edu/academics/graduate_school/assistantships.php
- rentcafe.com (custo de vida em Wichita)

Uso: python scripts/etl_enrich_wichita_state_round2.py (a partir da raiz).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = ("Wichita State", "KS", "Wichita")

UPDATES = {
    "Formato": "Presencial (mix de aulas de 16 semanas e híbridas de 8 semanas)",
    "Duração (meses)": 16,
    "Custo/Crédito (Intl) USD": 903.59,
    "Custo/Crédito (Residente) USD": 367.92,
    "Tuition Total (Intl) USD": 27108,
    "Tuition Total (Residente) USD": 11038,
    "Fees Adicionais USD": 4930,
    "Seguro Saúde (programa) USD": 3663,
    "Seguro Residencial/mês USD": 15,
    "Utilities/mês USD": 206,
    "Celular/mês USD": 50,
    "WiFi/mês USD": 60,
    "Transporte/mês USD": 153,
    "Outros Gastos/mês USD": 270,
    # Mesma fórmula de computeCustoVidaTotal (derived.ts): soma mensal 1.779 x 16 meses + seguro 3.663.
    "Custo Vida Total (programa) USD": 32127,
    "Tipos de Bolsa": (
        "Global Select (automática para GPA ≥ 3.30 na admissão, renovada até o fim do "
        "programa): tarifa cai de US$903,59 para US$551,89/crédito (~38%; US$16.557 de "
        "tuition no programa em vez de US$27.108); não combinável com GA. Bolsas Barton "
        "(até 1 fev) e fellowships de US$750-8.500 (até 1 jul), abertas a internacionais"
    ),
    "Graduate Assistantships": (
        "GTA/GRA de até 20h/semana: GTA isenta até 100% da tuition de até 12 créditos por "
        "semestre (não as fees); não-residente com 16+ h/semana paga tarifa de residente; "
        "stipend varia por departamento; muito disputadas; não combinável com a Global "
        "Select Scholarship"
    ),
    "Deadline Aplicação": (
        "Fall: prioritário 1 mar, secundário 30 abr; Spring: prioritário 1 ago, secundário "
        "30 set; sem admissão no Summer; aplicações processadas por ordem de chegada até "
        "lotar; bolsas Barton: 1 fev"
    ),
    "Rounds Aplicação": (
        "Prazo prioritário + secundário por semestre; vagas limitadas, processadas por ordem "
        "de chegada até lotar. Taxa internacional: US$75 até o prioritário, US$135 até o "
        "secundário, US$175 depois (doméstico US$60); WSU não concede isenção da taxa"
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
    print(f"Linha alvo confirmada: {row_idx} (Wichita State, Wichita, KS)")

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
