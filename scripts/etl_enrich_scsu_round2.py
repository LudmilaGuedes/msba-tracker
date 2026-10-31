# -*- coding: utf-8 -*-
"""
Data enrichment (rodada 2, 20/out/2026): atualiza APENAS a linha da
'Southern Connecticut State University (SCSU)' (School of Business, New
Haven, CT) no msba_programs_database.xlsx. Foco: STEM MS in Business
Analytics, presencial (on-ground).

A rodada 1 (etl_enrich_scsu.py, 03/out/2026) já tinha auditado tuition,
fees, bolsas/GA, prazos, inglês e documentos nas páginas oficiais; esta
rodada preenche o que ficou em Verificar.

CONFIRMADO SEM ALTERACAO: tuition fixa por semestre (tabela 2026-27):
intl/out-of-state US$6.537 (NEBHE), in-state US$4.358, fees US$3.122;
"typically two years" (4 semestres) -> US$26.148 / US$17.432 + US$12.488
de fees; 30 créditos; GPA 3.0; sem GMAT/GRE, cartas ou redação; Duolingo
105; taxa US$50; GA US$5.280/semestre + isenção de fees (não de tuition).

PREENCHIDO:
- Seguro Saúde: Verificar -> 1.920. A SCSU exige seguro para internacionais
  e recomenda o ISO Secure (parceira oficial): US$80/mês (preço da página
  ISO para a SCSU, idade 20; varia com a idade) x 24 meses.
- Transporte: Verificar -> 10. Alunos de pós com aula no campus compram o
  U-Pass CT (ônibus CTtransit e trens dentro de CT) por US$40/semestre.
- Aluguel Privativo: Verificar -> 1.100. Estimativa: quartos perto da SCSU
  anunciados de US$1.000 (banheiro compartilhado) a US$1.700; 1 quarto
  inteiro em New Haven custa US$1.600-2.000 (Trulia/Rentometer 2026).
- Utilities 260 (média do guia de custo de vida de New Haven, com gás);
  Seguro Residencial 15, Celular 50, WiFi 60, Outros 150 (valores padrão
  das outras linhas).
- Custo Vida Total: Verificar -> 48.600, fórmula do dashboard: soma mensal
  1.945 x 24 meses + seguro 1.920.
- Rounds Aplicação: a única isenção da taxa no catálogo é por necessidade
  financeira com critérios domésticos (SAT fee waiver, CONNTAC/CONNCAP);
  nada para internacionais nem por ir ao Open House.

MANTIDO COMO Verificar: Taxa Empregabilidade e Salário Médio (a SCSU não
publica outcomes do MSBA; o College Scorecard só tem o mestrado de
Business Administration, mediana de US$60.200 no 1º ano -- outro
programa), Rankings, Tamanho Turma, GPA Médio, Taxa de Aceitação.

Fontes (out/2026):
- inside.southernct.edu/international/study-at-scsu/matriculated-students/insurance
- isoa.org/schools/southern-connecticut-state-university
- hootloot.southernct.edu/benefits/upass.php
- catalog.southernct.edu/graduate/tuition-and-fees/waivers.html
- grad.southernct.edu/open-house/
- trulia.com, rentometer.com (aluguel em New Haven); guia de utilities de New Haven

Uso: python scripts/etl_enrich_scsu_round2.py (a partir da raiz).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = ("Southern Connecticut State University (SCSU)", "CT", "New Haven")

UPDATES = {
    "Aluguel Privativo/mês USD": 1100,
    "Seguro Saúde (programa) USD": 1920,
    "Seguro Residencial/mês USD": 15,
    "Utilities/mês USD": 260,
    "Celular/mês USD": 50,
    "WiFi/mês USD": 60,
    "Transporte/mês USD": 10,
    "Outros Gastos/mês USD": 150,
    # Mesma fórmula de computeCustoVidaTotal (derived.ts): soma mensal 1.945 x 24 meses + seguro 1.920.
    "Custo Vida Total (programa) USD": 48600,
    "Rounds Aplicação": (
        "Sem rounds -- prazos fixos por semestre; internacionais novos têm prazo "
        "antecipado (1 abr Fall / 1 out Spring, catálogo 2026-27). Completar o arquivo até "
        "o prazo dá prioridade; depois do prazo não se aceitam aplicações. Taxa de US$50 "
        "sem isenção para internacionais (a única isenção do catálogo é por necessidade "
        "financeira com critérios domésticos; o Open House não isenta)"
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
    print(f"Linha alvo confirmada: {row_idx} (SCSU, New Haven, CT)")

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
