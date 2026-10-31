# -*- coding: utf-8 -*-
"""
Data enrichment (20/out/2026): atualiza APENAS a linha da 'Loyola University
Chicago' (Quinlan School of Business, Chicago, IL) no
msba_programs_database.xlsx. Foco: MS in Business Analytics (MSBA), STEM.
(Não confundir com Loyola Marymount University, LMU, Los Angeles.)

CONFIRMADO SEM ALTERACAO (gpem.luc.edu, programa businessanalyticsms):
- 12 cursos / 36 créditos; US$1.840/crédito (2026-27, mesma tarifa para
  todos) = US$66.240; início Fall ou Spring; 1 a 2 anos conforme a carga;
  semestres de 15 semanas a partir do Fall 2027.
- Aplicação gratuita e rolling; sem cartas; transcripts oficiais + avaliação
  de credenciais estrangeiras (obrigatória); inglês da Quinlan: TOEFL 4.5
  (90) / IELTS 6.5 / PTE 61 / Duolingo 115.
- Bolsas: Dean's Merit automática na admissão (normalmente 1-2 cursos,
  única); Graduate Business Student Scholarship exige FAFSA (internacionais
  não elegíveis); Returning Rambler só para ex-alunos.
- Empregadores citados na página: Accenture, Aon, Boeing, Bosch, CME,
  Capital One, Dyson, Epic Systems, GE, Morningstar, PwC, West Monroe,
  Wells Fargo.

CORRIGIDO/PREENCHIDO:
- Formato: a página lista "Delivery format: Hybrid" (online + presencial no
  Water Tower Campus). Pela regra da ISSS da Loyola, só 1 curso online
  (3 créditos) por termo conta para o full-time de F-1 (9 créditos na pós);
  o restante precisa ser presencial.
- Ranking US News: 20 -> 73. O perfil atual da Quinlan na US News mostra
  #73 (tie) em Business Analytics (ranking de especialidade). O "#20 (2025)"
  citado pela Loyola não corresponde a esse ranking.
- Fees Adicionais: Verificar -> 1.252. Tabela do Bursar 2026-27 (Quinlan),
  3 termos (Fall, Spring, Summer) no ritmo de 12 meses: CTA U-Pass US$118
  + Student Development Fee US$131 + Technology Fee US$135 por termo, mais
  Matriculation Fee de US$100.
- Seguro Saúde: Verificar -> 3.869 (Student Health Insurance Premium anual,
  Bursar 2026-27).
- Custo de vida (Chicago): Aluguel 1.200 (quarto privativo em moradia
  compartilhada custa em média US$1.050 e US$1.300+ perto do centro; com
  banheiro privativo, estimativa de 1.200), Mercado 300, Utilities 260
  (meio da faixa de US$180-360 para apartamento), Transporte 0 (CTA U-Pass
  já cobrado nas fees), Seguro Residencial 15, Celular 50, WiFi 60, Outros
  150 (valores padrão das outras linhas).
- Custo Vida Total: Verificar -> 28.289, fórmula do dashboard: soma mensal
  2.035 x 12 meses + seguro 3.869.
- Rounds Aplicação: acrescentados os eventos (Open House 12/nov/2026 no
  Water Tower Campus; info session virtual do MSBA 1/dez/2026).

MANTIDO COMO Verificar: Taxa Empregabilidade, Salário Médio (a Quinlan não
publica outcomes do MSBA), Ranking QS, Tamanho Turma, GPA Mínimo/Médio,
GMAT Médio, Taxa de Aceitação.

Fontes (out/2026):
- gpem.luc.edu/portal/program?name=businessanalyticsms
- gpem.luc.edu/portal/admission?tab=international-requirements
- luc.edu/bursar/tuitionfees/2026-2027/business/
- luc.edu/isss/currentinternationalstudents.shtml (limite de cursos online para F-1)
- luc.edu/quinlan/academics/graduatedegrees/ms/choosingabusinessdataanalyticsprogram/
- usnews.com/best-graduate-schools/top-business-schools/loyola-university-chicago-01066
- citiesinsider.com (co-living em Chicago 2026); guias de utilities de Chicago

Uso: python scripts/etl_enrich_loyola_chicago.py (a partir da raiz).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = ("Loyola University Chicago", "IL", "Chicago")

UPDATES = {
    "Formato": (
        "Híbrido (presencial no Water Tower Campus + online); para F-1, só 1 curso online "
        "(3 créditos) por termo conta no full-time de 9 créditos -- o restante precisa ser "
        "presencial; também existe 100% online (não serve para visto)"
    ),
    "Ranking US News": 73,
    "Fees Adicionais USD": 1252,
    "Aluguel Privativo/mês USD": 1200,
    "Mercado/mês USD": 300,
    "Seguro Saúde (programa) USD": 3869,
    "Seguro Residencial/mês USD": 15,
    "Utilities/mês USD": 260,
    "Celular/mês USD": 50,
    "WiFi/mês USD": 60,
    "Transporte/mês USD": 0,
    "Outros Gastos/mês USD": 150,
    # Mesma fórmula de computeCustoVidaTotal (derived.ts): soma mensal 2.035 x 12 meses + seguro 3.869.
    "Custo Vida Total (programa) USD": 28289,
    "Rounds Aplicação": (
        "Admissão rolling para Fall e Spring; aplicação gratuita; eventos: Open House "
        "12/nov/2026 (Water Tower Campus) e info session virtual do MSBA 1/dez/2026"
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
    print(f"Linha alvo confirmada: {row_idx} (Loyola University Chicago, Chicago, IL)")

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
