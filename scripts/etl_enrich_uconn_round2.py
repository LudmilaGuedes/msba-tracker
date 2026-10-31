# -*- coding: utf-8 -*-
"""
Data enrichment (rodada 2, 20/out/2026): atualiza APENAS a linha da 'UConn'
(School of Business, Hartford/Stamford, CT) no msba_programs_database.xlsx.
Foco: MS in Business Analytics and Project Management (MSBAPM), presencial,
STEM.

CONFIRMADO SEM ALTERACAO (analytics.business.uconn.edu, out/2026):
- Course fee 2026-27: US$1.250/crédito x 37 = US$46.250, igual para
  residente, não-residente, online, presencial e internacional.
- Full-time em 3 semestres (Fall/Spring/Fall, ~16 meses); F-1 precisa de 9+
  créditos em cada Fall e Spring; visto só para entrada no Fall ou Spring.
- Prazos: Fall 2027 28/jul/2027 (aulas 30/ago); Spring 2027 16/dez/2026
  (só Hartford/online); Summer 1/mai/2027.
- Taxa US$75, isenta ao assistir uma info session ao vivo (online) e
  responder a pesquisa depois. Sem cartas; GMAT/GRE opcional; GPA 3.0 nos
  últimos 2 anos; 1 semestre de cálculo ou estatística (C ou melhor).
- Inglês: TOEFL 79 (ou 4.5 na nova escala) / IELTS 6.5 / PTE 60 / Duolingo 110.
- O programa não oferece GA/TA.

CORRIGIDO/PREENCHIDO:
- Fees Adicionais: 0 -> 1.350. A tabela de custo de attendance do CISS
  2026-27 inclui a visa compliance fee de US$450 por semestre para alunos
  com visto da UConn (3 semestres).
- Tipos de Bolsa / Quão Realista: a página de custos agora diz que as bolsas
  cobrem ~10-20% do custo do programa (~US$4.600-9.250); admitidos
  qualificados são convidados a aplicar pelo portal; prioridade por ordem de
  chegada. Realismo: Baixa -> Baixa-Moderada (bolsa parcial é plausível).
- Graduate Assistantships: o programa não oferece; a página sugere procurar
  GAs em outros departamentos da UConn.
- Ranking US News: Verificar -> 48 (Business Analytics, ranking de
  especialidade, tie; perfil da UConn na US News).
- Custo de vida (base Hartford, mais barata que Stamford e única com início
  no Spring): Aluguel 900 (quartos com banheiro privativo anunciados de
  US$620 a US$1.600; 1 quarto inteiro em média US$1.477), Utilities 215
  (média para apartamento básico em Hartford), Transporte 0 (U-Pass CT
  gratuito para alunos full-time da UConn Hartford), Seguro Residencial 15,
  Celular 50, WiFi 60, Outros 150 (valores padrão das outras linhas).
- Seguro Saúde: Verificar -> 4.531. Plano de 2025-26 (o de 2026-27 não foi
  publicado): anual US$3.214 (cobre Fall + Spring/Summer) + Fall do 2º ano
  (3.214 - 1.897 = 1.317). Estimativa.
- Custo Vida Total: Verificar -> 31.571, fórmula do dashboard: soma mensal
  1.690 x 16 meses + seguro 4.531.
- Rounds Aplicação: acrescentado o depósito de US$500 não reembolsável
  (até 2 semanas após aceitar a admissão).
- Principais Empregadores: acrescentado o maior salário reportado,
  US$125.000 (2024-25).

MANTIDO COMO Verificar: Taxa Empregabilidade e Salário Médio (a página de
carreiras só informa o maior salário; um "US$102.087 da Class of 2024"
citado por busca não aparece em nenhuma página oficial atual), Ranking QS,
Tamanho Turma, Taxa de Aceitação.

Fontes (out/2026):
- analytics.business.uconn.edu (cost-financial-aid, careers, apply, faq)
- international.global.uconn.edu/coa/ (CISS, 2026-27 Estimated Cost of Attendance)
- usnews.com/best-graduate-schools/top-business-schools/university-of-connecticut-01040
- studenthealth.media.uconn.edu (SHIP 2025-26)
- campushousing.uconn.edu/hartford/, socialwork.uconn.edu (U-Pass CT)
- roomies.com, usrentprices.com (aluguel em Hartford)

Uso: python scripts/etl_enrich_uconn_round2.py (a partir da raiz).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = ("UConn", "CT", "Hartford/Stamford")

UPDATES = {
    "Fees Adicionais USD": 1350,
    "Aluguel Privativo/mês USD": 900,
    "Seguro Saúde (programa) USD": 4531,
    "Seguro Residencial/mês USD": 15,
    "Utilities/mês USD": 215,
    "Celular/mês USD": 50,
    "WiFi/mês USD": 60,
    "Transporte/mês USD": 0,
    "Outros Gastos/mês USD": 150,
    # Mesma fórmula de computeCustoVidaTotal (derived.ts): soma mensal 1.690 x 16 meses + seguro 4.531.
    "Custo Vida Total (programa) USD": 31571,
    "Tipos de Bolsa": (
        "Merit scholarships parciais, ~10-20% do custo do programa (~US$4.600-9.250); "
        "admitidos qualificados são convidados a aplicar pelo portal; prioridade por ordem "
        "de chegada (aplicar cedo); o programa não oferece GA/TA"
    ),
    "Quão Realista (Intl)": (
        "Baixa-Moderada -- bolsa parcial (10-20%) é plausível para quem aplica cedo, mas "
        "é limitada e por convite; não há GA/TA no programa"
    ),
    "Graduate Assistantships": (
        "Não oferecidas pelo programa (e o MSBAPM não é elegível a GA waiver); a página "
        "sugere procurar GAs em outros departamentos da UConn"
    ),
    "Ranking US News": 48,
    "Rounds Aplicação": (
        "Sem rounds: um prazo por semestre; decisão em 4-6 semanas; bolsas por ordem de "
        "chegada (aplicar cedo); visto F-1 só para Fall/Spring e leva ~3 meses após a "
        "admissão; taxa de US$75 dispensada ao assistir a uma info session ao vivo (online) "
        "e responder a pesquisa; depósito de US$500 não reembolsável até 2 semanas após "
        "aceitar a admissão"
    ),
    "Principais Empregadores": (
        "Accenture, Amazon, Amtrak, Boehringer Ingelheim, Capgemini, Capital One, Cigna, "
        "Collins Aerospace, CVS Health, Hartford Healthcare, Nassau Financial Group, Optum, "
        "State Street, TD Bank, The Hartford, Walmart (página Careers do MSBAPM); maior "
        "salário reportado: US$125.000 (2024-25)"
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
    print(f"Linha alvo confirmada: {row_idx} (UConn, Hartford/Stamford, CT)")

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
