# -*- coding: utf-8 -*-
"""
Data enrichment, RODADA 4 (31/out/2026): atualiza APENAS a linha
'University of Miami' (University of Miami Business School, Coral Gables, FL)
no msba_programs_database.xlsx. NÃO confundir com 'Miami U Ohio' (Oxford, OH);
o match exige nome + estado + cidade.

CONFIRMADO SEM ALTERAÇÃO (mbs.miami.edu, out/2026):
- US$2.561/crédito x 33 = US$84.523 + fees US$1.438 = US$85.961 (ano 2027-28).
  A página tem pequenas inconsistências de soma (33 x 2.561 = 84.513;
  2 x 714 = 1.428), mas o total de US$85.961 bate com os valores listados, que
  foram mantidos.
- 10 meses, início em agosto, STEM; taxa de US$100; TOEFL 94 (só iBT),
  IELTS 7.0, DET 125; 2 essays; entrevista por convite (virtual, ~30 min).
- Bolsas de mérito automáticas: 74% do Fall 2025 receberam.
- 94% empregados em até 6 meses (2024).
- Seguro saúde 2026-27: US$4.483 por ano (folheto do Student Health Service).
- Prazos de 2026-27 (Round 1 em 1 out, prioridade de bolsa).

CORRIGIDO/PREENCHIDO:
- Formato: on campus em Coral Gables, full-time de 10 meses (também existe
  part-time).
- GMAT/GRE: "Not required"; notas não oficiais ajudam na bolsa máxima; se
  enviar, a nota oficial é exigida antes da orientação (códigos GMAT 7NV-S1-96,
  GRE 5815).
- Documentos: para diploma estrangeiro, avaliação WES + tradução certificada do
  histórico e do diploma + cópia do diploma + histórico oficial em envelope
  lacrado da universidade; o histórico não oficial pode ser enviado antes para
  acelerar a análise.
- Rounds Aplicação: o FAQ recomenda que internacionais apliquem até o Round 3,
  por causa do visto.
- Principais Empregadores: lista da página do programa (alumni employers),
  mantendo os da rodada anterior.
- TA/RA Disponível / Processo TA/RA / Graduate Assistantships: Verificar ->
  não divulgado; a Business School só oferece bolsa de mérito; internacionais
  não têm ajuda federal e só conseguem empréstimo privado com fiador americano.

MANTIDO COMO Verificar: Salário Médio (a escola não publica relatório de salário
do MSBA), Tamanho Turma, GPA Mínimo/Médio (só existe a média geral da pós da
Business School, 3.40 em 2026-27, que não é do MSBA), GMAT Médio, Taxa de
Aceitação, Ranking US News/QS (a página só cita ChaseDream 2020).
Seguro Residencial/Utilities/Celular/WiFi continuam Verificar: os valores
mensais de custo de vida vêm do orçamento oficial da UM (moradia, transporte e
pessoal/livros em "Outros"), que não separa esses itens; preencher
duplicaria custo.

Fontes (out/2026):
- mbs.miami.edu/graduate/find-and-compare-programs/business-analytics/ (index, requirements, cost-of-attendance)
- mbs-new.miami.edu/graduate/graduate-how-to-apply.html; gradbusiness.miami.edu/faq
- mbs-new.miami.edu/graduate/class-profile.html (perfil geral da pós, 2026-27)
- studenthealth.studentaffairs.miami.edu: 2026-2027 summary flyer (medical)

Uso: python scripts/etl_enrich_u_miami_round4.py (a partir da raiz).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = ("University of Miami", "FL", "Coral Gables")

UPDATES = {
    "Formato": "Presencial (on campus em Coral Gables); full-time de 10 meses (também existe part-time de 2 anos)",
    "TA/RA Disponível": "Não divulgado (a Business School não oferece TA/RA para o MSBA; o financiamento é só bolsa de mérito)",
    "Processo TA/RA": "Não se aplica; internacionais não têm ajuda federal e só conseguem empréstimo privado com fiador americano",
    "Graduate Assistantships": "Nenhuma divulgada para o MSBA",
    "Principais Empregadores": (
        "Abbott, American Express, Carnival, Chewy, Deloitte, EY, FedEx, Hertz, JP Morgan, KLX "
        "Aerospace, NBCUniversal, Nielsen, Norwegian Cruise Line, Perry Ellis, Royal Caribbean, RBI "
        "(Burger King), TransUnion, Vanguard, Visa, Walmart, Whole Foods (alumni employers na página "
        "do programa); também citados: McKinsey, BCG, Amazon, Cisco, Capital One, Citi, Mastercard, "
        "ChenMed, Chicago Cubs"
    ),
    "GMAT/GRE Obrigatório": (
        "Não exigido; notas não oficiais ajudam a obter a bolsa máxima; se enviar, a nota oficial "
        "(eletrônica, da agência) é exigida antes da orientação (GMAT 7NV-S1-96, GRE 5815)"
    ),
    "Rounds Aplicação": (
        "Round 1: 1 out 2026 (decisão até 15 nov; prioridade de bolsa); Round 2: 1 dez 2026 (decisão "
        "até 15 jan); Round 3: 1 fev 2027 (decisão até 15 mar); Round 4: 2 abr 2027 (decisão até "
        "15 mai); Final: 1 jun 2027. O FAQ recomenda que internacionais apliquem até o Round 3 por "
        "causa do visto. Taxa US$100 (isenção limitada para quem participa de evento informativo)"
    ),
    "Documentos Necessários": (
        "Aplicação online + taxa US$100; 2 essays (1 comum + 1 específico do programa); resume; "
        "2 cartas de recomendação obrigatórias (até 3), só pelo formulário da aplicação, com e-mail profissional do recomendador; transcripts "
        "de todas as instituições (o não oficial pode ser enviado antes para acelerar); diploma "
        "estrangeiro: avaliação WES + tradução certificada do histórico e do diploma + cópia do "
        "diploma + histórico oficial em envelope lacrado da universidade; TOEFL/IELTS/DET; GMAT/GRE "
        "opcional (ajuda na bolsa); entrevista por convite"
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
    print(f"Linha alvo confirmada: {row_idx} (University of Miami, Coral Gables, FL)")

    no_fill = PatternFill(fill_type=None)
    changes = []
    for header, new_value in UPDATES.items():
        cell = ws.cell(row=row_idx, column=col[header])
        if cell.value != new_value:
            changes.append((header, cell.value, new_value))
        cell.value = new_value
        cell.fill = no_fill

    wb.save(DB_PATH)

    print(f"\nCampos atualizados ({len(changes)}) (Antes -> Depois):")
    for header, old, new in changes:
        print(f"  - {header}:")
        print(f"      Antes:  {old!r}")
        print(f"      Depois: {new!r}")


if __name__ == "__main__":
    main()
