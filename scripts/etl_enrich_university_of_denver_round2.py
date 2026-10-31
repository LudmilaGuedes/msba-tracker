# -*- coding: utf-8 -*-
"""
Data enrichment (rodada 2, 20/out/2026): atualiza APENAS a linha da
'University of Denver' (Daniels College of Business, Denver, CO) no
msba_programs_database.xlsx. Foco: MS in Business Analytics, presencial, STEM.

CONFIRMADO SEM ALTERACAO:
- Tuition 2026-27: US$72.500 base para 58 créditos (quarters), mesma tarifa
  para todos (privada); ~US$1.250/crédito (calculado, não publicado).
- 15 disciplinas; full-time em 4 quarters (11,5 meses) ou part-time;
  aulas à noite (16h-22h); início Fall (meados de set) ou Spring (mar).
- 94% empregados full-time em até 6 meses (formados 2022-2024, página do
  programa). Bolsas de mérito e assistantships com consideração automática.
- Taxa US$100; GMAT/GRE opcional; GPA mínimo 2.5; sem cartas; 2 ensaios
  obrigatórios (350 e 250 palavras) + 1 opcional; resume; transcripts.
- Inglês (DU Bulletin): TOEFL 80 (antes de 21/jan/2026) ou 4.0 (nova
  escala) / IELTS 6.5 / Duolingo 115 / C1 Advanced 176.
- Seguro saúde SHIP 2026-27: US$4.400/ano. Transporte US$0 (RTD CollegePass).

CORRIGIDO/PREENCHIDO:
- Fees Adicionais: Verificar -> 1.552. Tabela oficial 2026-27 da DU:
  Health & Counseling Fee US$258/quarter x 4 + Graduate Activity Fee
  US$14/quarter x 4 + Technology Fee US$8/crédito x 58.
- Aluguel Privativo: 700 -> 950. O schema define quarto individual com
  banheiro privativo; anúncios perto da DU em 2026: US$850 (com utilities)
  e US$1.038/mês. Os US$700 eram de quarto sem banheiro privativo.
- Seguro Residencial 15, Celular 50, WiFi 60: valores padrão usados nas
  outras linhas. Utilities 186: estimativa da utility-rates.com para Denver
  (jul/2026; luz, água, esgoto e lixo). Outros Gastos 198: orçamento oficial
  da DU (personal/misc US$1.782 por 9 meses).
- Custo Vida Total: Verificar -> 25.508, fórmula do dashboard
  (computeCustoVidaTotal): soma mensal 1.759 x 12 meses + seguro 4.400.
- Ranking US News: Verificar -> 80 (Business Analytics, ranking de
  especialidade 2026-27, notícia oficial da DU de 7/abr/2026).
- GPA Médio Admitidos: Verificar -> 3.4 (FAQ de admissões da Daniels:
  "minimum overall GPA required is 2.5, with the average being 3.4").
- Principais Empregadores: lista dos mestrados da Daniels (Class of 2025,
  First Destination Outcomes da DU, p. 37). O relatório não separa o MSBA.
- Rounds Aplicação: isenção da taxa só para ex-alunos/alunos e funcionários
  da DU, militares dos EUA, Peace Corps/AmeriCorps/Teach for America,
  McNair e Fulbright Scholars; nada por participar de eventos.
- Entrevista: presencial ou online; quem é entrevistado recebe decisão
  acelerada em até 7 dias úteis.
- Graduate Assistantships: Duolingo não vale para a elegibilidade de GTA
  (DU Bulletin).

MANTIDO COMO Verificar:
- Salário Médio Pós: a DU só publica salário por faculdade, não por
  programa. Mestrados da Daniels, Class of 2025: média US$114.021 /
  mediana US$92.000 (121 salários, inclui MBA) -- fica no texto de
  Principais Empregadores, não na coluna numérica, pelo mesmo critério da
  rodada 1 (não misturar outcomes de programas diferentes).
- Ranking QS (o Daniels não aparece na tabela QS Business Analytics 2027
  consultada), Tamanho Turma, GMAT Médio, Taxa de Aceitação.

Fontes (out/2026):
- daniels.du.edu/business-information-analytics/masters/
- daniels.du.edu/graduate-admissions/ (+ /tuition-and-aid/, /faq/)
- du.edu/admission-aid/financial-aid-scholarships/graduate-financial-aid/cost-attend-du/tuition-fees
- du.edu/admission-aid/financial-aid-scholarships/graduate-cost-attendance
- bulletin.du.edu (English Language Proficiency)
- du.edu/news/du-grad-programs-continue-legacy-excellence (US News 2026-27)
- career.du.edu: Class of 2025 First Destination Outcomes (PDF, p. 34-37)
- offcampus.du.edu e housingboard.denverseminary.edu (anúncios); utility-rates.com/colorado/denver

Uso: python scripts/etl_enrich_university_of_denver_round2.py (a partir da raiz).
"""

import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.path.insert(0, "scripts")
from db_retention import rotate_backup, DB_PATH

sys.stdout.reconfigure(encoding="utf-8")

SHEET_DATA = "MSBA Programs"

MATCH = ("University of Denver", "CO", "Denver")

UPDATES = {
    "Fees Adicionais USD": 1552,
    "Aluguel Privativo/mês USD": 950,
    "Seguro Residencial/mês USD": 15,
    "Utilities/mês USD": 186,
    "Celular/mês USD": 50,
    "WiFi/mês USD": 60,
    "Outros Gastos/mês USD": 198,
    # Mesma fórmula de computeCustoVidaTotal (derived.ts): soma mensal 1.759 x 12 meses + seguro 4.400.
    "Custo Vida Total (programa) USD": 25508,
    "Ranking US News": 80,
    "GPA Médio Admitidos": 3.4,
    "Principais Empregadores": (
        "Mestrados da Daniels, Class of 2025 (o relatório da DU não separa o MSBA): "
        "Deloitte, PwC, University of Denver, EY, Lockheed Martin, KPMG, Plante Moran, "
        "Grant Thornton, Booz Allen Hamilton, Xcel Energy, JPMorganChase, Guidehouse, "
        "Fidelity Investments, Kaiser Permanente, DaVita, Crocs, CoBank, Arrow Electronics, "
        "Amazon; salário dos mestrados da Daniels: média US$114.021 / mediana US$92.000 "
        "(121 salários, inclui MBA)"
    ),
    "Rounds Aplicação": (
        "4 priority deadlines para o Fall (15 out; 1 fev; 15 mar; 1 mai) + deadline "
        "internacional 5 mai; após os priority deadlines, aplicações avaliadas se houver "
        "vaga. Taxa US$100 + US$50 de avaliação de diploma estrangeiro; isenção da taxa só "
        "para ex-alunos/alunos e funcionários da DU, militares dos EUA, Peace Corps/"
        "AmeriCorps/Teach for America, McNair e Fulbright Scholars (nada por participar "
        "de eventos)"
    ),
    "Entrevista": (
        "Por convite apenas, depois da aplicação completa (presencial ou online); quem é "
        "entrevistado recebe decisão acelerada em até 7 dias úteis"
    ),
    "Graduate Assistantships": (
        "Daniels considera automaticamente os admitidos para graduate assistantships "
        "(marcar interesse na aplicação); valores não publicados para o MSBA; GTA exige "
        "testes específicos de inglês -- Duolingo não vale para a elegibilidade de GTA"
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
    print(f"Linha alvo confirmada: {row_idx} (University of Denver, Denver, CO)")

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
